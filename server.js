'use strict';

require('dotenv').config();

const express = require('express');
const { Pool } = require('pg');
const crypto = require('crypto');
const path = require('path');
const rateLimit = require('express-rate-limit');
const ExcelJS = require('exceljs');

// ==========================================
// CONFIGURACIÓN
// ==========================================

const { DATABASE_URL, ADMIN_PASSWORD, SESSION_SECRET } = process.env;

if (!DATABASE_URL || !ADMIN_PASSWORD || !SESSION_SECRET) {
  console.error(
    'Faltan variables de entorno: DATABASE_URL, ADMIN_PASSWORD y SESSION_SECRET son obligatorias.'
  );
  process.exit(1);
}

const PORT = Number(process.env.PORT) || 8080;
const PRECIO_INSCRIPCION = 10;
const IS_PROD = process.env.NODE_ENV === 'production';

const pool = new Pool({
  connectionString: DATABASE_URL,
  max: 5,
  idleTimeoutMillis: 30000,
  connectionTimeoutMillis: 10000
});

pool.on('error', (err) => {
  console.error('Error inesperado en PostgreSQL:', err.message);
});

const app = express();

app.set('trust proxy', 1);
app.disable('x-powered-by');

app.use(express.json({ limit: '200kb' }));

app.use((req, res, next) => {
  res.set('X-Content-Type-Options', 'nosniff');
  res.set('Referrer-Policy', 'same-origin');
  res.set('X-Frame-Options', 'SAMEORIGIN');
  next();
});

// ==========================================
// LIMITACIÓN DE SOLICITUDES
// ==========================================

const writeLimiter = rateLimit({
  windowMs: 10 * 60 * 1000,
  limit: 60,
  standardHeaders: true,
  legacyHeaders: false,
  message: {
    error: 'Demasiados envíos. Espera unos minutos e inténtalo otra vez.'
  }
});

const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  limit: 10,
  standardHeaders: true,
  legacyHeaders: false,
  message: {
    error: 'Demasiados intentos. Espera 15 minutos.'
  }
});

// ==========================================
// UTILIDADES
// ==========================================

class HttpError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

const bad = (message) => new HttpError(400, message);

const clean = (value, max = 160) =>
  String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, max);

function wrap(fn) {
  return (req, res, next) =>
    Promise.resolve(fn(req, res, next)).catch(next);
}

// ==========================================
// DATOS DEL FORMULARIO
// ==========================================

const GRADOS = [
  { n: 'Inicial · 3 años', v: 'Inicial' },
  { n: 'Inicial · 4 años', v: 'Inicial' },
  { n: 'Inicial · 5 años', v: 'Inicial' },

  ...[1, 2, 3, 4, 5, 6].map((i) => ({
    n: `${i}° de Primaria`,
    v: 'Primaria'
  })),

  ...[1, 2, 3, 4, 5].map((i) => ({
    n: `${i}° de Secundaria`,
    v: 'Secundaria'
  }))
];

const GRADO_NIVEL = Object.fromEntries(
  GRADOS.map((g) => [g.n, g.v])
);

const CATEGORIAS = ['CATEGORÍA “A”', 'CATEGORÍA “B”'];

// Acepta también los textos largos antiguos y los guarda siempre en forma simple.
function normCategoria(valor) {
  if (/^CATEGOR[ÍI]A\s*[“"]A[”"]/i.test(valor)) return CATEGORIAS[0];
  if (/^CATEGOR[ÍI]A\s*[“"]B[”"]/i.test(valor)) return CATEGORIAS[1];
  return '';
}

// ==========================================
// VALIDACIÓN DE PAGOS
// ==========================================

function parseOperaciones(value) {
  const operations = clean(value, 120)
    .split(/[,;\s]+/)
    .filter(Boolean);

  if (
    !operations.length ||
    operations.length > 6 ||
    !operations.every((item) => /^[A-Za-z0-9-]{4,20}$/.test(item))
  ) {
    throw bad(
      'El N° de operación no es válido. Usa letras y números de 4 a 20 caracteres.'
    );
  }

  return operations.join(', ');
}

function parseMedio(value) {
  if (!['Yape', 'Transferencia', 'Efectivo'].includes(value)) {
    throw bad('Elige un medio de pago válido.');
  }

  return value;
}

// ==========================================
// BASE DE DATOS
// ==========================================

const SCHEMA = `
  CREATE TABLE IF NOT EXISTS libres (
    id SERIAL PRIMARY KEY,
    dni TEXT NOT NULL UNIQUE,
    grado TEXT NOT NULL,
    nivel TEXT NOT NULL,
    nombre TEXT NOT NULL,
    telefono TEXT,
    aula TEXT,
    categoria TEXT,
    monto INTEGER NOT NULL DEFAULT 10,
    operacion TEXT NOT NULL DEFAULT '',
    medio TEXT NOT NULL DEFAULT 'Yape',
    pagador TEXT,
    verificado BOOLEAN NOT NULL DEFAULT FALSE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
  );

  ALTER TABLE libres ADD COLUMN IF NOT EXISTS telefono TEXT;
  ALTER TABLE libres ADD COLUMN IF NOT EXISTS aula TEXT;
  ALTER TABLE libres ADD COLUMN IF NOT EXISTS categoria TEXT;
  ALTER TABLE libres ADD COLUMN IF NOT EXISTS pagador TEXT;

  ALTER TABLE libres
    ADD COLUMN IF NOT EXISTS verificado BOOLEAN NOT NULL DEFAULT FALSE;

  ALTER TABLE libres
    ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT now();

  ALTER TABLE libres
    ADD COLUMN IF NOT EXISTS monto INTEGER NOT NULL DEFAULT 10;

  ALTER TABLE libres
    ADD COLUMN IF NOT EXISTS operacion TEXT NOT NULL DEFAULT '';

  ALTER TABLE libres
    ADD COLUMN IF NOT EXISTS medio TEXT NOT NULL DEFAULT 'Yape';

  ALTER TABLE libres DROP CONSTRAINT IF EXISTS libres_medio_check;

  ALTER TABLE libres
    ADD CONSTRAINT libres_medio_check
    CHECK (medio IN ('Yape', 'Transferencia', 'Efectivo'));

  CREATE INDEX IF NOT EXISTS idx_libres_created_at
    ON libres(created_at);

  ALTER TABLE libres ADD COLUMN IF NOT EXISTS puntaje NUMERIC(6,2);
  ALTER TABLE libres ADD COLUMN IF NOT EXISTS hora_entrega TIME;

  UPDATE libres SET categoria = 'CATEGORÍA “A”'
    WHERE categoria LIKE 'CATEGORÍA “A”%' AND categoria <> 'CATEGORÍA “A”';
  UPDATE libres SET categoria = 'CATEGORÍA “B”'
    WHERE categoria LIKE 'CATEGORÍA “B”%' AND categoria <> 'CATEGORÍA “B”';

  CREATE TABLE IF NOT EXISTS config (
    clave TEXT PRIMARY KEY,
    valor TEXT NOT NULL
  );
`;

// ==========================================
// RUTA DE SALUD DEL SERVIDOR
// ==========================================

app.get('/healthz', wrap(async (_req, res) => {
  await pool.query('SELECT 1');

  res.json({
    ok: true
  });
}));

// ==========================================
// REGISTRO DE ALUMNOS
// ==========================================

app.post('/api/libres', writeLimiter, wrap(async (req, res) => {
  const body = req.body || {};

  const dni = clean(body.dni, 14);
  const grado = clean(body.grado, 40);
  const nivel = GRADO_NIVEL[grado];

  const nombre = clean(body.nombre, 120);
  const telefono = clean(body.tel, 20);
  const categoria = normCategoria(clean(body.categoria, 120));
  const pagador = clean(body.pagador, 120);

  // Validaciones

  if (!/^\d{8}$/.test(dni)) {
    throw bad('El DNI debe tener 8 dígitos.');
  }

  if (!nivel) {
    throw bad('Elige un grado válido.');
  }

  if (!nombre) {
    throw bad('Escribe los apellidos y nombres del alumno.');
  }

  if (!/^9\d{8}$/.test(telefono)) {
    throw bad('El celular debe tener 9 dígitos y empezar con 9.');
  }

  if (!CATEGORIAS.includes(categoria)) {
    throw bad('Elige una categoría válida.');
  }

  const medio = parseMedio(body.medio);

  if (!pagador) {
    throw bad('Escribe el nombre de quien hizo el pago.');
  }

  const operacion =
    medio === 'Efectivo'
      ? ''
      : parseOperaciones(body.operacion);

  try {
    const result = await pool.query(
      `INSERT INTO libres
        (
          dni,
          grado,
          nivel,
          nombre,
          telefono,
          categoria,
          monto,
          operacion,
          medio,
          pagador
        )
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)
       RETURNING id`,
      [
        dni,
        grado,
        nivel,
        nombre,
        telefono,
        categoria,
        PRECIO_INSCRIPCION,
        operacion,
        medio,
        pagador
      ]
    );

    res.status(201).json({
      ok: true,
      id: result.rows[0].id,
      monto: PRECIO_INSCRIPCION
    });

  } catch (err) {
    if (err.code === '23505') {
      throw new HttpError(
        409,
        `El DNI ${dni} ya tiene una inscripción registrada.`
      );
    }

    throw err;
  }
}));

// ==========================================
// AUTENTICACIÓN DEL ADMINISTRADOR
// ==========================================

const sign = (value) =>
  crypto
    .createHmac('sha256', SESSION_SECRET)
    .update(value)
    .digest('hex');

function safeEqual(a, b) {
  const left = crypto
    .createHash('sha256')
    .update(String(a ?? ''))
    .digest();

  const right = crypto
    .createHash('sha256')
    .update(String(b ?? ''))
    .digest();

  return crypto.timingSafeEqual(left, right);
}

function getCookies(req) {
  const cookies = {};

  for (const part of (req.headers.cookie || '').split(';')) {
    const index = part.indexOf('=');

    if (index > 0) {
      cookies[part.slice(0, index).trim()] =
        decodeURIComponent(part.slice(index + 1).trim());
    }
  }

  return cookies;
}

function isAdmin(req) {
  const token = getCookies(req).adm;

  if (!token) return false;

  const [expires, signature] = token.split('.');

  if (!expires || !signature || !Number.isFinite(Number(expires))) {
    return false;
  }

  if (Number(expires) < Date.now()) {
    return false;
  }

  return safeEqual(signature, sign(expires));
}

function needAdmin(req, res, next) {
  if (isAdmin(req)) {
    return next();
  }

  return res.status(401).json({
    error: 'No autorizado.'
  });
}

// LOGIN

app.post('/api/admin/login', loginLimiter, (req, res) => {
  if (!safeEqual(req.body && req.body.password, ADMIN_PASSWORD)) {
    return res.status(401).json({
      error: 'Contraseña incorrecta.'
    });
  }

  const expires = String(
    Date.now() + 12 * 60 * 60 * 1000
  );

  res.setHeader(
    'Set-Cookie',
    `adm=${expires}.${sign(expires)}; Path=/; HttpOnly; SameSite=Strict; Max-Age=43200${IS_PROD ? '; Secure' : ''}`
  );

  return res.json({
    ok: true
  });
});

// LOGOUT

app.post('/api/admin/logout', (_req, res) => {
  res.setHeader(
    'Set-Cookie',
    `adm=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0${IS_PROD ? '; Secure' : ''}`
  );

  res.json({
    ok: true
  });
});

// ==========================================
// OBTENER INSCRIPCIONES
// ==========================================

async function loadInscripciones() {
  const result = await pool.query(
    `SELECT
       id,
       dni,
       grado,
       nivel,
       nombre,
       categoria,
       telefono,
       monto,
       operacion,
       medio,
       pagador,
       verificado,
       puntaje,
       to_char(hora_entrega, 'HH24:MI:SS') AS hora,
       created_at
     FROM libres
     ORDER BY created_at DESC, id DESC`
  );

  return result.rows.map((row) => ({
    _id: row.id,
    dni: row.dni,
    grado: row.grado,
    nivel: row.nivel,
    nombre: row.nombre,
    categoria: row.categoria || '',
    tel: row.telefono || '',
    monto: Number(row.monto) || 0,
    operacion: row.operacion || '',
    medio: row.medio || 'Yape',
    pagador: row.pagador || '',
    verificado: Boolean(row.verificado),
    puntaje: row.puntaje === null ? null : Number(row.puntaje),
    hora: row.hora || '',
    fecha: row.created_at
  }));
}

// ==========================================
// DATOS DEL PANEL ADMINISTRATIVO
// ==========================================

app.get('/api/admin/data', needAdmin, wrap(async (_req, res) => {
  res.set('Cache-Control', 'no-store');

  const libres = await loadInscripciones();

  res.json({
    fichas: [],
    libres
  });
}));

// ==========================================
// VERIFICACIÓN DE PAGOS
// ==========================================

app.patch('/api/admin/verify', needAdmin, wrap(async (req, res) => {
  const { t, id, v } = req.body || {};

  if (
    t !== 'libres' ||
    !Number.isInteger(id) ||
    id <= 0 ||
    typeof v !== 'boolean'
  ) {
    throw bad('Datos de verificación no válidos.');
  }

  const result = await pool.query(
    `UPDATE libres
     SET verificado = $1
     WHERE id = $2
     RETURNING id`,
    [v, id]
  );

  if (!result.rowCount) {
    throw new HttpError(404, 'No se encontró la inscripción.');
  }

  res.json({
    ok: true
  });
}));

// ==========================================
// EXPORTACIÓN A EXCEL
// ==========================================

app.get('/api/admin/export.xlsx', needAdmin, wrap(async (_req, res) => {
  const inscripciones = await loadInscripciones();

  const workbook = new ExcelJS.Workbook();

  workbook.creator = 'Tarpuy - VIII Olimpiada Tarpuyana 2026';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Inscripciones', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  sheet.columns = [
    { header: 'N°', key: 'numero', width: 7 },
    { header: 'DNI', key: 'dni', width: 14 },
    { header: 'Grado', key: 'grado', width: 20 },
    { header: 'Nivel', key: 'nivel', width: 15 },
    { header: 'Categoría', key: 'categoria', width: 48 },
    { header: 'Apellidos y nombres', key: 'nombre', width: 38 },
    { header: 'Celular / WhatsApp', key: 'tel', width: 20 },
    { header: 'Monto (S/)', key: 'monto', width: 13 },
    { header: 'Medio de pago', key: 'medio', width: 18 },
    { header: 'N° de operación', key: 'operacion', width: 24 },
    { header: 'Nombre de quien pagó', key: 'pagador', width: 30 },
    { header: 'Estado de pago', key: 'estado', width: 18 },
    { header: 'Puntaje', key: 'puntaje', width: 12 },
    { header: 'Hora de entrega', key: 'hora', width: 16 },
    { header: 'Fecha de inscripción', key: 'fecha', width: 24 }
  ];

  inscripciones.forEach((row, index) => {
    sheet.addRow({
      numero: index + 1,
      ...row,
      estado: row.verificado ? 'Verificado' : 'Pendiente',
      fecha: row.fecha
        ? new Date(row.fecha).toLocaleString('es-PE')
        : ''
    });
  });

  const header = sheet.getRow(1);

  header.font = {
    bold: true,
    color: { argb: 'FFFFFFFF' }
  };

  header.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF14532D' }
  };

  header.alignment = {
    vertical: 'middle',
    wrapText: true
  };

  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: sheet.columns.length }
  };

  res.set({
    'Content-Type':
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

    'Content-Disposition':
      'attachment; filename="inscripciones-olimpiada-tarpuyana-2026.xlsx"',

    'Cache-Control': 'no-store'
  });

  await workbook.xlsx.write(res);
  res.end();
}));

// ==========================================
// EXCEL: PAGOS PENDIENTES DE VERIFICAR
// ==========================================

const GRADOS_ORDEN = [
  'Inicial · 3 años',
  'Inicial · 4 años',
  'Inicial · 5 años',
  ...[1, 2, 3, 4, 5, 6].map((n) => `${n}° de Primaria`),
  ...[1, 2, 3, 4, 5].map((n) => `${n}° de Secundaria`)
];

const ordenGrado = (g) => {
  const i = GRADOS_ORDEN.indexOf(g);
  return i < 0 ? 99 : i;
};

const splitOps = (v) => String(v || '').split(/[,;\s]+/).filter(Boolean);

app.get('/api/admin/export-pendientes.xlsx', needAdmin, wrap(async (_req, res) => {
  const todas = await loadInscripciones();

  // Cuántas veces aparece cada N° de operación (para avisar repetidos).
  const veces = {};

  todas.forEach((r) => {
    splitOps(r.operacion).forEach((o) => {
      veces[o] = (veces[o] || 0) + 1;
    });
  });

  const pendientes = todas
    .filter((r) => !r.verificado)
    .sort(
      (a, b) =>
        ordenGrado(a.grado) - ordenGrado(b.grado) ||
        a.grado.localeCompare(b.grado, 'es') ||
        a.categoria.localeCompare(b.categoria, 'es') ||
        a.nombre.localeCompare(b.nombre, 'es')
    );

  const workbook = new ExcelJS.Workbook();

  workbook.creator = 'Tarpuy - VIII Olimpiada Tarpuyana 2026';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Pagos pendientes', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  sheet.columns = [
    { header: 'N°', key: 'numero', width: 7 },
    { header: 'DNI', key: 'dni', width: 14 },
    { header: 'Grado', key: 'grado', width: 20 },
    { header: 'Categoría', key: 'categoria', width: 18 },
    { header: 'Apellidos y nombres', key: 'nombre', width: 38 },
    { header: 'Celular / WhatsApp', key: 'tel', width: 20 },
    { header: 'Monto (S/)', key: 'monto', width: 13 },
    { header: 'Medio de pago', key: 'medio', width: 16 },
    { header: 'N° de operación', key: 'operacion', width: 24 },
    { header: '¿Operación repetida?', key: 'repetida', width: 20 },
    { header: 'Nombre de quien pagó', key: 'pagador', width: 30 },
    { header: 'Fecha de inscripción', key: 'fecha', width: 24 },
    { header: 'Observaciones', key: 'obs', width: 30 }
  ];

  sheet.getColumn('dni').numFmt = '@';

  pendientes.forEach((row, index) => {
    const repetida = splitOps(row.operacion).some((o) => veces[o] > 1);

    const fila = sheet.addRow({
      numero: index + 1,
      dni: row.dni,
      grado: row.grado,
      categoria: row.categoria,
      nombre: row.nombre,
      tel: row.tel,
      monto: row.monto,
      medio: row.medio,
      operacion: row.operacion,
      repetida: repetida ? 'SÍ' : 'No',
      pagador: row.pagador,
      fecha: row.fecha ? new Date(row.fecha).toLocaleString('es-PE') : '',
      obs: ''
    });

    if (repetida) {
      fila.getCell('repetida').font = { bold: true, color: { argb: 'FFB91C1C' } };
    }
  });

  const header = sheet.getRow(1);

  header.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  header.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF14532D' }
  };
  header.alignment = { vertical: 'middle', wrapText: true };

  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: sheet.columns.length }
  };

  res.set({
    'Content-Type':
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition':
      'attachment; filename="pagos-pendientes-tarpuyana-2026.xlsx"',
    'Cache-Control': 'no-store'
  });

  await workbook.xlsx.write(res);
  res.end();
}));

// ==========================================
// RESULTADOS DEL CONCURSO
// ==========================================

const pad2 = (n) => String(n).padStart(2, '0');

// Puntaje máximo según el nivel del alumno.
const PUNTAJE_MAX = { Inicial: 200, Primaria: 150, Secundaria: 150 };

function parsePuntaje(value, nivel) {
  const max = PUNTAJE_MAX[nivel] ?? 150;
  const n = Number(String(value ?? '').trim().replace(',', '.'));

  if (!Number.isFinite(n) || n < 0) {
    throw bad('El puntaje debe ser un número válido.');
  }

  if (n > max) {
    throw bad(`El puntaje máximo en ${nivel || 'este nivel'} es ${max}.`);
  }

  return Math.round(n * 100) / 100;
}

// Acepta "10:42:15", "10:42", Date de Excel o fracción de día.
function parseHora(value) {
  if (value === null || value === undefined || value === '') return null;

  let h;
  let m;
  let s;

  if (value instanceof Date) {
    h = value.getUTCHours();
    m = value.getUTCMinutes();
    s = value.getUTCSeconds();
  } else if (typeof value === 'number') {
    const total = Math.round((value % 1) * 86400);
    h = Math.floor(total / 3600) % 24;
    m = Math.floor((total % 3600) / 60);
    s = total % 60;
  } else {
    const match = String(value).trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);

    if (!match) throw bad('La hora debe tener el formato HH:MM:SS.');

    h = Number(match[1]);
    m = Number(match[2]);
    s = Number(match[3] || 0);
  }

  if (h > 23 || m > 59 || s > 59) {
    throw bad('La hora de entrega no es válida.');
  }

  return `${pad2(h)}:${pad2(m)}:${pad2(s)}`;
}

async function resultadosPublicados() {
  const r = await pool.query(
    "SELECT valor FROM config WHERE clave = 'resultados_publicados'"
  );

  return r.rows.length > 0 && r.rows[0].valor === '1';
}

// ---- Ranking (misma regla que la página): mayor puntaje; en empate, menor hora ----
const catKeyS = (c) =>
  /LEIBNIZ|[“"]A[”"]/i.test(c || '') ? 'A' : /GAUSS|[“"]B[”"]/i.test(c || '') ? 'B' : '';

const horaOrden = (h) => h || '99:99:99';

function rankear(lista) {
  const a = [...lista].sort(
    (x, y) =>
      y.puntaje - x.puntaje ||
      horaOrden(x.hora).localeCompare(horaOrden(y.hora)) ||
      x.nombre.localeCompare(y.nombre, 'es')
  );

  a.forEach((r, i) => {
    const p = a[i - 1];
    r.merito =
      !p || p.puntaje !== r.puntaje || horaOrden(p.hora) !== horaOrden(r.hora)
        ? i + 1
        : p.merito;
  });

  return a;
}

async function cargarAlumnosResultados() {
  const r = await pool.query(
    `SELECT dni, nombre, grado, nivel, categoria, puntaje,
            to_char(hora_entrega, 'HH24:MI:SS') AS hora
     FROM libres`
  );

  return r.rows.map((x) => ({
    dni: x.dni,
    nombre: x.nombre,
    grado: x.grado,
    nivel: x.nivel,
    categoria: x.categoria || '',
    cat: catKeyS(x.categoria),
    puntaje: x.puntaje === null ? null : Number(x.puntaje),
    hora: x.hora || ''
  }));
}

// Público: solo los 5 primeros puestos de cada grado y categoría, más los
// totales por grado (para Interaulas). El resto de alumnos NO se envía:
// se consulta uno por uno con /api/resultados/buscar.
app.get('/api/resultados', wrap(async (req, res) => {
  res.set('Cache-Control', 'no-store');

  const publicado = await resultadosPublicados();
  const admin = isAdmin(req);

  if (!publicado && !admin) {
    return res.json({ publicado: false, vistaPrevia: false, items: [], aulas: [] });
  }

  const todos = await cargarAlumnosResultados();
  const items = [];

  for (const grado of [...new Set(todos.map((x) => x.grado))]) {
    for (const cat of ['A', 'B']) {
      const grupo = todos.filter(
        (x) => x.puntaje !== null && x.grado === grado && x.cat === cat
      );

      rankear(grupo)
        .filter((x) => x.merito <= 5)
        .forEach((x) =>
          items.push({
            nombre: x.nombre,
            grado: x.grado,
            nivel: x.nivel,
            categoria: x.categoria,
            puntaje: x.puntaje,
            hora: x.hora
          })
        );
    }
  }

  const porGrado = {};

  todos
    .filter((x) => x.puntaje !== null)
    .forEach((x) => {
      const g = (porGrado[x.grado] = porGrado[x.grado] || { grado: x.grado, n: 0, suma: 0 });
      g.n += 1;
      g.suma += x.puntaje;
    });

  res.json({
    publicado,
    vistaPrevia: !publicado && admin,
    items,
    aulas: Object.values(porGrado)
  });
}));

// ---- Búsqueda personal: por DNI exacto o por apellidos y nombres ----
const buscarLimiter = rateLimit({
  windowMs: 60 * 1000,
  limit: 25,
  standardHeaders: true,
  legacyHeaders: false,
  message: {
    error: 'Demasiadas búsquedas seguidas. Espera un minuto e inténtalo otra vez.'
  }
});

const normTexto = (s) =>
  String(s || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

app.get('/api/resultados/buscar', buscarLimiter, wrap(async (req, res) => {
  res.set('Cache-Control', 'no-store');

  const publicado = await resultadosPublicados();

  if (!publicado && !isAdmin(req)) {
    return res.json({ resultados: [], noPublicado: true });
  }

  const q = clean(req.query.q, 80);
  const soloDigitos = q.replace(/\s/g, '');
  const esDni = /^\d+$/.test(soloDigitos);

  if (esDni && soloDigitos.length !== 8) {
    throw bad('El DNI debe tener 8 dígitos.');
  }

  const tokens = normTexto(q).split(' ').filter(Boolean);

  if (!esDni && (tokens.length < 2 || tokens.join('').length < 5)) {
    throw bad('Escribe al menos dos palabras: apellidos y nombres del alumno.');
  }

  const todos = await cargarAlumnosResultados();

  const coincide = (x) => {
    if (esDni) return x.dni === soloDigitos;

    const palabras = normTexto(x.nombre).split(' ');

    return tokens.every((tk) => palabras.some((w) => w.startsWith(tk)));
  };

  const encontrados = todos.filter(coincide).slice(0, 8);

  const resultados = encontrados.map((x) => {
    const base = { nombre: x.nombre, grado: x.grado, nivel: x.nivel, cat: x.cat };

    if (x.puntaje === null) {
      return { ...base, puntaje: null, hora: '', max: PUNTAJE_MAX[x.nivel] ?? 150 };
    }

    const enNivel = rankear(
      todos.filter((y) => y.puntaje !== null && y.nivel === x.nivel && y.cat === x.cat)
    );
    const enGrado = rankear(
      todos.filter((y) => y.puntaje !== null && y.grado === x.grado && y.cat === x.cat)
    );
    const yo = (lista) => lista.find((y) => y.dni === x.dni);

    return {
      ...base,
      puntaje: x.puntaje,
      hora: x.hora,
      max: PUNTAJE_MAX[x.nivel] ?? 150,
      meritoNivel: yo(enNivel).merito,
      totalNivel: enNivel.length,
      meritoGrado: yo(enGrado).merito,
      totalGrado: enGrado.length
    };
  });

  res.json({ resultados });
}));

app.get('/api/admin/publicado', needAdmin, wrap(async (_req, res) => {
  res.set('Cache-Control', 'no-store');
  res.json({ publicado: await resultadosPublicados() });
}));

app.put('/api/admin/publicar', needAdmin, wrap(async (req, res) => {
  const publicado = Boolean(req.body && req.body.publicado);

  await pool.query(
    `INSERT INTO config (clave, valor)
     VALUES ('resultados_publicados', $1)
     ON CONFLICT (clave) DO UPDATE SET valor = EXCLUDED.valor`,
    [publicado ? '1' : '0']
  );

  res.json({ ok: true, publicado });
}));

// Guardar (o quitar) el puntaje de un alumno a partir de su DNI.
app.put('/api/admin/puntaje', needAdmin, wrap(async (req, res) => {
  const body = req.body || {};
  const dni = clean(body.dni, 14);

  if (!/^\d{8}$/.test(dni)) throw bad('El DNI debe tener 8 dígitos.');

  const found = await pool.query('SELECT nivel FROM libres WHERE dni = $1', [dni]);

  if (!found.rowCount) {
    throw new HttpError(404, `No hay ningún alumno inscrito con el DNI ${dni}.`);
  }

  const quitar = body.puntaje === null || body.puntaje === '';
  const puntaje = quitar ? null : parsePuntaje(body.puntaje, found.rows[0].nivel);
  const hora = quitar ? null : parseHora(body.hora);

  const r = await pool.query(
    `UPDATE libres
     SET puntaje = $2, hora_entrega = $3
     WHERE dni = $1
     RETURNING nombre`,
    [dni, puntaje, hora]
  );

  if (!r.rowCount) {
    throw new HttpError(404, `No hay ningún alumno inscrito con el DNI ${dni}.`);
  }

  res.json({ ok: true, nombre: r.rows[0].nombre });
}));

// Quitar puntaje y hora de forma masiva: por lista de DNI o de todo el concurso.
app.post('/api/admin/quitar-puntajes', needAdmin, wrap(async (req, res) => {
  const body = req.body || {};
  let r;

  if (body.todos === true) {
    if (body.confirmar !== 'ELIMINAR') throw bad('Falta la confirmación.');

    r = await pool.query(
      `UPDATE libres SET puntaje = NULL, hora_entrega = NULL
       WHERE puntaje IS NOT NULL OR hora_entrega IS NOT NULL`
    );
  } else {
    const dnis = Array.isArray(body.dnis) ? [...new Set(body.dnis.map(String))] : [];

    if (!dnis.length) throw bad('No seleccionaste ningún alumno.');
    if (dnis.length > 5000 || dnis.some((d) => !/^\d{8}$/.test(d))) {
      throw bad('La lista de DNI no es válida.');
    }

    r = await pool.query(
      `UPDATE libres SET puntaje = NULL, hora_entrega = NULL
       WHERE dni = ANY($1::text[])
         AND (puntaje IS NOT NULL OR hora_entrega IS NOT NULL)`,
      [dnis]
    );
  }

  res.json({ ok: true, quitados: r.rowCount });
}));

// Plantilla Excel con todos los inscritos, lista para llenar Puntaje y Hora.
app.get('/api/admin/plantilla-puntajes.xlsx', needAdmin, wrap(async (_req, res) => {
  const lista = await loadInscripciones();

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Puntajes', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  sheet.columns = [
    { header: 'DNI', key: 'dni', width: 14 },
    { header: 'Alumno', key: 'nombre', width: 40 },
    { header: 'Grado', key: 'grado', width: 20 },
    { header: 'Categoría', key: 'categoria', width: 48 },
    { header: 'Puntaje máx.', key: 'max', width: 14 },
    { header: 'Puntaje', key: 'puntaje', width: 12 },
    { header: 'Hora de entrega (HH:MM:SS)', key: 'hora', width: 28 }
  ];

  sheet.getColumn('dni').numFmt = '@';
  sheet.getColumn('hora').numFmt = '@';

  [...lista]
    .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'))
    .forEach((x) => {
      sheet.addRow({
        dni: x.dni,
        nombre: x.nombre,
        grado: x.grado,
        categoria: x.categoria,
        max: PUNTAJE_MAX[x.nivel] ?? 150,
        puntaje: x.puntaje,
        hora: x.hora
      });
    });

  const header = sheet.getRow(1);
  header.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  header.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF14532D' }
  };

  res.set({
    'Content-Type':
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition':
      'attachment; filename="plantilla-puntajes-tarpuyana-2026.xlsx"',
    'Cache-Control': 'no-store'
  });

  await workbook.xlsx.write(res);
  res.end();
}));

const normHeader = (value) =>
  String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();

function cellValue(cell) {
  let v = cell.value;

  if (v && typeof v === 'object' && !(v instanceof Date)) {
    if ('result' in v) v = v.result;
    else if ('text' in v) v = v.text;
    else if (Array.isArray(v.richText)) v = v.richText.map((t) => t.text).join('');
  }

  return v;
}

// Subida masiva: Excel con columnas DNI, Puntaje y Hora de entrega.
app.post(
  '/api/admin/puntajes-masivo',
  needAdmin,
  express.raw({ type: () => true, limit: '5mb' }),
  wrap(async (req, res) => {
    if (!Buffer.isBuffer(req.body) || !req.body.length) {
      throw bad('Sube un archivo Excel (.xlsx).');
    }

    const workbook = new ExcelJS.Workbook();

    try {
      await workbook.xlsx.load(req.body);
    } catch (_err) {
      throw bad('No se pudo leer el archivo. Debe ser un Excel (.xlsx).');
    }

    const sheet = workbook.worksheets[0];

    if (!sheet) throw bad('El Excel no tiene hojas.');

    let headerRow = 0;
    let cDni = 0;
    let cPun = 0;
    let cHora = 0;

    for (let r = 1; r <= Math.min(5, sheet.rowCount) && !headerRow; r += 1) {
      sheet.getRow(r).eachCell((cell, col) => {
        const h = normHeader(cellValue(cell));

        if (h === 'dni') cDni = col;
        else if ((h.startsWith('puntaje') && !h.includes('max')) || h === 'puntos' || h === 'nota') cPun = col;
        else if (h.startsWith('hora')) cHora = col;
      });

      if (cDni && cPun) headerRow = r;
    }

    if (!headerRow) {
      throw bad('No encontré las columnas "DNI" y "Puntaje". Descarga la plantilla y úsala.');
    }

    const filas = [];
    const errores = [];
    const niveles = new Map(
      (await pool.query('SELECT dni, nivel FROM libres')).rows.map((x) => [x.dni, x.nivel])
    );

    for (let r = headerRow + 1; r <= sheet.rowCount; r += 1) {
      const row = sheet.getRow(r);
      const rawDni = cellValue(row.getCell(cDni));
      const rawPun = cellValue(row.getCell(cPun));
      const rawHora = cHora ? cellValue(row.getCell(cHora)) : null;

      const vacio = (v) => v === null || v === undefined || String(v).trim() === '';

      if (vacio(rawDni) || (vacio(rawPun) && vacio(rawHora))) continue;

      try {
        const dni = String(rawDni).trim().replace(/\.0+$/, '').padStart(8, '0');

        if (!/^\d{8}$/.test(dni)) throw bad(`DNI no válido (${rawDni}).`);
        if (!niveles.has(dni)) throw bad(`El DNI ${dni} no está inscrito.`);
        if (vacio(rawPun)) throw bad('Falta el puntaje.');

        filas.push({
          fila: r,
          dni,
          puntaje: parsePuntaje(rawPun, niveles.get(dni)),
          hora: parseHora(rawHora)
        });
      } catch (err) {
        errores.push({ fila: r, mensaje: err.message });
      }

      if (filas.length + errores.length > 5000) {
        throw bad('El archivo tiene demasiadas filas (máximo 5000).');
      }
    }

    let actualizados = 0;
    const client = await pool.connect();

    try {
      await client.query('BEGIN');

      for (const f of filas) {
        const r = await client.query(
          `UPDATE libres SET puntaje = $2, hora_entrega = $3 WHERE dni = $1`,
          [f.dni, f.puntaje, f.hora]
        );

        if (r.rowCount) actualizados += 1;
        else errores.push({ fila: f.fila, mensaje: `El DNI ${f.dni} no está inscrito.` });
      }

      await client.query('COMMIT');
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }

    errores.sort((a, b) => a.fila - b.fila);

    res.json({
      ok: true,
      actualizados,
      totalErrores: errores.length,
      errores: errores.slice(0, 50)
    });
  })
);

// ==========================================
// ARCHIVOS ESTÁTICOS
// ==========================================

app.use(
  express.static(
    path.join(__dirname, 'public'),
    { maxAge: '1h' }
  )
);

app.use('/api', (_req, res) => {
  res.status(404).json({
    error: 'Ruta API no encontrada.'
  });
});

// ==========================================
// MANEJO DE ERRORES
// ==========================================

app.use((err, _req, res, _next) => {
  if (err instanceof HttpError) {
    return res.status(err.status).json({
      error: err.message
    });
  }

  if (err && err.type === 'entity.parse.failed') {
    return res.status(400).json({
      error: 'El contenido enviado no es un JSON válido.'
    });
  }

  console.error('Error del servidor:', err);

  return res.status(500).json({
    error: 'Error del servidor. Inténtalo de nuevo en un momento.'
  });
});

// ==========================================
// INICIAR SERVIDOR
// ==========================================

(async () => {
  await pool.query(SCHEMA);

  app.listen(PORT, '0.0.0.0', () => {
    console.log(
      `VIII Olimpiada Tarpuyana escuchando en el puerto ${PORT}`
    );
  });
})().catch((err) => {
  console.error('No se pudo iniciar el servidor:', err);
  process.exit(1);
});
