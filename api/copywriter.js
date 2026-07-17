// Vercel Serverless Function
// POST /api/copywriter
//
// Agente E (UX Copywriter) — lee reviews reales de distintas plataformas
// (archivos exportados Y/o links de páginas de venta), detecta qué valoran
// más los clientes, y genera copy de conversión anclado en esos insights,
// por retailer o como pieza genérica para reversionar.
//
// División de responsabilidades (mismo principio que Agentes A/F/G):
// Claude NUNCA calcula un porcentaje. Claude solo hace dos cosas: (1)
// clasificar cada review individual en los temas que él mismo detecta, y
// (2) redactar copy sobre los temas ya identificados. El conteo de cuántas
// reviews caen en cada tema, y el % que eso representa, se calcula acá en
// JS a partir de esa clasificación — de forma determinística y auditable
// (cada tema queda con la lista real de ids de reviews que lo sustentan).

const MAX_TEXTO_LARGO = 4000;
const MAX_TEXTO_REVIEWS = 60000;
const MAX_FILAS_TABLA = 8000;
// Cuantas más reviews se le pidan clasificar a Claude en una sola corrida,
// más larga es la respuesta que tiene que generar (un ítem de clasificación
// por review) y más chance de pasarse del maxDuration de la función
// serverless. 110 es un techo que en la práctica da tiempo de sobra para
// terminar dentro de los 90s configurados en vercel.json mantiendo una
// muestra representativa.
const MAX_REVIEWS = 110;
const MAX_TEMAS = 8;
const MAX_LINKS = 8;
const LINK_TIMEOUT_MS = 12000;
const CLAUDE_TIMEOUT_MS = 65000;

/* ================= PARSERS (archivos) ================= */

function bufferDesdeBase64(base64) {
  const data = String(base64).split(',').pop();
  return Buffer.from(data, 'base64');
}

function parseTabular(buffer, nombreArchivo) {
  let XLSX;
  try { XLSX = require('xlsx'); } catch (e) { return { error: 'No se pudo cargar el parser de planillas (xlsx).' }; }
  try {
    const wb = XLSX.read(buffer, { type: 'buffer' });
    const hoja = wb.Sheets[wb.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(hoja, { defval: '' });
    return {
      headers: filas.length ? Object.keys(filas[0]) : [],
      filas: filas.slice(0, MAX_FILAS_TABLA),
      totalFilasOriginal: filas.length,
      truncado: filas.length > MAX_FILAS_TABLA
    };
  } catch (e) { return { error: `No se pudo leer "${nombreArchivo}" como planilla: ${e.message}` }; }
}

async function parsePDF(buffer, nombreArchivo, textCap) {
  let pdfParse;
  try { pdfParse = require('pdf-parse'); } catch (e) { return { error: 'No se pudo cargar el parser de PDF.' }; }
  try {
    const data = await pdfParse(buffer);
    return { texto: (data.text || '').slice(0, textCap) };
  } catch (e) { return { error: `No se pudo leer "${nombreArchivo}" como PDF: ${e.message}` }; }
}

// Word, PowerPoint, OpenDocument y RTF — vía officeparser (misma librería
// para las cuatro cosas, sin dependencias nativas).
async function parseOfficeDoc(buffer, nombreArchivo, ext, textCap) {
  let officeParser;
  try { officeParser = require('officeparser'); } catch (e) { return { error: 'No se pudo cargar el parser de documentos de Office.' }; }
  try {
    const ast = await officeParser.parseOffice(buffer, { fileType: ext });
    const texto = typeof ast.toText === 'function' ? ast.toText() : '';
    return { texto: (texto || '').slice(0, textCap) };
  } catch (e) { return { error: `No se pudo leer "${nombreArchivo}" (${ext}): ${e.message}` }; }
}

// textCap: los documentos de contexto (brief, specs, producto) usan
// MAX_TEXTO_LARGO porque son texto de apoyo. Las fuentes de reviews en
// TXT/MD/PDF usan un cap mucho más grande (MAX_TEXTO_REVIEWS) porque ahí sí
// son los datos primarios del análisis.
async function parseArchivo(archivo, textCap = MAX_TEXTO_LARGO) {
  const nombre = archivo.nombre || 'archivo';
  const ext = (nombre.split('.').pop() || '').toLowerCase();
  const buffer = bufferDesdeBase64(archivo.base64);
  if (['csv', 'xlsx', 'xls'].includes(ext)) return { kind: 'tabular', nombre, ...parseTabular(buffer, nombre) };
  if (ext === 'pdf') return { kind: 'texto', nombre, ...(await parsePDF(buffer, nombre, textCap)) };
  if (['txt', 'md', 'json'].includes(ext)) return { kind: 'texto', nombre, texto: buffer.toString('utf-8').slice(0, textCap) };
  if (['docx', 'pptx', 'odt', 'odp', 'ods', 'rtf'].includes(ext)) return { kind: 'texto', nombre, ...(await parseOfficeDoc(buffer, nombre, ext, textCap)) };
  return { kind: 'desconocido', nombre, error: `Formato de "${nombre}" no reconocido — probá CSV, XLSX, PDF, Word, PowerPoint, TXT o MD.` };
}

/* ================= PARSER DE LINKS (páginas de venta / reseñas) ================= */

function hostnameDe(url) {
  try { return new URL(url).hostname.replace(/^www\./, ''); }
  catch (e) { return String(url || '').slice(0, 60); }
}

// No hay browser headless acá (función serverless liviana) — se lee el
// HTML tal como lo devuelve el server. Funciona bien en sitios que
// renderizan las reviews del lado del servidor (frecuente en fichas de
// producto de retailers y marketplaces); si una página arma sus reviews
// 100% con JavaScript del lado del cliente, puede no traer nada útil — en
// ese caso conviene exportar las reviews como archivo en vez de pegar el link.
async function parseLinkHTML(url, textCap) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), LINK_TIMEOUT_MS);
  try {
    const r = await fetch(url, {
      signal: controller.signal,
      redirect: 'follow',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'es-419,es;q=0.9'
      }
    });
    if (!r.ok) return { error: `La página respondió con estado ${r.status} — puede estar bloqueando accesos automáticos.` };
    const html = await r.text();
    const texto = html
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<!--[\s\S]*?-->/g, ' ')
      .replace(/<(br|\/p|\/div|\/li|\/tr|\/h[1-6])\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&quot;/gi, '"')
      .replace(/&#0?39;/gi, "'")
      .replace(/[ \t]{2,}/g, ' ')
      .replace(/\n{2,}/g, '\n')
      .trim();
    if (!texto) return { error: 'La página no devolvió contenido de texto legible.' };
    return { texto: texto.slice(0, textCap) };
  } catch (e) {
    if (e.name === 'AbortError') return { error: 'La página tardó demasiado en responder (timeout de 12s).' };
    return { error: `No se pudo leer la página: ${e.message}` };
  } finally {
    clearTimeout(timeout);
  }
}

function normalizar(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}

const REVIEW_ALIAS = {
  texto: ['texto', 'comentario', 'review', 'resena', 'reseña', 'contenido', 'mensaje', 'opinion', 'opinión', 'comment', 'review_text', 'body'],
  rating: ['rating', 'calificacion', 'calificación', 'estrellas', 'puntaje', 'score', 'stars']
};

function detectarColumnasReviews(headers) {
  const norm = headers.map(h => ({ original: h, norm: normalizar(h) }));
  const mapping = {};
  for (const campo of Object.keys(REVIEW_ALIAS)) {
    const candidatos = REVIEW_ALIAS[campo];
    const match = norm.find(h => candidatos.includes(h.norm)) || norm.find(h => candidatos.some(c => h.norm.includes(c)));
    if (match) mapping[campo] = match.original;
  }
  return mapping;
}

function numeroDesde(valor) {
  if (typeof valor === 'number') return valor;
  const limpio = String(valor || '').replace(/[^\d.,-]/g, '').replace(/\.(?=\d{3},)/g, '').replace(',', '.');
  const n = parseFloat(limpio);
  return isNaN(n) ? 0 : n;
}

function partirTextoEnComentarios(texto) {
  return String(texto || '')
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(l => l.length >= 8 && l.length <= 600);
}

const TIPO_A_PLATAFORMA = {
  reviewsPropias: 'Tienda propia',
  reviewsML: 'Mercado Libre',
  reviewsAmazon: 'Amazon',
  reviewsApp: 'App (Store / Google Play)',
  redesSociales: 'Redes sociales'
};

/* ================= MUESTREO DETERMINÍSTICO ================= */

// Si hay más reviews de las que conviene mandarle a Claude en una sola
// corrida, se recorta de forma proporcional por plataforma (nunca al azar
// ni descartando una plataforma entera) para no sesgar el análisis hacia
// la fuente con más volumen.
function muestrearReviews(reviews, cap) {
  if (reviews.length <= cap) return reviews;
  const plataformas = [...new Set(reviews.map(r => r.plataforma))];
  const resultado = [];
  for (const p of plataformas) {
    const deEstaPlataforma = reviews.filter(r => r.plataforma === p);
    const cuota = Math.max(1, Math.floor(cap * deEstaPlataforma.length / reviews.length));
    resultado.push(...deEstaPlataforma.slice(0, cuota));
  }
  return resultado.slice(0, cap);
}

/* ================= COBERTURA (equivalente a "precisión" en G) ================= */

function calcularCobertura(reviewsUsadas, totalCargadas, plataformasConDatos) {
  const n = reviewsUsadas.length;
  if (n === 0) return { score: 0, nivel: 'sin datos' };
  let score;
  if (n < 20) score = 30;
  else if (n < 60) score = 55;
  else if (n < 150) score = 75;
  else score = 88;
  const diversidad = plataformasConDatos.length;
  if (diversidad >= 3) score += 8;
  else if (diversidad === 1) score -= 10;
  score = Math.max(15, Math.min(97, score));
  const nivel = score >= 75 ? 'alta' : score >= 50 ? 'media' : 'baja';
  return { score, nivel, n, totalCargadas, diversidad };
}

function sugerenciaCobertura({ n, totalCargadas, plataformasFaltantes }) {
  const partes = [];
  if (n < 60) partes.push(`Sumar más reviews (hoy hay ${n} analizadas) daría una base más confiable para los porcentajes.`);
  if (plataformasFaltantes.length) partes.push(`Todavía falta cargar reviews de: ${plataformasFaltantes.join(', ')}.`);
  if (totalCargadas > n) partes.push(`Se analizó una muestra de ${n} de ${totalCargadas} reviews cargadas, repartida proporcionalmente por plataforma, para mantener la corrida liviana.`);
  if (!partes.length) partes.push('Buena cobertura — varias plataformas y volumen suficiente de reviews.');
  return partes.join(' ');
}

/* ================= RETAILERS OBJETIVO ================= */

function parseRetailers(texto) {
  const partes = String(texto || '')
    .split(/[,;/]| y |&/i)
    .map(s => s.trim())
    .filter(Boolean);
  if (!partes.length) return [{ slug: 'generico', nombre: 'Creativo genérico (a reversionar)' }];
  const vistos = new Set();
  const out = [];
  for (const nombre of partes) {
    const slug = normalizar(nombre).replace(/[^a-z0-9]/g, '');
    if (!slug || vistos.has(slug)) continue;
    vistos.add(slug);
    out.push({ slug, nombre });
  }
  return out.length ? out : [{ slug: 'generico', nombre: 'Creativo genérico (a reversionar)' }];
}

/* ================= PROMPT Y LLAMADA A CLAUDE ================= */

function construirPrompt({ cliente, producto, objetivo, retailers, briefTexto, specsTexto, perfilTexto, productoTexto, creativoBase }, reviews) {
  const bloques = [];
  bloques.push(`CLIENTE: ${cliente || 'sin nombre'}\nPRODUCTO / LÍNEA: ${producto || 'no especificado'}\nOBJETIVO DE CAMPAÑA: ${objetivo || 'no especificado'}\nRETAILERS OBJETIVO: ${JSON.stringify(retailers)}`);

  if (productoTexto) bloques.push(`\n=== DOCUMENTOS DEL PRODUCTO (fichas técnicas, brand book, info oficial) ===\n${productoTexto}`);
  if (briefTexto) bloques.push(`\n=== BRIEF DE CAMPAÑA / LINEAMIENTOS DE MARCA ===\n${briefTexto}`);
  if (specsTexto) bloques.push(`\n=== SPECS DE RETAILER (si se subieron) ===\n${specsTexto}`);
  if (perfilTexto) bloques.push(`\n=== PERFIL DE CONSUMIDOR (Agente F, si se subió) ===\n${perfilTexto}`);
  if (creativoBase && creativoBase.trim()) bloques.push(`\n=== CREATIVO BASE A REVERSIONAR ===\n${creativoBase.trim()}`);

  bloques.push(`\n=== REVIEWS A CLASIFICAR (id, plataforma, texto) — NO calcules porcentajes, solo etiquetá cada una ===\n${JSON.stringify(reviews.map(r => ({ id: r.id, plataforma: r.plataforma, texto: r.texto })))}`);

  return bloques.join('\n');
}

const SYSTEM_PROMPT = `Sos el motor de clasificación y redacción del Agente E (UX Copywriter) de LiftyFive, una agencia de retail media AI-first.

Tu trabajo tiene dos partes:

PARTE 1 — CLASIFICACIÓN (etiquetado, no matemática):
Vas a recibir una lista de reviews reales, cada una con un "id" y su plataforma de origen (puede ser una tienda, un marketplace, una app, redes sociales, o el dominio de una página de venta leída directamente por link). Tenés que:
1. Detectar como máximo 8 temas/insights recurrentes — lo que los clientes más valoran o más critican. Cada tema lleva un nombre corto y un sentimiento ("positivo", "negativo" o "neutro"). Usá lo que la gente REALMENTE dice, no categorías genéricas de manual de marketing.
2. Para CADA review de la lista, indicar a qué tema(s) pertenece (puede ser ninguno, uno o varios). NUNCA calculés porcentajes ni conteos vos mismo — eso lo hace el sistema después, contando cuántas reviews etiquetaste con cada tema. Tu única responsabilidad acá es clasificar bien, review por review.

PARTE 2 — COPY (redacción, sobre los temas que vos mismo detectaste en la Parte 1):
Con los temas ya identificados — y usando también los DOCUMENTOS DEL PRODUCTO si se incluyeron, para no contradecir ningún atributo o claim oficial — escribí copy de conversión para cada retailer en RETAILERS OBJETIVO (usá el "slug" de cada uno como key del objeto "copy"):
- "titulos_banner": EXACTAMENTE 2 títulos de banner alternativos, cada uno con el "tema_id" del tema en que se basa y una "razon" de 1 oración explicando por qué ese mensaje funciona para ese retailer y ese objetivo de campaña.
- "brand_store_desc": una descripción de Brand Store (2-3 oraciones).
- "titulo_imagen": un título corto para overlay de imagen (máximo 8 palabras).

Reglas estrictas:
- Nunca uses como tema_id principal de un título o descripción un tema con sentimiento "negativo" — un tema negativo solo se puede mencionar como alerta interna en el resumen ejecutivo, nunca como mensaje de venta.
- Nunca inventes un tema_id que no hayas declarado vos mismo en "temas_detectados", ni un review_id que no esté en la lista que te pasaron.
- El copy tiene que sonar distinto entre retailers si el tono o el comprador típico de cada uno es distinto — no repitas el mismo texto para todos.
- Si RETAILERS OBJETIVO trae un solo elemento con slug "generico", escribí una pieza más flexible, sin nombrar un retailer puntual, pensada para reversionar después en cualquier canal.
- Si hay un CREATIVO BASE A REVERSIONAR, tu copy nuevo tiene que partir de esa base y adaptarla — no ignorarla y escribir algo completamente distinto.
- Si hay DOCUMENTOS DEL PRODUCTO, nunca redactes un claim que los contradiga (ej. un ingrediente, certificación o beneficio que el documento no confirma).
- No devuelvas texto fuera del JSON. Respondé ÚNICAMENTE con un JSON válido, sin bloques de markdown, con este schema exacto:

{
  "temas_detectados": [
    { "id": "tema_1", "nombre": "string corto", "sentimiento": "positivo" }
  ],
  "clasificacion": [
    { "review_id": 0, "temas": ["tema_1", "tema_3"] }
  ],
  "copy": {
    "<slug_retailer>": {
      "titulos_banner": [
        { "texto": "string", "tema_id": "string", "razon": "string" },
        { "texto": "string", "tema_id": "string", "razon": "string" }
      ],
      "brand_store_desc": "string",
      "titulo_imagen": "string"
    }
  },
  "resumen_ejecutivo": "string (2-3 oraciones sobre el panorama general de insights y el copy generado)"
}`;

async function llamarClaude(prompt, apiKey) {
  // Timeout propio, menor al maxDuration de la función serverless. Así, si
  // Claude se cuelga generando una respuesta muy larga, cortamos nosotros
  // con un error entendible en vez de que Vercel mate la función y devuelva
  // una página de error en texto plano que el frontend no puede parsear.
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), CLAUDE_TIMEOUT_MS);
  let r;
  try {
    r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      signal: controller.signal,
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      body: JSON.stringify({ model: 'claude-sonnet-5', max_tokens: 8000, system: SYSTEM_PROMPT, messages: [{ role: 'user', content: prompt }] })
    });
  } catch (e) {
    if (e.name === 'AbortError') throw new Error('Claude tardó demasiado en responder para esta cantidad de reviews. Probá con menos reviews por corrida (por ejemplo, subiendo un archivo más chico o menos plataformas a la vez).');
    throw e;
  } finally {
    clearTimeout(timeout);
  }
  const data = await r.json();
  if (!r.ok) throw new Error(data?.error?.message || 'Error llamando a la API de Claude.');

  const bloqueTexto = (data?.content || []).find(b => b.type === 'text');
  const texto = bloqueTexto?.text || '';
  if (!texto) {
    console.error('Claude respondió sin bloque de texto. stop_reason:', data?.stop_reason);
    throw new Error(`Claude no devolvió texto en la respuesta (stop_reason: ${data?.stop_reason || 'desconocido'}).`);
  }
  const limpio = texto.trim().replace(/^```(json)?/i, '').replace(/```$/, '').trim();
  try {
    return JSON.parse(limpio);
  } catch (e) {
    console.error('JSON inválido de Claude. stop_reason:', data?.stop_reason, 'respuesta completa:', texto);
    throw new Error(`Claude no devolvió un JSON válido (stop_reason: ${data?.stop_reason || 'desconocido'}). Respuesta cruda: ` + (limpio.slice(0, 500) || '(vacía)'));
  }
}

/* ================= HANDLER ================= */

module.exports = async (req, res) => {
  if (req.method !== 'POST') { res.status(405).json({ error: 'Método no permitido. Usá POST.' }); return; }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) { res.status(500).json({ error: 'Falta configurar ANTHROPIC_API_KEY en las variables de entorno de Vercel.' }); return; }

  let body;
  try { body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
  catch (e) { res.status(400).json({ error: 'Body inválido.' }); return; }

  const { cliente, producto, objetivoCampana, retailersObjetivo, creativoBase, archivos = [], links = [] } = body || {};

  try {
    // --- Reviews por plataforma (archivos exportados) ---
    const tiposReviews = ['reviewsPropias', 'reviewsML', 'reviewsAmazon', 'reviewsApp', 'redesSociales'];
    let reviewsRaw = [];
    let fuentesResumen = [];

    for (const tipo of tiposReviews) {
      const archivosTipo = archivos.filter(a => a.tipo === tipo);
      const plataforma = TIPO_A_PLATAFORMA[tipo];
      if (!archivosTipo.length) continue;
      let count = 0;
      for (const archivo of archivosTipo) {
        const parsed = await parseArchivo(archivo, MAX_TEXTO_REVIEWS);
        if (parsed.error) { fuentesResumen.push({ plataforma, nombre: archivo.nombre, usada: false, detalle: parsed.error }); continue; }
        if (parsed.kind === 'tabular') {
          const mapping = detectarColumnasReviews(parsed.headers);
          if (!mapping.texto) { fuentesResumen.push({ plataforma, nombre: archivo.nombre, usada: false, detalle: 'No se detectó una columna de texto de review.' }); continue; }
          for (const fila of parsed.filas) {
            const texto = String(fila[mapping.texto] || '').trim();
            if (!texto) continue;
            const rating = mapping.rating ? numeroDesde(fila[mapping.rating]) : null;
            reviewsRaw.push({ texto: texto.slice(0, 500), rating, plataforma });
            count++;
          }
        } else if (parsed.kind === 'texto') {
          const comentarios = partirTextoEnComentarios(parsed.texto);
          comentarios.forEach(c => { reviewsRaw.push({ texto: c, rating: null, plataforma }); count++; });
        }
      }
      fuentesResumen.push({ plataforma, nombre: archivosTipo.map(a => a.nombre).join(', '), usada: count > 0, detalle: `${count} reviews detectadas.` });
    }

    // --- Reviews desde links de venta (páginas leídas en vivo) ---
    const linksLimpios = [...new Set((Array.isArray(links) ? links : [])
      .map(l => String(l || '').trim())
      .filter(l => /^https?:\/\//i.test(l)))].slice(0, MAX_LINKS);

    if (linksLimpios.length) {
      const resultadosLinks = await Promise.all(linksLimpios.map(async (url) => {
        const hostname = hostnameDe(url);
        const parsed = await parseLinkHTML(url, MAX_TEXTO_REVIEWS);
        return { url, hostname, parsed };
      }));
      for (const { url, hostname, parsed } of resultadosLinks) {
        if (parsed.error) { fuentesResumen.push({ plataforma: hostname, nombre: url, usada: false, detalle: parsed.error }); continue; }
        const comentarios = partirTextoEnComentarios(parsed.texto);
        comentarios.forEach(c => { reviewsRaw.push({ texto: c, rating: null, plataforma: hostname }); });
        fuentesResumen.push({
          plataforma: hostname, nombre: url, usada: comentarios.length > 0,
          detalle: comentarios.length ? `${comentarios.length} fragmentos de texto leídos de la página.` : 'La página respondió pero no se detectó texto aprovechable (probablemente carga las reviews con JavaScript del lado del cliente).'
        });
      }
    }

    if (!reviewsRaw.length) {
      res.status(400).json({ error: 'Necesito al menos una fuente de reviews — un archivo exportado (tienda propia, Mercado Libre, Amazon, app o redes) o un link a una página de venta — para poder identificar insights reales. Sin eso no hay de dónde sacar el mensaje.' });
      return;
    }

    const totalCargadas = reviewsRaw.length;
    const muestreadas = muestrearReviews(reviewsRaw, MAX_REVIEWS);
    const reviews = muestreadas.map((r, i) => ({ ...r, id: i }));
    const reviewsPorId = {}; reviews.forEach(r => { reviewsPorId[r.id] = r; });

    const plataformasConDatos = [...new Set(reviews.map(r => r.plataforma))];
    const plataformasFaltantes = Object.values(TIPO_A_PLATAFORMA).filter(p => !plataformasConDatos.includes(p));

    // --- Documentos de contexto (texto) ---
    async function textoDe(files) {
      let acumulado = '';
      for (const f of files) {
        const p = await parseArchivo(f);
        if (p.kind === 'texto' && p.texto) acumulado += `\n--- ${f.nombre} ---\n${p.texto}`;
        else if (p.kind === 'tabular') acumulado += `\n--- ${f.nombre} (planilla) ---\n${JSON.stringify(p.filas.slice(0, 50))}`;
      }
      return acumulado.slice(0, MAX_TEXTO_LARGO);
    }
    const productoTexto = await textoDe(archivos.filter(a => a.tipo === 'documentosProducto'));
    const briefTexto = await textoDe(archivos.filter(a => a.tipo === 'briefCampana'));
    const specsTexto = await textoDe(archivos.filter(a => a.tipo === 'specsRetailer'));
    const perfilTexto = await textoDe(archivos.filter(a => a.tipo === 'perfilConsumidor'));

    const retailers = parseRetailers(retailersObjetivo);

    const prompt = construirPrompt({ cliente, producto, objetivo: objetivoCampana, retailers, briefTexto, specsTexto, perfilTexto, productoTexto, creativoBase }, reviews);
    const resultado = await llamarClaude(prompt, apiKey);

    // --- Validar temas y calcular conteos/porcentajes de forma determinística ---
    const temasValidos = (Array.isArray(resultado.temas_detectados) ? resultado.temas_detectados : [])
      .slice(0, MAX_TEMAS)
      .filter(t => t && t.id && t.nombre);
    const temaIds = new Set(temasValidos.map(t => t.id));

    const conteoPorTema = {};
    temaIds.forEach(id => { conteoPorTema[id] = []; });

    const clasifValida = Array.isArray(resultado.clasificacion) ? resultado.clasificacion : [];
    for (const c of clasifValida) {
      const rid = c.review_id;
      if (!(rid in reviewsPorId)) continue;
      const temas = Array.isArray(c.temas) ? c.temas : [];
      for (const tid of temas) {
        if (temaIds.has(tid)) conteoPorTema[tid].push(rid);
      }
    }

    const totalAnalizadas = reviews.length;
    const temasFinal = temasValidos.map(t => {
      const ids = [...new Set(conteoPorTema[t.id] || [])];
      const pct = totalAnalizadas ? Math.round((ids.length / totalAnalizadas) * 1000) / 10 : 0;
      const citaId = ids.length ? ids[0] : null;
      const cita = citaId !== null ? reviewsPorId[citaId] : null;
      return {
        id: t.id,
        nombre: t.nombre,
        sentimiento: ['positivo', 'negativo', 'neutro'].includes(t.sentimiento) ? t.sentimiento : 'neutro',
        menciones: ids.length,
        pct,
        citaTexto: cita ? cita.texto : '',
        citaPlataforma: cita ? cita.plataforma : ''
      };
    }).sort((a, b) => b.pct - a.pct);

    const temasPorId = {}; temasFinal.forEach(t => { temasPorId[t.id] = t; });

    // --- Copy final, con el % real inyectado desde JS (nunca el que "recuerde" Claude) ---
    const copyRaw = resultado.copy || {};
    const copyFinal = {};
    for (const r of retailers) {
      const bloque = copyRaw[r.slug];
      if (!bloque) continue;
      const titulos = Array.isArray(bloque.titulos_banner) ? bloque.titulos_banner.slice(0, 2) : [];
      copyFinal[r.slug] = {
        nombre: r.nombre,
        titulos_banner: titulos.map(tb => {
          const tema = temasPorId[tb.tema_id];
          return {
            texto: tb.texto || '',
            razon: tb.razon || '',
            tema_nombre: tema ? tema.nombre : null,
            tema_pct: tema ? tema.pct : null,
            tema_sentimiento: tema ? tema.sentimiento : null
          };
        }),
        brand_store_desc: bloque.brand_store_desc || '',
        titulo_imagen: bloque.titulo_imagen || ''
      };
    }

    const cobertura = calcularCobertura(reviews, totalCargadas, plataformasConDatos);
    const comoMejorar = sugerenciaCobertura({ n: reviews.length, totalCargadas, plataformasFaltantes });

    res.status(200).json({
      resumen_ejecutivo: resultado.resumen_ejecutivo || '',
      cobertura: { ...cobertura, plataformasConDatos, plataformasFaltantes, comoMejorar },
      temas: temasFinal,
      copy: copyFinal,
      retailers: retailers.map(r => ({ slug: r.slug, nombre: r.nombre })),
      fuentes: fuentesResumen,
      totalReviewsAnalizadas: totalAnalizadas,
      totalReviewsCargadas: totalCargadas
    });
  } catch (err) {
    console.error('Error en /api/copywriter:', err);
    res.status(500).json({ error: 'Error inesperado generando el copy: ' + err.message });
  }
};
