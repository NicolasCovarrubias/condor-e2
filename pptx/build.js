// Genera "E2 ppt-G68.pptx" editable a partir del diseño del deck HTML (1920x1080 px -> 13,333 x 7,5 in).
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.SKILL_DIR + "/scripts/apply_theme.js");

const HERE = __dirname;
const DECK = "C:/Users/ncova/condor-e2";
const IMG = DECK + "/img/";
const OUT = path.join(HERE, "E2 ppt-G68.pptx");

const THEME = {
  name: "CONDOR Control y Calma",
  headFontFace: "Plus Jakarta Sans",
  bodyFontFace: "Inter",
  colors: {
    dk1: "292524", lt1: "FFFFFF", dk2: "44403C", lt2: "FAFAF9",
    accent1: "059669", accent2: "3B82F6", accent3: "78716C", accent4: "D6D3D1",
    accent5: "E3F3EC", accent6: "E8EFFC", hlink: "059669", folHlink: "3B82F6",
  },
};
const HEX = { ink: "292524", ink2: "44403C", gray: "78716C", line: "E7E5E4", green: "059669", blue: "3B82F6", bar: "D6D3D1", tbd: "2563EB", tbdBg: "E8EFFC" };

const I = px => +(px / 144).toFixed(4);      // px del deck HTML -> pulgadas
const P = (px, min = 8) => Math.max(min, Math.round(px * 0.5 * 2) / 2); // px -> pt
const HEAD = "+mj-lt";                         // fuente de títulos del tema (Plus Jakarta Sans)

// Notas del presentador: se reutilizan las del deck HTML, en orden.
const html = fs.readFileSync(DECK + "/index.html", "utf8");
const NOTES = [...html.matchAll(/<section class="slide[^"]*"[^>]*data-notes="([^"]*)"/g)].map(m => m[1]);

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "CONDOR · Entrega 2";
pres.author = "Grupo 68";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

// ---------- Layouts ----------
const LOGO_DARK = path.join(HERE, "condor_292524.png");
const LOGO_WHITE = path.join(HERE, "condor_FFFFFF.png");

pres.defineSlideMaster({
  title: "PORTADA",
  background: { color: C.background2 },
  objects: [
    { rect: { x: I(1200), y: 0, w: I(720), h: 7.5, fill: { color: C.accent5 } } },
    { image: { x: I(120), y: I(150), w: I(48), h: I(16), path: LOGO_DARK } },
    { text: { text: "CONDOR", options: { x: I(178), y: I(140), w: I(200), h: I(36), fontFace: HEAD, fontSize: 11, bold: true, charSpacing: 3, color: C.text1, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: I(120), y: I(390), w: I(1000), h: I(480), fontSize: 54, bold: true, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "CONTENIDO",
  background: { color: C.background2 },
  objects: [
    { rect: { x: I(120), y: I(44), w: I(58), h: I(58), fill: { color: C.accent1 } } },
    { line: { x: I(120), y: I(126), w: I(1680), h: 0, line: { color: C.text1, width: 1.5 } } },
    { image: { x: I(1622), y: I(66), w: I(42), h: I(14), path: LOGO_DARK } },
    { text: { text: "CONDOR", options: { x: I(1670), y: I(56), w: I(130), h: I(34), fontFace: HEAD, fontSize: 10, bold: true, charSpacing: 3, color: C.text1, align: "right", margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: I(120), y: I(165), w: I(1600), h: I(170), fontSize: 32, bold: true, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: I(120), y: I(44), w: I(58), h: I(58), color: "FFFFFF", fontFace: "Inter", fontSize: 10, bold: true, align: "center", valign: "middle" },
});

// ---------- Helpers ----------
let N = 0;
const SECTIONS = new Set();
function newSlide(master, section, label) {
  if (!SECTIONS.has(section)) { pres.addSection({ title: section }); SECTIONS.add(section); }
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  s.addNotes(NOTES[N++] || "");
  if (label) txt(s, label.toUpperCase(), 198, 44, 600, 58, { fontSize: 8, color: C.text2, charSpacing: 1, valign: "middle", objectName: "Sección" });
  return s;
}
function txt(s, text, x, y, w, h, o = {}) {
  s.addText(text, { x: I(x), y: I(y), w: I(w), h: I(h), isTextBox: true, margin: 0, valign: "top", color: C.text1, fontSize: 10, ...o });
}
function rect(s, x, y, w, h, o = {}) { s.addShape(pres.shapes.RECTANGLE, { x: I(x), y: I(y), w: I(w), h: I(h), line: { type: "none" }, ...o }); }
function hline(s, x, y, w, color = C.text1, width = 1.5) { s.addShape(pres.shapes.LINE, { x: I(x), y: I(y), w: I(w), h: 0, line: { color, width } }); }
function vline(s, x, y, h, color = C.text1, width = 1) { s.addShape(pres.shapes.LINE, { x: I(x), y: I(y), w: 0, h: I(h), line: { color, width } }); }
const sup = n => ({ text: String(n), options: { superscript: true, color: C.accent1 } });
const tbd = t => ({ text: t, options: { color: HEX.tbd, highlight: HEX.tbdBg } });
function foot(s, text) { txt(s, text, 120, 1018, 1680, 30, { fontSize: 7, color: C.accent3, objectName: "Fuente" }); }
function label(s, text, x, y, w, color = C.text2) { txt(s, text.toUpperCase(), x, y, w, 30, { fontSize: 8, bold: true, color, charSpacing: 1 }); }

// =============== 01 PORTADA ===============
{
  const s = newSlide("PORTADA", "Presentación");
  s.addText([
    { text: "Que ninguna", options: { breakLine: true } },
    { text: "suscripción se", options: { breakLine: true } },
    { text: "cobre sin que", options: { breakLine: true } },
    { text: "la veas." },
  ], { placeholder: "title" });
  hline(s, 120, 932, 980, C.accent4, 0.75);
  txt(s, "GRUPO 68", 120, 952, 120, 30, { fontSize: 8, bold: true, charSpacing: 1 });
  txt(s, "Agustín Tupper · Baltazar Herreros · Benjamin Garrido · Florencia Candia · Nicolás Covarrubias", 240, 950, 860, 34, { fontSize: 9, color: C.accent3 });
  // Tarjeta de cobros
  s.addShape(pres.shapes.RECTANGLE, { x: I(1280), y: I(305), w: I(560), h: I(470), fill: { color: C.background1 }, line: { type: "none" }, shadow: { type: "outer", blur: 12, offset: 4, angle: 90, color: "28326E", opacity: 0.14 }, objectName: "Tarjeta de cobros" });
  txt(s, "COBROS DE OCTUBRE", 1308, 322, 300, 30, { fontSize: 8, bold: true, charSpacing: 1, valign: "middle" });
  txt(s, "CONDOR", 1612, 322, 200, 30, { fontSize: 8, bold: true, charSpacing: 1, align: "right", valign: "middle" });
  hline(s, 1280, 366, 560, C.text1, 1.25);
  const rows = [
    ["NETFLIX.COM", "$10.090", "SUSCRIPCIÓN", "k", "mensual"],
    ["FITNESS CLUB", "$29.249", "NUEVA SUSCRIPCIÓN", "c", "1er cobro"],
    ["SPOTIFY", "$6.490", "SUBIÓ 22,7%", "s", "antes $5.290"],
    ["UBER *TRIP", "$4.380", "COMPRA PUNTUAL", "g", ""],
  ];
  rows.forEach(([m, a, tag, kind, d], i) => {
    const y0 = 366 + i * 102;
    txt(s, m, 1308, y0 + 22, 300, 30, { fontSize: 11, bold: true, valign: "middle" });
    txt(s, a, 1600, y0 + 22, 212, 30, { fontSize: 11, align: "right", valign: "middle" });
    const tw = Math.round(tag.length * 9.6 + 22);
    if (kind === "g") txt(s, tag, 1308, y0 + 60, 260, 24, { fontSize: 7, bold: true, color: C.accent3, charSpacing: 1, valign: "middle" });
    else {
      const fill = kind === "k" ? "F0EFED" : kind === "c" ? HEX.green : HEX.blue;
      const col = kind === "k" ? C.text2 : C.background1;
      s.addText(tag, { x: I(1308), y: I(y0 + 58), w: I(tw), h: I(26), shape: pres.shapes.RECTANGLE, fill: { color: fill }, color: col, fontSize: 7, bold: true, charSpacing: 1, align: "center", valign: "middle", margin: 0 });
    }
    txt(s, d, 1600, y0 + 58, 212, 26, { fontSize: 8, color: C.accent3, align: "right", valign: "middle" });
    if (i < 3) hline(s, 1280, y0 + 102, 560, C.accent4, 0.75);
  });
}

// =============== 02 PROBLEMA ===============
{
  const s = newSlide("CONTENIDO", "Presentación", "Problema");
  s.addText("Te enteras del cobro cuando la plata ya salió.", { placeholder: "title" });
  hline(s, 120, 282, 1680, C.accent4, 0.75);
  vline(s, 960, 282, 700, C.accent4, 0.75);
  const col = (x, lab, who, num, numColor, cap, ref) => {
    label(s, lab, x, 330, 400);
    txt(s, who, x, 355, 780, 50, { fontFace: HEAD, fontSize: 18, bold: true });
    txt(s, num, x, 415, 800, 130, { fontFace: HEAD, fontSize: 58, bold: true, color: numColor, valign: "middle" });
    txt(s, [{ text: cap }, sup(ref)], x, 555, 780, 60, { fontSize: 10, color: C.text2 });
  };
  col(120, "Lo sufre", "El titular de la cuenta", "~37 mil", C.text1, "reclamos al año por cobros indebidos en el mercado financiero", 1);
  col(1021, "Lo paga", "El banco emisor", "$72.000 MM", C.accent2, "perdidos por fraude externo en un solo mes, marzo de 2024", 3);
  s.addText([
    { text: "LEY 20.009", options: { fontSize: 7, bold: true, charSpacing: 1, breakLine: true } },
    { text: "lo que el cliente no reconoce, el banco lo devuelve", options: { fontSize: 10 } },
  ], { x: I(750), y: I(832), w: I(420), h: I(118), shape: pres.shapes.RECTANGLE, fill: { color: C.accent1 }, color: C.background1, align: "center", valign: "middle", margin: 8, paraSpaceAfter: 4, objectName: "Puente legal" });
  foot(s, "[1] SERNAC, Ranking Financiero 2022, estimación anual del equipo · [3] CMF, Nota Técnica Ley de Fraudes, oct-2025");
}

// =============== 03 VIAJE ===============
{
  const s = newSlide("CONTENIDO", "Presentación", "Viaje");
  s.addText("Lo que cambiamos en el camino.", { placeholder: "title" });
  hline(s, 120, 408, 1680, C.text1, 1.5);
  const rows = [
    ["01", "Detener cualquier cobro", "Ver y avisar primero", [{ text: "Feedback del equipo", options: { breakLine: true } }, { text: "+ Visa Stop Payment Service" }, sup(13)]],
    ["02", "Que pague el usuario", "Que pague el banco", [tbd("entrevistas"), { text: " + Ley 20.009" }]],
    ["03", "Reconocer por el nombre", "Reconocer por la regularidad", [{ text: "Simulación", options: { breakLine: true } }, { text: "11.384 cobros" }]],
    ["04", "Que el sistema decida solo", "Que pregunte cuando duda", "Gemelo digital"],
  ];
  rows.forEach(([n, was, now, ev], i) => {
    const y0 = 408 + i * 104;
    txt(s, n, 120, y0, 60, 104, { fontSize: 8, color: C.accent3, valign: "middle" });
    txt(s, was, 190, y0, 560, 104, { fontFace: HEAD, fontSize: 17, bold: true, color: C.accent3, strike: "sngStrike", valign: "middle" });
    s.addText("→", { x: I(767), y: I(y0 + 32), w: I(40), h: I(40), shape: pres.shapes.RECTANGLE, fill: { color: C.accent1 }, color: C.background1, fontSize: 11, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, now, 837, y0, 680, 104, { fontFace: HEAD, fontSize: 19, bold: true, valign: "middle" });
    txt(s, ev, 1500, y0, 300, 104, { fontSize: 8, color: C.accent3, align: "right", valign: "middle" });
    hline(s, 120, y0 + 104, 1680, C.accent4, 0.75);
  });
  foot(s, "Filas 03 y 04: documentadas en el filtro y el dataset del prototipo.");
}

// =============== 04 SOLUCIÓN ===============
{
  const s = newSlide("CONTENIDO", "Presentación", "Solución");
  s.addText("Todas tus suscripciones a la vista, desde el primer cobro.", { placeholder: "title" });
  hline(s, 196, 560, 1096, C.text1, 1.25);
  const nodes = [
    [120, "01", HEX.ink, "Llega un cobro", "comercio · categoría (MCC) · monto · señales de la red"],
    [511, "02", HEX.green, "¿Volverá a cobrar?", "cómo cobra ese comercio a todos los clientes del banco"],
    [901, "03", HEX.green, "¿Qué tipo es?", "suscripción, cuenta o incierto"],
  ];
  nodes.forEach(([x, n, fill, h, p]) => {
    s.addText(n, { x: I(x), y: I(522), w: I(76), h: I(76), shape: pres.shapes.RECTANGLE, fill: { color: fill }, color: C.background1, fontSize: 12, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, h, x, 620, 370, 50, { fontFace: HEAD, fontSize: 17, bold: true });
    txt(s, p, x, 675, 320, 70, { fontSize: 9.5, color: C.accent3 });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: I(1292), y: I(478), w: I(508), h: I(294), fill: { color: C.background1 }, line: { color: C.text1, width: 1.25 }, objectName: "Resultados" });
  const outs = [["+", HEX.green, "FFFFFF", "La suma a tu lista", "confianza alta"], ["?", "F0EFED", HEX.ink, "Te pregunta", "confianza intermedia"], ["↑", HEX.blue, "FFFFFF", "Te avisa si sube", "cambio de precio"]];
  outs.forEach(([ic, fill, col, b, sub], i) => {
    const y0 = 478 + i * 98;
    s.addText(ic, { x: I(1318), y: I(y0 + 21), w: I(56), h: I(56), shape: pres.shapes.RECTANGLE, fill: { color: fill }, color: col, fontSize: 12, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, b, 1392, y0 + 22, 390, 34, { fontFace: HEAD, fontSize: 13, bold: true });
    txt(s, sub, 1392, y0 + 56, 390, 24, { fontSize: 8, color: C.accent3 });
    if (i < 2) hline(s, 1292, y0 + 98, 508, C.accent4, 0.75);
  });
  hline(s, 120, 926, 1680, C.accent4, 0.75);
  label(s, "Siguiente paso", 120, 955, 170, C.accent1);
  txt(s, [{ text: "Detener el cobro antes de que ocurra: los bancos ya pueden bloquear cobros recurrentes con Visa" }, sup(13), { text: ". Falta validarlo con un emisor." }], 293, 950, 1500, 40, { fontSize: 10.5, color: C.text2 });
  foot(s, "Lógica y umbrales: código del prototipo CONDOR.");
}

// =============== 05 ESTADO DEL ARTE ===============
{
  const s = newSlide("CONTENIDO", "Presentación", "Estado del arte");
  s.addText("Ver suscripciones ya existe. Lo nuevo es cómo se reconocen.", { placeholder: "title" });
  hline(s, 120, 370, 1680, C.text1, 1.5);
  const COLS = [120, 540, 740, 930];
  ["Alternativa", "Dónde", "Dentro del banco", "Cómo reconoce la suscripción"].forEach((h, i) => txt(s, h.toUpperCase(), COLS[i], 372, 400, 46, { fontSize: 7, bold: true, color: C.text2, charSpacing: 1, valign: "middle" }));
  hline(s, 120, 418, 1680, C.accent4, 0.75);
  const rows5 = [
    ["BancoEstado", "Mis Suscripciones", "Chile, hoy", true, "Solo comercios con tarjeta tokenizada", 12],
    ["Nubank", "Gestor de suscripciones", "Brasil, 2025", true, "Cobros recurrentes de la tarjeta", 7],
    ["Visa", "Subscription Manager", "2026 → LatAm", true, "Gestión en 100+ comercios asociados", 6],
    ["Rocket Money", "app aparte", "EE.UU.", false, "Historial de cuentas vinculadas", 8],
  ];
  rows5.forEach(([n, sub, where, ok, how, ref], i) => {
    const y0 = 418 + i * 100;
    txt(s, [{ text: n, options: { fontFace: HEAD, fontSize: 12.5, bold: true, breakLine: true } }, { text: sub, options: { fontSize: 7.5, color: C.accent3 } }], COLS[0], y0, 400, 100, { valign: "middle" });
    txt(s, where, COLS[1], y0, 190, 100, { fontSize: 9, color: C.text2, valign: "middle" });
    s.addText(ok ? "✓" : "✕", { x: I(COLS[2]), y: I(y0 + 35), w: I(30), h: I(30), shape: pres.shapes.RECTANGLE, fill: { color: ok ? HEX.ink : "FFFFFF" }, line: { color: ok ? HEX.ink : HEX.bar, width: 1.25 }, color: ok ? C.background1 : C.accent3, fontSize: 9, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, [{ text: how }, sup(ref)], COLS[3], y0, 870, 100, { fontSize: 10.5, color: C.text2, valign: "middle" });
    hline(s, 120, y0 + 100, 1680, C.accent4, 0.75);
  });
  rect(s, 96, 818, 1728, 84, { fill: { color: C.accent1 }, objectName: "Fila CONDOR" });
  txt(s, "CONDOR", COLS[0], 818, 400, 84, { fontFace: HEAD, fontSize: 12.5, bold: true, color: C.background1, valign: "middle" });
  txt(s, "Prototipo", COLS[1], 818, 190, 84, { fontSize: 9, color: C.background1, valign: "middle" });
  s.addText("✓", { x: I(COLS[2]), y: I(845), w: I(30), h: I(30), shape: pres.shapes.RECTANGLE, fill: { color: "FFFFFF" }, color: C.accent1, fontSize: 9, bold: true, align: "center", valign: "middle", margin: 0 });
  txt(s, "Cómo cobra el comercio a todo el banco, desde el 1er cobro", COLS[3], 818, 870, 84, { fontSize: 10.5, color: C.background1, valign: "middle" });
  label(s, "Diferencia de fondo", 120, 922, 600);
  txt(s, "No depende de que el comercio se declare: aprende de cómo cobra.", 120, 950, 860, 40, { fontFace: HEAD, fontSize: 11, bold: true });
  label(s, "Difícil de copiar", 990, 922, 600);
  txt(s, "El modelo de comercios mejora con cada cliente del banco.", 990, 950, 810, 40, { fontFace: HEAD, fontSize: 11, bold: true });
  foot(s, "[6] Visa, mar-2026 · [7] Nubank, dic-2025 · [8] rocketmoney.com · [12] BancoEstado, Mis Suscripciones");
}

// =============== 06 PROTOTIPO ===============
{
  const s = newSlide("CONTENIDO", "Presentación", "Prototipo");
  s.addText("Así funciona el prototipo.", { placeholder: "title" });
  const step = (x, y, n, t) => {
    txt(s, n, x, y - 10, 40, 56, { fontFace: HEAD, fontSize: 24, bold: true, color: C.accent1, valign: "bottom" });
    txt(s, t, x + 38, y, 700, 44, { fontFace: HEAD, fontSize: 13, bold: true, valign: "bottom" });
  };
  const frame = (x, y, w, h) => rect(s, x, y, w, h, { fill: { color: "F5F7FA" }, line: { color: C.text1, width: 1.1 } });
  step(120, 282, "1", "Llega el primer cobro y lo reconoce");
  step(982, 282, "2", "Si duda, te pregunta");
  step(120, 646, "3", "Todas tus suscripciones a la vista");
  step(982, 646, "4", "Te avisa si sube el precio");
  txt(s, "en el prototipo, el cobro espera tu decisión", 1400, 660, 400, 30, { fontSize: 8, color: C.accent3, align: "right", valign: "bottom" });
  // 1: cobro + detección
  frame(120, 336, 818, 279);
  s.addImage({ path: IMG + "p1a_cobro.png", x: I(128), y: I(372), w: I(802), h: I(121.5), altText: "Primer cobro de FITNESS CLUB por $29.249, aún sin clasificar" });
  txt(s, "↓ CONDOR", 128, 499, 802, 22, { fontSize: 7, bold: true, color: C.accent1, align: "center", charSpacing: 1, valign: "middle" });
  s.addImage({ path: IMG + "p1b_detectada.png", x: I(128), y: I(527), w: I(802), h: I(48.4), altText: "Detectamos una nueva suscripción: FITNESS CLUB" });
  // 2: pregunta
  frame(982, 336, 818, 279);
  s.addImage({ path: IMG + "p2_pregunta.png", x: I(1043), y: I(344), w: I(696), h: I(263), altText: "Posible suscripción detectada, con opciones Sí, No estoy seguro y No" });
  // 3: lista (recorte superior)
  frame(120, 700, 818, 280);
  s.addImage({ path: IMG + "p3b_lista.png", x: I(121), y: I(701), w: I(816), h: I(816 * 1120 / 1936), sizing: { type: "crop", x: 0, y: 0, w: I(816), h: I(278) }, altText: "Lista de suscripciones activas: Spotify supervisada, Netflix de confianza" });
  // 4: alza
  frame(982, 700, 818, 280);
  s.addImage({ path: IMG + "p4_alza.png", x: I(1048), y: I(708), w: I(686), h: I(264), altText: "Spotify: cambio de precio detectado de $5.290 a $6.490" });
  txt(s, "Capturas reales del gemelo digital (Streamlit). Real: filtro, lógica y app. Simulado: transacciones y conexión con un banco.", 120, 1018, 1300, 30, { fontSize: 7, color: C.accent3 });
  txt(s, [{ text: "En vivo: " }, tbd("link Streamlit")], 1500, 1018, 300, 30, { fontSize: 7, color: C.accent3, align: "right" });
}

// =============== 07 TESTEO ===============
{
  const s = newSlide("CONTENIDO", "Presentación", "Testeo");
  s.addText("Repetir no basta: lo que separa es la regularidad.", { placeholder: "title" });
  // Izquierda: dudas
  label(s, "Duda 1 · ¿Se reconoce desde el 1er cobro?", 120, 285, 520, C.accent1);
  txt(s, "11/11", 120, 318, 480, 150, { fontFace: HEAD, fontSize: 72, bold: true, valign: "middle" });
  txt(s, "comercios bien clasificados en simulación (11.384 cobros sintéticos)", 120, 478, 480, 60, { fontSize: 10, color: C.text2 });
  hline(s, 120, 572, 480, C.text1, 1.25);
  label(s, "Duda 2 · ¿El titular lo entiende? · encuesta a 8 usuarios", 120, 590, 520, C.accent1);
  txt(s, [{ text: "4,0", options: { fontSize: 25 } }, { text: "/5", options: { fontSize: 12, color: C.accent3 } }], 120, 650, 210, 60, { fontFace: HEAD, bold: true, color: C.accent1, valign: "middle" });
  txt(s, "facilidad para navegar y entender las alertas, tras usar el prototipo", 340, 652, 280, 56, { fontSize: 8, color: C.text2 });
  s.addChart(pres.charts.BAR, [{ name: "Personas", labels: ["1", "2", "3", "4", "5"], values: [0, 0, 3, 2, 3] }], {
    x: I(330), y: I(708), w: I(260), h: I(96), barDir: "col", chartColors: [HEX.green], barGapWidthPct: 40,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 7, dataLabelColor: HEX.gray, dataLabelFontFace: "+mn-lt",
    valAxisHidden: true, valAxisMaxVal: 4, valAxisMinVal: 0, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelFontSize: 7, catAxisLabelColor: HEX.gray, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false, showLegend: false, objectName: "Distribución de notas",
  });
  txt(s, "6 de 8", 120, 810, 210, 56, { fontFace: HEAD, fontSize: 25, bold: true, color: C.accent1, valign: "middle" });
  txt(s, "ya vivieron un cobro sin uso o un alza sin aviso", 340, 812, 280, 50, { fontSize: 8, color: C.text2 });
  txt(s, "Muestra pequeña de contactos del equipo, con instrucciones de uso: es un indicio, no una medición.", 120, 875, 480, 40, { fontSize: 7, color: C.accent3 });
  // Derecha: gráfico de repetición
  label(s, "% de clientes que repite el pago", 700, 285, 600);
  const leg = [[1440, HEX.green, "Suscripción"], [1580, HEX.ink, "Cuenta"], [1680, HEX.bar, "No recurrente"]];
  leg.forEach(([x, c, t]) => { rect(s, x, 293, 12, 12, { fill: { color: c } }); txt(s, t, x + 18, 286, 120, 26, { fontSize: 7.5, color: C.accent3, valign: "middle" }); });
  const cats = ["Seguro Auto", "Electricidad", "Spotify", "Netflix", "Agua", "Adobe", "Fitness Club", "Jumbo", "Uber", "Starbucks", "Amazon"];
  const vals = [98.6, 97.9, 95.8, 95.5, 93.8, 92.6, 89.2, 86.2, 84.2, 74.0, 53.6];
  const kind = ["b", "b", "s", "s", "b", "s", "s", "n", "n", "n", "n"];
  s.addChart(pres.charts.BAR, [{ name: "% de clientes que repite", labels: cats, values: vals }], {
    x: I(690), y: I(318), w: I(1125), h: I(420), barDir: "bar", catAxisOrientation: "maxMin",
    chartColors: kind.map(k => (k === "s" ? HEX.green : k === "b" ? HEX.ink : HEX.bar)), barGapWidthPct: 45,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0\"%\"", dataLabelFontSize: 8, dataLabelColor: HEX.ink2, dataLabelFontFace: "+mn-lt",
    valAxisHidden: true, valAxisMaxVal: 110, valAxisMinVal: 0, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    catAxisLabelFontSize: 9, catAxisLabelColor: HEX.ink2, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false, showLegend: false,
    layout: { x: 0.14, y: 0.0, w: 0.86, h: 1.0 }, objectName: "Repetición por comercio",
  });
  // Recuadro de los tres casos cercanos (filas 7 a 9 de 11)
  const rowH = 420 / 11;
  s.addShape(pres.shapes.RECTANGLE, { x: I(690), y: I(318 + 6 * rowH - 2), w: I(1125), h: I(3 * rowH + 4), fill: { type: "none" }, line: { color: C.accent2, width: 1.5 }, objectName: "Casos cercanos" });
  s.addText("CASI IGUAL, DISTINTO TIPO", { x: I(1570), y: I(318 + 9 * rowH + 4), w: I(245), h: I(26), shape: pres.shapes.RECTANGLE, fill: { color: C.accent2 }, color: C.background1, fontSize: 7, bold: true, charSpacing: 0.5, align: "center", valign: "middle", margin: 0 });
  txt(s, [{ text: "Simulación: dataset sintético del prototipo (2.876 clientes, 11 comercios). Encuesta: informe E2, sección 4.1" }, sup(14)], 120, 1018, 1680, 30, { fontSize: 7, color: C.accent3 });
}

// =============== 08 OPORTUNIDAD ===============
{
  const s = newSlide("CONTENIDO", "Presentación", "Oportunidad");
  s.addText("Por qué ahora y quién paga.", { placeholder: "title" });
  hline(s, 120, 458, 1680, C.text1, 1.5);
  [[120, "2020", "El banco debe devolver todo cargo no reconocido"], [680, "2024", "La reforma premia prevenir antes que devolver"], [1240, "2025–26", "BancoEstado, Nubank y Visa ya muestran suscripciones"]].forEach(([x, y, t]) => {
    s.addText(y, { x: I(x), y: I(486), w: I(y.length * 13 + 22), h: I(30), shape: pres.shapes.RECTANGLE, fill: { color: C.accent1 }, color: C.background1, fontSize: 8, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, t, x, 534, 480, 80, { fontFace: HEAD, fontSize: 15, bold: true });
  });
  const stat = (x, w, num, color, cap) => {
    txt(s, num, x, 800, w, 120, { fontFace: HEAD, fontSize: 50, bold: true, color, valign: "bottom" });
    txt(s, cap, x, 928, w, 60, { fontSize: 10, color: C.text2 });
  };
  stat(120, 680, "$11.000 MM", C.accent2, [{ text: "al mes sigue perdiendo la banca por fraude externo, aun después de la reforma (dic-2024)" }, sup(3)]);
  stat(850, 460, "17 + 10", C.text1, [{ text: "bancos y emisores de retail financiero " }, tbd("fuente")]);
  stat(1360, 440, "SaaS", C.accent1, "licencia por cuenta protegida, pagada por Riesgo o Canales Digitales");
  foot(s, "[3] CMF, Nota Técnica Ley de Fraudes (oct-2025) · Leyes 21.234 (2020) y 21.673 (2024) · [6] Visa · [7] Nubank · [12] BancoEstado");
}

// =============== 09 RUTA AL MVP ===============
{
  const s = newSlide("CONTENIDO", "Presentación", "Ruta al MVP");
  s.addText("Del gemelo digital a un piloto.", { placeholder: "title" });
  hline(s, 120, 380, 1680, C.text1, 1.5);
  const steps = [
    ["01", "Datos reales", "calibrar umbrales con históricos de un emisor"],
    ["02", "Conexión al banco", "leer cada autorización en tiempo real"],
    ["03", "Piloto de visibilidad", [{ text: "un emisor digital, " }, tbd("[N] cuentas"), { text: ", 3 meses" }]],
    ["04", "Detener cobros", [{ text: "si el emisor lo permite: Visa ya ofrece bloquearlos" }, sup(13)]],
  ];
  steps.forEach(([n, h, p], i) => {
    const x = 120 + i * 420;
    s.addText(n, { x: I(x), y: I(408), w: I(48), h: I(48), shape: pres.shapes.RECTANGLE, fill: { color: i === 3 ? HEX.green : HEX.ink }, color: C.background1, fontSize: 9, bold: true, align: "center", valign: "middle", margin: 0 });
    txt(s, h, x, 474, 390, 50, { fontFace: HEAD, fontSize: 17, bold: true });
    txt(s, p, x, 528, 340, 70, { fontSize: 9.5, color: C.accent3 });
  });
  label(s, "Pendientes que declaramos", 120, 738, 600, "2563EB");
  ["Qué % de los reclamos de un banco viene de cobros recurrentes", "Al menos 5 pruebas de usuario con el gemelo digital", "Que el banco lo compraría a un tercero, y a qué precio"].forEach((t, i) => {
    const y = 778 + i * 40;
    s.addShape(pres.shapes.RECTANGLE, { x: I(120), y: I(y + 4), w: I(16), h: I(16), fill: { type: "none" }, line: { color: C.accent2, width: 1.25 } });
    txt(s, t, 150, y, 840, 28, { fontSize: 10, color: C.text2, valign: "middle" });
  });
  rect(s, 1040, 740, 760, 240, { fill: { color: C.accent1 }, objectName: "Cierre" });
  s.addImage({ path: LOGO_WHITE, x: I(1082), y: I(780), w: I(90), h: I(30) });
  txt(s, "CONDOR: cada suscripción a la vista desde el primer cobro.", 1082, 836, 680, 120, { fontFace: HEAD, fontSize: 22, bold: true, color: C.background1 });
}

// =============== 10 CIERRE ===============
{
  const s = newSlide("CONTENIDO", "Presentación", "Cierre");
  txt(s, "¿Cuántas suscripciones pagaste este mes sin darte cuenta?", 120, 330, 860, 360, { fontFace: HEAD, fontSize: 40, bold: true, valign: "middle" });
  txt(s, "CONDOR quiere que esa respuesta sea siempre cero.", 120, 730, 860, 40, { fontSize: 12, color: C.text2 });
  vline(s, 1040, 290, 576, C.text1, 1);
  const acro = [["C", "Control", " de"], ["O", "Órdenes", " y"], ["N", "Notificación", " de"], ["D", "Débitos", ""], ["O", "Opacos", " y"], ["R", "Recurrentes", ""]];
  acro.forEach(([l, w, c], i) => {
    const y = 296 + i * 96;
    txt(s, l, 1100, y, 110, 96, { fontFace: HEAD, fontSize: 42, bold: true, color: C.accent1, valign: "middle" });
    txt(s, [{ text: w, options: { bold: true, color: C.text1 } }, { text: c, options: { color: C.accent3 } }], 1222, y, 560, 96, { fontSize: 15, valign: "middle" });
  });
}

// =============== ANEXO A · FUENTES ===============
{
  const s = newSlide("CONTENIDO", "Anexos", "Anexo · Fuentes");
  s.addText("Fuentes", { placeholder: "title" });
  const refs = [
    ["1", "SERNAC, Ranking Financiero 2022: 19.370 reclamos por cobros (1er sem.) y 30,6% de 58.690 (2º sem.). Vía Diario Financiero y La Tercera. Total anual estimado por el equipo.", "https://www.df.cl/reclamos-contra-el-mercado-financiero-aumentan-11-y-superan-las-53-mil"],
    ["2", "SERNAC (jun-2025), estudio \"Uso de Plataformas de Streaming en Chile\": gasto promedio $22.809 al mes. Vía The Clinic.", "https://www.theclinic.cl/2025/06/12/estudio-revelo-que-el-gasto-promedio-mensual-de-los-chilenos-en-plataformas-de-streaming-alcanza-los-22-809/"],
    ["3", "Cortés, J. y Ormazábal, F. (oct-2025). Nota Técnica \"Antecedentes sobre la aplicación de la Ley de Fraudes\". CMF.", "https://www.cmfchile.cl/portal/estadisticas/617/articles-100079_doc_pdf.pdf"],
    ["4", "Ley 20.009, modificada por la Ley 21.234 (2020) y la Ley 21.673 (2024). Resumen de cambios en [3].", ""],
    ["5", "The Clinic (abr-2024). SERNAC oficia a Rappi por cobros indebidos de membresía.", "https://www.theclinic.cl/2024/04/03/sernac-rappi-cobros-indebidos-oficio/"],
    ["6", "Visa (26-mar-2026). Visa launches Enhanced Subscription Manager.", "https://corporate.visa.com/en/sites/visa-perspectives/newsroom/visa-enhanced-subscription-manager-launch.html"],
    ["7", "Nubank (dic-2025). Nubank launches subscription manager to help control recurring charges.", "https://nu.com/en/newsroom/consumers/nubank-launches-subscription-manager-to-help-control-recurring-charges"],
    ["8", "Rocket Money, sitio oficial del producto.", "https://www.rocketmoney.com"],
    ["9", "CMF (mar-2026). Calificación anual de bancos de importancia sistémica (6 bancos).", "https://www.cmfchile.cl/portal/prensa/625/w4-article-109087.html"],
    ["10", "Ley 21.521 (2023), Ley Fintec, y Ley 21.719 (2024), de protección de datos personales.", ""],
    ["11", "Código, dataset sintético, evaluación y capturas del prototipo CONDOR (repositorio del equipo, octubre 2026).", ""],
    ["12", "BancoEstado, \"Mis Suscripciones en mi App BancoEstado\": muestra suscripciones de comercios que usan tokenización.", "https://www.bancoestado.cl/content/bancoestado-public/cl/es/home/home/app-bancoestado/mis-suscripciones-app---bancoestado-personas.html"],
    ["13", "Visa, Visa Stop Payment Service: permite a los emisores detener pagos con tarjeta guardada, incluidos los recurrentes, antes de autorizarlos.", "https://developer.visa.com/capabilities/visa-stop-payment-service/overview"],
    ["14", "Encuesta de testeo del prototipo a 8 usuarios, octubre 2026. Informe E2 del grupo 68, sección 4.1.", ""],
  ];
  const col = list => list.flatMap(([n, t, url], i) => {
    const last = i === list.length - 1;
    const runs = [{ text: `[${n}] `, options: { bold: true, color: C.accent1 } }, { text: t, options: { breakLine: !!url || !last } }];
    if (url) runs.push({ text: url.replace(/^https?:\/\/(www\.)?/, "").slice(0, 70) + (url.length > 80 ? "…" : ""), options: { hyperlink: { url }, color: C.accent1, fontSize: 7, breakLine: !last } });
    return runs;
  });
  txt(s, col(refs.slice(0, 7)), 120, 260, 800, 720, { fontSize: 8.5, color: C.text2, paraSpaceAfter: 6 });
  txt(s, col(refs.slice(7)), 990, 260, 810, 720, { fontSize: 8.5, color: C.text2, paraSpaceAfter: 6 });
  foot(s, "Las cifras marcadas como pendientes no tienen fuente verificada todavía.");
}

// =============== ANEXO B · USO DE IA ===============
{
  const s = newSlide("CONTENIDO", "Anexos", "Anexo · Uso de IA");
  s.addText("Declaración de uso de inteligencia artificial", { placeholder: "title" });
  label(s, "Herramienta y propósito", 120, 290, 700, C.accent1);
  txt(s, [
    { text: "Claude (Anthropic), vía Claude Code. ", options: { bold: true, color: C.text1 } },
    { text: "Usada para diseñar y programar esta presentación, proponer su estructura según la pauta y la guía extendida de la E2, buscar y verificar fuentes públicas, y capturar pantallas del prototipo.", options: { breakLine: true } },
    tbd("Agregar otros usos (p. ej., apoyo en el código del prototipo)"),
  ], 120, 330, 760, 400, { fontSize: 11.5, color: C.text2, paraSpaceAfter: 12 });
  label(s, "Cómo influyó y responsabilidad", 990, 290, 700, C.accent1);
  txt(s, "La IA propuso estructura, redacción y diseño. El equipo revisó, corrigió y validó cada contenido y cifra. Las entrevistas, las decisiones del proyecto y el prototipo son trabajo del equipo, que responde por la exactitud de todo el material.", 990, 330, 810, 400, { fontSize: 11.5, color: C.text2 });
  foot(s, "Según la política de uso de IA del programa ING2030 2026-2.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("OK", OUT, "slides:", N, "notas:", NOTES.length);
})().catch(e => { console.error(e); process.exit(1); });
