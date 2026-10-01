const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const NAVY = "0B2545";
const RED = "D62828";
const SKY = "E8EEF6";
const INK = "1F2933";
const MUTED = "5B6770";
const WHITE = "FFFFFF";
const HF = "Cambria";
const BF = "Calibri";

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function iconCircle(slide, img, x, y, d, fill) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
  const p = d * 0.26;
  slide.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}

function footer(slide, n, dark) {
  slide.addText(`CorreosChile · Trabajo híbrido   |   ${n} / 6`, {
    x: 0.5, y: 5.2, w: 5, h: 0.3, fontFace: BF, fontSize: 9,
    color: dark ? "9FB3C8" : MUTED, margin: 0, isTextBox: true,
  });
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "CorreosChile ante el trabajo híbrido";
  pres.author = "Álvaro Quintero y Brando Barreto";

  const I = {
    env: await icon(fa.FaEnvelopeOpenText, WHITE),
    build: await icon(fa.FaBuilding, WHITE),
    cogs: await icon(fa.FaCogs, WHITE),
    users: await icon(fa.FaUsers, WHITE),
    pin: await icon(fa.FaMapMarkedAlt, WHITE),
    chair: await icon(fa.FaChair, WHITE),
    sign: await icon(fa.FaFileSignature, WHITE),
    cal: await icon(fa.FaCalendarCheck, WHITE),
    ruler: await icon(fa.FaRulerCombined, WHITE),
    flag: await icon(fa.FaFlagCheckered, WHITE),
  };

  // ---------- 1. Portada / encuadre ----------
  let s = pres.addSlide();
  s.background = { color: NAVY };
  iconCircle(s, I.env, 0.5, 0.55, 0.75, RED);
  s.addText("Lunes: 100% de escritorios ocupados.\nViernes: ¿cuántos quedan vacíos?", {
    x: 0.5, y: 1.45, w: 6.2, h: 1.5, fontFace: HF, fontSize: 24, bold: true, color: WHITE,
    margin: 0, valign: "top", isTextBox: true,
  });
  s.addText("CorreosChile ante el trabajo híbrido: espacios y procesos administrativos", {
    x: 0.5, y: 3.25, w: 6.0, h: 0.7, fontFace: BF, fontSize: 16, color: "C9D6E3",
    margin: 0, valign: "top", isTextBox: true,
  });
  s.addText("Álvaro Quintero · Brando Barreto  |  Carrera de Finanzas", {
    x: 0.5, y: 4.55, w: 6, h: 0.35, fontFace: BF, fontSize: 12, color: WHITE, margin: 0, isTextBox: true,
  });
  // Ficha de la empresa
  const facts = [
    ["1981", "empresa del Estado de Chile"],
    ["≈ 4.700", "trabajadores (dic. 2024)"],
    ["≈ 200", "sucursales, de Arica a la Antártica"],
  ];
  facts.forEach(([big, small], i) => {
    const y = 1.35 + i * 1.12;
    s.addText(big, { x: 7.0, y, w: 2.5, h: 0.55, fontFace: HF, fontSize: 28, bold: true, color: WHITE, margin: 0, isTextBox: true });
    s.addText(small, { x: 7.0, y: y + 0.52, w: 2.6, h: 0.35, fontFace: BF, fontSize: 11, color: "C9D6E3", margin: 0, isTextBox: true });
  });
  footer(s, 1, true);
  s.addNotes(
    "(≈30 s) Imaginen la casa matriz de CorreosChile un lunes: todos los puestos ocupados. Ahora el viernes: muchos vacíos, pero igual pagamos luz, arriendo y aseo por ellos. " +
    "CorreosChile es una empresa del Estado creada en 1981, con cerca de 4.700 trabajadores y unas 200 sucursales desde Arica hasta la Antártica. " +
    "Hoy les traemos un análisis breve: cómo el trabajo híbrido cambia la forma en que usamos nuestros espacios y procesos administrativos."
  );

  // ---------- 2. Desafío ----------
  s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("El desafío: oficinas pensadas para 5 días, equipos que van 3 o 4", {
    x: 0.5, y: 0.4, w: 9, h: 0.9, fontFace: HF, fontSize: 26, bold: true, color: NAVY, margin: 0, valign: "top", isTextBox: true,
  });
  const cards = [
    [I.build, "Espacios", "Puestos fijos que quedan vacíos = costo fijo sin uso"],
    [I.cogs, "Procesos", "Firmas, visados y archivo aún dependen del papel"],
    [I.users, "Personas", "Coordinar equipos entre casa matriz y sucursales"],
  ];
  cards.forEach(([img, t, d], i) => {
    const x = 0.5 + i * 3.1;
    s.addShape("roundRect", { x, y: 1.75, w: 2.8, h: 2.55, fill: { color: SKY }, line: { color: SKY }, rectRadius: 0.12 });
    iconCircle(s, img, x + 0.3, 2.0, 0.7, i === 0 ? RED : NAVY);
    s.addText(t, { x: x + 0.3, y: 2.85, w: 2.3, h: 0.45, fontFace: HF, fontSize: 18, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText(d, { x: x + 0.3, y: 3.3, w: 2.3, h: 0.85, fontFace: BF, fontSize: 13, color: INK, margin: 0, valign: "top", isTextBox: true });
  });
  s.addText("Alcance: áreas administrativas. La atención en sucursales y el reparto siguen siendo presenciales.", {
    x: 0.5, y: 4.55, w: 9, h: 0.35, fontFace: BF, fontSize: 11, italic: true, color: MUTED, margin: 0, isTextBox: true,
  });
  footer(s, 2, false);
  s.addNotes(
    "(≈35 s) El desafío central es un desajuste: nuestras oficinas administrativas están diseñadas para que todos vayan los 5 días, pero el mercado ya se movió a 3 o 4 días presenciales. " +
    "Eso impacta en tres frentes: espacios que pagamos y no usamos, procesos que todavía exigen papel y firma en persona, y la coordinación de equipos repartidos entre casa matriz y sucursales. " +
    "Ojo: hablamos de las áreas administrativas; la atención al público y el reparto son, por naturaleza, presenciales."
  );

  // ---------- 3. Gráfico ----------
  s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("En Chile, el híbrido ya es la norma", {
    x: 0.5, y: 0.4, w: 9, h: 0.6, fontFace: HF, fontSize: 28, bold: true, color: NAVY, margin: 0, isTextBox: true,
  });
  s.addChart(pres.charts.BAR, [{
    name: "% de empresas",
    labels: ["Híbrido con días fijos", "Híbrido flexible (elige el trabajador)", "100% presencial", "100% remoto"],
    values: [49, 38, 13, 0],
  }], {
    x: 0.4, y: 1.15, w: 5.9, h: 3.75, barDir: "bar",
    chartColors: [NAVY], catAxisOrientation: "maxMin",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"',
    dataLabelColor: INK, dataLabelFontSize: 12, dataLabelFontFace: BF,
    catAxisLabelColor: INK, catAxisLabelFontSize: 11, catAxisLabelFontFace: BF,
    valAxisHidden: true, valAxisMaxVal: 60, valAxisMinVal: 0,
    valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showLegend: false, barGapWidthPct: 60,
    showTitle: true, title: "Modelo de trabajo de las empresas en Chile, 2025",
    titleFontSize: 12, titleColor: MUTED, titleFontFace: BF,
  });
  s.addText("87%", { x: 6.7, y: 1.3, w: 2.8, h: 0.95, fontFace: HF, fontSize: 60, bold: true, color: RED, margin: 0, isTextBox: true });
  s.addText("de las empresas trabaja con algún modelo híbrido", {
    x: 6.7, y: 2.25, w: 2.8, h: 0.6, fontFace: BF, fontSize: 14, color: INK, margin: 0, valign: "top", isTextBox: true,
  });
  s.addText("32%", { x: 6.7, y: 3.05, w: 2.8, h: 0.6, fontFace: HF, fontSize: 32, bold: true, color: NAVY, margin: 0, isTextBox: true });
  s.addText("de las híbridas exige 4 días presenciales", {
    x: 6.7, y: 3.65, w: 2.8, h: 0.55, fontFace: BF, fontSize: 13, color: INK, margin: 0, valign: "top", isTextBox: true,
  });
  s.addText("Fuente: Robert Half Chile, encuesta de modelos de trabajo 2025 (citada por Emol y El Mostrador, 09-05-2025).", {
    x: 0.5, y: 4.95, w: 9, h: 0.25, fontFace: BF, fontSize: 9, color: MUTED, margin: 0, isTextBox: true,
  });
  s.addText("3 / 6", { x: 8.5, y: 5.22, w: 1, h: 0.25, fontFace: BF, fontSize: 9, color: MUTED, align: "right", margin: 0, isTextBox: true });
  s.addNotes(
    "(≈40 s) Este es el dato real que investigamos. Según la encuesta 2025 de Robert Half Chile, el 49% de las empresas usa un modelo híbrido con días fijos en la oficina, " +
    "un 38% deja que el trabajador elija cuántos días va, solo un 13% volvió a la presencialidad total y ninguna opera 100% remoto. " +
    "En suma, 87% de las empresas ya es híbrida, y la fórmula más común es 4 días presenciales. Para nosotros esto significa que, aunque CorreosChile adopte un híbrido moderado, " +
    "cada día habrá puestos sin usar: es un costo que podemos gestionar."
  );

  // ---------- 4. Organizador gráfico (esquema de llaves) ----------
  s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("Qué cambia con el híbrido", {
    x: 0.5, y: 0.4, w: 9, h: 0.6, fontFace: HF, fontSize: 28, bold: true, color: NAVY, margin: 0, isTextBox: true,
  });
  // Nodo central
  s.addShape("roundRect", { x: 0.5, y: 2.2, w: 2.0, h: 1.2, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.12 });
  s.addText("Trabajo híbrido en CorreosChile", {
    x: 0.6, y: 2.2, w: 1.8, h: 1.2, fontFace: HF, fontSize: 15, bold: true, color: WHITE, align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  // Llave
  s.addShape("leftBrace", { x: 2.7, y: 1.25, w: 0.35, h: 3.1, line: { color: RED, width: 2.5 } });
  const branches = [
    [I.build, "Espacios", ["Puestos compartidos", "Reserva", "Menos m²"]],
    [I.cogs, "Procesos", ["Firma electrónica", "Flujos digitales", "Nube"]],
    [I.users, "Personas", ["Días ancla", "Metas claras", "Desconexión"]],
  ];
  branches.forEach(([img, t, items], i) => {
    const y = 1.2 + i * 1.12;
    iconCircle(s, img, 3.25, y + 0.12, 0.55, i === 0 ? RED : NAVY);
    s.addText(t, { x: 3.95, y: y + 0.12, w: 1.4, h: 0.55, fontFace: HF, fontSize: 17, bold: true, color: NAVY, valign: "middle", margin: 0, isTextBox: true });
    s.addText(items.join("  ·  "), {
      x: 5.4, y: y + 0.12, w: 4.2, h: 0.55, fontFace: BF, fontSize: 12.5, color: INK, valign: "middle", margin: 0, isTextBox: true,
    });
  });
  footer(s, 4, false);
  s.addNotes(
    "(≈40 s) Ordenamos el análisis en un esquema de llaves. El trabajo híbrido en CorreosChile toca tres ámbitos. " +
    "Espacios: pasar de puestos fijos a puestos compartidos con reserva, lo que permite usar menos metros cuadrados y bajar costos fijos. " +
    "Procesos: firma electrónica, flujos digitales y archivo en la nube, para que un trámite no se detenga porque alguien no está en la oficina. " +
    "Personas: días ancla en que el equipo coincide, metas medidas por resultados y reglas claras de desconexión."
  );

  // ---------- 5. Conclusiones ----------
  s = pres.addSlide();
  s.background = { color: SKY };
  s.addText("Conclusiones", {
    x: 0.5, y: 0.4, w: 9, h: 0.6, fontFace: HF, fontSize: 28, bold: true, color: NAVY, margin: 0, isTextBox: true,
  });
  const concl = [
    ["01", "El híbrido llegó para quedarse", "87% de las empresas en Chile ya lo aplica."],
    ["02", "Un puesto vacío es un costo", "Pagamos m², luz y aseo aunque nadie lo use."],
    ["03", "Sin procesos digitales, no funciona", "El papel obliga a estar en la oficina."],
  ];
  concl.forEach(([n, t, d], i) => {
    const y = 1.3 + i * 1.2;
    s.addShape("roundRect", { x: 0.5, y, w: 7.4, h: 0.95, fill: { color: WHITE }, line: { color: WHITE }, rectRadius: 0.1 });
    s.addText(n, { x: 0.75, y, w: 0.9, h: 0.95, fontFace: HF, fontSize: 28, bold: true, color: RED, valign: "middle", margin: 0, isTextBox: true });
    s.addText(t, { x: 1.7, y: y + 0.13, w: 6.0, h: 0.38, fontFace: HF, fontSize: 17, bold: true, color: NAVY, margin: 0, isTextBox: true });
    s.addText(d, { x: 1.7, y: y + 0.5, w: 6.0, h: 0.32, fontFace: BF, fontSize: 13, color: INK, margin: 0, isTextBox: true });
  });
  footer(s, 5, false);
  s.addNotes(
    "(≈35 s) Tres conclusiones. Uno: el híbrido no es una moda, ya es la norma en Chile. " +
    "Dos: desde la mirada financiera, cada puesto vacío es un costo fijo que no genera valor. " +
    "Tres: el híbrido solo funciona si los procesos administrativos son digitales; si dependen del papel, la gente igual tiene que venir."
  );

  // ---------- 6. Recomendación ----------
  s = pres.addSlide();
  s.background = { color: NAVY };
  s.addText("Recomendación: piloto híbrido de 90 días en casa matriz", {
    x: 0.5, y: 0.4, w: 9, h: 0.9, fontFace: HF, fontSize: 26, bold: true, color: WHITE, margin: 0, valign: "top", isTextBox: true,
  });
  const steps = [
    [I.ruler, "Mes 1", "Medir", "ocupación real de puestos"],
    [I.chair, "Mes 2", "Compartir", "puestos con reserva"],
    [I.sign, "Mes 3", "Digitalizar", "3 trámites con firma electrónica"],
    [I.flag, "Día 90", "Decidir", "con datos: m² y costos"],
  ];
  steps.forEach(([img, when, t, d], i) => {
    const x = 0.5 + i * 2.3;
    iconCircle(s, img, x, 1.65, 0.75, i === 3 ? RED : "1D4E89");
    if (i < 3) s.addShape("line", { x: x + 0.85, y: 2.025, w: 1.35, h: 0, line: { color: "5C7A99", width: 1.5, endArrowType: "triangle" } });
    s.addText(when, { x, y: 2.55, w: 2.0, h: 0.3, fontFace: BF, fontSize: 11, color: "9FB3C8", margin: 0, isTextBox: true });
    s.addText(t, { x, y: 2.85, w: 2.0, h: 0.42, fontFace: HF, fontSize: 18, bold: true, color: WHITE, margin: 0, isTextBox: true });
    s.addText(d, { x, y: 3.27, w: 1.95, h: 0.6, fontFace: BF, fontSize: 12.5, color: "C9D6E3", margin: 0, valign: "top", isTextBox: true });
  });
  s.addText("Menos papel, menos m² vacíos, mismo servicio a Chile.", {
    x: 0.5, y: 4.3, w: 9, h: 0.45, fontFace: HF, fontSize: 18, italic: true, bold: true, color: WHITE, margin: 0, isTextBox: true,
  });
  footer(s, 6, true);
  s.addNotes(
    "(≈40 s) Nuestra recomendación para el área: un piloto de 90 días en casa matriz, sin grandes inversiones. " +
    "Mes 1: medir cuántos puestos se usan realmente cada día. Mes 2: implementar puestos compartidos con un sistema simple de reserva. " +
    "Mes 3: digitalizar tres trámites internos con firma electrónica. Al día 90, decidimos con datos cuántos metros cuadrados podemos liberar y cuánto ahorramos. " +
    "Cierre: menos papel, menos metros cuadrados vacíos, y el mismo servicio a Chile. Muchas gracias."
  );

  await pres.writeFile({ fileName: "/home/user/claude/presentacion/CorreosChile_Trabajo_Hibrido.pptx" });
  console.log("done");
})();
