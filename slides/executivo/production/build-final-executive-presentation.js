const pptxgen = require("pptxgenjs");
const {
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require("./pptxgenjs_helpers/layout");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "OpenAI Codex";
pptx.company = "EEB Araujo Figueiredo";
pptx.subject = "Araujo Innovation Lab - Apresentacao Executiva Final";
pptx.title = "Araujo Innovation Lab - Apresentacao Executiva Final";
pptx.lang = "pt-BR";
pptx.theme = {
  headFontFace: "Georgia",
  bodyFontFace: "Arial",
  lang: "pt-BR",
};

const W = 13.333;
const H = 7.5;

const C = {
  bg: "F7F7F5",
  white: "FFFFFF",
  ink: "1F2C31",
  body: "34434A",
  muted: "6E7A80",
  line: "D7DEDA",
  lineDark: "9AA8A2",
  blue: "1E3A3A",
  blueSoft: "EAF0EE",
  blueSoft2: "F1F5F3",
  gold: "C9A227",
  goldSoft: "F6F1DD",
  greenSoft: "EEF4F0",
  grayBox: "F0F2F1",
};

function addSlideBase(slide, page, sectionLabel) {
  slide.background = { color: C.bg };
  slide.addShape(pptx.ShapeType.line, {
    x: 0.66,
    y: 0.4,
    w: 0.9,
    h: 0,
    line: { color: C.gold, pt: 1.6 },
  });
  if (sectionLabel) {
    slide.addText(sectionLabel, {
      x: 0.7,
      y: 0.15,
      w: 2.6,
      h: 0.22,
      fontFace: "Arial",
      fontSize: 8.5,
      color: C.muted,
      bold: true,
      charSpace: 1.1,
      allCaps: true,
      margin: 0,
    });
  }
  slide.addShape(pptx.ShapeType.line, {
    x: 0.66,
    y: 7.02,
    w: 12.0,
    h: 0,
    line: { color: C.line, pt: 1 },
  });
  slide.addText(`Documento Institucional  |  2026  |  ${page.toString().padStart(2, "0")}`, {
    x: 0.7,
    y: 7.06,
    w: 4.2,
    h: 0.18,
    fontFace: "Arial",
    fontSize: 8,
    color: C.muted,
    margin: 0,
  });
  slide.addText("Araujo Innovation Lab  |  EEB Araujo Figueiredo", {
    x: 8.6,
    y: 7.06,
    w: 4.0,
    h: 0.18,
    fontFace: "Arial",
    fontSize: 8,
    color: C.muted,
    align: "right",
    margin: 0,
  });
}

function addTitle(slide, title, subtitle) {
  slide.addText(title, {
    x: 0.85,
    y: 0.72,
    w: 8.7,
    h: 0.55,
    fontFace: "Georgia",
    fontSize: 24,
    bold: true,
    color: C.ink,
    margin: 0,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.88,
      y: 1.22,
      w: 8.8,
      h: 0.34,
      fontFace: "Arial",
      fontSize: 11.5,
      color: C.muted,
      margin: 0,
    });
  }
}

function addTextBlock(slide, x, y, w, h, text, opts = {}) {
  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontFace: opts.fontFace || "Arial",
    fontSize: opts.fontSize || 11,
    color: opts.color || C.body,
    bold: !!opts.bold,
    italic: !!opts.italic,
    align: opts.align || "left",
    valign: opts.valign || "top",
    margin: opts.margin !== undefined ? opts.margin : 0.05,
    breakLine: opts.breakLine,
    fit: "shrink",
  });
}

function addBulletList(slide, x, y, w, items, opts = {}) {
  const fontSize = opts.fontSize || 11;
  const color = opts.color || C.body;
  const gap = opts.gap || 0.27;
  items.forEach((item, index) => {
    slide.addText([{ text: item, options: { bullet: { indent: 10 } } }], {
      x,
      y: y + index * gap,
      w,
      h: 0.22,
      fontFace: "Arial",
      fontSize,
      color,
      margin: 0,
      breakLine: false,
      fit: "shrink",
    });
  });
}

function addCard(slide, x, y, w, h, title, bodyLines, opts = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.05,
    line: { color: opts.line || C.line, pt: 1 },
    fill: { color: opts.fill || C.white },
  });
  if (opts.badge) {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.12,
      y: y + 0.13,
      w: 0.34,
      h: 0.34,
      line: { color: opts.badgeLine || C.blue, pt: 1.2 },
      fill: { color: opts.badgeFill || C.white },
    });
    slide.addText(opts.badge, {
      x: x + 0.12,
      y: y + 0.19,
      w: 0.34,
      h: 0.12,
      fontFace: "Arial",
      fontSize: 9,
      color: opts.badgeColor || C.blue,
      bold: true,
      align: "center",
      margin: 0,
    });
  }
  slide.addText(title, {
    x: x + (opts.badge ? 0.54 : 0.14),
    y: y + 0.12,
    w: w - (opts.badge ? 0.66 : 0.28),
    h: 0.24,
    fontFace: "Arial",
    fontSize: 12,
    bold: true,
    color: C.ink,
    margin: 0,
    fit: "shrink",
  });
  addBulletList(slide, x + 0.16, y + 0.44, w - 0.28, bodyLines, {
    fontSize: 10.5,
    gap: opts.bulletGap || 0.24,
  });
}

function addQuoteBand(slide, text, opts = {}) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: opts.x || 0.85,
    y: opts.y || 5.95,
    w: opts.w || 11.6,
    h: opts.h || 0.6,
    rectRadius: 0.03,
    line: { color: opts.line || C.line, pt: 1 },
    fill: { color: opts.fill || C.grayBox },
  });
  slide.addText(text, {
    x: (opts.x || 0.85) + 0.18,
    y: (opts.y || 5.95) + 0.12,
    w: (opts.w || 11.6) - 0.36,
    h: (opts.h || 0.6) - 0.18,
    fontFace: "Arial",
    fontSize: opts.fontSize || 11,
    color: C.body,
    italic: true,
    align: "center",
    margin: 0,
    fit: "shrink",
  });
}

function addLabelPill(slide, x, y, w, text, fill = C.blueSoft, color = C.blue) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h: 0.28,
    rectRadius: 0.08,
    line: { color: fill, pt: 0.5 },
    fill: { color: fill },
  });
  slide.addText(text, {
    x,
    y: y + 0.06,
    w,
    h: 0.1,
    fontFace: "Arial",
    fontSize: 8.5,
    bold: true,
    color,
    align: "center",
    margin: 0,
  });
}

function addSimpleIconCircle(slide, x, y, label, opts = {}) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x,
    y,
    w: opts.w || 0.42,
    h: opts.h || 0.42,
    line: { color: opts.line || C.blue, pt: 1.2 },
    fill: { color: opts.fill || C.white },
  });
  slide.addText(label, {
    x,
    y: y + 0.085,
    w: opts.w || 0.42,
    h: 0.12,
    fontFace: "Arial",
    fontSize: 9,
    bold: true,
    color: opts.color || C.blue,
    align: "center",
    margin: 0,
  });
}

function finalizeSlide(slide) {
  warnIfSlideHasOverlaps(slide, pptx, {
    muteContainment: true,
    ignoreDecorativeShapes: true,
  });
  warnIfSlideElementsOutOfBounds(slide, pptx);
}

// Slide 1
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  slide.addShape(pptx.ShapeType.line, {
    x: 0.84,
    y: 0.85,
    w: 0.95,
    h: 0,
    line: { color: C.gold, pt: 1.7 },
  });
  slide.addText("ARAUJO INNOVATION LAB", {
    x: 0.88,
    y: 1.08,
    w: 5.9,
    h: 0.3,
    fontFace: "Arial",
    fontSize: 11,
    bold: true,
    color: C.muted,
    charSpace: 1.5,
  });
  slide.addText("Araujo\nInnovation Lab", {
    x: 0.88,
    y: 1.62,
    w: 4.9,
    h: 1.6,
    fontFace: "Georgia",
    fontSize: 25,
    bold: true,
    color: C.ink,
    margin: 0,
    fit: "shrink",
  });
  slide.addText("Curso Tecnico em Marketing Digital", {
    x: 0.92,
    y: 3.35,
    w: 5.2,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 13,
    bold: true,
    color: C.blue,
    margin: 0,
  });
  slide.addText("Apresentacao Executiva Institucional", {
    x: 0.92,
    y: 3.7,
    w: 4.8,
    h: 0.24,
    fontFace: "Arial",
    fontSize: 10.8,
    color: C.muted,
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 9.45,
    y: 1.2,
    w: 2.9,
    h: 4.65,
    line: { color: C.line, pt: 1 },
    fill: { color: "FAFBFA", transparency: 0 },
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 9.75,
    y: 1.55,
    w: 2.3,
    h: 0,
    line: { color: C.lineDark, pt: 0.8 },
  });
  slide.addText("EEB Araujo Figueiredo", {
    x: 9.78,
    y: 1.8,
    w: 2.0,
    h: 0.5,
    fontFace: "Arial",
    fontSize: 14,
    bold: true,
    color: C.ink,
    margin: 0,
  });
  slide.addText("Urubici - SC", {
    x: 9.8,
    y: 2.75,
    w: 1.8,
    h: 0.22,
    fontFace: "Arial",
    fontSize: 11,
    color: C.muted,
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 9.78,
    y: 4.75,
    w: 1.8,
    h: 0,
    line: { color: C.gold, pt: 1.2 },
  });
  slide.addText("Documento Institucional\n2026", {
    x: 9.8,
    y: 4.95,
    w: 1.9,
    h: 0.42,
    fontFace: "Arial",
    fontSize: 10,
    color: C.body,
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 0.88,
    y: 6.45,
    w: 11.35,
    h: 0,
    line: { color: C.line, pt: 1 },
  });
  slide.addText("Documento Institucional  -  Apresentacao Executiva  -  2026", {
    x: 0.9,
    y: 6.58,
    w: 6,
    h: 0.18,
    fontFace: "Arial",
    fontSize: 8.5,
    color: C.muted,
    margin: 0,
  });
  finalizeSlide(slide);
}

// Slide 2
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 2, "Contexto");
  addTitle(slide, "A Transformacao da Educacao Tecnica na Economia Digital", "Urgencia publica, relevancia profissional e oportunidade regional");
  slide.addShape(pptx.ShapeType.line, {
    x: 6.64,
    y: 1.85,
    w: 0,
    h: 3.8,
    line: { color: C.line, pt: 1 },
  });
  addTextBlock(slide, 0.92, 1.9, 5.2, 0.24, "O cenario", { fontSize: 12, bold: true, color: C.blue });
  addBulletList(slide, 0.95, 2.28, 5.0, [
    "Pequenos negocios dependem cada vez mais de presenca digital para captar clientes.",
    "A comunicacao digital cresce como frente de trabalho em capitais e no interior.",
    "IA, atendimento digital e leitura de dados ja fazem parte da rotina de mercado.",
  ], { gap: 0.46, fontSize: 11.2 });
  addTextBlock(slide, 6.95, 1.9, 5.0, 0.24, "A necessidade educacional", { fontSize: 12, bold: true, color: C.blue });
  addBulletList(slide, 6.98, 2.28, 5.0, [
    "A formacao tecnica precisa gerar repertorio aplicavel, e nao apenas familiaridade conceitual.",
    "Estudantes precisam sair com portfolio, criterio e maturidade profissional crescente.",
    "A escola tecnica pode atuar como polo de inovacao educacional e desenvolvimento regional.",
  ], { gap: 0.46, fontSize: 11.2 });
  addQuoteBand(slide, "Nao se trata de modismo tecnologico. Trata-se de preparar estudantes para a realidade profissional que ja esta em curso.");
  finalizeSlide(slide);
}

// Slide 3
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 3, "Relevancia");
  addTitle(slide, "Marketing Digital: Da Comunicacao a Estrategia de Negocio", "Uma area aplicada, regionalmente aderente e profissionalmente robusta");
  const cards = [
    {
      x: 0.92, y: 1.95, w: 5.62, h: 1.48, badge: "01", title: "Relevancia de Mercado",
      body: ["Funcao estrategica em empresas de diferentes portes.", "Demanda crescente por perfis tecnicos, criativos e analiticos."],
    },
    {
      x: 6.78, y: 1.95, w: 5.62, h: 1.48, badge: "02", title: "Transformacao da Comunicacao",
      body: ["Clientes buscam relacao, confianca e experiencia.", "A descoberta de negocios locais acontece primeiro no digital."],
    },
    {
      x: 0.92, y: 3.74, w: 5.62, h: 1.48, badge: "03", title: "Presenca Digital como Base",
      body: ["Google Meu Negocio, reviews e redes definem visibilidade.", "Ausencia digital reduz relevancia na percepcao do consumidor."],
    },
    {
      x: 6.78, y: 3.74, w: 5.62, h: 1.48, badge: "04", title: "Oportunidade Regional",
      body: ["Turismo, gastronomia e servicos dependem de boa comunicacao.", "O territorio oferece campo real para aprendizagem aplicada."],
      fill: C.goldSoft,
      line: C.gold,
      badgeLine: C.gold,
      badgeColor: C.gold,
    },
  ];
  cards.forEach((c) => addCard(slide, c.x, c.y, c.w, c.h, c.title, c.body, c));
  addQuoteBand(slide, "Em Urubici, comunicar melhor nao e apenas diferencial. Em muitos casos, e condicao de competitividade.");
  finalizeSlide(slide);
}

// Slide 4
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 4, "Visao");
  addTitle(slide, "Uma Visao de Formacao Aplicada, Responsavel e Continua", "");
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 2.0, y: 1.55, w: 9.1, h: 2.05,
    line: { color: C.blueSoft, pt: 1 },
    fill: { color: C.blueSoft2 },
  });
  addTextBlock(slide, 2.4, 2.05, 8.3, 1.1,
    '"Ser um laboratorio institucional de referencia regional em educacao aplicada, marketing tecnico e transformacao digital, conectando escola, estudantes, mercado e tecnologia de forma consistente, etica e escalavel."',
    { fontFace: "Georgia", fontSize: 18, color: C.ink, italic: true, align: "center" });
  addCard(slide, 0.92, 4.25, 3.9, 1.45, "Educacao com Responsabilidade", [
    "Teoria, pratica e reflexao etica em equilibrio.",
    "IA e automacao como apoio pedagogico.",
  ], { badge: "E1" });
  addCard(slide, 4.72, 4.25, 3.9, 1.45, "Mercado como Campo de Aprendizagem", [
    "Projetos, visitas e experiencias de observacao.",
    "Portfolio construido antes do termino do curso.",
  ], { badge: "E2" });
  addCard(slide, 8.52, 4.25, 3.9, 1.45, "Continuidade Institucional", [
    "Memoria, documentacao e apoio da escola.",
    "Modelo capaz de amadurecer ao longo do tempo.",
  ], { badge: "E3" });
  finalizeSlide(slide);
}

// Slide 5
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 5, "Estrutura");
  addTitle(slide, "O Laboratorio como Ecossistema de Formacao", "Aprender, produzir e registrar evolucao");
  addCard(slide, 0.92, 2.0, 3.9, 2.05, "Aprender e Experimentar", [
    "Espaco configurado para pratica orientada.",
    "Complexidade adequada a cada etapa de formacao.",
  ], { badge: "A" });
  addCard(slide, 4.72, 2.0, 3.9, 2.05, "Aplicar e Produzir", [
    "Tres anos de progressao com projetos integradores.",
    "Competencias reunidas em entregas concretas.",
  ], { badge: "P" });
  addCard(slide, 8.52, 2.0, 3.9, 2.05, "Conectar e Preservar", [
    "Territorio como fonte de experiencias e parcerias.",
    "Registro essencial de aprendizados e decisoes.",
  ], { badge: "C" });
  addQuoteBand(slide, "O laboratorio nao substitui o curriculo oficial. Ele ajuda a torna-lo mais vivo, mais legivel e mais conectado a realidade.");
  finalizeSlide(slide);
}

// Slide 6
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 6, "Ecossistema");
  addTitle(slide, "Escola, Estudantes, Mercado e Comunidade em Relacao Continua", "Uma troca coordenada e sustentavel");
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 5.45, y: 2.72, w: 2.45, h: 1.2,
    line: { color: C.blue, pt: 1.4 },
    fill: { color: C.blueSoft },
  });
  addTextBlock(slide, 5.75, 3.03, 1.85, 0.45, "Araujo Innovation Lab", { fontSize: 16, bold: true, color: C.blue, align: "center" });
  const nodes = [
    { x: 0.9, y: 2.6, title: "Escola", lines: ["Direcao e coordenacao", "Estrutura e apoio institucional"], cx: 2.35, cy: 3.15 },
    { x: 10.05, y: 2.6, title: "Mercado Local", lines: ["Demandas e experiencias", "Parcerias educacionais"], cx: 10.95, cy: 3.15 },
    { x: 3.2, y: 4.72, title: "Estudantes", lines: ["Aprendizagem progressiva", "Portfolio e repertorio"], cx: 4.0, cy: 5.2 },
    { x: 7.62, y: 4.72, title: "Comunidade", lines: ["Confianca e reconhecimento", "Impacto social e regional"], cx: 8.4, cy: 5.2 },
  ];
  nodes.forEach((n) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: n.x, y: n.y, w: 2.35, h: 1.02,
      line: { color: C.lineDark, pt: 1 },
      fill: { color: C.white },
    });
    addTextBlock(slide, n.x + 0.16, n.y + 0.14, 2.0, 0.2, n.title, { fontSize: 11.5, bold: true, color: C.ink });
    addTextBlock(slide, n.x + 0.16, n.y + 0.4, 2.0, 0.44, `${n.lines[0]}\n${n.lines[1]}`, { fontSize: 9.7, color: C.body });
  });
  const lines = [
    [3.25, 3.1, 2.2, 0.22],
    [7.9, 3.1, 2.15, 0.22],
    [6.0, 3.9, -1.55, 0.95],
    [7.35, 3.9, 1.4, 0.95],
  ];
  lines.forEach(([x, y, w, h]) => {
    slide.addShape(pptx.ShapeType.line, { x, y, w, h, line: { color: C.lineDark, pt: 1.1, beginArrowType: "none", endArrowType: "triangle" } });
  });
  addQuoteBand(slide, "Cada parte cumpre um papel diferente. O valor do projeto esta justamente na capacidade de articular essas relacoes sem perder foco pedagogico.");
  finalizeSlide(slide);
}

// Slide 7
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 7, "Aprendizagem");
  addTitle(slide, "Matriz 6D de Aprendizagem", "Uma das assinaturas visuais do projeto");
  const dims = [
    ["Teoria", "Base conceitual e visao de contexto"],
    ["Laboratorio", "Pratica guiada com apoio e feedback"],
    ["Projetos", "Entregas estruturadas e casos aplicados"],
    ["Experiencias", "Observacoes e leitura de territorio"],
    ["Mercado", "Contato progressivo com demandas reais"],
    ["Portfolio", "Evidencias de trajetoria e maturidade"],
  ];
  let idx = 0;
  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < 3; col++) {
      const x = 0.92 + col * 4.04;
      const y = 1.95 + row * 1.72;
      const [title, desc] = dims[idx];
      slide.addShape(pptx.ShapeType.roundRect, {
        x, y, w: 3.68, h: 1.34,
        line: { color: C.line, pt: 1 },
        fill: { color: idx % 2 === 0 ? C.white : C.blueSoft2 },
      });
      addSimpleIconCircle(slide, x + 0.16, y + 0.16, (idx + 1).toString(), { line: C.gold, color: C.gold });
      addTextBlock(slide, x + 0.65, y + 0.15, 2.8, 0.2, title, { fontSize: 12, bold: true, color: C.ink });
      addTextBlock(slide, x + 0.18, y + 0.54, 3.16, 0.42, desc, { fontSize: 10.2, color: C.body });
      idx++;
    }
  }
  addLabelPill(slide, 1.4, 5.72, 2.15, "1o ano  -  base");
  addLabelPill(slide, 5.55, 5.72, 2.25, "2o ano  -  integracao", C.goldSoft, C.gold);
  addLabelPill(slide, 9.75, 5.72, 2.2, "3o ano  -  consolidacao");
  slide.addShape(pptx.ShapeType.line, {
    x: 1.85, y: 6.25, w: 8.95, h: 0,
    line: { color: C.lineDark, pt: 1.2, beginArrowType: "none", endArrowType: "triangle" },
  });
  finalizeSlide(slide);
}

// Slide 8
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 8, "Laboratorio");
  addTitle(slide, "Um Ambiente Vivo de Experimentacao Profissional", "Tecnologia e pratica como apoio a aprendizagem");
  const fronts = [
    ["Branding e Identidade Visual", "Marca, linguagem visual e coerencia institucional"],
    ["Conteudo e Narrativa", "Escrita estrategica e producao de comunicacao"],
    ["Presenca Digital", "Redes sociais, Google Meu Negocio e visibilidade local"],
    ["Reputacao e Relacionamento", "Reviews, atendimento e confianca publica"],
    ["Tecnologia como Apoio", "IA, automacao simples e produtividade"],
    ["Metricas e Operacao", "Leitura basica de resultados e organizacao digital"],
  ];
  idx = 0;
  for (let row = 0; row < 3; row++) {
    for (let col = 0; col < 2; col++) {
      const x = 0.92 + col * 6.14;
      const y = 1.88 + row * 1.23;
      slide.addShape(pptx.ShapeType.roundRect, {
        x, y, w: 5.34, h: 0.96,
        line: { color: C.line, pt: 1 },
        fill: { color: C.white },
      });
      addSimpleIconCircle(slide, x + 0.16, y + 0.16, (idx + 1).toString(), { line: C.blue, color: C.blue });
      addTextBlock(slide, x + 0.64, y + 0.15, 4.4, 0.18, fronts[idx][0], { fontSize: 11.3, bold: true, color: C.ink });
      addTextBlock(slide, x + 0.64, y + 0.42, 4.4, 0.26, fronts[idx][1], { fontSize: 10, color: C.body });
      idx++;
    }
  }
  addLabelPill(slide, 1.0, 5.86, 3.5, "Experimentacao segura");
  addLabelPill(slide, 4.95, 5.86, 3.5, "Complexidade gradual", C.goldSoft, C.gold);
  addLabelPill(slide, 8.9, 5.86, 3.5, "Etica digital");
  finalizeSlide(slide);
}

// Slide 9
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 9, "Interdisciplinaridade");
  addTitle(slide, "O Marketing Digital na Intersecao das Disciplinas", "Convergencia curricular com aplicacao concreta");
  const cards = [
    ["Lingua Portuguesa", "Copywriting, tom de voz e revisao estrategica"],
    ["Matematica", "KPIs, conversao, orcamento e leitura de resultados"],
    ["Artes / Educacao Visual", "Branding, identidade visual e direcao de arte"],
    ["Tecnologia / Informatica", "IA, automacao simples, plataformas e organizacao digital"],
    ["Sociologia / Ciencias Humanas", "Segmentacao, comportamento do consumidor e responsabilidade comunicacional"],
  ];
  cards.forEach((c, i) => {
    const y = 1.92 + i * 0.82;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 1.0, y, w: 11.2, h: 0.62,
      line: { color: C.line, pt: 1 },
      fill: { color: i % 2 === 0 ? C.white : C.blueSoft2 },
    });
    addTextBlock(slide, 1.22, y + 0.14, 3.0, 0.14, c[0], { fontSize: 11, bold: true, color: C.ink });
    slide.addShape(pptx.ShapeType.line, {
      x: 4.32, y: y + 0.31, w: 0.62, h: 0,
      line: { color: C.gold, pt: 1.1, endArrowType: "triangle" },
    });
    addTextBlock(slide, 5.1, y + 0.14, 6.7, 0.2, c[1], { fontSize: 10.2, color: C.body });
  });
  addQuoteBand(slide, "A interdisciplinaridade, aqui, nao e adorno conceitual. E a forma pratica de dar profundidade ao curso tecnico.");
  finalizeSlide(slide);
}

// Slide 10
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 10, "Governanca");
  addTitle(slide, "Governanca Educacional", "Coordenacao clara, apoio institucional e operacao sustentavel");
  const bands = [
    { y: 1.95, h: 1.0, fill: C.blueSoft2, title: "Faixa 1  -  Lideranca Institucional", items: ["Direcao", "Orientacao Pedagogica"] },
    { y: 3.18, h: 1.05, fill: C.white, title: "Faixa 2  -  Coordenacao Pedagogica", items: ["Professora Titular", "Professor Orientador"] },
    { y: 4.48, h: 1.14, fill: C.blueSoft2, title: "Faixa 3  -  Operacao Formativa", items: ["Professores Interdisciplinares", "Estudantes"] },
  ];
  bands.forEach((b) => {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.95, y: b.y, w: 9.2, h: b.h,
      line: { color: C.line, pt: 1 },
      fill: { color: b.fill },
    });
    addTextBlock(slide, 1.18, b.y + 0.14, 3.0, 0.18, b.title, { fontSize: 11.4, bold: true, color: C.blue });
    b.items.forEach((item, idx2) => {
      const bx = 1.2 + idx2 * 4.1;
      slide.addShape(pptx.ShapeType.roundRect, {
        x: bx, y: b.y + 0.42, w: 3.4, h: 0.38,
        line: { color: C.lineDark, pt: 0.8 },
        fill: { color: C.white },
      });
      addTextBlock(slide, bx + 0.12, b.y + 0.53, 3.12, 0.08, item, { fontSize: 9.6, color: C.ink, bold: true, align: "center" });
    });
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 10.45, y: 3.0, w: 1.9, h: 1.52,
    line: { color: C.gold, pt: 1 },
    fill: { color: C.goldSoft },
  });
  addTextBlock(slide, 10.63, 3.18, 1.54, 0.18, "Parceiros\nExternos", { fontSize: 11, bold: true, color: C.ink, align: "center" });
  addTextBlock(slide, 10.6, 3.72, 1.58, 0.34, "Experiencias,\nobservacao e projetos", { fontSize: 9.2, color: C.body, align: "center" });
  slide.addShape(pptx.ShapeType.line, {
    x: 9.9, y: 3.78, w: 0.55, h: 0,
    line: { color: C.gold, pt: 1.1, endArrowType: "triangle" },
  });
  addQuoteBand(slide, "A governanca existe para apoiar a execucao pedagogica, preservar continuidade e permitir ajustes progressivos.");
  finalizeSlide(slide);
}

// Slide 11
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 11, "Implementacao");
  addTitle(slide, "Implementacao Progressiva", "Estruturacao, maturacao e expansao com pacing realista");
  const years = [
    ["Ano 1", "Estruturacao", ["Base operacional e cultura do laboratorio", "Introducao gradual a ferramentas", "Primeiros portfolios e projetos"]],
    ["Ano 2", "Maturacao", ["Projetos mais estruturados", "Leitura de desempenho e rotina", "Primeiras experiencias com escopo controlado"]],
    ["Ano 3", "Expansao", ["Estagios e imersoes profissionais", "Projetos para clientes reais", "Modelo validado e continuidade assegurada"]],
  ];
  years.forEach((yr, i) => {
    const x = 0.95 + i * 4.08;
    const fill = i === 1 ? C.goldSoft : C.white;
    const line = i === 1 ? C.gold : C.line;
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y: 2.0, w: 3.58, h: 3.35,
      line: { color: line, pt: 1 },
      fill: { color: fill },
    });
    addSimpleIconCircle(slide, x + 1.5, 2.18, `${i + 1}`, { w: 0.55, h: 0.55, line: i === 1 ? C.gold : C.blue, color: i === 1 ? C.gold : C.blue });
    addTextBlock(slide, x + 0.88, 2.82, 1.8, 0.18, yr[0], { fontSize: 11, bold: true, color: C.muted, align: "center" });
    addTextBlock(slide, x + 0.5, 3.08, 2.6, 0.22, yr[1], { fontSize: 13, bold: true, color: C.ink, align: "center" });
    addBulletList(slide, x + 0.28, 3.48, 3.0, yr[2], { gap: 0.42, fontSize: 10.2 });
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 2.75, y: 5.72, w: 7.75, h: 0,
    line: { color: C.lineDark, pt: 1.2, beginArrowType: "none", endArrowType: "triangle" },
  });
  finalizeSlide(slide);
}

// Slide 12
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 12, "MVP");
  addTitle(slide, "MVP Operacional", "Implantacao sustentavel com complexidade controlada");
  const quads = [
    ["Infraestrutura Existente", ["Computadores ou sala funcional", "Internet e ferramentas acessiveis", "Repositorio digital organizado"]],
    ["Equipe e Papeis", ["Professora titular e professor orientador", "Apoio da coordenacao e direcao", "Ritmo de implementacao definido"]],
    ["Processos e Registro", ["Templates e fluxo de aprovacao", "Documentacao essencial", "Acompanhamento institucional"]],
    ["Primeiros Resultados", ["Mapa do ecossistema local", "Projetos integradores iniciais", "Portfolio e cultura do laboratorio"]],
  ];
  quads.forEach((q, i) => {
    const x = 0.95 + (i % 2) * 6.0;
    const y = 1.95 + Math.floor(i / 2) * 1.7;
    addCard(slide, x, y, 5.42, 1.44, q[0], q[1], { badge: `${i + 1}`, fill: i === 3 ? C.goldSoft : C.white, line: i === 3 ? C.gold : C.line });
  });
  addQuoteBand(slide, "Comecar menor, com clareza e consistencia, e o que permite crescer melhor depois.", { fill: C.blueSoft2, line: C.lineDark });
  finalizeSlide(slide);
}

// Slide 13
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 13, "Territorio");
  addTitle(slide, "O Territorio como Sala de Aula", "Urubici e regiao como campo real de aprendizagem aplicada");
  const segs = [
    ["Pousadas e Turismo", "Visibilidade local, reviews e conteudo sazonal"],
    ["Restaurantes", "Cardapios digitais, redes sociais e relacionamento"],
    ["Comercio Local", "Presenca digital e alcance regional"],
    ["Servicos Regionais", "Posicionamento e captacao na regiao"],
    ["Empreendedorismo Individual", "Primeira presenca digital e marca"],
    ["Eventos e Temporada", "Cobertura e comunicacao institucional"],
  ];
  idx = 0;
  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < 3; col++) {
      const x = 0.92 + col * 4.04;
      const y = 1.94 + row * 1.63;
      slide.addShape(pptx.ShapeType.roundRect, {
        x, y, w: 3.68, h: 1.25,
        line: { color: C.line, pt: 1 },
        fill: { color: C.white },
      });
      addSimpleIconCircle(slide, x + 0.16, y + 0.16, `${idx + 1}`, { line: C.blue, color: C.blue });
      addTextBlock(slide, x + 0.64, y + 0.14, 2.7, 0.18, segs[idx][0], { fontSize: 11.1, bold: true, color: C.ink });
      addTextBlock(slide, x + 0.18, y + 0.5, 3.16, 0.36, segs[idx][1], { fontSize: 9.8, color: C.body });
      idx++;
    }
  }
  addLabelPill(slide, 1.1, 5.74, 3.35, "Supervisao docente");
  addLabelPill(slide, 5.02, 5.74, 3.35, "Parceria educacional", C.goldSoft, C.gold);
  addLabelPill(slide, 8.94, 5.74, 3.35, "Debriefing e registro");
  finalizeSlide(slide);
}

// Slide 14
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 14, "Impacto");
  addTitle(slide, "O que Muda: Impactos para Estudantes, Escola e Comunidade", "Beneficios concretos e observaveis");
  const cols = [
    ["Estudantes", ["Portfolio iniciado ainda no curso", "Ferramentas e IA com criterio", "Maturidade profissional e empregabilidade"]],
    ["Escola", ["Fortalecimento do posicionamento regional", "Reconhecimento em educacao tecnica aplicada", "Memoria que protege o projeto"]],
    ["Comunidade", ["Melhoria da comunicacao digital local", "Talentos formados no proprio territorio", "Valor publico da educacao tecnica"]],
  ];
  cols.forEach((col, i) => {
    const x = 0.95 + i * 4.08;
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y: 2.0, w: 3.58, h: 3.1,
      line: { color: C.line, pt: 1 },
      fill: { color: i === 1 ? C.blueSoft2 : C.white },
    });
    addTextBlock(slide, x + 0.18, 2.18, 3.2, 0.2, col[0], { fontSize: 12, bold: true, color: C.blue, align: "center" });
    addBulletList(slide, x + 0.2, 2.65, 3.0, col[1], { gap: 0.56, fontSize: 10.4 });
  });
  addQuoteBand(slide, "Ao final do terceiro ano, o sinal mais claro de sucesso sera a capacidade de atender ao territorio com qualidade profissional e maturidade institucional.");
  finalizeSlide(slide);
}

// Slide 15
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 15, "Diferenciais");
  addTitle(slide, "Por que Este Modelo Funciona", "Diferenciais estruturais para duracao e evolucao");
  const diffs = [
    ["Clareza Institucional", "Papeis, prioridades e decisoes com leitura simples."],
    ["Crescimento Progressivo", "Complexidade distribuida no tempo e implementacao compatível com a escola."],
    ["Relacao Real com o Territorio", "Mercado local como campo de aprendizagem, nao como pressao prematura."],
    ["Tecnologia como Apoio", "IA e automacao a servico da aprendizagem e da produtividade."],
  ];
  diffs.forEach((d, i) => {
    const y = 2.0 + i * 0.95;
    slide.addText(String(i + 1).padStart(2, "0"), {
      x: 1.0, y: y + 0.08, w: 0.4, h: 0.2,
      fontFace: "Georgia", fontSize: 19, color: C.gold, bold: true, margin: 0, align: "center"
    });
    slide.addText(d[0], {
      x: 1.6, y, w: 4.2, h: 0.2,
      fontFace: "Arial", fontSize: 12, bold: true, color: C.ink, margin: 0
    });
    slide.addText(d[1], {
      x: 1.62, y: y + 0.28, w: 9.8, h: 0.24,
      fontFace: "Arial", fontSize: 10.4, color: C.body, margin: 0
    });
    slide.addShape(pptx.ShapeType.line, {
      x: 1.6, y: y + 0.68, w: 9.8, h: 0,
      line: { color: C.line, pt: 0.8 },
    });
  });
  addQuoteBand(slide, "O diferencial nao esta em parecer avancado. Esta em conseguir evoluir com consistencia.", { fill: C.goldSoft, line: C.gold });
  finalizeSlide(slide);
}

// Slide 16
{
  const slide = pptx.addSlide();
  addSlideBase(slide, 16, "Futuro");
  addTitle(slide, "A Evolucao Natural: Do Laboratorio a Referencia Regional", "Compromisso com evolucao responsavel");
  const xs = [2.0, 6.0, 10.0];
  const marcos = [
    ["Ano 3", "Consolidacao", ["Primeira turma completa", "Modelo validado e documentado"]],
    ["Ano 4", "Expansao", ["Ampliacao avaliada com cautela", "Reconhecimento regional crescente"]],
    ["Ano 5", "Referencia", ["Cultura institucionalizada", "Parcerias e ex-estudantes no ecossistema"]],
  ];
  slide.addShape(pptx.ShapeType.line, {
    x: 2.3, y: 3.65, w: 7.8, h: 0,
    line: { color: C.lineDark, pt: 1.3 },
  });
  marcos.forEach((m, i) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: xs[i], y: 3.43, w: 0.44, h: 0.44,
      line: { color: i === 1 ? C.gold : C.blue, pt: 1.3 },
      fill: { color: C.white },
    });
    addTextBlock(slide, xs[i] - 0.4, 2.45, 1.25, 0.16, m[0], { fontSize: 10.5, bold: true, color: C.muted, align: "center" });
    addTextBlock(slide, xs[i] - 0.75, 2.8, 1.95, 0.22, m[1], { fontSize: 12, bold: true, color: C.ink, align: "center" });
    addBulletList(slide, xs[i] - 1.0, 4.12, 2.5, m[2], { gap: 0.36, fontSize: 9.8 });
  });
  addQuoteBand(slide, "O futuro nao e previsao grandiosa. E compromisso com base organizada, confianca territorial e capacidade de aprender com a pratica.");
  finalizeSlide(slide);
}

// Slide 17
{
  const slide = pptx.addSlide();
  slide.background = { color: "F5F6F4" };
  slide.addShape(pptx.ShapeType.line, {
    x: 0.92, y: 0.9, w: 1.0, h: 0,
    line: { color: C.gold, pt: 1.7 },
  });
  slide.addText("ENCERRAMENTO", {
    x: 0.96, y: 1.15, w: 2.2, h: 0.2,
    fontFace: "Arial", fontSize: 9.5, bold: true, color: C.muted, charSpace: 1.4,
  });
  slide.addText("Educar para o Presente.\nPreparar para o Futuro.", {
    x: 1.55, y: 2.0, w: 10.2, h: 1.2,
    fontFace: "Georgia", fontSize: 24, bold: true, color: C.ink, align: "center", margin: 0,
  });
  slide.addText("A modernidade educacional nao esta na quantidade de tecnologia.\nEsta na qualidade da integracao entre escola, estudante e realidade.", {
    x: 1.55, y: 3.55, w: 10.2, h: 0.75,
    fontFace: "Arial", fontSize: 13, color: C.body, align: "center", italic: true, margin: 0,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 3.15, y: 5.15, w: 7.0, h: 0,
    line: { color: C.line, pt: 1 },
  });
  slide.addText("EEB Araujo Figueiredo  |  Curso Tecnico em Marketing Digital  |  Urubici - SC", {
    x: 1.8, y: 5.42, w: 9.8, h: 0.2,
    fontFace: "Arial", fontSize: 10.2, color: C.muted, align: "center", margin: 0,
  });
  slide.addText("Documento Institucional  -  2026", {
    x: 4.5, y: 5.74, w: 4.2, h: 0.18,
    fontFace: "Arial", fontSize: 9.4, color: C.muted, align: "center", margin: 0,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 3.2, y: 6.28, w: 6.95, h: 0.48,
    line: { color: C.line, pt: 0.8 },
    fill: { color: C.white },
  });
  slide.addText("Este e um documento vivo. Sua proxima versao sera escrita pelos resultados que ajudar a gerar.", {
    x: 3.38, y: 6.42, w: 6.55, h: 0.14,
    fontFace: "Arial", fontSize: 8.8, color: C.body, align: "center", margin: 0,
  });
  finalizeSlide(slide);
}

async function main() {
  await pptx.writeFile({
    fileName: "araujo-innovation-lab-final-executive-presentation.pptx",
  });
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
