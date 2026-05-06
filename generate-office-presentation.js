const PptxGenJS = require("pptxgenjs");

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Ahmed";
pptx.company = "NLP Learning Track";
pptx.subject = "Introduction to Natural Language Processing";
pptx.title = "Day 1 - NLP Introduction";
pptx.lang = "en-US";

const COLORS = {
  bg: "F3EEFF",
  primary: "2C3E50",
  secondary: "5234B7",
  dark: "1C2833",
  muted: "7F8C8D",
  white: "FFFFFF",
  accent: "9E59CD",
  warm: "E67E22",
  border: "BDC3C7",
};

function addHeader(slide, title, subtitle) {
  slide.background = { color: COLORS.bg };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.33,
    h: 0.9,
    fill: { color: COLORS.primary },
    line: { color: COLORS.primary },
  });
  slide.addText(title, {
    x: 0.6,
    y: 0.22,
    w: 10.5,
    h: 0.4,
    color: COLORS.white,
    fontSize: 20,
    bold: true,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.6,
      y: 1.0,
      w: 12.0,
      h: 0.4,
      color: COLORS.muted,
      fontSize: 13,
    });
  }
}

function addFooter(slide, page) {
  slide.addText(`NLP Day 1 | ${page}`, {
    x: 0.6,
    y: 7.15,
    w: 12.0,
    h: 0.3,
    color: "9CA3AF",
    fontSize: 10,
    align: "right",
  });
}

function bulletText(lines) {
  return lines.map((line) => ({ text: line, options: { bullet: { indent: 14 } } }));
}

// Slide 1: Title
{
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.primary };
  slide.addText("Day 1: Introduction to NLP", {
    x: 0.8,
    y: 2.0,
    w: 11.8,
    h: 0.8,
    color: COLORS.white,
    fontSize: 40,
    bold: true,
    align: "center",
  });
  slide.addText("Theory and practical examples (no coding)", {
    x: 1.3,
    y: 3.1,
    w: 10.8,
    h: 0.5,
    color: "DBEAFE",
    fontSize: 18,
    align: "center",
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 4.9,
    y: 4.2,
    w: 3.5,
    h: 0.7,
    radius: 0.08,
    fill: { color: COLORS.accent },
    line: { color: COLORS.accent },
  });
  slide.addText("Beginner -> Intermediate", {
    x: 4.9,
    y: 4.42,
    w: 3.5,
    h: 0.26,
    align: "center",
    color: COLORS.white,
    bold: true,
    fontSize: 14,
  });
}

// Slide 2: What is NLP
{
  const slide = pptx.addSlide();
  addHeader(slide, "What Is NLP and Why Is It Hard?", "Core concepts and challenges");
  slide.addText("Natural Language Processing enables computers to understand and generate human language.", {
    x: 0.8,
    y: 1.6,
    w: 12.0,
    h: 0.7,
    fontSize: 16,
    color: COLORS.dark,
  });
  slide.addText(bulletText([
    "Language includes meaning, context, intent, and emotion.",
    "Ambiguity is the biggest challenge.",
    "Humans resolve ambiguity quickly using context; machines need explicit modeling.",
    "Common issues: sarcasm, idioms, misspellings, slang, negation, and coreference."
  ]), {
    x: 1.0,
    y: 2.4,
    w: 11.4,
    h: 2.6,
    fontSize: 15,
    color: COLORS.dark,
    breakLine: true,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.9,
    y: 5.4,
    w: 11.8,
    h: 1.3,
    radius: 0.06,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.border },
  });
  slide.addText("Example ambiguity: \"I saw the man with the telescope\" can mean either you used a telescope, or the man had one.", {
    x: 1.2,
    y: 5.75,
    w: 11.2,
    h: 0.7,
    fontSize: 14,
    italic: true,
    color: COLORS.secondary,
  });
  addFooter(slide, 2);
}

// Slide 3: NLP applications
{
  const slide = pptx.addSlide();
  addHeader(slide, "Real-World NLP Applications", "Where NLP appears in daily products");
  slide.addText(bulletText([
    "Search engines: understanding user intent behind queries.",
    "Virtual assistants: speech recognition and command interpretation.",
    "Machine translation: converting meaning across languages.",
    "Email systems: spam filtering and smart reply generation.",
    "Healthcare and finance: extracting structured insights from documents.",
    "Social media moderation: detecting hate speech, abuse, and misinformation."
  ]), {
    x: 0.9,
    y: 1.6,
    w: 12.0,
    h: 4.8,
    fontSize: 16,
    color: COLORS.dark,
    breakLine: true,
  });
  addFooter(slide, 3);
}

// Slide 4: Full pipeline
{
  const slide = pptx.addSlide();
  addHeader(slide, "The Full NLP Pipeline", "From raw text to structured data");
  const steps = [
    "1) Text Cleaning",
    "2) Tokenization",
    "3) Stop Words Handling",
    "4) Normalization",
    "5) Stemming or Lemmatization",
    "6) POS Tagging",
    "7) Named Entity Recognition",
  ];
  steps.forEach((step, idx) => {
    const y = 1.7 + idx * 0.74;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 1.0,
      y,
      w: 5.0,
      h: 0.55,
      radius: 0.06,
      fill: { color: idx % 2 === 0 ? "FFFFFF" : "F8FAFC" },
      line: { color: COLORS.border },
    });
    slide.addText(step, {
      x: 1.2,
      y: y + 0.15,
      w: 4.6,
      h: 0.25,
      fontSize: 13,
      bold: true,
      color: COLORS.primary,
    });
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.4,
    y: 1.8,
    w: 5.8,
    h: 4.6,
    radius: 0.08,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.border },
  });
  slide.addText("Example Flow", {
    x: 6.7,
    y: 2.05,
    w: 5.2,
    h: 0.3,
    bold: true,
    color: COLORS.secondary,
    fontSize: 15,
  });
  slide.addText("Input: \"Apple CEO Tim Cook said iPhone sales rose 20% last quarter.\"\nOutput: entities, key terms, grammatical roles, and normalized tokens ready for analysis.", {
    x: 6.7,
    y: 2.45,
    w: 5.1,
    h: 2.3,
    fontSize: 13,
    color: COLORS.dark,
    valign: "top",
  });
  addFooter(slide, 4);
}

// Slide 5: Tokenization and stop words
{
  const slide = pptx.addSlide();
  addHeader(slide, "Tokenization and Stop Words", "How text is split and filtered");
  slide.addText("Tokenization types:", {
    x: 0.9,
    y: 1.6,
    w: 4.5,
    h: 0.3,
    bold: true,
    fontSize: 15,
    color: COLORS.primary,
  });
  slide.addText(bulletText([
    "Word tokens: split by words and punctuation.",
    "Character tokens: every character as a token.",
    "Subword tokens: best for LLMs and unseen words.",
    "Sentence tokens: split by sentence boundaries."
  ]), {
    x: 1.0,
    y: 2.0,
    w: 5.7,
    h: 2.6,
    fontSize: 13,
    color: COLORS.dark,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.8,
    y: 1.6,
    w: 5.6,
    h: 3.1,
    radius: 0.06,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.border },
  });
  slide.addText("Example", {
    x: 7.1,
    y: 1.9,
    w: 1.8,
    h: 0.3,
    bold: true,
    color: COLORS.secondary,
    fontSize: 14,
  });
  slide.addText("Sentence: \"This is the book that talks about artificial intelligence\"\nAfter stop-word reduction: \"book talks artificial intelligence\"", {
    x: 7.1,
    y: 2.3,
    w: 5.0,
    h: 1.9,
    fontSize: 13,
    color: COLORS.dark,
  });
  slide.addText("Important: Do not remove stop words in tasks where meaning of words like \"not\" is critical.", {
    x: 1.0,
    y: 5.3,
    w: 11.6,
    h: 0.9,
    fontSize: 14,
    color: COLORS.warm,
    italic: true,
  });
  addFooter(slide, 5);
}

// Slide 6: Stemming vs Lemmatization
{
  const slide = pptx.addSlide();
  addHeader(slide, "Stemming vs Lemmatization", "Speed versus linguistic accuracy");
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.9,
    y: 1.6,
    w: 5.8,
    h: 3.8,
    radius: 0.08,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.border },
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.7,
    y: 1.6,
    w: 5.8,
    h: 3.8,
    radius: 0.08,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.border },
  });
  slide.addText("Stemming", {
    x: 1.2,
    y: 1.95,
    w: 2.2,
    h: 0.3,
    fontSize: 16,
    bold: true,
    color: COLORS.primary,
  });
  slide.addText(bulletText([
    "Rule-based truncation.",
    "Very fast.",
    "May output non-real words.",
    "Best for large-scale indexing."
  ]), {
    x: 1.2,
    y: 2.35,
    w: 5.0,
    h: 2.6,
    fontSize: 13,
    color: COLORS.dark,
  });
  slide.addText("Lemmatization", {
    x: 7.0,
    y: 1.95,
    w: 3.2,
    h: 0.3,
    fontSize: 16,
    bold: true,
    color: COLORS.secondary,
  });
  slide.addText(bulletText([
    "Dictionary and grammar aware.",
    "More accurate, slightly slower.",
    "Outputs valid base forms.",
    "Best for chatbots and QA tasks."
  ]), {
    x: 7.0,
    y: 2.35,
    w: 5.0,
    h: 2.6,
    fontSize: 13,
    color: COLORS.dark,
  });
  slide.addText("Examples: studies -> study, mice -> mouse, better -> good", {
    x: 1.0,
    y: 5.8,
    w: 11.6,
    h: 0.5,
    fontSize: 14,
    italic: true,
    color: COLORS.warm,
  });
  addFooter(slide, 6);
}

// Slide 7: POS and NER
{
  const slide = pptx.addSlide();
  addHeader(slide, "POS Tagging and Named Entity Recognition", "Understanding grammar and extracting facts");
  slide.addText("POS Tagging labels each word by grammatical role (NOUN, VERB, ADJ, etc.).", {
    x: 0.9,
    y: 1.7,
    w: 11.8,
    h: 0.45,
    fontSize: 14,
    color: COLORS.dark,
  });
  slide.addText("NER identifies entities like PERSON, ORG, DATE, MONEY, and LOCATION.", {
    x: 0.9,
    y: 2.2,
    w: 11.8,
    h: 0.45,
    fontSize: 14,
    color: COLORS.dark,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.9,
    y: 2.9,
    w: 11.8,
    h: 2.7,
    radius: 0.08,
    fill: { color: "FFFFFF" },
    line: { color: COLORS.border },
  });
  slide.addText("Example sentence: \"Apple CEO Tim Cook announced a $430 billion investment in the United States.\"\n\nPossible extracted entities:\n- ORG: Apple\n- PERSON: Tim Cook\n- MONEY: $430 billion\n- GPE: United States", {
    x: 1.2,
    y: 3.2,
    w: 11.1,
    h: 2.2,
    fontSize: 13,
    color: COLORS.secondary,
  });
  addFooter(slide, 7);
}

// Slide 8: Wrap-up
{
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.bg };
  slide.addText("Day 1 Summary", {
    x: 0.8,
    y: 1.3,
    w: 11.8,
    h: 0.8,
    align: "center",
    fontSize: 38,
    bold: true,
    color: COLORS.primary,
  });
  slide.addText(bulletText([
    "NLP is about converting language into structured, machine-usable information.",
    "The pipeline starts with cleaning and ends with rich linguistic features.",
    "Tokenization, lemmatization, POS, and NER are foundational building blocks.",
    "Choosing the right preprocessing depends on the target task and quality needs."
  ]), {
    x: 1.2,
    y: 2.4,
    w: 10.9,
    h: 2.6,
    fontSize: 16,
    color: COLORS.dark,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 4.6,
    y: 5.5,
    w: 4.1,
    h: 0.75,
    radius: 0.09,
    fill: { color: COLORS.warm },
    line: { color: COLORS.warm },
  });
  slide.addText("Next: Text Representations", {
    x: 4.6,
    y: 5.75,
    w: 4.1,
    h: 0.3,
    align: "center",
    fontSize: 14,
    bold: true,
    color: COLORS.white,
  });
}

pptx
  .writeFile({ fileName: "Office-Presentation-English.pptx" })
  .then(() => {
    console.log("Presentation created: Office-Presentation-English.pptx");
  })
  .catch((err) => {
    console.error("Failed to create presentation:", err);
    process.exitCode = 1;
  });
