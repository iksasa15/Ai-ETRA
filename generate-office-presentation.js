const PptxGenJS = require("pptxgenjs");

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Ahmed";
pptx.title = "Empty Presentation";
pptx.lang = "en-US";

const slide = pptx.addSlide();
slide.background = { color: "FFFFFF" };
slide.addText("Day 1 - Introduction to Natural Language Processing (NLP)", {
  x: 0.8,
  y: 2.8,
  w: 11.8,
  h: 0.8,
  align: "center",
  fontSize: 34,
  bold: true,
  color: "1F2937",
});

const slide2 = pptx.addSlide();
slide2.background = { color: "FFFFFF" };
slide2.addText("1. What Is NLP and Why Is It Hard?", {
  x: 0.7,
  y: 0.6,
  w: 12.0,
  h: 0.5,
  fontSize: 28,
  bold: true,
  color: "1F2937",
});
slide2.addText("1.1 Definition", {
  x: 0.8,
  y: 1.4,
  w: 4.0,
  h: 0.4,
  fontSize: 20,
  bold: true,
  color: "374151",
});
slide2.addText(
  "Natural Language Processing (NLP) is a subfield of Artificial Intelligence focused on enabling computers to understand, generate, and interact with human language.",
  {
    x: 0.8,
    y: 2.0,
    w: 11.7,
    h: 1.0,
    fontSize: 16,
    color: "111827",
    valign: "top",
  }
);

const slide3 = pptx.addSlide();
slide3.background = { color: "FFFFFF" };
slide3.addText("1.2 Why Is Language Hard for Computers?", {
  x: 0.7,
  y: 0.6,
  w: 12.0,
  h: 0.5,
  fontSize: 28,
  bold: true,
  color: "1F2937",
});
slide3.addText("Ambiguity is the biggest challenge in NLP.", {
  x: 0.8,
  y: 1.35,
  w: 11.8,
  h: 0.4,
  fontSize: 17,
  bold: true,
  color: "374151",
});
slide3.addShape(pptx.ShapeType.roundRect, {
  x: 0.8,
  y: 1.85,
  w: 11.8,
  h: 1.5,
  radius: 0.05,
  fill: { color: "F9FAFB" },
  line: { color: "D1D5DB" },
});
slide3.addText(
  "Sentence: \"I saw the man with the telescope\"\nInterpretation 1: I used a telescope to see the man.\nInterpretation 2: I saw a man who was carrying a telescope.",
  {
    x: 1.1,
    y: 2.08,
    w: 11.1,
    h: 1.1,
    fontSize: 14,
    color: "111827",
    valign: "top",
  }
);
slide3.addText("Types of Ambiguity", {
  x: 0.8,
  y: 3.65,
  w: 4.0,
  h: 0.35,
  fontSize: 18,
  bold: true,
  color: "1F2937",
});
slide3.addShape(pptx.ShapeType.roundRect, {
  x: 0.8,
  y: 4.1,
  w: 11.8,
  h: 2.8,
  radius: 0.05,
  fill: { color: "FFFFFF" },
  line: { color: "D1D5DB" },
});
slide3.addText("Lexical Ambiguity", {
  x: 1.0, y: 4.35, w: 3.0, h: 0.25, fontSize: 13, bold: true, color: "1F2937"
});
slide3.addText("Example: \"bank\"", {
  x: 3.5, y: 4.35, w: 2.5, h: 0.25, fontSize: 13, color: "374151"
});
slide3.addText("Interpretation: financial bank or river bank?", {
  x: 6.0, y: 4.35, w: 6.2, h: 0.25, fontSize: 13, color: "374151"
});

slide3.addText("Syntactic Ambiguity", {
  x: 1.0, y: 4.95, w: 3.0, h: 0.25, fontSize: 13, bold: true, color: "1F2937"
});
slide3.addText("Example: \"I saw the child by the window\"", {
  x: 3.5, y: 4.95, w: 2.5, h: 0.25, fontSize: 13, color: "374151"
});
slide3.addText("Interpretation: who exactly was by the window?", {
  x: 6.0, y: 4.95, w: 6.2, h: 0.25, fontSize: 13, color: "374151"
});

slide3.addText("Semantic Ambiguity", {
  x: 1.0, y: 5.55, w: 3.0, h: 0.25, fontSize: 13, bold: true, color: "1F2937"
});
slide3.addText("Example: \"The mouse eats the cat\"", {
  x: 3.5, y: 5.55, w: 2.5, h: 0.25, fontSize: 13, color: "374151"
});
slide3.addText("Interpretation: unusual meaning, who is eating whom?", {
  x: 6.0, y: 5.55, w: 6.2, h: 0.25, fontSize: 13, color: "374151"
});

slide3.addText("Pragmatic Ambiguity", {
  x: 1.0, y: 6.15, w: 3.0, h: 0.25, fontSize: 13, bold: true, color: "1F2937"
});
slide3.addText("Example: \"Can you open the door?\"", {
  x: 3.5, y: 6.15, w: 2.5, h: 0.25, fontSize: 13, color: "374151"
});
slide3.addText("Interpretation: literal question or polite request?", {
  x: 6.0, y: 6.15, w: 6.2, h: 0.25, fontSize: 13, color: "374151"
});

const slide4 = pptx.addSlide();
slide4.background = { color: "FFFFFF" };
slide4.addText("Other NLP Challenges", {
  x: 0.7,
  y: 0.7,
  w: 12.0,
  h: 0.5,
  fontSize: 30,
  bold: true,
  color: "1F2937",
});
slide4.addText("Beyond ambiguity, language has many complex signals:", {
  x: 0.8,
  y: 1.45,
  w: 11.8,
  h: 0.4,
  fontSize: 16,
  color: "374151",
});

slide4.addShape(pptx.ShapeType.roundRect, {
  x: 0.9,
  y: 2.0,
  w: 11.7,
  h: 4.8,
  radius: 0.07,
  fill: { color: "F9FAFB" },
  line: { color: "D1D5DB" },
});

slide4.addText(
  [
    { text: "Sarcasm: ", options: { bold: true, color: "111827" } },
    { text: "\"Great! We are late again.\" (meaning is opposite to the words)\n", options: { color: "374151" } },
    { text: "Dialects: ", options: { bold: true, color: "111827" } },
    { text: "\"Zain\" can mean \"good\" in Gulf Arabic, but it can also be a person name in Persian.\n", options: { color: "374151" } },
    { text: "Abbreviations: ", options: { bold: true, color: "111827" } },
    { text: "One short form can have multiple meanings depending on context.\n", options: { color: "374151" } },
    { text: "Spelling Errors: ", options: { bold: true, color: "111827" } },
    { text: "Small spelling mistakes can change meaning and reduce model accuracy.\n", options: { color: "374151" } },
    { text: "Evolving Language: ", options: { bold: true, color: "111827" } },
    { text: "Slang, proverbs, and idiomatic expressions constantly change usage.", options: { color: "374151" } },
  ],
  {
    x: 1.2,
    y: 2.35,
    w: 11.0,
    h: 4.0,
    fontSize: 15,
    breakLine: true,
    valign: "top",
  }
);

const slide5 = pptx.addSlide();
slide5.background = { color: "FFFFFF" };
slide5.addText("1.3 Real-World NLP Applications", {
  x: 0.7,
  y: 0.7,
  w: 12.0,
  h: 0.5,
  fontSize: 30,
  bold: true,
  color: "1F2937",
});
slide5.addText("Examples of NLP in products and industries:", {
  x: 0.8,
  y: 1.45,
  w: 11.8,
  h: 0.4,
  fontSize: 16,
  color: "374151",
});
slide5.addShape(pptx.ShapeType.roundRect, {
  x: 0.9,
  y: 2.0,
  w: 11.7,
  h: 4.9,
  radius: 0.07,
  fill: { color: "F9FAFB" },
  line: { color: "D1D5DB" },
});
slide5.addText(
  [
    { text: "Google Search -> ", options: { bold: true, color: "111827" } },
    { text: "Understanding user search intent\n", options: { color: "374151" } },
    { text: "ChatGPT / Claude -> ", options: { bold: true, color: "111827" } },
    { text: "Generating text and answers\n", options: { color: "374151" } },
    { text: "Gmail -> ", options: { bold: true, color: "111827" } },
    { text: "Spam filtering\n", options: { color: "374151" } },
    { text: "Siri / Alexa -> ", options: { bold: true, color: "111827" } },
    { text: "Speech recognition\n", options: { color: "374151" } },
    { text: "Google Translate -> ", options: { bold: true, color: "111827" } },
    { text: "Machine translation\n", options: { color: "374151" } },
    { text: "Bloomberg -> ", options: { bold: true, color: "111827" } },
    { text: "Financial market sentiment analysis\n", options: { color: "374151" } },
    { text: "Hospitals -> ", options: { bold: true, color: "111827" } },
    { text: "Extracting information from medical records\n", options: { color: "374151" } },
    { text: "Twitter/X -> ", options: { bold: true, color: "111827" } },
    { text: "Detecting fake news", options: { color: "374151" } },
  ],
  {
    x: 1.2,
    y: 2.35,
    w: 11.0,
    h: 4.2,
    fontSize: 15,
    breakLine: true,
    valign: "top",
  }
);

const slide6 = pptx.addSlide();
slide6.background = { color: "FFFFFF" };
slide6.addText("2. The Complete NLP Pipeline", {
  x: 0.7,
  y: 0.6,
  w: 12.0,
  h: 0.5,
  fontSize: 30,
  bold: true,
  color: "1F2937",
});
slide6.addText("The pipeline is a sequence of transformations that converts raw text into data a model can understand.", {
  x: 0.8,
  y: 1.3,
  w: 12.0,
  h: 0.5,
  fontSize: 15,
  color: "374151",
});

const pipelineSteps = [
  "Raw Text",
  "[1] Text Cleaning",
  "[2] Tokenization",
  "[3] Stop Words Removal",
  "[4] Normalization",
  "[5] Stemming / Lemmatization",
  "[6] POS Tagging",
  "[7] Named Entity Recognition (NER)",
  "Model-Ready Structured Data",
];

pipelineSteps.forEach((step, idx) => {
  const y = 1.95 + idx * 0.56;
  const isEdge = idx === 0 || idx === pipelineSteps.length - 1;
  slide6.addShape(pptx.ShapeType.roundRect, {
    x: 3.0,
    y,
    w: 7.3,
    h: 0.42,
    radius: 0.05,
    fill: { color: isEdge ? "EDE9FE" : "F9FAFB" },
    line: { color: "D1D5DB" },
  });
  slide6.addText(step, {
    x: 3.2,
    y: y + 0.1,
    w: 6.9,
    h: 0.22,
    align: "center",
    fontSize: 12,
    bold: isEdge,
    color: "111827",
  });
  if (idx < pipelineSteps.length - 1) {
    slide6.addText("↓", {
      x: 6.5,
      y: y + 0.42,
      w: 0.3,
      h: 0.15,
      align: "center",
      fontSize: 12,
      color: "6B7280",
    });
  }
});
slide2.addText("Human language is not just data - it carries:", {
  x: 0.8,
  y: 3.25,
  w: 8.0,
  h: 0.4,
  fontSize: 16,
  bold: true,
  color: "1F2937",
});
slide2.addText(
  [
    { text: "Semantics: ", options: { bold: true } },
    { text: "What does the word mean?" },
    { text: "\nContext: ", options: { bold: true } },
    { text: "What does the word mean here?" },
    { text: "\nIntent: ", options: { bold: true } },
    { text: "What does the speaker want?" },
    { text: "\nSentiment: ", options: { bold: true } },
    { text: "How does the speaker feel?" },
  ],
  {
    x: 1.1,
    y: 3.8,
    w: 11.2,
    h: 2.2,
    fontSize: 15,
    color: "111827",
    breakLine: true,
    valign: "top",
  }
);

pptx
  .writeFile({ fileName: "Office-Presentation-English.pptx" })
  .then(() => {
    console.log("Presentation created: Office-Presentation-English.pptx");
  })
  .catch((err) => {
    console.error("Failed to create presentation:", err);
    process.exitCode = 1;
  });
