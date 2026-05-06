const PptxGenJS = require("pptxgenjs");

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Ahmed";
pptx.company = "Office Team";
pptx.subject = "Office Introduction";
pptx.title = "Modern Office Presentation";
pptx.lang = "en-US";

const COLORS = {
  // Extracted from Share A1_W1_Day2_Slides.pptx
  bg: "F3EEFF",
  primary: "2C3E50",
  secondary: "5234B7",
  dark: "1C2833",
  muted: "7F8C8D",
  white: "FFFFFF",
  accent: "9E59CD",
  warm: "E67E22",
  success: "27AE60",
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
    w: 8.5,
    h: 0.4,
    color: COLORS.white,
    fontSize: 20,
    bold: true,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.6,
      y: 1.0,
      w: 12.2,
      h: 0.4,
      color: COLORS.muted,
      fontSize: 13,
    });
  }
}

function addFooter(slide, page) {
  slide.addText(`Office Presentation | ${page}`, {
    x: 0.6,
    y: 7.2,
    w: 12.0,
    h: 0.3,
    color: "9CA3AF",
    fontSize: 10,
    align: "right",
  });
}

// Slide 1: Title
{
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.primary };
  slide.addText("Welcome to Our Office", {
    x: 0.8,
    y: 2.1,
    w: 11.8,
    h: 0.8,
    color: COLORS.white,
    fontSize: 42,
    bold: true,
    align: "center",
  });
  slide.addText("A quick overview of our workplace, culture, and operations", {
    x: 1.1,
    y: 3.2,
    w: 11.2,
    h: 0.5,
    color: "DBEAFE",
    fontSize: 18,
    align: "center",
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 5.5,
    y: 4.3,
    w: 2.4,
    h: 0.7,
    radius: 0.1,
    fill: { color: COLORS.accent },
    line: { color: COLORS.accent },
  });
  slide.addText("2026", {
    x: 5.5,
    y: 4.48,
    w: 2.4,
    h: 0.3,
    color: COLORS.white,
    bold: true,
    align: "center",
    fontSize: 16,
  });
}

// Slide 2: Who We Are
{
  const slide = pptx.addSlide();
  addHeader(slide, "Who We Are", "Our mission and core values");
  slide.addText(
    [
      { text: "Mission: ", options: { bold: true, color: COLORS.dark } },
      { text: "Deliver excellent services with efficiency and integrity." },
      { text: "\nVision: ", options: { bold: true, color: COLORS.dark } },
      { text: "Build a smart, collaborative, and future-ready workplace." },
    ],
    {
      x: 0.8,
      y: 1.8,
      w: 12.0,
      h: 1.6,
      fontSize: 18,
      color: "334155",
      valign: "top",
    }
  );

  const values = ["Professionalism", "Teamwork", "Innovation", "Customer Focus"];
  values.forEach((value, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const x = 0.9 + col * 6.25;
    const y = 3.8 + row * 1.35;
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w: 5.5,
      h: 0.95,
      radius: 0.08,
      fill: { color: col === 0 ? "F3EEFF" : "F9F5FF" },
      line: { color: "D8CCF4" },
    });
    slide.addText(value, {
      x: x + 0.35,
      y: y + 0.28,
      w: 4.8,
      h: 0.4,
      fontSize: 16,
      bold: true,
      color: COLORS.secondary,
      align: "left",
    });
  });
  addFooter(slide, 2);
}

// Slide 3: Departments
{
  const slide = pptx.addSlide();
  addHeader(slide, "Main Departments", "How our office is structured");
  const departments = [
    ["Operations", "Ensures smooth day-to-day workflow."],
    ["Sales", "Builds client relationships and growth."],
    ["Human Resources", "Supports people and culture."],
    ["Finance", "Manages budgets and financial planning."],
  ];

  departments.forEach((dept, idx) => {
    const y = 1.7 + idx * 1.3;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.9,
      y,
      w: 11.6,
      h: 1.0,
      radius: 0.06,
      fill: { color: idx % 2 === 0 ? "FFFFFF" : "F3EEFF" },
      line: { color: COLORS.border },
    });
    slide.addText(dept[0], {
      x: 1.2,
      y: y + 0.16,
      w: 3.0,
      h: 0.35,
      fontSize: 15,
      bold: true,
      color: COLORS.primary,
    });
    slide.addText(dept[1], {
      x: 4.0,
      y: y + 0.2,
      w: 7.9,
      h: 0.35,
      fontSize: 14,
      color: COLORS.dark,
    });
  });
  addFooter(slide, 3);
}

// Slide 4: Daily Workflow
{
  const slide = pptx.addSlide();
  addHeader(slide, "Daily Workflow", "A typical productive day");
  const timeline = [
    ["09:00", "Team Stand-up & Priority Setting"],
    ["11:00", "Client Communication & Follow-ups"],
    ["13:00", "Project Execution & Collaboration"],
    ["16:00", "Reporting and Next-Day Planning"],
  ];

  timeline.forEach((item, idx) => {
    const y = 1.9 + idx * 1.25;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 1.0,
      y: y + 0.1,
      w: 0.85,
      h: 0.85,
      fill: { color: COLORS.secondary },
      line: { color: COLORS.secondary },
    });
    slide.addText(String(idx + 1), {
      x: 1.0,
      y: y + 0.34,
      w: 0.85,
      h: 0.25,
      color: COLORS.white,
      bold: true,
      align: "center",
      fontSize: 14,
    });
    slide.addText(item[0], {
      x: 2.1,
      y: y + 0.18,
      w: 1.3,
      h: 0.3,
      color: COLORS.primary,
      bold: true,
      fontSize: 13,
    });
    slide.addText(item[1], {
      x: 3.3,
      y: y + 0.18,
      w: 8.8,
      h: 0.3,
      color: COLORS.dark,
      fontSize: 14,
    });
  });
  addFooter(slide, 4);
}

// Slide 5: Closing
{
  const slide = pptx.addSlide();
  slide.background = { color: COLORS.bg };
  slide.addText("Thank You", {
    x: 0.8,
    y: 2.2,
    w: 11.8,
    h: 0.9,
    align: "center",
    fontSize: 44,
    bold: true,
    color: COLORS.primary,
  });
  slide.addText("We look forward to achieving great results together.", {
    x: 1.2,
    y: 3.3,
    w: 11.0,
    h: 0.5,
    align: "center",
    fontSize: 19,
    color: COLORS.dark,
  });
  slide.addText("Contact: office@example.com", {
    x: 1.2,
    y: 4.2,
    w: 11.0,
    h: 0.4,
    align: "center",
    fontSize: 14,
    color: COLORS.muted,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 5.1,
    y: 4.8,
    w: 3.1,
    h: 0.45,
    radius: 0.07,
    fill: { color: COLORS.warm },
    line: { color: COLORS.warm },
  });
  slide.addText("Let's Build Success", {
    x: 5.1,
    y: 4.9,
    w: 3.1,
    h: 0.2,
    align: "center",
    fontSize: 12,
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
