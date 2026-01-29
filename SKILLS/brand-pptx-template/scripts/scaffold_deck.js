#!/usr/bin/env node
// Create a minimal branded base deck with common slide types.

const pptxgen = require("pptxgenjs");

const args = process.argv.slice(2);
const getArg = (name, fallback) => {
  const idx = args.indexOf(name);
  if (idx === -1) return fallback;
  return args[idx + 1] ?? fallback;
};

const variant = getArg("--variant", "light").toLowerCase();
const company = getArg("--company", "Company Name");
const out = getArg("--out", "template-base.pptx");

const COLORS = {
  light: {
    background: "FFFFFF",
    text: "111827",
    accent: "1F6BFF",
  },
  dark: {
    background: "0B0B0E",
    text: "F8FAFC",
    accent: "38D3FF",
  },
};

const theme = COLORS[variant] || COLORS.light;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = company;
pres.company = company;
pres.title = `${company} Template`;

const addTitle = (slide, title) => {
  slide.addText(title, {
    x: 0.6,
    y: 0.4,
    w: 12,
    h: 0.8,
    fontSize: 36,
    color: theme.text,
    bold: true,
  });
};

const addFooter = (slide) => {
  slide.addText(company, {
    x: 0.6,
    y: 6.8,
    w: 12,
    h: 0.3,
    fontSize: 10,
    color: theme.text,
    opacity: 60,
  });
};

const addSlide = (title, body) => {
  const slide = pres.addSlide();
  slide.background = { color: theme.background };
  addTitle(slide, title);
  slide.addText(body, {
    x: 0.6,
    y: 1.4,
    w: 12,
    h: 4.8,
    fontSize: 16,
    color: theme.text,
  });
  addFooter(slide);
};

// Title
{
  const slide = pres.addSlide();
  slide.background = { color: theme.background };
  slide.addText(company, {
    x: 0.8,
    y: 2.2,
    w: 12,
    h: 1.0,
    fontSize: 44,
    color: theme.text,
    bold: true,
  });
  slide.addShape(pres.ShapeType.line, {
    x: 0.8,
    y: 3.6,
    w: 4,
    h: 0,
    line: { color: theme.accent, width: 2 },
  });
  slide.addText("Presentation Title", {
    x: 0.8,
    y: 4.0,
    w: 12,
    h: 0.5,
    fontSize: 20,
    color: theme.text,
  });
}

addSlide("Section Divider", "Section description goes here.");
addSlide("Content Slide", "Body text goes here.");
addSlide("Two Column", "Left column\n\nRight column");
addSlide("Feature Grid", "Feature 1\nFeature 2\nFeature 3\nFeature 4");
addSlide("Metrics", "Metric 1\nMetric 2\nMetric 3");
addSlide("Quote", "\"Quote or testimonial\"\n- Attribution");
addSlide("Timeline", "Step 1\nStep 2\nStep 3");
addSlide("Closing", "Thank you\ncontact@example.com");

pres.writeFile({ fileName: out });
