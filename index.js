const express = require("express");
const pptxgen = require("pptxgenjs");

const app = express();
app.use(express.json({ limit: "10mb" }));

// ---------- Design tokens (repris du design Scope) ----------
const COLORS = {
  bg: "0A0713",
  card: "170F28",
  cardBorder: "2A2340",
  purple: "B19DFD",
  purpleDark: "8371C2",
  purpleCircle: "21163B",
  green: "4ADE80",
  white: "FFFFFF",
  textMuted: "9B95AE",
  textFaint: "6B6580"
};
const FONT = "Arial";
const W = 13.333;
const H = 7.5;

app.get("/", (req, res) => {
  res.send("Service PPTX Scope v2 — en ligne. POST /generate-pptx pour générer un fichier.");
});

// ---------- Utilitaires de mise en forme ----------

// Découpe un texte "**gras**" en runs pptxgenjs { text, options }
function boldRuns(line, baseOpts) {
  const parts = line.split(/(\*\*[^*]+\*\*)/g).filter((p) => p.length > 0);
  return parts.map((part) => {
    if (part.startsWith("**") && part.endsWith("**")) {
      return { text: part.slice(2, -2), options: { ...baseOpts, bold: true, color: COLORS.white } };
    }
    return { text: part, options: { ...baseOpts } };
  });
}

// Transforme le texte markdown-lite de Claude en blocs typés
function parseToBlocks(raw, sectionTitle) {
  if (!raw) return [];
  let text = raw.trim();

  // Retire le titre en doublon en tête de texte (ex: "# 1. Définition...")
  const lines = text.split("\n");
  if (lines.length && /^#{1,2}\s+/.test(lines[0])) {
    const firstLineClean = lines[0].replace(/^#{1,2}\s+/, "").trim().toLowerCase();
    const titleClean = (sectionTitle || "").replace(/^\d+\.\s*/, "").trim().toLowerCase();
    if (firstLineClean.includes(titleClean.slice(0, 15)) || titleClean.includes(firstLineClean.slice(0, 15))) {
      lines.shift();
    }
  }
  text = lines.join("\n").trim();

  const blocks = [];
  const rawLines = text.split("\n").map((l) => l.trim()).filter((l) => l.length > 0);

  for (const line of rawLines) {
    if (/^###\s+/.test(line)) {
      blocks.push({ type: "h3", text: line.replace(/^###\s+/, "") });
    } else if (/^##\s+/.test(line)) {
      blocks.push({ type: "h2", text: line.replace(/^##\s+/, "") });
    } else if (/^#\s+/.test(line)) {
      blocks.push({ type: "h2", text: line.replace(/^#\s+/, "") });
    } else if (/^[-•]\s+/.test(line)) {
      blocks.push({ type: "bullet", text: line.replace(/^[-•]\s+/, "") });
    } else {
      blocks.push({ type: "p", text: line });
    }
  }
  return blocks;
}

// Estime le "poids" en lignes d'un bloc pour la pagination
function blockWeight(block) {
  const charsPerLine = 95;
  if (block.type === "h2") return 2.5;
  if (block.type === "h3") return 2;
  const lines = Math.ceil(block.text.length / charsPerLine);
  return lines + (block.type === "p" ? 0.6 : 0.3);
}

// Découpe une liste de blocs en pages (~22 lignes de budget par slide)
function paginateBlocks(blocks, budget = 21) {
  const pages = [];
  let current = [];
  let total = 0;
  for (const block of blocks) {
    const w = blockWeight(block);
    if (total + w > budget && current.length > 0) {
      pages.push(current);
      current = [];
      total = 0;
    }
    current.push(block);
    total += w;
  }
  if (current.length > 0) pages.push(current);
  return pages.length ? pages : [[]];
}

// Construit le tableau de texte enrichi (runs) pour pptxgenjs à partir des blocs d'une page
function blocksToTextRuns(blocks) {
  const runs = [];
  blocks.forEach((block, idx) => {
    const isFirst = idx === 0;
    if (block.type === "h2") {
      runs.push({
        text: block.text,
        options: {
          bold: true,
          color: COLORS.purple,
          fontSize: 15,
          breakLine: true,
          paraSpaceBefore: isFirst ? 0 : 14,
          paraSpaceAfter: 4
        }
      });
    } else if (block.type === "h3") {
      runs.push({
        text: block.text,
        options: {
          bold: true,
          color: COLORS.white,
          fontSize: 12.5,
          breakLine: true,
          paraSpaceBefore: isFirst ? 0 : 10,
          paraSpaceAfter: 3
        }
      });
    } else if (block.type === "bullet") {
      const r = boldRuns(block.text, { fontSize: 11.5, color: COLORS.textMuted });
      r[0].text = "•  " + r[0].text;
      r[0].options.paraSpaceBefore = isFirst ? 0 : 4;
      r[r.length - 1].options.breakLine = true;
      runs.push(...r);
    } else {
      const r = boldRuns(block.text, { fontSize: 11.5, color: COLORS.textMuted });
      r[0].options.paraSpaceBefore = isFirst ? 0 : 8;
      r[r.length - 1].options.breakLine = true;
      runs.push(...r);
    }
  });
  return runs;
}

// ---------- Construction des slides ----------

function addFooter(slide, pageNumber, footerLabel) {
  slide.addShape("line", {
    x: 0.5, y: 7.05, w: W - 1, h: 0,
    line: { color: COLORS.cardBorder, width: 0.75 }
  });
  slide.addText(footerLabel || "SCOPE — Étude générée automatiquement", {
    x: 0.5, y: 7.12, w: 8, h: 0.3,
    fontSize: 9, color: COLORS.textFaint, fontFace: FONT
  });
  slide.addText(String(pageNumber).padStart(2, "0"), {
    x: W - 1.3, y: 7.12, w: 0.8, h: 0.3,
    fontSize: 9, color: COLORS.textFaint, fontFace: FONT, align: "right"
  });
}

function addCoverSlide(pres, { titre, sousTitre, kicker }) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.bg };

  // Cercle décoratif
  slide.addShape("ellipse", {
    x: W - 4.2, y: -2.2, w: 6.5, h: 6.5,
    fill: { color: COLORS.purpleCircle }, line: { type: "none" }
  });

  slide.addText((kicker || "ÉTUDE GÉNÉRÉE PAR SCOPE").toUpperCase(), {
    x: 0.75, y: 1.15, w: 10, h: 0.4,
    fontSize: 12, bold: true, color: COLORS.purple, fontFace: FONT, charSpacing: 2
  });

  slide.addText(titre || "Étude", {
    x: 0.75, y: 1.65, w: 10.5, h: 1.6,
    fontSize: 34, bold: true, color: COLORS.white, fontFace: FONT, valign: "top"
  });

  slide.addText(sousTitre || "Généré par Scope", {
    x: 0.75, y: 3.15, w: 10, h: 0.5,
    fontSize: 13, color: COLORS.textMuted, fontFace: FONT
  });

  slide.addShape("line", {
    x: 0.75, y: 4.1, w: W - 1.5, h: 0,
    line: { color: COLORS.cardBorder, width: 0.75 }
  });

  const dateStr = new Date().toLocaleDateString("fr-FR", { year: "numeric", month: "long" });
  slide.addText("DATE DE GÉNÉRATION", { x: 0.75, y: 4.35, w: 4, h: 0.3, fontSize: 9, color: COLORS.textFaint, fontFace: FONT, charSpacing: 1 });
  slide.addText(dateStr, { x: 0.75, y: 4.65, w: 4, h: 0.35, fontSize: 13, bold: true, color: COLORS.white, fontFace: FONT });
}

function addSommaireSlide(pres, sections) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.bg };

  slide.addText("Sommaire", {
    x: 0.75, y: 0.5, w: 8, h: 0.8,
    fontSize: 28, italic: true, color: COLORS.white, fontFace: FONT
  });

  const startY = 1.55;
  const rowH = 0.62;
  const maxRows = Math.min(sections.length, 9);
  const colGap = 0.18;

  sections.slice(0, maxRows).forEach((s, i) => {
    const y = startY + i * rowH;
    slide.addShape("roundRect", {
      x: 0.75, y, w: W - 1.5, h: rowH - colGap,
      rectRadius: 0.06,
      fill: { color: COLORS.card }, line: { color: COLORS.cardBorder, width: 0.75 }
    });
    slide.addText(String(i + 1).padStart(2, "0"), {
      x: 1.0, y, w: 0.8, h: rowH - colGap,
      fontSize: 15, bold: true, color: COLORS.purple, fontFace: FONT, valign: "middle"
    });
    const cleanTitle = (s.titre || "").replace(/^\d+\.\s*/, "");
    slide.addText(cleanTitle, {
      x: 1.75, y: y + 0.04, w: W - 3, h: (rowH - colGap) / 2 + 0.05,
      fontSize: 13, bold: true, color: COLORS.white, fontFace: FONT, valign: "bottom"
    });
  });

  addFooter(slide, 2);
}

function addSectionSlides(pres, section, globalIndexLabel) {
  const blocks = parseToBlocks(section.texte, section.titre);
  const pages = paginateBlocks(blocks);

  pages.forEach((pageBlocks, pageIdx) => {
    const slide = pres.addSlide();
    slide.background = { color: COLORS.bg };

    const cleanTitle = (section.titre || "Section").replace(/^\d+\.\s*/, "");
    const suffix = pages.length > 1 ? ` (suite ${pageIdx + 1}/${pages.length})` : "";

    slide.addText(cleanTitle + suffix, {
      x: 0.75, y: 0.45, w: W - 1.5, h: 0.6,
      fontSize: 22, bold: true, color: COLORS.white, fontFace: FONT
    });

    slide.addShape("roundRect", {
      x: 0.75, y: 1.25, w: W - 1.5, h: 5.55,
      rectRadius: 0.06,
      fill: { color: COLORS.card }, line: { color: COLORS.cardBorder, width: 0.75 }
    });

    const runs = blocksToTextRuns(pageBlocks);
    if (runs.length > 0) {
      slide.addText(runs, {
        x: 1.05, y: 1.5, w: W - 2.1, h: 5.05,
        valign: "top", fontFace: FONT, lineSpacingMultiple: 1.15
      });
    } else {
      slide.addText("Donnée non disponible.", {
        x: 1.05, y: 1.5, w: W - 2.1, h: 1,
        fontSize: 12, italic: true, color: COLORS.textFaint, fontFace: FONT
      });
    }

    addFooter(slide, globalIndexLabel + pageIdx);
  });

  return pages.length;
}

function addSourcesSlide(pres, sources, pageNumber) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.bg };

  slide.addText("Sources", {
    x: 0.75, y: 0.5, w: 8, h: 0.7,
    fontSize: 24, bold: true, color: COLORS.white, fontFace: FONT
  });

  slide.addShape("line", {
    x: 0.75, y: 1.25, w: W - 1.5, h: 0,
    line: { color: COLORS.cardBorder, width: 0.75 }
  });

  slide.addText("MÉTHODOLOGIE & TRAÇABILITÉ", {
    x: 0.75, y: 1.45, w: 8, h: 0.3,
    fontSize: 10, bold: true, color: COLORS.purple, fontFace: FONT, charSpacing: 1.5
  });

  const list = Array.isArray(sources) ? sources.slice(0, 14) : [];
  const runs = [];
  list.forEach((s, i) => {
    runs.push({
      text: `${s.titre || "Source"} `,
      options: { bold: true, color: COLORS.white, fontSize: 10.5, breakLine: false, paraSpaceBefore: i === 0 ? 0 : 7 }
    });
    runs.push({
      text: `— ${s.url || ""}`,
      options: { color: COLORS.textMuted, fontSize: 9.5, breakLine: true, hyperlink: s.url ? { url: s.url } : undefined }
    });
  });

  if (runs.length > 0) {
    slide.addText(runs, {
      x: 0.75, y: 1.9, w: W - 1.5, h: 4.6,
      valign: "top", fontFace: FONT, lineSpacingMultiple: 1.1
    });
  }

  slide.addText("Généré automatiquement par Scope à partir de sources web vérifiées. Les chiffres non trouvés dans les sources sont explicitement signalés comme non disponibles plutôt qu'estimés.", {
    x: 0.75, y: 6.5, w: W - 1.5, h: 0.4,
    fontSize: 8.5, italic: true, color: COLORS.textFaint, fontFace: FONT
  });

  addFooter(slide, pageNumber);
}

// ---------- Route principale ----------

app.post("/generate-pptx", async (req, res) => {
  try {
    const { sections, sources, titre, sousTitre, kicker } = req.body;

    if (!Array.isArray(sections)) {
      return res.status(400).json({ error: "Le champ 'sections' est requis et doit être un tableau." });
    }

    const pres = new pptxgen();
    pres.defineLayout({ name: "SCOPE_16x9", width: W, height: H });
    pres.layout = "SCOPE_16x9";

    addCoverSlide(pres, { titre, sousTitre, kicker });
    addSommaireSlide(pres, sections);

    let pageCounter = 3;
    sections.forEach((s) => {
      const nbPages = addSectionSlides(pres, s, pageCounter);
      pageCounter += nbPages;
    });

    addSourcesSlide(pres, sources, pageCounter);

    const buffer = await pres.write({ outputType: "nodebuffer" });

    res.set({
      "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "Content-Disposition": `attachment; filename="etude-${Date.now()}.pptx"`
    });
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message, stack: err.stack });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Service PPTX Scope v2 démarré sur le port ${PORT}`);
});
