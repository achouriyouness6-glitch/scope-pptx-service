const express = require("express");
const pptxgen = require("pptxgenjs");

const app = express();

// Augmente la limite car les sources (extraits d'articles) peuvent être volumineuses
app.use(express.json({ limit: "10mb" }));

app.get("/", (req, res) => {
  res.send("Service PPTX Scope — en ligne. POST /generate-pptx pour générer un fichier.");
});

app.post("/generate-pptx", async (req, res) => {
  try {
    const { sections, sources, titre } = req.body;

    if (!Array.isArray(sections)) {
      return res.status(400).json({ error: "Le champ 'sections' est requis et doit être un tableau." });
    }

    const pres = new pptxgen();

    // Slide de titre
    let titleSlide = pres.addSlide();
    titleSlide.addText(titre || "Étude générée par Scope", {
      x: 0.5,
      y: 2,
      w: 9,
      fontSize: 28,
      bold: true,
      align: "center"
    });

    // Une slide par section
    sections.forEach((s) => {
      const sl = pres.addSlide();
      sl.addText(s.titre || "Section sans titre", {
        x: 0.4,
        y: 0.3,
        w: 9.2,
        fontSize: 20,
        bold: true
      });
      sl.addText(s.texte || "Contenu non disponible", {
        x: 0.4,
        y: 1,
        w: 9.2,
        h: 5.5,
        fontSize: 11,
        valign: "top"
      });
    });

    // Slide finale : sources
    if (Array.isArray(sources) && sources.length > 0) {
      const sourceSlide = pres.addSlide();
      sourceSlide.addText("Sources", {
        x: 0.4,
        y: 0.3,
        w: 9.2,
        fontSize: 20,
        bold: true
      });
      const sourceLines = sources
        .map((s) => `• ${s.titre || "Source"} — ${s.url || ""}`)
        .join("\n");
      sourceSlide.addText(sourceLines, {
        x: 0.4,
        y: 1,
        w: 9.2,
        h: 5.5,
        fontSize: 9,
        valign: "top"
      });
    }

    const buffer = await pres.write({ outputType: "nodebuffer" });

    res.set({
      "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "Content-Disposition": `attachment; filename="etude-${Date.now()}.pptx"`
    });
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Service PPTX démarré sur le port ${PORT}`);
});
