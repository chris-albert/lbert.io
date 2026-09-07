// Résumé exporters. Every format is generated from the same RESUME object the
// page renders, so the downloads can never drift from what the site shows.
//
// Markdown needs no dependencies. DOCX (docx) and PDF (jsPDF) are fetched from a
// CDN the first time they are requested.
(function () {
  const DOCX_URL = "https://cdn.jsdelivr.net/npm/docx@9.5.1/dist/index.iife.js";
  const JSPDF_URL = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/4.2.1/jspdf.umd.min.js";

  const ACCENT = "B06020";
  const MUTED = "6A6A6A";
  const RULE = "D0D0D0";

  const hostOf = url => url.replace(/^https?:\/\//, "").replace(/\/$/, "");
  const slug = name => name.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") + "-resume";

  // ---------------------------------------------------------------- Markdown
  function markdown(r) {
    const out = [];
    out.push("# " + r.name, "", r.tagline, "");
    out.push([
      "[" + r.email + "](mailto:" + r.email + ")",
      "[" + hostOf(r.github) + "](" + r.github + ")",
      "[" + hostOf(r.site) + "](" + r.site + ")"
    ].join(" · "), "");

    out.push("## Experience", "");
    for (const j of r.roles) {
      out.push("### " + j.role + " — " + j.company, "", "*" + j.dates + "*", "");
      for (const b of j.bullets) out.push("- " + b);
      out.push("", "_" + j.tags.join(" · ") + "_", "");
    }

    out.push("## Skills", "");
    for (const s of r.stack) out.push("- **" + s.label + ":** " + s.items);
    out.push("");

    out.push("## Education", "");
    for (const e of r.education) out.push("- " + (e.url ? "[" + e.label + "](" + e.url + ")" : e.label));
    out.push("");

    return out.join("\n");
  }

  // -------------------------------------------------------------------- DOCX
  function docxDocument(docx, r) {
    const { Document, Paragraph, TextRun, ExternalHyperlink, LevelFormat, AlignmentType, BorderStyle, TabStopType } = docx;
    // US Letter in twips (1/1440 in), 0.75in margins.
    const PAGE_WIDTH = 12240, MARGIN = 1080, TEXT_WIDTH = PAGE_WIDTH - MARGIN * 2;

    const link = (text, url) => new ExternalHyperlink({ children: [new TextRun({ text, style: "Hyperlink" })], link: url });
    const bullet = children => new Paragraph({ children, numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 } });
    const section = title => new Paragraph({ text: title, style: "Section" });

    const children = [
      new Paragraph({ text: r.name, style: "Name" }),
      new Paragraph({ children: [new TextRun({ text: r.tagline, color: MUTED, size: 22 })], spacing: { after: 120 } }),
      new Paragraph({
        children: [
          link(r.email, "mailto:" + r.email), new TextRun("   ·   "),
          link(hostOf(r.github), r.github), new TextRun("   ·   "),
          link(hostOf(r.site), r.site)
        ]
      }),
      section("Experience")
    ];

    for (const j of r.roles) {
      children.push(new Paragraph({
        children: [
          new TextRun({ text: j.role, bold: true, size: 23 }),
          new TextRun({ text: "   " + j.company, color: ACCENT }),
          new TextRun({ text: "\t" + j.dates, color: MUTED })
        ],
        tabStops: [{ type: TabStopType.RIGHT, position: TEXT_WIDTH }],
        spacing: { before: 220, after: 80 }
      }));
      for (const b of j.bullets) children.push(bullet([new TextRun(b)]));
      children.push(new Paragraph({
        children: [new TextRun({ text: j.tags.join(" · "), italics: true, color: MUTED, size: 18 })],
        spacing: { before: 40 }
      }));
    }

    children.push(section("Skills"));
    for (const s of r.stack) {
      children.push(new Paragraph({
        children: [new TextRun({ text: s.label + ": ", bold: true }), new TextRun(s.items)],
        spacing: { after: 60 }
      }));
    }

    children.push(section("Education"));
    for (const e of r.education) children.push(bullet([e.url ? link(e.label, e.url) : new TextRun(e.label)]));

    return new Document({
      creator: r.name,
      title: r.name + " — Résumé",
      styles: {
        default: { document: { run: { font: "Calibri", size: 21 } } },
        paragraphStyles: [
          { id: "Name", name: "Name", basedOn: "Normal", run: { size: 44, bold: true }, paragraph: { spacing: { after: 80 } } },
          {
            id: "Section", name: "Section", basedOn: "Normal",
            run: { size: 19, bold: true, color: ACCENT, allCaps: true, characterSpacing: 30 },
            paragraph: { spacing: { before: 320, after: 120 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 4 } } }
          }
        ]
      },
      numbering: {
        config: [{
          reference: "bullets",
          levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 220 } } } }]
        }]
      },
      sections: [{ properties: { page: { size: { width: PAGE_WIDTH, height: 15840 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } }, children }]
    });
  }

  // --------------------------------------------------------------------- PDF
  function pdfDocument(jsPDF, r) {
    const doc = new jsPDF({ unit: "pt", format: "letter" });
    const W = doc.internal.pageSize.getWidth();
    const H = doc.internal.pageSize.getHeight();
    const M = 54;
    const maxW = W - M * 2;
    const LH = 1.35;
    const hex = h => [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16));
    const ink = [34, 34, 34], muted = hex(MUTED), accent = hex(ACCENT), rule = hex(RULE);
    let y = M;

    const ensure = h => { if (y + h > H - M) { doc.addPage(); y = M; } };
    const font = (size, style, color) => { doc.setFont("helvetica", style || "normal"); doc.setFontSize(size); doc.setTextColor(...(color || ink)); };
    // Wrapped text block at (x, y); advances y by its height.
    const block = (str, size, o = {}) => {
      font(size, o.style, o.color);
      const lines = doc.splitTextToSize(str, o.width || maxW);
      const h = lines.length * size * LH;
      ensure(h);
      doc.text(lines, o.x || M, y, { baseline: "top", lineHeightFactor: LH });
      y += h;
      return h;
    };
    const section = title => {
      ensure(60);
      y += 22;
      block(title.toUpperCase(), 9, { style: "bold", color: accent });
      doc.setDrawColor(...rule);
      doc.setLineWidth(0.6);
      doc.line(M, y + 2, W - M, y + 2);
      y += 14;
    };

    // Header
    block(r.name, 24, { style: "bold" });
    y += 4;
    block(r.tagline, 11, { color: muted });
    y += 8;
    font(9.5, "normal", accent);
    let x = M;
    [[r.email, "mailto:" + r.email], [hostOf(r.github), r.github], [hostOf(r.site), r.site]].forEach(([label, url], i) => {
      if (i) { doc.setTextColor(...muted); doc.text("·", x + 6, y, { baseline: "top" }); x += 18; doc.setTextColor(...accent); }
      doc.text(label, x, y, { baseline: "top" });
      const w = doc.getTextWidth(label);
      doc.link(x, y, w, 11, { url });
      x += w;
    });
    y += 9.5 * LH;

    // Experience: dates in a left column, details on the right (as on the site).
    section("Experience");
    const col = 96;
    r.roles.forEach((j, i) => {
      if (i) {
        y += 10;
        ensure(50);
        doc.setDrawColor(...rule);
        doc.setLineWidth(0.4);
        doc.line(M, y, W - M, y);
        y += 12;
      } else {
        ensure(50);
      }
      const top = y;
      block(j.role, 11.5, { style: "bold", x: M + col, width: maxW - col });
      y += 1;
      block(j.company, 9.5, { color: accent, x: M + col, width: maxW - col });
      font(9, "normal", muted);
      doc.text(j.dates, M, top, { baseline: "top" });
      y += 6;
      for (const b of j.bullets) {
        font(10, "normal", muted);
        const lines = doc.splitTextToSize(b, maxW - col - 12);
        ensure(lines.length * 10 * LH);
        doc.text("•", M + col, y, { baseline: "top" });
        doc.text(lines, M + col + 12, y, { baseline: "top", lineHeightFactor: LH });
        y += lines.length * 10 * LH + 2;
      }
      y += 4;
      block(j.tags.join("  ·  "), 8.5, { color: muted, x: M + col, width: maxW - col });
    });

    section("Skills");
    for (const s of r.stack) {
      ensure(14);
      font(10, "bold", ink);
      doc.text(s.label, M, y, { baseline: "top" });
      block(s.items, 10, { x: M + col, width: maxW - col, color: muted });
      y += 4;
    }

    section("Education");
    for (const e of r.education) {
      font(10, "normal", e.url ? accent : muted);
      const lines = doc.splitTextToSize(e.label, maxW - 12);
      ensure(lines.length * 10 * LH);
      doc.setTextColor(...muted);
      doc.text("•", M, y, { baseline: "top" });
      doc.setTextColor(...(e.url ? accent : muted));
      doc.text(lines, M + 12, y, { baseline: "top", lineHeightFactor: LH });
      if (e.url) doc.link(M + 12, y, doc.getTextWidth(lines[0]), 11, { url: e.url });
      y += lines.length * 10 * LH + 3;
    }

    return doc;
  }

  // ----------------------------------------------------------------- browser
  const loadScript = (src, isReady) => new Promise((resolve, reject) => {
    if (isReady()) return resolve();
    const s = document.createElement("script");
    s.src = src;
    s.onload = () => (isReady() ? resolve() : reject(new Error("Library did not initialise: " + src)));
    s.onerror = () => reject(new Error("Failed to load " + src));
    document.head.appendChild(s);
  });

  const save = (blob, filename) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  async function download(format, r) {
    const base = slug(r.name);
    if (format === "md") {
      save(new Blob([markdown(r)], { type: "text/markdown;charset=utf-8" }), base + ".md");
    } else if (format === "docx") {
      await loadScript(DOCX_URL, () => typeof window.docx !== "undefined");
      save(await window.docx.Packer.toBlob(docxDocument(window.docx, r)), base + ".docx");
    } else if (format === "pdf") {
      await loadScript(JSPDF_URL, () => !!(window.jspdf && window.jspdf.jsPDF));
      save(pdfDocument(window.jspdf.jsPDF, r).output("blob"), base + ".pdf");
    } else {
      throw new Error("Unknown format: " + format);
    }
  }

  const api = { markdown, docxDocument, pdfDocument, download, slug };
  if (typeof window !== "undefined") window.ResumeExport = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})();
