// Cordoba Research Group - Research Documentation Tool
// Restores original behaviours (blank-line preservation, co-author required once added,
// reset prompt) while adding preview, sources, doc ID, PDF export, bullet helper, house style toggles.

let coAuthorCount = 0;
let sourceCount = 0;

// ---------------------------
// Utilities
// ---------------------------
function pad2(n) { return String(n).padStart(2, "0"); }
function pad3(n) { return String(n).padStart(3, "0"); }
function safeText(v) { return (v || "").toString(); }

function escapeHtml(str) {
  return (str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function formatDateTime(date) {
  const months = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];
  const month = months[date.getMonth()];
  const day = date.getDate();
  const year = date.getFullYear();

  let hours = date.getHours();
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const ampm = hours >= 12 ? "PM" : "AM";
  hours = hours % 12 || 12;

  return `${month} ${day}, ${year} ${hours}:${minutes} ${ampm}`;
}

function noteTypeCode(noteType) {
  const map = {
    "General Note": "GEN",
    "Equity Research": "EQ",
    "Macro Research": "MACRO",
    "Fixed Income Research": "FI",
    "Commodity Insights": "COM"
  };
  return map[noteType] || "GEN";
}

// ---------------------------
// Document ID (local sequence)
// ---------------------------
function getDocSequenceKey(noteType, yyyy, mm) {
  const code = noteTypeCode(noteType);
  return `crg_seq_${code}_${yyyy}_${mm}`;
}

function incrementAndGetSeq(noteType, yyyy, mm) {
  const key = getDocSequenceKey(noteType, yyyy, mm);
  const current = parseInt(localStorage.getItem(key) || "0", 10);
  const next = current + 1;
  localStorage.setItem(key, String(next));
  return next;
}

function generateDocumentId(noteType, dateObj, forceNew = false) {
  const code = noteTypeCode(noteType);
  const yyyy = dateObj.getFullYear();
  const mm = pad2(dateObj.getMonth() + 1);

  const sessionKey = "crg_docid_session";
  if (!forceNew) {
    const existing = localStorage.getItem(sessionKey);
    if (existing && existing.startsWith(`CRG-${code}-${yyyy}-${mm}-`)) return existing;
  }

  const seq = incrementAndGetSeq(noteType, yyyy, mm);
  const docId = `CRG-${code}-${yyyy}-${mm}-${pad3(seq)}`;
  localStorage.setItem(sessionKey, docId);
  return docId;
}

function clearSessionDocId() {
  localStorage.removeItem("crg_docid_session");
}

function updateDocIdFromType(forceNew = false) {
  const noteType = document.getElementById("noteType").value;
  const docIdInput = document.getElementById("documentId");
  if (!noteType) {
    docIdInput.value = "";
    clearSessionDocId();
    renderPreview();
    return;
  }
  const now = new Date();
  docIdInput.value = generateDocumentId(noteType, now, forceNew);
  renderPreview();
}

// ---------------------------
// Co-authors (original behaviour preserved)
// If you add a row, inputs are REQUIRED (same as original).
// ---------------------------
function addCoAuthorRow() {
  coAuthorCount++;
  const list = document.getElementById("coAuthorsList");

  const row = document.createElement("div");
  row.className = "row-card";
  row.id = `coauthor-${coAuthorCount}`;
  row.innerHTML = `
    <div class="row-grid-coauthor">
      <div>
        <label class="mini-label">Last Name</label>
        <input type="text" class="coauthor-lastname" placeholder="Last Name" required>
      </div>
      <div>
        <label class="mini-label">First Name</label>
        <input type="text" class="coauthor-firstname" placeholder="First Name" required>
      </div>
      <div>
        <label class="mini-label">Phone</label>
        <input type="text" class="coauthor-phone" placeholder="e.g., 44-7398344190" required>
      </div>
      <button type="button" class="btn btn-ghost btn-sm" onclick="removeCoAuthor(${coAuthorCount})">Remove</button>
    </div>
  `;
  list.appendChild(row);
}

function removeCoAuthor(id) {
  const el = document.getElementById(`coauthor-${id}`);
  if (el) el.remove();
  renderPreview();
}

document.getElementById("addCoAuthor").addEventListener("click", () => {
  addCoAuthorRow();
  renderPreview();
});

// ---------------------------
// Sources (optional)
// ---------------------------
function addSourceRow() {
  sourceCount++;
  const list = document.getElementById("sourcesList");

  const row = document.createElement("div");
  row.className = "row-card";
  row.id = `source-${sourceCount}`;
  row.innerHTML = `
    <div class="row-grid-source">
      <div>
        <label class="mini-label">Source name</label>
        <input type="text" class="src-name" placeholder="e.g., IMF WEO">
      </div>
      <div>
        <label class="mini-label">URL</label>
        <input type="text" class="src-url" placeholder="https://...">
      </div>
      <div>
        <label class="mini-label">Date accessed</label>
        <input type="text" class="src-date" placeholder="YYYY-MM-DD">
      </div>
      <div>
        <label class="mini-label">Key line (optional)</label>
        <input type="text" class="src-line" placeholder="Short note on what matters...">
      </div>
      <button type="button" class="btn btn-ghost btn-sm" onclick="removeSource(${sourceCount})">Remove</button>
    </div>
  `;
  list.appendChild(row);
}

function removeSource(id) {
  const el = document.getElementById(`source-${id}`);
  if (el) el.remove();
  renderPreview();
}

document.getElementById("addSource").addEventListener("click", () => {
  addSourceRow();
  renderPreview();
});

// mini label css injection
const miniStyle = document.createElement("style");
miniStyle.textContent = `
  .mini-label{
    display:block;
    font-size:12px;
    color: rgba(15,23,42,.62);
    font-weight:600;
    margin-bottom:6px;
  }
`;
document.head.appendChild(miniStyle);

// ---------------------------
// Bullet helper (does NOT remove blank lines)
// Original Word behaviour preserved: blank lines remain blank
// ---------------------------
function convertToBulletsInTextarea() {
  const el = document.getElementById("keyTakeaways");
  const raw = safeText(el.value).replace(/\r\n/g, "\n");
  const lines = raw.split("\n").map(l => l.replace(/^[-*•]\s*/, ""));

  // Keep empties exactly
  el.value = lines
    .map(l => (l.trim() === "" ? "" : `- ${l.trim()}`))
    .join("\n");
}

document.getElementById("convertBullets").addEventListener("click", () => {
  convertToBulletsInTextarea();
  renderPreview();
});

// ---------------------------
// Smart paragraphing that still preserves blank lines
// - For preview and Word: split on DOUBLE blank lines, but keep explicit blank paragraphs.
// - Within each paragraph block, single line breaks are kept as-is in Word mode?:
//   We'll preserve author intent by keeping single lines as separate paragraphs ONLY if they entered them.
//   (We do this by preserving line breaks inside blocks as separate paragraphs.)
// ---------------------------
function splitPreservingBlankLines(text) {
  // This matches original: every line becomes its own paragraph; blank lines become blank paragraphs.
  const raw = safeText(text).replace(/\r\n/g, "\n");
  return raw.split("\n"); // includes empty strings
}

function splitIntoParagraphBlocks(text) {
  // This matches the requested "smart paragraphing" but preserves blank lines:
  // - double line breaks create visual separation, but we keep empty lines.
  // Implementation: keep original line-splitting, but treat consecutive blanks as paragraph separators.
  return splitPreservingBlankLines(text);
}

// ---------------------------
// Collect data
// ---------------------------
function collectData() {
  const noteType = document.getElementById("noteType").value;
  const title = document.getElementById("title").value;
  const topic = document.getElementById("topic").value;

  const authorLastName = document.getElementById("authorLastName").value;
  const authorFirstName = document.getElementById("authorFirstName").value;
  const authorPhone = document.getElementById("authorPhone").value;

  const keyTakeaways = document.getElementById("keyTakeaways").value;
  const analysis = document.getElementById("analysis").value;
  const content = document.getElementById("content").value;
  const cordobaView = document.getElementById("cordobaView").value;

  const imageFiles = document.getElementById("imageUpload").files;

  const tightSpacing = document.getElementById("tightSpacing").checked;
  const smallCaptions = document.getElementById("smallCaptions").checked;

  const documentId = document.getElementById("documentId").value;

  const now = new Date();
  const dateTimeString = formatDateTime(now);

  // Co-authors (only include complete rows)
  const coAuthors = [];
  document.querySelectorAll("#coAuthorsList .row-card").forEach(row => {
    const lastName = safeText(row.querySelector(".coauthor-lastname")?.value).trim();
    const firstName = safeText(row.querySelector(".coauthor-firstname")?.value).trim();
    const phone = safeText(row.querySelector(".coauthor-phone")?.value).trim();
    if (lastName && firstName && phone) coAuthors.push({ lastName, firstName, phone });
  });

  // Sources (include if any field present)
  const sources = [];
  document.querySelectorAll("#sourcesList .row-card").forEach(row => {
    const name = safeText(row.querySelector(".src-name")?.value).trim();
    const url = safeText(row.querySelector(".src-url")?.value).trim();
    const date = safeText(row.querySelector(".src-date")?.value).trim();
    const line = safeText(row.querySelector(".src-line")?.value).trim();
    if (name || url || date || line) sources.push({ name, url, date, line });
  });

  return {
    noteType, title, topic,
    authorLastName, authorFirstName, authorPhone,
    coAuthors,
    keyTakeaways, analysis, content, cordobaView,
    sources,
    imageFiles,
    dateTimeString,
    documentId,
    tightSpacing,
    smallCaptions
  };
}

// ---------------------------
// Live Preview (mirrors Word structure)
// ---------------------------
function renderPreview() {
  const d = collectData();

  const headerLine = `Cordoba Research Group | ${d.noteType || "—"} | ${d.dateTimeString || "—"}${d.documentId ? " | " + d.documentId : ""}`;
  document.getElementById("previewHeaderLine").textContent = headerLine;

  const takeawaysLines = splitPreservingBlankLines(d.keyTakeaways)
    .map(l => l.replace(/^[-*•]\s*/, "").trim());

  const analysisLines = splitIntoParagraphBlocks(d.analysis);
  const contentLines = splitIntoParagraphBlocks(d.content);
  const cordobaLines = splitIntoParagraphBlocks(d.cordobaView);

  const figures = Array.from(d.imageFiles || []).map((f, i) => {
    const caption = (f.name || "").replace(/\.[^/.]+$/, "");
    return `Figure ${i + 1}: ${caption}`;
  });

  const primaryAuthor = `${(d.authorLastName || "").toUpperCase()}, ${(d.authorFirstName || "").toUpperCase()}${d.authorPhone ? ` (${d.authorPhone})` : ""}`;
  const coAuthorsText = d.coAuthors.map(a => `${(a.lastName || "").toUpperCase()}, ${(a.firstName || "").toUpperCase()}${a.phone ? ` (${a.phone})` : ""}`).join(" · ");

  const preview = document.getElementById("preview");
  preview.innerHTML = `
    <div class="preview-card">
      <div class="preview-head">
        <div class="preview-title">${escapeHtml(d.title || "—")}</div>
        <div class="preview-meta">
          <div><strong>TOPIC:</strong> ${escapeHtml(d.topic || "—")}</div>
          <div><strong>PRIMARY:</strong> ${escapeHtml(primaryAuthor || "—")}</div>
          <div><strong>DOC ID:</strong> ${escapeHtml(d.documentId || "—")}</div>
        </div>
        ${coAuthorsText ? `<div class="preview-meta" style="margin-top:8px;"><div><strong>CO-AUTHORS:</strong> ${escapeHtml(coAuthorsText)}</div></div>` : ""}
      </div>

      <div class="preview-body">
        <div class="preview-h">Key Takeaways</div>
        ${
          takeawaysLines.some(x => x.length > 0)
          ? `<ul class="preview-ul">
              ${takeawaysLines.map(t => t === "" ? `<li style="list-style:none;margin-left:-18px;height:8px;"></li>` : `<li>${escapeHtml(t)}</li>`).join("")}
            </ul>`
          : `<div class="preview-p muted">—</div>`
        }

        <div class="preview-h">Analysis and Commentary</div>
        ${analysisLines.length ? analysisLines.map(l => l.trim() === "" ? `<div style="height:10px"></div>` : `<div class="preview-p">${escapeHtml(l)}</div>`).join("") : `<div class="preview-p muted">—</div>`}

        ${contentLines.some(l => l.trim() !== "") ? `<div class="preview-h">Additional Content</div>${contentLines.map(l => l.trim() === "" ? `<div style="height:10px"></div>` : `<div class="preview-p">${escapeHtml(l)}</div>`).join("")}` : ""}

        ${cordobaLines.some(l => l.trim() !== "") ? `<div class="preview-h">The Cordoba View</div>${cordobaLines.map(l => l.trim() === "" ? `<div style="height:10px"></div>` : `<div class="preview-p">${escapeHtml(l)}</div>`).join("")}` : ""}

        ${figures.length ? `<div class="preview-h">Figures and Charts</div><ul class="preview-ul">${figures.map(f => `<li>${escapeHtml(f)}</li>`).join("")}</ul>` : ""}

        ${
          d.sources.length
          ? `<div class="preview-h">Sources</div>
             <ol class="preview-ul" style="margin-left:22px;">
              ${d.sources.map(s => {
                const parts = [];
                if (s.name) parts.push(`<strong>${escapeHtml(s.name)}</strong>`);
                if (s.url) parts.push(`<span class="mono">${escapeHtml(s.url)}</span>`);
                if (s.date) parts.push(`(Accessed ${escapeHtml(s.date)})`);
                const head = parts.join(" — ");
                const line = s.line ? `<div class="preview-p" style="margin:6px 0 0;color:rgba(15,23,42,.78)">${escapeHtml(s.line)}</div>` : "";
                return `<li style="margin:0 0 10px">${head}${line}</li>`;
              }).join("")}
             </ol>`
          : ""
        }
      </div>
    </div>
  `;
}

// watch changes
[
  "noteType","title","topic","authorLastName","authorFirstName","authorPhone",
  "keyTakeaways","analysis","content","cordobaView",
  "tightSpacing","smallCaptions","imageUpload"
].forEach(id => {
  const el = document.getElementById(id);
  if (!el) return;
  el.addEventListener("input", renderPreview);
  el.addEventListener("change", renderPreview);
});
document.getElementById("coAuthorsList").addEventListener("input", renderPreview);
document.getElementById("sourcesList").addEventListener("input", renderPreview);

// note type triggers doc id generation
document.getElementById("noteType").addEventListener("change", () => updateDocIdFromType(true));
document.getElementById("regenDocId").addEventListener("click", () => updateDocIdFromType(true));

// ---------------------------
// Images (same as original)
// ---------------------------
async function addImages(files, captionFontHalfPoints) {
  const imageParagraphs = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const arrayBuffer = await file.arrayBuffer();
    const fileNameWithoutExt = file.name.replace(/\.[^/.]+$/, "");

    imageParagraphs.push(
      new docx.Paragraph({
        children: [
          new docx.ImageRun({
            data: arrayBuffer,
            transformation: { width: 600, height: 450 }
          })
        ],
        spacing: { before: 200, after: 100 },
        alignment: docx.AlignmentType.CENTER
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: `Figure ${i + 1}: ${fileNameWithoutExt}`,
            italics: true,
            size: captionFontHalfPoints,
            font: "Book Antiqua"
          })
        ],
        spacing: { after: 300 },
        alignment: docx.AlignmentType.CENTER
      })
    );
  }

  return imageParagraphs;
}

// ---------------------------
// Word document creation (restores original blank-line preservation)
// ---------------------------
function buildParagraphsPreserveBlankLines(text, spacingAfter) {
  const lines = splitPreservingBlankLines(text);
  return lines.map(line => {
    if (line.trim() === "") {
      return new docx.Paragraph({ text: "", spacing: { after: spacingAfter } });
    }
    return new docx.Paragraph({ text: line, spacing: { after: spacingAfter } });
  });
}

function buildTakeawayBulletsPreserveBlankLines(text, spacingAfter) {
  const lines = splitPreservingBlankLines(text);
  return lines.map(line => {
    if (line.trim() === "") {
      return new docx.Paragraph({ text: "", spacing: { after: spacingAfter } });
    }
    const clean = line.replace(/^[-*•]\s*/, "").trim();
    return new docx.Paragraph({
      text: clean,
      bullet: { level: 0 },
      spacing: { after: spacingAfter }
    });
  });
}

function buildSourcesNumbered(sources, spacingAfter) {
  if (!sources || !sources.length) return [];
  return sources.map((s, idx) => {
    const parts = [];
    if (s.name) parts.push(s.name);
    if (s.url) parts.push(s.url);
    if (s.date) parts.push(`Accessed ${s.date}`);
    const head = parts.join(" — ");
    const line = s.line ? ` | ${s.line}` : "";
    return new docx.Paragraph({
      text: `${idx + 1}. ${head}${line}`,
      spacing: { after: spacingAfter }
    });
  });
}

async function createDocument(data) {
  const {
    noteType, title, topic,
    authorLastName, authorFirstName, authorPhone,
    coAuthors,
    keyTakeaways, analysis, content, cordobaView,
    sources,
    imageFiles,
    dateTimeString,
    documentId,
    tightSpacing,
    smallCaptions
  } = data;

  const baseAfter = tightSpacing ? 110 : 150;
  const headingAfter = tightSpacing ? 160 : 200;
  const bulletAfter = tightSpacing ? 80 : 100;
  const captionFont = smallCaptions ? 14 : 18;

  const takeawayBullets = buildTakeawayBulletsPreserveBlankLines(keyTakeaways, bulletAfter);
  const analysisParagraphs = buildParagraphsPreserveBlankLines(analysis, baseAfter);

  // IMPORTANT: original behaviour always appended content paragraphs (even if empty)
  const contentParagraphs = buildParagraphsPreserveBlankLines(content, baseAfter);

  const cordobaParagraphs = buildParagraphsPreserveBlankLines(cordobaView, baseAfter);
  const imageParagraphs = await addImages(imageFiles, captionFont);

  const infoTable = new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    borders: {
      top: { style: docx.BorderStyle.NONE },
      bottom: { style: docx.BorderStyle.NONE },
      left: { style: docx.BorderStyle.NONE },
      right: { style: docx.BorderStyle.NONE },
      insideHorizontal: { style: docx.BorderStyle.NONE },
      insideVertical: { style: docx.BorderStyle.NONE }
    },
    rows: [
      new docx.TableRow({
        children: [
          new docx.TableCell({
            children: [
              new docx.Paragraph({
                text: title,
                bold: true,
                size: 28,
                spacing: { after: 80 }
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    text: documentId ? `Document ID: ${documentId}` : "",
                    italics: true,
                    size: 18,
                    font: "Book Antiqua"
                  })
                ],
                spacing: { after: 120 }
              })
            ],
            width: { size: 60, type: docx.WidthType.PERCENTAGE },
            verticalAlign: docx.VerticalAlign.TOP
          }),
          new docx.TableCell({
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    text: `${authorLastName.toUpperCase()}, ${authorFirstName.toUpperCase()} (${authorPhone})`,
                    bold: true,
                    size: 28
                  })
                ],
                alignment: docx.AlignmentType.RIGHT,
                spacing: { after: 80 }
              })
            ],
            width: { size: 40, type: docx.WidthType.PERCENTAGE },
            verticalAlign: docx.VerticalAlign.TOP
          })
        ]
      }),
      new docx.TableRow({
        children: [
          new docx.TableCell({
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({ text: "TOPIC: ", bold: true, size: 28 }),
                  new docx.TextRun({ text: topic, bold: true, size: 28 })
                ],
                spacing: { after: 120 }
              })
            ],
            width: { size: 60, type: docx.WidthType.PERCENTAGE },
            verticalAlign: docx.VerticalAlign.TOP
          }),
          new docx.TableCell({
            children: coAuthors.length > 0
              ? coAuthors.map(c => new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: `${c.lastName.toUpperCase()}, ${c.firstName.toUpperCase()} (${c.phone})`,
                      bold: true,
                      size: 28
                    })
                  ],
                  alignment: docx.AlignmentType.RIGHT,
                  spacing: { after: 80 }
                }))
              : [new docx.Paragraph({ text: "" })],
            width: { size: 40, type: docx.WidthType.PERCENTAGE },
            verticalAlign: docx.VerticalAlign.TOP
          })
        ]
      })
    ]
  });

  const children = [
    infoTable,
    new docx.Paragraph({
      border: {
        bottom: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 }
      },
      spacing: { after: 260 }
    }),

    new docx.Paragraph({
      children: [new docx.TextRun({ text: "Key Takeaways", bold: true, size: 24, font: "Book Antiqua" })],
      spacing: { after: headingAfter }
    }),
    ...takeawayBullets,

    new docx.Paragraph({ spacing: { after: 240 } }),

    new docx.Paragraph({
      children: [new docx.TextRun({ text: "Analysis and Commentary", bold: true, size: 24, font: "Book Antiqua" })],
      spacing: { after: headingAfter }
    }),
    ...analysisParagraphs,

    // original: content flows directly after analysis (always included)
    ...contentParagraphs
  ];

  if (cordobaView.trim()) {
    children.push(
      new docx.Paragraph({ spacing: { after: 240 } }),
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "The Cordoba View", bold: true, size: 24, font: "Book Antiqua" })],
        spacing: { after: headingAfter }
      }),
      ...cordobaParagraphs
    );
  }

  if (imageParagraphs.length > 0) {
    children.push(
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "Figures and Charts", bold: true, size: 24, font: "Book Antiqua" })],
        spacing: { before: 320, after: headingAfter }
      }),
      ...imageParagraphs
    );
  }

  if (sources && sources.length) {
    children.push(
      new docx.Paragraph({ spacing: { after: 240 } }),
      new docx.Paragraph({
        children: [new docx.TextRun({ text: "Sources", bold: true, size: 24, font: "Book Antiqua" })],
        spacing: { after: headingAfter }
      }),
      ...buildSourcesNumbered(sources, baseAfter)
    );
  }

  const headerText = `Cordoba Research Group | ${noteType} | ${dateTimeString}${documentId ? " | " + documentId : ""}`;

  return new docx.Document({
    styles: {
      default: {
        document: {
          run: { font: "Book Antiqua", size: 20, color: "000000" },
          paragraph: { spacing: { after: baseAfter } }
        }
      }
    },
    sections: [{
      properties: {
        page: {
          margin: { top: 720, right: 720, bottom: 720, left: 720 },
          pageSize: {
            orientation: docx.PageOrientation.LANDSCAPE,
            width: 15840,
            height: 12240
          }
        }
      },
      headers: {
        default: new docx.Header({
          children: [
            new docx.Paragraph({
              children: [new docx.TextRun({ text: headerText, size: 16, font: "Book Antiqua" })],
              alignment: docx.AlignmentType.RIGHT,
              spacing: { after: 80 },
              border: {
                bottom: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 }
              }
            })
          ]
        })
      },
      footers: {
        default: new docx.Footer({
          children: [
            new docx.Paragraph({
              border: {
                top: { color: "000000", space: 1, style: docx.BorderStyle.SINGLE, size: 6 }
              },
              spacing: { after: 0 }
            }),
            new docx.Paragraph({
              children: [
                new docx.TextRun({ text: "\t" }),
                new docx.TextRun({
                  text: "Cordoba Research Group Internal Information",
                  size: 16,
                  font: "Book Antiqua",
                  italics: true
                }),
                new docx.TextRun({ text: "\t" }),
                new docx.TextRun({
                  children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES],
                  size: 16,
                  font: "Book Antiqua",
                  italics: true
                })
              ],
              spacing: { before: 0, after: 0 },
              tabStops: [
                { type: docx.TabStopType.CENTER, position: 5000 },
                { type: docx.TabStopType.RIGHT, position: 10000 }
              ]
            })
          ]
        })
      },
      children
    }]
  });
}

// ---------------------------
// PDF export (print preview window)
// ---------------------------
function exportPDF() {
  const d = collectData();
  const headerLine = `Cordoba Research Group | ${d.noteType || ""} | ${d.dateTimeString}${d.documentId ? " | " + d.documentId : ""}`;

  const figures = Array.from(d.imageFiles || []).map((f, i) => {
    const caption = (f.name || "").replace(/\.[^/.]+$/, "");
    return `Figure ${i + 1}: ${caption}`;
  });

  const html = `
<!doctype html>
<html>
<head>
  <meta charset="utf-8"/>
  <title>${escapeHtml(d.title || "CRG Note")}</title>
  <style>
    @page { size: A4 landscape; margin: 12mm; }
    body{ font-family:"Book Antiqua","Palatino Linotype",Palatino,serif; color:#000; line-height:1.35; }
    .header{ text-align:right; font-size:10px; padding-bottom:6px; border-bottom:1px solid #000; margin-bottom:10px; }
    .titleRow{ display:flex; justify-content:space-between; gap:16px; align-items:flex-start; }
    .title{ font-size:18px; font-weight:700; }
    .docid{ font-size:10px; font-style:italic; margin-top:4px; }
    .author{ font-size:12px; font-weight:700; text-align:right; }
    .topicRow{ display:flex; justify-content:space-between; gap:16px; margin-top:8px; padding-bottom:8px; border-bottom:1px solid #000; }
    .topic{ font-size:12px; font-weight:700; }
    h2{ font-size:13px; margin:14px 0 6px; }
    ul{ margin:0 0 10px 18px; }
    li{ margin:0 0 4px; }
    p{ margin:0 0 8px; white-space:pre-wrap; }
    .mono{ font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono","Courier New", monospace; }
  </style>
</head>
<body>
  <div class="header">${escapeHtml(headerLine)}</div>

  <div class="titleRow">
    <div>
      <div class="title">${escapeHtml(d.title || "")}</div>
      <div class="docid">${d.documentId ? escapeHtml("Document ID: " + d.documentId) : ""}</div>
    </div>
    <div class="author">${escapeHtml((d.authorLastName || "").toUpperCase())}, ${escapeHtml((d.authorFirstName || "").toUpperCase())}${d.authorPhone ? escapeHtml(" (" + d.authorPhone + ")") : ""}</div>
  </div>

  <div class="topicRow">
    <div class="topic">TOPIC: ${escapeHtml(d.topic || "")}</div>
    <div class="topic">${escapeHtml(d.coAuthors.map(a => `${(a.lastName||"").toUpperCase()}, ${(a.firstName||"").toUpperCase()}${a.phone ? " ("+a.phone+")" : ""}`).join(" · "))}</div>
  </div>

  <h2>Key Takeaways</h2>
  <ul>
    ${splitPreservingBlankLines(d.keyTakeaways).map(l => {
      const clean = l.replace(/^[-*•]\s*/, "").trim();
      if (!clean) return `<li style="list-style:none;height:8px;"></li>`;
      return `<li>${escapeHtml(clean)}</li>`;
    }).join("")}
  </ul>

  <h2>Analysis and Commentary</h2>
  ${splitPreservingBlankLines(d.analysis).map(l => l.trim() === "" ? `<p></p>` : `<p>${escapeHtml(l)}</p>`).join("")}

  ${safeText(d.content).trim() ? `<h2>Additional Content</h2>${splitPreservingBlankLines(d.content).map(l => l.trim() === "" ? `<p></p>` : `<p>${escapeHtml(l)}</p>`).join("")}` : ""}

  ${safeText(d.cordobaView).trim() ? `<h2>The Cordoba View</h2>${splitPreservingBlankLines(d.cordobaView).map(l => l.trim() === "" ? `<p></p>` : `<p>${escapeHtml(l)}</p>`).join("")}` : ""}

  ${figures.length ? `<h2>Figures and Charts</h2><ul>${figures.map(f => `<li>${escapeHtml(f)}</li>`).join("")}</ul>` : ""}

  ${
    d.sources.length
      ? `<h2>Sources</h2><ol>
          ${d.sources.map(s => {
            const parts = [];
            if (s.name) parts.push(`<strong>${escapeHtml(s.name)}</strong>`);
            if (s.url) parts.push(`<span class="mono">${escapeHtml(s.url)}</span>`);
            if (s.date) parts.push(`(Accessed ${escapeHtml(s.date)})`);
            const head = parts.join(" — ");
            const line = s.line ? `<div style="margin-top:4px;color:#333">${escapeHtml(s.line)}</div>` : "";
            return `<li style="margin:0 0 10px">${head}${line}</li>`;
          }).join("")}
        </ol>`
      : ""
  }

  <script>
    window.onload = () => window.print();
  </script>
</body>
</html>
  `;

  const w = window.open("", "_blank");
  w.document.open();
  w.document.write(html);
  w.document.close();
}

document.getElementById("pdfBtn").addEventListener("click", exportPDF);

// ---------------------------
// Submit -> Word download (restores original reset prompt)
// ---------------------------
document.getElementById("researchForm").addEventListener("submit", async (e) => {
  e.preventDefault();

  const button = document.getElementById("wordBtn");
  const messageDiv = document.getElementById("message");

  button.disabled = true;
  button.classList.add("loading");
  button.textContent = "Generating Document...";
  messageDiv.className = "message";
  messageDiv.textContent = "";

  try {
    if (typeof docx === "undefined") throw new Error("docx library not loaded. Please refresh.");
    if (typeof saveAs === "undefined") throw new Error("FileSaver library not loaded. Please refresh.");

    // Ensure doc ID exists
    if (!document.getElementById("documentId").value) updateDocIdFromType(true);

    const data = collectData();
    const doc = await createDocument(data);
    const blob = await docx.Packer.toBlob(doc);

    const safeTitle = (data.title || "crg_note").replace(/[^a-z0-9]/gi, "_").toLowerCase();
    const safeType = (data.noteType || "note").replace(/\s+/g, "_").toLowerCase();
    const safeId = (data.documentId || "").replace(/[^a-z0-9-]/gi, "_").toLowerCase();
    const fileName = `${safeTitle}_${safeType}_${safeId}.docx`;

    saveAs(blob, fileName);

    messageDiv.className = "message success";
    messageDiv.textContent = `✓ Document "${fileName}" generated successfully!`;

    // Original UX: prompt to reset
    setTimeout(() => {
      if (confirm("Document generated! Would you like to create another document?")) {
        document.getElementById("researchForm").reset();
        document.getElementById("coAuthorsList").innerHTML = "";
        document.getElementById("sourcesList").innerHTML = "";
        coAuthorCount = 0;
        sourceCount = 0;
        clearSessionDocId();
        updateDocIdFromType(false);
        messageDiv.className = "message";
        messageDiv.textContent = "";
        renderPreview();
      }
    }, 1200);

  } catch (err) {
    console.error(err);
    messageDiv.className = "message error";
    messageDiv.textContent = `✗ Error: ${err.message}`;
  } finally {
    button.disabled = false;
    button.classList.remove("loading");
    button.textContent = "Generate Word Document";
  }
});

// ---------------------------
// Time label + init
// ---------------------------
function tickNowLabel() {
  document.getElementById("nowLabel").textContent = formatDateTime(new Date());
}
tickNowLabel();
setInterval(tickNowLabel, 30_000);

updateDocIdFromType(false);
renderPreview();

console.log("CRG app.js loaded (fixed)");
