// Cordoba Research Group - Research Documentation Tool
// Keeps original functionality, adds: live preview, sources, house style toggles,
// smart paragraphing, bullet conversion, doc ID generation, PDF export via print window.

// ---------------------------
// Helpers
// ---------------------------
function pad3(n) { return String(n).padStart(3, "0"); }
function pad2(n) { return String(n).padStart(2, "0"); }

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

function safeText(v) {
  return (v || "").toString().trim();
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

// Smart paragraphing:
// - Split into paragraphs by blank lines (double line breaks)
// - Within a paragraph, single line breaks become spaces
function smartParagraphBlocks(text) {
  const raw = safeText(text);
  if (!raw) return [];
  const blocks = raw
    .replace(/\r\n/g, "\n")
    .split(/\n\s*\n/g)
    .map(b => b.replace(/\n+/g, " ").trim())
    .filter(Boolean);
  return blocks;
}

// Key takeaways lines (bullets)
function takeawayLines(text) {
  const raw = safeText(text);
  if (!raw) return [];
  return raw
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map(l => l.replace(/^[-*•]\s*/, "").trim())
    .filter(l => l.length > 0);
}

// Convert current Key Takeaways textarea content into clean bullet-style lines
function convertToBulletsInTextarea() {
  const el = document.getElementById("keyTakeaways");
  const lines = takeawayLines(el.value);
  el.value = lines.map(l => `- ${l}`).join("\n");
}

// ---------------------------
// Document ID generation (local sequence)
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

  // If not forcing new, reuse stored ID for the current form session if present
  const sessionKey = "crg_docid_session";
  if (!forceNew) {
    const existing = localStorage.getItem(sessionKey);
    if (existing && existing.startsWith(`CRG-${code}-${yyyy}-${mm}-`)) {
      return existing;
    }
  }

  const seq = incrementAndGetSeq(noteType, yyyy, mm);
  const docId = `CRG-${code}-${yyyy}-${mm}-${pad3(seq)}`;
  localStorage.setItem(sessionKey, docId);
  return docId;
}

function clearSessionDocId() {
  localStorage.removeItem("crg_docid_session");
}

// ---------------------------
// Dynamic Co-authors
// ---------------------------
let coAuthorCount = 0;

function addCoAuthorRow(prefill = {}) {
  coAuthorCount++;
  const coAuthorsList = document.getElementById("coAuthorsList");

  const row = document.createElement("div");
  row.className = "row-card";
  row.id = `coauthor-${coAuthorCount}`;
  row.innerHTML = `
    <div class="row-grid-coauthor">
      <div>
        <label class="mini-label">Last Name</label>
        <input type="text" class="coauthor-lastname" placeholder="Last Name" value="${prefill.lastName || ""}">
      </div>
      <div>
        <label class="mini-label">First Name</label>
        <input type="text" class="coauthor-firstname" placeholder="First Name" value="${prefill.firstName || ""}">
      </div>
      <div>
        <label class="mini-label">Phone</label>
        <input type="text" class="coauthor-phone" placeholder="e.g., 44-7398344190" value="${prefill.phone || ""}">
      </div>
      <button type="button" class="btn btn-ghost btn-sm" onclick="removeCoAuthor(${coAuthorCount})">Remove</button>
    </div>
  `;
  coAuthorsList.appendChild(row);
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
// Dynamic Sources
// ---------------------------
let sourceCount = 0;

function addSourceRow(prefill = {}) {
  sourceCount++;
  const sourcesList = document.getElementById("sourcesList");

  const row = document.createElement("div");
  row.className = "row-card";
  row.id = `source-${sourceCount}`;
  row.innerHTML = `
    <div class="row-grid-source">
      <div>
        <label class="mini-label">Source name</label>
        <input type="text" class="src-name" placeholder="e.g., IMF WEO" value="${prefill.name || ""}">
      </div>
      <div>
        <label class="mini-label">URL</label>
        <input type="text" class="src-url" placeholder="https://..." value="${prefill.url || ""}">
      </div>
      <div>
        <label class="mini-label">Date accessed</label>
        <input type="text" class="src-date" placeholder="YYYY-MM-DD" value="${prefill.date || ""}">
      </div>
      <div>
        <label class="mini-label">Key line (optional)</label>
        <input type="text" class="src-line" placeholder="Short note on what matters..." value="${prefill.line || ""}">
      </div>
      <button type="button" class="btn btn-ghost btn-sm" onclick="removeSource(${sourceCount})">Remove</button>
    </div>
  `;
  sourcesList.appendChild(row);
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

// ---------------------------
// Mini label styling (injected; keeps CSS simple)
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
// Live Preview
// ---------------------------
function collectFormDataLite() {
  const noteType = safeText(document.getElementById("noteType").value);
  const title = safeText(document.getElementById("title").value);
  const topic = safeText(document.getElementById("topic").value);

  const authorLastName = safeText(document.getElementById("authorLastName").value);
  const authorFirstName = safeText(document.getElementById("authorFirstName").value);
  const authorPhone = safeText(document.getElementById("authorPhone").value);

  const keyTakeaways = safeText(document.getElementById("keyTakeaways").value);
  const analysis = safeText(document.getElementById("analysis").value);
  const content = safeText(document.getElementById("content").value);
  const cordobaView = safeText(document.getElementById("cordobaView").value);

  const documentId = safeText(document.getElementById("documentId").value);
  const now = new Date();
  const dateTimeString = formatDateTime(now);

  const tightSpacing = document.getElementById("tightSpacing").checked;
  const smallCaptions = document.getElementById("smallCaptions").checked;

  const coAuthors = [];
  document.querySelectorAll("#coAuthorsList .row-card").forEach(row => {
    const lastName = safeText(row.querySelector(".coauthor-lastname")?.value);
    const firstName = safeText(row.querySelector(".coauthor-firstname")?.value);
    const phone = safeText(row.querySelector(".coauthor-phone")?.value);
    if (lastName || firstName || phone) coAuthors.push({ lastName, firstName, phone });
  });

  const sources = [];
  document.querySelectorAll("#sourcesList .row-card").forEach(row => {
    const name = safeText(row.querySelector(".src-name")?.value);
    const url = safeText(row.querySelector(".src-url")?.value);
    const date = safeText(row.querySelector(".src-date")?.value);
    const line = safeText(row.querySelector(".src-line")?.value);
    if (name || url || date || line) sources.push({ name, url, date, line });
  });

  const imageFiles = document.getElementById("imageUpload").files;

  return {
    noteType, title, topic,
    authorLastName, authorFirstName, authorPhone,
    coAuthors,
    keyTakeaways, analysis, content, cordobaView,
    sources,
    imageFiles,
    documentId,
    dateTimeString,
    tightSpacing,
    smallCaptions
  };
}

function renderPreview() {
  const p = collectFormDataLite();

  // Header line preview
  const headerLine = `Cordoba Research Group | ${p.noteType || "—"} | ${p.dateTimeString || "—"}${p.documentId ? " | " + p.documentId : ""}`;
  document.getElementById("previewHeaderLine").textContent = headerLine;

  const takeaways = takeawayLines(p.keyTakeaways);
  const analysisParas = smartParagraphBlocks(p.analysis);
  const contentParas = smartParagraphBlocks(p.content);
  const cordobaParas = smartParagraphBlocks(p.cordobaView);

  const figures = Array.from(p.imageFiles || []).map((f, i) => {
    const caption = (f.name || "").replace(/\.[^/.]+$/, "");
    return `Figure ${i + 1}: ${caption}`;
  });

  const sources = p.sources || [];

  const title = p.title || "—";
  const topic = p.topic || "—";

  const primaryAuthor = (p.authorLastName || p.authorFirstName || p.authorPhone)
    ? `${(p.authorLastName || "").toUpperCase()}, ${(p.authorFirstName || "").toUpperCase()}${p.authorPhone ? ` (${p.authorPhone})` : ""}`
    : "—";

  const coAuthorsText = p.coAuthors.length
    ? p.coAuthors
        .filter(a => a.lastName || a.firstName || a.phone)
        .map(a => `${(a.lastName || "").toUpperCase()}, ${(a.firstName || "").toUpperCase()}${a.phone ? ` (${a.phone})` : ""}`)
    : [];

  const previewEl = document.getElementById("preview");
  previewEl.innerHTML = `
    <div class="preview-card">
      <div class="preview-head">
        <div class="preview-title">${escapeHtml(title)}</div>
        <div class="preview-meta">
          <div><strong>TOPIC:</strong> ${escapeHtml(topic)}</div>
          <div><strong>PRIMARY:</strong> ${escapeHtml(primaryAuthor)}</div>
          <div><strong>DOC ID:</strong> ${escapeHtml(p.documentId || "—")}</div>
        </div>
        ${
          coAuthorsText.length
            ? `<div class="preview-meta" style="margin-top:8px;">
                <div><strong>CO-AUTHORS:</strong> ${escapeHtml(coAuthorsText.join(" · "))}</div>
              </div>`
            : ""
        }
      </div>

      <div class="preview-body">
        <div class="preview-h">Key Takeaways</div>
        ${
          takeaways.length
            ? `<ul class="preview-ul">${takeaways.map(t => `<li>${escapeHtml(t)}</li>`).join("")}</ul>`
            : `<div class="preview-p muted">—</div>`
        }

        <div class="preview-h">Analysis and Commentary</div>
        ${
          analysisParas.length
            ? analysisParas.map(par => `<div class="preview-p">${escapeHtml(par)}</div>`).join("")
            : `<div class="preview-p muted">—</div>`
        }

        ${
          contentParas.length
            ? `<div class="preview-h">Additional Content</div>${contentParas.map(par => `<div class="preview-p">${escapeHtml(par)}</div>`).join("")}`
            : ""
        }

        ${
          cordobaParas.length
            ? `<div class="preview-h">The Cordoba View</div>${cordobaParas.map(par => `<div class="preview-p">${escapeHtml(par)}</div>`).join("")}`
            : ""
        }

        ${
          figures.length
            ? `<div class="preview-h">Figures and Charts</div>
               <ul class="preview-ul">${figures.map(f => `<li>${escapeHtml(f)}</li>`).join("")}</ul>`
            : ""
        }

        ${
          sources.length
            ? `<div class="preview-h">Sources</div>
               <ol class="preview-ul" style="margin-left:22px;">
                ${sources.map(s => {
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

// Escape for preview HTML
function escapeHtml(str) {
  return (str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

// Update preview on input
const previewWatchedIds = [
  "noteType","title","topic","authorLastName","authorFirstName","authorPhone",
  "keyTakeaways","analysis","content","cordobaView",
  "tightSpacing","smallCaptions","imageUpload"
];
previewWatchedIds.forEach(id => {
  const el = document.getElementById(id);
  if (!el) return;
  el.addEventListener("input", renderPreview);
  el.addEventListener("change", renderPreview);
});

// Re-render preview when dynamic lists change
document.getElementById("coAuthorsList").addEventListener("input", renderPreview);
document.getElementById("sourcesList").addEventListener("input", renderPreview);

// ---------------------------
// Auto doc ID logic on note type change
// ---------------------------
function updateDocIdFromType(forceNew = false) {
  const noteType = safeText(document.getElementById("noteType").value);
  const docIdInput = document.getElementById("documentId");

  if (!noteType) {
    docIdInput.value = "";
    clearSessionDocId();
    renderPreview();
    return;
  }
  const now = new Date();
  const id = generateDocumentId(noteType, now, forceNew);
  docIdInput.value = id;
  renderPreview();
}

document.getElementById("noteType").addEventListener("change", () => {
  // when type changes, generate a new session ID for that type
  updateDocIdFromType(true);
});

document.getElementById("regenDocId").addEventListener("click", () => {
  updateDocIdFromType(true);
});

// ---------------------------
// Key Takeaways Convert button
// ---------------------------
document.getElementById("convertBullets").addEventListener("click", () => {
  convertToBulletsInTextarea();
  renderPreview();
});

// ---------------------------
// Images -> docx paragraphs
// ---------------------------
async function addImages(files, captionFontHalfPoints) {
  const imageParagraphs = [];
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    try {
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
    } catch (err) {
      console.error("Image processing error:", err);
    }
  }
  return imageParagraphs;
}

// ---------------------------
// Word document creation
// ---------------------------
function docxParagraphsFromSmartText(text, spacingAfter) {
  const blocks = smartParagraphBlocks(text);
  if (!blocks.length) return [];
  return blocks.map(b => new docx.Paragraph({ text: b, spacing: { after: spacingAfter } }));
}

function docxBulletParagraphsFromLines(lines, spacingAfter) {
  if (!lines.length) return [new docx.Paragraph({ text: "", spacing: { after: spacingAfter } })];
  return lines.map(line => new docx.Paragraph({
    text: line,
    bullet: { level: 0 },
    spacing: { after: spacingAfter }
  }));
}

function docxNumberedSources(sources, spacingAfter) {
  if (!sources || !sources.length) return [];

  return sources.map((s, idx) => {
    const parts = [];
    if (s.name) parts.push(s.name);
    if (s.url) parts.push(s.url);
    if (s.date) parts.push(`Accessed ${s.date}`);
    const header = parts.join(" — ");

    const line = s.line ? ` | ${s.line}` : "";
    const full = `${idx + 1}. ${header}${line}`;

    return new docx.Paragraph({
      text: full,
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
