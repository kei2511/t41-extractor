/**
 * Ekstraktor Ruang Lingkup T.4.1 ke T.4.2 (Hybrid: DOCX & PDF)
 * - Mendukung ekstraksi langsung dari dokumen DOCX (100% presisi format, bold, italic, spasi)
 * - Mendukung ekstraksi cerdas dari PDF (deteksi font & koordinat posisi)
 * - Menyalin ke clipboard dengan Rich Text (HTML) & Plain Text
 * - Menginjeksi teks berformat ke tabel target di template T.4.2 (.docx)
 */

if (typeof pdfjsLib !== 'undefined') {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'libs/pdf.worker.min.js';
}

// DOM Elements
const filePdfInput = document.getElementById('file-pdf');
const pdfFilenameLabel = document.getElementById('pdf-filename');
const resultEditor = document.getElementById('result-text');
const itemCountBadge = document.getElementById('item-count');
const sourceBadge = document.getElementById('source-badge');
const btnClear = document.getElementById('btn-clear');
const btnCopy = document.getElementById('btn-copy');
const btnDownloadTxt = document.getElementById('btn-download-txt');
const fileDocxInput = document.getElementById('file-docx');
const docxFilenameLabel = document.getElementById('docx-filename');
const btnInjectDocx = document.getElementById('btn-inject-docx');
const toast = document.getElementById('toast');

// Global State
let currentInputFile = null;
let currentFileType = null; // 'docx' or 'pdf'
let currentDocxFile = null;
let currentParagraphsData = []; // Array of paragraphs, each with runs [{text, bold, italic}]

// Storage polyfill (Web localStorage vs Chrome Extension Storage)
const storage = {
  async get(keys) {
    if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
      return new Promise((resolve) => chrome.storage.local.get(keys, resolve));
    }
    const result = {};
    for (const k of keys) {
      const val = localStorage.getItem(k);
      if (val !== null) {
        try { result[k] = JSON.parse(val); } catch (e) { result[k] = val; }
      }
    }
    return result;
  },
  async set(items) {
    if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
      return new Promise((resolve) => chrome.storage.local.set(items, resolve));
    }
    for (const [k, v] of Object.entries(items)) {
      localStorage.setItem(k, typeof v === 'object' ? JSON.stringify(v) : v);
    }
  },
  async remove(keys) {
    if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
      return new Promise((resolve) => chrome.storage.local.remove(keys, resolve));
    }
    for (const k of keys) {
      localStorage.removeItem(k);
    }
  }
};

// Count items and update badge
function updateItemCount(text) {
  if (!text || !text.trim()) {
    itemCountBadge.textContent = '0 item';
    return;
  }
  const lines = text.split('\n').filter(l => l.trim().length > 0);
  const count = lines.filter(l => /^\d+\./.test(l.trim())).length;
  itemCountBadge.textContent = count > 0 ? `${count} fungsi terdeteksi` : `${lines.length} baris`;
}

// Update source badge
function updateSourceBadge(type) {
  if (!sourceBadge) return;
  if (type === 'docx') {
    sourceBadge.textContent = 'DOCX (Akurat 100%)';
    sourceBadge.className = 'source-badge badge-docx';
    sourceBadge.style.display = 'inline-flex';
  } else if (type === 'pdf') {
    sourceBadge.textContent = 'PDF (Ekstraksi Cerdas)';
    sourceBadge.className = 'source-badge badge-pdf';
    sourceBadge.style.display = 'inline-flex';
  } else {
    sourceBadge.style.display = 'none';
  }
}

// Convert HTML inside editor to array of paragraph lines / text
function getEditorPlainText() {
  return resultEditor.innerText || resultEditor.textContent || '';
}

// Convert editor content into Paragraph Data Model for DOCX injection
function getParagraphsFromEditor() {
  const pContainers = Array.from(resultEditor.querySelectorAll('.p-item, div, p'));
  const paragraphs = [];

  const elementsToProcess = pContainers.length > 0 ? pContainers : [resultEditor];

  for (const el of elementsToProcess) {
    if (el === resultEditor && pContainers.length > 0) continue;

    // Detect level from class or style
    let level = 0;
    if (el.classList && el.classList.contains('level-1')) level = 1;
    else if (el.classList && el.classList.contains('level-2')) level = 2;
    else if (el.style && (el.style.paddingLeft || el.style.marginLeft)) {
      const pad = parseInt(el.style.paddingLeft || el.style.marginLeft, 10);
      if (pad >= 40) level = 2;
      else if (pad >= 15) level = 1;
    }

    const runs = [];
    function traverse(node, isBold, isItalic) {
      if (node.nodeType === Node.TEXT_NODE) {
        if (node.textContent) {
          runs.push({
            text: node.textContent,
            bold: isBold,
            italic: isItalic
          });
        }
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        const tag = node.tagName.toLowerCase();
        const bold = isBold || tag === 'b' || tag === 'strong' || node.style.fontWeight === 'bold' || parseInt(node.style.fontWeight, 10) >= 600;
        const italic = isItalic || tag === 'i' || tag === 'em' || node.style.fontStyle === 'italic';
        for (const child of node.childNodes) {
          traverse(child, bold, italic);
        }
      }
    }

    traverse(el, false, false);
    const fullText = runs.map(r => r.text).join('').trim();
    if (fullText.length > 0) {
      // If no explicit class but text starts with sub-item (e.g. "a.", "b.", "•"), assign level 1
      if (level === 0 && (/^[a-zA-Z][\.\)]\s+/.test(fullText) || /^\([a-zA-Z0-9]+\)\s+/.test(fullText))) {
        level = 1;
      }
      paragraphs.push({ runs, level });
    }
  }

  // Fallback if empty
  if (paragraphs.length === 0) {
    const rawLines = (resultEditor.innerText || '').split('\n');
    for (const line of rawLines) {
      const trimmed = line.trim();
      if (!trimmed) continue;
      const isSub = /^[a-zA-Z][\.\)]\s+/.test(trimmed) || line.startsWith('   ') || line.startsWith('\t');
      paragraphs.push({
        runs: [{ text: trimmed, bold: false, italic: false }],
        level: isSub ? 1 : 0
      });
    }
  }

  return paragraphs;
}

// Render paragraph model into editor HTML
function renderParagraphsToEditor(paragraphs) {
  let html = '';
  for (let i = 0; i < paragraphs.length; i++) {
    const pData = paragraphs[i];
    const runs = pData.runs || pData;
    const level = typeof pData.level !== 'undefined' ? pData.level : (
      (runs.length > 0 && /^[a-zA-Z][\.\)]\s+/.test(runs.map(r => r.text).join('').trim())) ? 1 : 0
    );

    let pHtml = '';
    for (const r of runs) {
      let t = escapeHtml(r.text);
      if (r.bold && r.italic) t = `<b><i>${t}</i></b>`;
      else if (r.bold) t = `<b>${t}</b>`;
      else if (r.italic) t = `<i>${t}</i>`;
      pHtml += t;
    }

    const levelClass = level > 0 ? ` level-${level}` : '';
    html += `<div class="p-item${levelClass}">${pHtml || '<br>'}</div>`;
  }
  resultEditor.innerHTML = html;
}

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// Save session
async function saveSession() {
  const html = resultEditor.innerHTML;
  const text = getEditorPlainText();
  if (!text || text.startsWith("Sedang membaca") || text.startsWith("Gagal mengekstrak")) {
    return;
  }
  await storage.set({
    savedResultHtml: html,
    savedResultText: text,
    savedPdfName: currentInputFile ? currentInputFile.name : '',
    savedFileType: currentFileType,
    savedItemCount: itemCountBadge.textContent
  });
}

// Restore saved session
async function restoreSession() {
  try {
    const data = await storage.get(['savedResultHtml', 'savedResultText', 'savedPdfName', 'savedFileType', 'savedItemCount']);
    if (data && (data.savedResultHtml || data.savedResultText)) {
      if (data.savedResultHtml) {
        resultEditor.innerHTML = data.savedResultHtml;
      } else {
        resultEditor.innerText = data.savedResultText;
      }

      if (data.savedPdfName) {
        currentFileType = data.savedFileType || (data.savedPdfName.endsWith('.docx') ? 'docx' : 'pdf');
        pdfFilenameLabel.innerHTML = `File sebelumnya: <strong>${data.savedPdfName}</strong>`;
        updateSourceBadge(currentFileType);
      }
      if (data.savedItemCount) {
        itemCountBadge.textContent = data.savedItemCount;
      } else {
        updateItemCount(getEditorPlainText());
      }
      btnCopy.disabled = false;
      btnDownloadTxt.disabled = false;
      if (currentDocxFile) {
        btnInjectDocx.disabled = false;
      }
    }
  } catch (e) {
    console.warn("Could not restore session:", e);
  }
}

// Show Toast feedback
function showToast(message) {
  toast.textContent = message;
  toast.classList.add('show');
  setTimeout(() => {
    toast.classList.remove('show');
  }, 2500);
}

// Drag & Drop Handlers
function setupDragDrop(zoneId, inputId) {
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(inputId);

  ['dragenter', 'dragover'].forEach(name => {
    zone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      zone.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach(name => {
    zone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      zone.classList.remove('drag-over');
    });
  });

  zone.addEventListener('drop', (e) => {
    if (e.dataTransfer && e.dataTransfer.files.length > 0) {
      input.files = e.dataTransfer.files;
      const event = new Event('change', { bubbles: true });
      input.dispatchEvent(event);
    }
  });
}

setupDragDrop('drop-zone-pdf', 'file-pdf');
setupDragDrop('drop-zone-docx', 'file-docx');

// Input File (DOCX / PDF) Handler
filePdfInput.addEventListener('change', async (e) => {
  if (!e.target.files || e.target.files.length === 0) return;
  const file = e.target.files[0];
  currentInputFile = file;

  const isDocx = /\.docx$/i.test(file.name);
  currentFileType = isDocx ? 'docx' : 'pdf';
  updateSourceBadge(currentFileType);

  pdfFilenameLabel.innerHTML = `File terpilih: <strong>${file.name}</strong>`;
  resultEditor.innerHTML = `<em>Sedang membaca dan mengekstrak ruang lingkup dari ${isDocx ? 'Word (.docx)' : 'PDF'}...</em>`;
  itemCountBadge.textContent = "Memproses...";

  try {
    const arrayBuffer = await file.arrayBuffer();
    let paragraphs = [];

    if (isDocx) {
      paragraphs = await extractScopeFromDocx(arrayBuffer);
    } else {
      paragraphs = await extractScopeFromPdf(arrayBuffer);
    }

    currentParagraphsData = paragraphs;
    renderParagraphsToEditor(paragraphs);

    const plainText = getEditorPlainText();
    updateItemCount(plainText);

    btnCopy.disabled = plainText.trim().length === 0;
    btnDownloadTxt.disabled = plainText.trim().length === 0;
    if (currentDocxFile) {
      btnInjectDocx.disabled = plainText.trim().length === 0;
    }

    await saveSession();
    showToast(`Ruang lingkup berhasil diekstrak dari ${isDocx ? 'DOCX (100% Presisi)' : 'PDF'}! ✅`);
  } catch (err) {
    console.error(err);
    resultEditor.innerHTML = `<span style="color:#ef4444;">Gagal mengekstrak file: ${escapeHtml(err.message)}</span>`;
    itemCountBadge.textContent = "Error";
    updateSourceBadge(null);
  }
});

// Auto-save edited text in editor
resultEditor.addEventListener('input', () => {
  const text = getEditorPlainText().trim();
  updateItemCount(text);
  btnCopy.disabled = text.length === 0;
  btnDownloadTxt.disabled = text.length === 0;
  if (currentDocxFile) {
    btnInjectDocx.disabled = text.length === 0;
  }
  saveSession();
});

// Clear session button
if (btnClear) {
  btnClear.addEventListener('click', async () => {
    resultEditor.innerHTML = '';
    currentInputFile = null;
    currentFileType = null;
    currentParagraphsData = [];
    pdfFilenameLabel.innerHTML = 'Pilih atau Drag & Drop file <strong>T.4.1 (.docx diutamakan / .pdf)</strong>';
    filePdfInput.value = '';
    itemCountBadge.textContent = '0 item';
    updateSourceBadge(null);
    btnCopy.disabled = true;
    btnDownloadTxt.disabled = true;
    btnInjectDocx.disabled = true;

    await storage.remove(['savedResultHtml', 'savedResultText', 'savedPdfName', 'savedFileType', 'savedItemCount']);
    showToast("Sesi berhasil dibersihkan! 🔄");
  });
}

// ---------------------------------------------------------------------------
// 1. EXTRACT FROM DOCX (Preserve 100% Formatting: Bold, Italic, Indents & Word Numbering)
// ---------------------------------------------------------------------------

function toLetters(num) {
  let s = '';
  while (num > 0) {
    const rem = (num - 1) % 26;
    s = String.fromCharCode(97 + rem) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s;
}

function toRoman(num) {
  const map = [
    [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'],
    [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'],
    [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']
  ];
  let res = '';
  for (const [val, roman] of map) {
    while (num >= val) {
      res += roman;
      num -= val;
    }
  }
  return res;
}

function parseNumberingDefinitions(numXmlStr) {
  if (!numXmlStr) return { numMap: {}, abstractNumMap: {} };

  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(numXmlStr, "application/xml");
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  const abstractNumMap = {};
  const abstractElements = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, "abstractNum"));
  for (const ab of abstractElements) {
    const abId = ab.getAttributeNS(W_NS, "abstractNumId");
    const levels = {};
    const lvlElements = Array.from(ab.getElementsByTagNameNS(W_NS, "lvl"));
    for (const lvl of lvlElements) {
      const ilvl = lvl.getAttributeNS(W_NS, "ilvl") || '0';
      const numFmtEl = lvl.getElementsByTagNameNS(W_NS, "numFmt")[0];
      const lvlTextEl = lvl.getElementsByTagNameNS(W_NS, "lvlText")[0];
      const startEl = lvl.getElementsByTagNameNS(W_NS, "start")[0];

      levels[ilvl] = {
        fmt: numFmtEl ? (numFmtEl.getAttributeNS(W_NS, "val") || 'decimal') : 'decimal',
        lvlText: lvlTextEl ? (lvlTextEl.getAttributeNS(W_NS, "val") || `%${parseInt(ilvl, 10) + 1}.`) : `%${parseInt(ilvl, 10) + 1}.`,
        start: startEl ? parseInt(startEl.getAttributeNS(W_NS, "val"), 10) || 1 : 1
      };
    }
    abstractNumMap[abId] = levels;
  }

  const numMap = {};
  const numElements = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, "num"));
  for (const n of numElements) {
    const numId = n.getAttributeNS(W_NS, "numId");
    const abRefEl = n.getElementsByTagNameNS(W_NS, "abstractNumId")[0];
    if (numId && abRefEl) {
      numMap[numId] = abRefEl.getAttributeNS(W_NS, "val");
    }
  }

  return { numMap, abstractNumMap };
}

function formatLevelNumber(val, fmt) {
  if (fmt === 'lowerLetter') return toLetters(val);
  if (fmt === 'upperLetter') return toLetters(val).toUpperCase();
  if (fmt === 'lowerRoman') return toRoman(val);
  if (fmt === 'upperRoman') return toRoman(val).toUpperCase();
  if (fmt === 'bullet') return '•';
  return String(val);
}

async function extractScopeFromDocx(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const xmlStr = await zip.file("word/document.xml").async("text");

  let numberingInfo = { numMap: {}, abstractNumMap: {} };
  const numFile = zip.file("word/numbering.xml");
  if (numFile) {
    try {
      const numXmlStr = await numFile.async("text");
      numberingInfo = parseNumberingDefinitions(numXmlStr);
    } catch (e) {
      console.warn("Could not parse word/numbering.xml:", e);
    }
  }

  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlStr, "application/xml");
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  // Tracker for list numbering counter per numId & ilvl
  const listCounters = {};

  // Collect all paragraphs in document order
  const pElements = Array.from(xmlDoc.getElementsByTagNameNS(W_NS, "p"));
  const allDocParagraphs = [];

  for (const p of pElements) {
    const runs = [];
    const childNodes = Array.from(p.childNodes);

    // Check if paragraph has Word List Numbering (w:numPr)
    let prefixText = '';
    const pPr = p.getElementsByTagNameNS(W_NS, 'pPr')[0];
    if (pPr) {
      const numPr = pPr.getElementsByTagNameNS(W_NS, 'numPr')[0];
      if (numPr) {
        const numIdEl = numPr.getElementsByTagNameNS(W_NS, 'numId')[0];
        const ilvlEl = numPr.getElementsByTagNameNS(W_NS, 'ilvl')[0];

        const numId = numIdEl ? numIdEl.getAttributeNS(W_NS, 'val') : null;
        const ilvl = ilvlEl ? (ilvlEl.getAttributeNS(W_NS, 'val') || '0') : '0';

        if (numId && numberingInfo.numMap[numId]) {
          const abId = numberingInfo.numMap[numId];
          const abDef = numberingInfo.abstractNumMap[abId];
          const lvlDef = abDef ? abDef[ilvl] : null;

          if (lvlDef) {
            const counterKey = `${numId}_${ilvl}`;
            if (typeof listCounters[counterKey] === 'undefined') {
              listCounters[counterKey] = lvlDef.start || 1;
            } else {
              listCounters[counterKey]++;
            }

            const currentVal = listCounters[counterKey];
            const fmtStr = lvlDef.lvlText || '%1.';

            if (lvlDef.fmt === 'bullet') {
              prefixText = '• ';
            } else {
              // Substitute placeholders like %1, %2, %3...
              let formattedPrefix = fmtStr.replace(/%(\d+)/g, (m, p1) => {
                const targetLvl = String(parseInt(p1, 10) - 1);
                const val = targetLvl === ilvl ? currentVal : (listCounters[`${numId}_${targetLvl}`] || 1);
                const targetFmt = (abDef && abDef[targetLvl]) ? abDef[targetLvl].fmt : lvlDef.fmt;
                return formatLevelNumber(val, targetFmt);
              });

              prefixText = `${formattedPrefix} `;
            }
          } else {
            prefixText = '• ';
          }
        } else {
          prefixText = '• ';
        }
      }
    }

    // Determine nesting level (0 = top, 1 = sub-point like a./b./c., 2 = sub-sub)
    let pLevel = 0;
    const pPrEl = p.getElementsByTagNameNS(W_NS, 'pPr')[0];
    if (pPrEl) {
      const numPr = pPrEl.getElementsByTagNameNS(W_NS, 'numPr')[0];
      if (numPr) {
        const ilvlEl = numPr.getElementsByTagNameNS(W_NS, 'ilvl')[0];
        pLevel = ilvlEl ? parseInt(ilvlEl.getAttributeNS(W_NS, 'val') || '0', 10) : 0;
      } else {
        const indEl = pPrEl.getElementsByTagNameNS(W_NS, 'ind')[0];
        if (indEl) {
          const leftVal = parseInt(indEl.getAttributeNS(W_NS, 'left') || '0', 10);
          if (leftVal >= 720) pLevel = 1;
          if (leftVal >= 1440) pLevel = 2;
        }
      }
    }

    // Add prefix run if any
    if (prefixText) {
      // If the paragraph already begins with explicit number/letter/bullet (e.g. "1.", "a.", "•"),
      // don't prepend duplicate numbering!
      const rawText = childNodes.map(c => c.textContent || '').join('').trim();
      const alreadyNumbered = /^((\d+|[a-zA-Z])[\.\)]|•|\-)\s+/.test(rawText);
      if (!alreadyNumbered) {
        runs.push({
          text: prefixText,
          bold: false,
          italic: false
        });
      }
    }

    for (const child of childNodes) {
      if (child.localName === 'r') {
        const rPr = child.getElementsByTagNameNS(W_NS, 'rPr')[0];
        let bold = false;
        let italic = false;

        if (rPr) {
          const bEl = rPr.getElementsByTagNameNS(W_NS, 'b')[0];
          if (bEl && bEl.getAttributeNS(W_NS, 'val') !== '0' && bEl.getAttributeNS(W_NS, 'val') !== 'false') {
            bold = true;
          }
          const iEl = rPr.getElementsByTagNameNS(W_NS, 'i')[0];
          if (iEl && iEl.getAttributeNS(W_NS, 'val') !== '0' && iEl.getAttributeNS(W_NS, 'val') !== 'false') {
            italic = true;
          }
        }

        const tElements = Array.from(child.getElementsByTagNameNS(W_NS, 't'));
        for (const t of tElements) {
          runs.push({
            text: t.textContent || '',
            bold: bold,
            italic: italic
          });
        }
      }
    }

    const fullText = runs.map(r => r.text).join('').trim();
    if (fullText.length > 0) {
      // If level wasn't set by numPr but text looks like sub-level ("a.", "b.", "1)", etc.), mark as level 1
      if (pLevel === 0 && (/^[a-zA-Z][\.\)]\s+/.test(fullText) || /^\([a-zA-Z0-9]+\)\s+/.test(fullText))) {
        pLevel = 1;
      }
      allDocParagraphs.push({ text: fullText, runs: runs, level: pLevel });
    }
  }

  // Anchor Search: Start at 1.1 or Fungsi-fungsi yang diuji, Stop at 1.2 or next chapter
  let startIdx = -1;
  for (let i = 0; i < allDocParagraphs.length; i++) {
    const txt = allDocParagraphs[i].text;
    if (/Fungsi\s*-\s*fungsi\s+yang\s+di\s*uji/i.test(txt) || /(^|\b)1\.1\.?/i.test(txt)) {
      startIdx = i;
      break;
    }
  }
  if (startIdx === -1) {
    for (let i = 0; i < allDocParagraphs.length; i++) {
      if (/(^|\b)1\.\s*Cakupan/i.test(allDocParagraphs[i].text)) {
        startIdx = i;
        break;
      }
    }
  }
  if (startIdx === -1) startIdx = 0;

  const extracted = [];
  let hasStarted = false;

  for (let i = startIdx; i < allDocParagraphs.length; i++) {
    const item = allDocParagraphs[i];
    const txt = item.text.trim();

    // Stop Anchor: Any paragraph containing "Fungsi-fungsi yang tidak diuji",
    // or starting with 1.2, 3.2, 2., etc.
    if (hasStarted && (
      /Fungsi\s*-\s*fungsi\s+yang\s+tidak\s+di\s*uji/i.test(txt) ||
      /(^|\b)\d+\.2\.?/i.test(txt) ||
      /^[2-9]\.\s+(Aspek|Lingkungan|Tugas|Hasil|Jadwal)/i.test(txt)
    )) {
      break;
    }

    // Skip the section header itself:
    // e.g. "1.1.", "1.1. Fungsi-fungsi yang diuji", "3.1. Fungsi-fungsi yang diuji",
    // or introductory sentence like "Adapun fungsi-fungsi yang akan dibangun... adalah sebagai berikut:"
    if (/Fungsi\s*-\s*fungsi\s+yang\s+di\s*uji/i.test(txt) ||
        /(^|\b)\d+\.\d+\.?\s*(Ruang\s+Lingkup)?$/i.test(txt) ||
        /^(Ruang\s+Lingkup)$/i.test(txt)) {
      continue;
    }
    
    // Skip any introductory sentences before the actual list items
    // (e.g. "...telah diuji...", "...adalah sebagai berikut:", "sebagai berikut:", "...Modul Laporan...adalah")
    if (!hasStarted) {
      if (/sebagai\s+berikut\s*:?$/i.test(txt) ||
          /adalah\s*:?$/i.test(txt) ||
          /diuji\s+untuk/i.test(txt) ||
          /telah\s+diuji/i.test(txt) ||
          /akan\s+dibangun/i.test(txt) ||
          /(adapun|berikut)\s+ini/i.test(txt)) {
        continue;
      }

      // Check if this paragraph is indeed the START of the functions list:
      // Must start with "1." or "Input:" or "Fungsi Input" or similar
      const isActualFirstItem = /^\d+\.\s+/i.test(txt) ||
                                /^(Fungsi\s+)?(Input|Proses|Output|General|Fitur):?/i.test(txt) ||
                                /^Modul\s+[A-Z0-9]/i.test(txt) && /:/i.test(txt);
      if (!isActualFirstItem) {
        continue; // Skip any residual introductory narrative text
      }
    }

    // Trigger start on actual function content
    hasStarted = true;

    // Filter noise like T.4.1 page header/footer
    if (/^T\.4\.\d+$/i.test(txt)) continue;
    if (/^\d+$/.test(txt) && !/^\d+\./.test(txt)) continue;

    extracted.push({ runs: item.runs, level: item.level || 0 });
  }

  return extracted.length > 0 ? extracted : allDocParagraphs.map(p => ({ runs: p.runs, level: p.level || 0 }));
}

// ---------------------------------------------------------------------------
// 2. EXTRACT FROM PDF (Rich Parser with Font & Coordinate Awareness)
// ---------------------------------------------------------------------------
async function extractScopeFromPdf(arrayBuffer) {
  const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
  const pdfDoc = await loadingTask.promise;

  const rawTokens = [];

  for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const textContent = await page.getTextContent();

    for (const item of textContent.items) {
      const str = item.str.trim();
      if (str.length > 0) {
        const fontName = (item.fontName || '').toLowerCase();
        const isBold = fontName.includes('bold') || fontName.includes('bd') || fontName.includes('b0') || fontName.includes('heavy') || fontName.includes('black');
        const isItalic = fontName.includes('italic') || fontName.includes('it') || fontName.includes('oblique');

        rawTokens.push({
          str: item.str,
          fontName: item.fontName,
          bold: isBold,
          italic: isItalic,
          x: item.transform[4],
          y: item.transform[5]
        });
      }
    }
  }

  return parseScopeTokensRich(rawTokens);
}

function parseScopeTokensRich(tokens) {
  let startIdx = -1;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i].str;
    if (/1\.1\.?/i.test(t) || /Fungsi\s*-\s*fungsi\s+yang\s+di\s*uji/i.test(t)) {
      startIdx = i;
      break;
    }
  }
  if (startIdx === -1) {
    for (let i = 0; i < tokens.length; i++) {
      if (/1\.\s*Cakupan/i.test(tokens[i].str)) {
        startIdx = i;
        break;
      }
    }
  }
  if (startIdx === -1) startIdx = 0;

  const selected = [];
  let hasStarted = false;

  for (let i = startIdx; i < tokens.length; i++) {
    const tok = tokens[i];
    const t = tok.str.trim();
    if (!t) continue;

    // Stop Anchor
    if (hasStarted && (
      /1\.2\.?/i.test(t) ||
      /Fungsi\s*-\s*fungsi\s+yang\s+tidak\s+di\s*uji/i.test(t) ||
      /^[2-9]\.\s+(Aspek|Lingkungan|Tugas|Hasil|Jadwal)/i.test(t)
    )) {
      break;
    }

    // Skip section headers & intro sentences before actual list begins
    if (!hasStarted) {
      if (/Fungsi\s*-\s*fungsi\s+yang\s+di\s*uji/i.test(t) ||
          /(^|\b)\d+\.\d+\.?/i.test(t) ||
          /sebagai\s+berikut/i.test(t) ||
          /diuji\s+untuk/i.test(t) ||
          /adalah/i.test(t) ||
          /akan\s+dibangun/i.test(t)) {
        continue;
      }

      if (/^\d+\./.test(t) && !t.startsWith('1.1')) {
        hasStarted = true;
      } else if (['Input', 'Input:', 'Proses', 'Proses:', 'Output', 'Output:'].includes(t) || /^(Fungsi\s+)?(Input|Proses|Output)/i.test(t)) {
        hasStarted = true;
      }
    }

    if (/^T\.4\.\d+$/i.test(t)) continue;
    if (/^\d+$/.test(t) && !/^\d+\./.test(t)) continue;

    if (!hasStarted) continue;
    selected.push(tok);
  }

  // Combine standalone ':' with previous token
  const fixed = [];
  for (const tok of selected) {
    const t = tok.str.trim();
    if (t === ':' && fixed.length > 0) {
      const prev = fixed[fixed.length - 1];
      prev.str = prev.str.trim() + ':';
    } else {
      fixed.push({ ...tok, str: tok.str.trim() });
    }
  }

  // Group tokens into items based on category/numbers
  const categoryHeaders = ['Input:', 'Proses:', 'Output:', 'General:', 'Fitur:'];
  const rawGroups = [];
  let currentGroup = [];

  for (const tok of fixed) {
    const t = tok.str;
    const isCategory = categoryHeaders.includes(t) || (t.endsWith(':') && t.split(' ').length <= 2 && !t.startsWith('1.1'));
    const isTop = /^\d+\./.test(t) && !t.startsWith('1.1');
    const isSub = /^[a-z]\./.test(t);

    if (isCategory || isTop || isSub) {
      if (currentGroup.length > 0) {
        rawGroups.push(currentGroup);
        currentGroup = [];
      }
    }
    currentGroup.push(tok);
  }
  if (currentGroup.length > 0) {
    rawGroups.push(currentGroup);
  }

  // Polish each group into paragraph runs
  const paragraphs = [];
  for (const group of rawGroups) {
    const runs = [];
    for (let idx = 0; idx < group.length; idx++) {
      const tok = group[idx];
      let prefixSpace = false;

      if (idx > 0) {
        const prevTok = group[idx - 1];
        const prevStr = prevTok.str;
        const currStr = tok.str;

        if ((prevStr.length <= 2 && /^[a-zA-Z]+$/.test(prevStr)) || ['Penyemp', 'Ko', 'Li', 'M'].includes(prevStr)) {
          prefixSpace = false;
        } else if (currStr.startsWith('-') || prevStr.endsWith('-')) {
          prefixSpace = false;
        } else {
          prefixSpace = true;
        }
      }

      let cleanedStr = (prefixSpace ? ' ' : '') + tok.str;
      cleanedStr = cleanedStr.replace(/\\\(|\(/g, '(').replace(/\\\)|\)/g, ')');
      cleanedStr = cleanedStr.replace(/\(CIU\\/g, '(CIU)').replace(/\(TKT\\/g, '(TKT)');

      runs.push({
        text: cleanedStr,
        bold: tok.bold,
        italic: tok.italic
      });
    }

    if (runs.length > 0) {
      const fullText = runs.map(r => r.text).join('').trim();
      // Indentasi level 1 jika:
      // - Berupa sub-huruf seperti 'a.', 'b.', 'c.'
      // - Berupa list '(1)', '(a)', atau bullet '•', '-'
      // - Berupa poin nomor (misal '1. Unggah...') yang berada di bawah header kategori (seperti 'Fungsi Input', 'Input:')
      const isSubLetter = /^[a-zA-Z][\.\)]\s+/.test(fullText) || /^\([a-zA-Z0-9]+\)\s+/.test(fullText);
      const isBullet = /^[•\-\*]\s+/.test(fullText);
      const isCategoryHeader = /^(Fungsi\s+)?(Input|Proses|Output|General|Fitur):?$/i.test(fullText) || (fullText.endsWith(':') && fullText.split(' ').length <= 3 && !/^\d+\./.test(fullText));
      
      let level = 0;
      if (isSubLetter || isBullet) {
        level = 1;
      }
      paragraphs.push({ runs: runs, level: level, isCategory: isCategoryHeader });
    }
  }

  // Pass kedua: jika ada pola kategori (Input/Proses/Output) lalu di bawahnya ada nomor (1. 2. 3.),
  // buat nomor di bawah kategori ikut menjorok jika dokumen bertipe hierarkis kategori -> nomor
  let insideCategory = false;
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const fullText = p.runs.map(r => r.text).join('').trim();
    if (p.isCategory) {
      insideCategory = true;
      p.level = 0;
    } else if (insideCategory) {
      if (/^\d+\.\s+(Fungsi\s+)?(Input|Proses|Output)/i.test(fullText)) {
        insideCategory = true;
        p.level = 0;
      } else if (/^[a-zA-Z][\.\)]\s+/.test(fullText)) {
        p.level = 1;
      }
    }
  }

  return paragraphs;
}

// ---------------------------------------------------------------------------
// 3. COPY TO CLIPBOARD (Rich Text HTML with Indents + Plain Text Fallback)
// ---------------------------------------------------------------------------
btnCopy.addEventListener('click', async () => {
  const plainText = getEditorPlainText().trim();
  if (!plainText) return;

  // Generate clean HTML with inline styling for clipboard so Microsoft Word and Google Docs render indents
  const paragraphs = getParagraphsFromEditor();
  let exportHtml = '<div>';
  for (const pItem of paragraphs) {
    const runs = pItem.runs || pItem;
    const level = typeof pItem.level !== 'undefined' ? pItem.level : 0;
    const indentPx = level === 1 ? '24px' : (level >= 2 ? '48px' : '0px');

    let rowHtml = '';
    for (const r of runs) {
      let t = escapeHtml(r.text);
      if (r.bold && r.italic) t = `<b><i>${t}</i></b>`;
      else if (r.bold) t = `<b>${t}</b>`;
      else if (r.italic) t = `<i>${t}</i>`;
      rowHtml += t;
    }

    const styleAttr = indentPx !== '0px' ? ` style="margin-left: ${indentPx}; margin-bottom: 2px;"` : ' style="margin-bottom: 2px;"';
    exportHtml += `<div${styleAttr}>${rowHtml || '&nbsp;'}</div>`;
  }
  exportHtml += '</div>';

  try {
    if (navigator.clipboard && navigator.clipboard.write) {
      const blobHtml = new Blob([exportHtml], { type: 'text/html' });
      const blobText = new Blob([plainText], { type: 'text/plain' });
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/html': blobHtml,
          'text/plain': blobText
        })
      ]);
      showToast("Tersalin ke Clipboard dengan Format Menjorok! Siap dipaste ke Word ✔");
      return;
    }
  } catch (err) {
    console.warn("ClipboardItem API failed, fallback to standard copy:", err);
  }

  // Fallback
  try {
    await navigator.clipboard.writeText(plainText);
    showToast("Tersalin ke Clipboard (Plain Text)! ✔");
  } catch (err) {
    const range = document.createRange();
    range.selectNodeContents(resultEditor);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand('copy');
    showToast("Tersalin ke Clipboard! ✔");
  }
});

// Download TXT
btnDownloadTxt.addEventListener('click', () => {
  const text = getEditorPlainText().trim();
  if (!text) return;

  const baseName = currentInputFile ? currentInputFile.name.replace(/\.(docx|pdf)$/i, '') : 'T41_Ruang_Lingkup';
  const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Ruang_Lingkup_${baseName}.txt`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast("File .TXT berhasil didownload! 📥");
});

// DOCX File Selection Handler (Template T.4.2)
fileDocxInput.addEventListener('change', (e) => {
  if (!e.target.files || e.target.files.length === 0) return;
  currentDocxFile = e.target.files[0];
  docxFilenameLabel.innerHTML = `Template T.4.2: <strong>${currentDocxFile.name}</strong>`;

  const text = getEditorPlainText().trim();
  if (text.length > 0 && !text.startsWith("Sedang membaca")) {
    btnInjectDocx.disabled = false;
  }
  showToast("Template T.4.2 siap untuk diinjeksi! 📄");
});

// ---------------------------------------------------------------------------
// 4. INJECT INTO DOCX (With Preserved Bold / Italic Runs)
// ---------------------------------------------------------------------------
btnInjectDocx.addEventListener('click', async () => {
  if (!currentDocxFile) {
    alert("Silakan upload file template T.4.2 (.docx) terlebih dahulu.");
    return;
  }
  const plainText = getEditorPlainText().trim();
  if (!plainText) {
    alert("Tidak ada teks ruang lingkup yang bisa diinjeksi.");
    return;
  }

  btnInjectDocx.disabled = true;
  btnInjectDocx.textContent = "Sedang menginjeksi dokumen...";

  try {
    const arrayBuffer = await currentDocxFile.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    // Read word/document.xml
    const docXmlStr = await zip.file("word/document.xml").async("text");
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(docXmlStr, "application/xml");

    const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const tables = xmlDoc.getElementsByTagNameNS(W_NS, "tbl");

    let targetCell = null;
    for (let t = 0; t < tables.length; t++) {
      const tbl = tables[t];
      const rows = tbl.getElementsByTagNameNS(W_NS, "tr");
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        const cells = row.getElementsByTagNameNS(W_NS, "tc");
        if (cells.length >= 2) {
          const c1Text = cells[0].textContent || "";
          if (c1Text.includes("Fungsi-fungsi yang diuji")) {
            targetCell = cells[1];
            break;
          }
        }
      }
      if (targetCell) break;
    }

    if (!targetCell) {
      throw new Error("Tidak menemukan sel 'Fungsi-fungsi yang diuji' pada tabel dokumen T.4.2");
    }

    // Preserve tcPr
    const tcPr = targetCell.getElementsByTagNameNS(W_NS, "tcPr")[0];

    // Clear all child nodes of targetCell
    while (targetCell.firstChild) {
      targetCell.removeChild(targetCell.firstChild);
    }
    if (tcPr) {
      targetCell.appendChild(tcPr);
    }

    // Get structured paragraphs with runs (from editor)
    const paragraphs = getParagraphsFromEditor();

    for (const pItem of paragraphs) {
      const pRuns = pItem.runs || pItem;
      const pLevel = typeof pItem.level !== 'undefined' ? pItem.level : 0;

      const p = xmlDoc.createElementNS(W_NS, "w:p");
      const pPr = xmlDoc.createElementNS(W_NS, "w:pPr");
      const pStyle = xmlDoc.createElementNS(W_NS, "w:pStyle");
      pStyle.setAttributeNS(W_NS, "w:val", "TableParagraph");
      pPr.appendChild(pStyle);

      // Apply indentation if level > 0 (menjorok ke dalam)
      // Level 1 = 360 twips (0.25 inch / ~0.63 cm), Level 2 = 720 twips
      if (pLevel > 0) {
        const ind = xmlDoc.createElementNS(W_NS, "w:ind");
        const leftTwips = pLevel === 1 ? "400" : (pLevel >= 2 ? "800" : "0");
        ind.setAttributeNS(W_NS, "w:left", leftTwips);
        pPr.appendChild(ind);
      }

      p.appendChild(pPr);

      for (const runData of pRuns) {
        const r = xmlDoc.createElementNS(W_NS, "w:r");
        
        // Add run formatting if bold or italic
        if (runData.bold || runData.italic) {
          const rPr = xmlDoc.createElementNS(W_NS, "w:rPr");
          if (runData.bold) {
            const b = xmlDoc.createElementNS(W_NS, "w:b");
            rPr.appendChild(b);
          }
          if (runData.italic) {
            const i = xmlDoc.createElementNS(W_NS, "w:i");
            rPr.appendChild(i);
          }
          r.appendChild(rPr);
        }

        const t = xmlDoc.createElementNS(W_NS, "w:t");
        t.textContent = runData.text;
        if (runData.text.startsWith(' ') || runData.text.endsWith(' ') || runData.text.includes('  ')) {
          t.setAttribute("xml:space", "preserve");
        }
        r.appendChild(t);
        p.appendChild(r);
      }

      targetCell.appendChild(p);
    }

    // Serialize and save back
    const serializer = new XMLSerializer();
    const newDocXml = serializer.serializeToString(xmlDoc);
    zip.file("word/document.xml", newDocXml);

    const updatedBlob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(updatedBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Updated_${currentDocxFile.name}`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast("Dokumen T.4.2 berhasil diinjeksi dengan format rapi! 🎉");
  } catch (err) {
    console.error(err);
    alert(`Gagal menginjeksi DOCX: ${err.message}`);
  } finally {
    btnInjectDocx.disabled = false;
    btnInjectDocx.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>
      <span>Injeksi & Download File T.4.2 (.docx)</span>
    `;
  }
});

// Initialize and restore previous session
restoreSession();
