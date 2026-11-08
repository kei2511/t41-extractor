/**
 * Popup Logic for Ekstraktor Ruang Lingkup T.4.1 Chrome Extension
 */

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'libs/pdf.worker.min.js';
}

// Elements
const filePdfInput = document.getElementById('file-pdf');
const dropZonePdf = document.getElementById('drop-zone-pdf');
const pdfFilenameLabel = document.getElementById('pdf-filename');

const fileDocxInput = document.getElementById('file-docx');
const dropZoneDocx = document.getElementById('drop-zone-docx');
const docxFilenameLabel = document.getElementById('docx-filename');

const resultTextarea = document.getElementById('result-text');
const itemCountBadge = document.getElementById('item-count');
const btnClear = document.getElementById('btn-clear');
const btnCopy = document.getElementById('btn-copy');
const btnDownloadTxt = document.getElementById('btn-download-txt');
const btnInjectDocx = document.getElementById('btn-inject-docx');
const toast = document.getElementById('toast');

let currentPdfName = '';
let currentDocxFile = null;

// Storage Helper (supports chrome.storage.local with localStorage fallback)
const storage = {
  async get(keys) {
    if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
      return new Promise((resolve) => chrome.storage.local.get(keys, resolve));
    }
    const res = {};
    for (const k of keys) {
      const val = localStorage.getItem(k);
      if (val !== null) {
        try { res[k] = JSON.parse(val); } catch (e) { res[k] = val; }
      }
    }
    return res;
  },
  async set(items) {
    if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
      return new Promise((resolve) => chrome.storage.local.set(items, resolve));
    }
    for (const [k, v] of Object.entries(items)) {
      localStorage.setItem(k, JSON.stringify(v));
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

// Save current session to storage
async function saveSession() {
  const text = resultTextarea.value;
  if (!text || text.startsWith("Sedang membaca") || text.startsWith("Gagal mengekstrak")) {
    return;
  }
  await storage.set({
    savedResultText: text,
    savedPdfName: currentPdfName,
    savedItemCount: itemCountBadge.textContent
  });
}

// Restore saved session from storage
async function restoreSession() {
  try {
    const data = await storage.get(['savedResultText', 'savedPdfName', 'savedItemCount']);
    if (data && data.savedResultText && data.savedResultText.trim().length > 0) {
      resultTextarea.value = data.savedResultText;
      if (data.savedPdfName) {
        currentPdfName = data.savedPdfName;
        pdfFilenameLabel.innerHTML = `File sebelumnya: <strong>${data.savedPdfName}</strong>`;
      }
      if (data.savedItemCount) {
        itemCountBadge.textContent = data.savedItemCount;
      } else {
        updateItemCount(data.savedResultText);
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

// Setup Drag & Drop Handlers
function setupDragDrop(zone, input, label, acceptExt) {
  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('drag-over');
  });

  zone.addEventListener('dragleave', () => {
    zone.classList.remove('drag-over');
  });

  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      if (file.name.toLowerCase().endsWith(acceptExt)) {
        input.files = e.dataTransfer.files;
        input.dispatchEvent(new Event('change'));
      } else {
        alert(`Harap upload file berekstensi ${acceptExt}`);
      }
    }
  });
}

setupDragDrop(dropZonePdf, filePdfInput, pdfFilenameLabel, '.pdf');
setupDragDrop(dropZoneDocx, fileDocxInput, docxFilenameLabel, '.docx');

// PDF File Selection Handler
filePdfInput.addEventListener('change', async (e) => {
  if (!e.target.files || e.target.files.length === 0) return;
  const file = e.target.files[0];
  currentPdfName = file.name;
  pdfFilenameLabel.innerHTML = `File terpilih: <strong>${file.name}</strong>`;

  resultTextarea.value = "Sedang membaca dan mengekstrak ruang lingkup dari PDF...";
  itemCountBadge.textContent = "Memproses...";

  try {
    const arrayBuffer = await file.arrayBuffer();
    const extractedText = await extractScopeFromPdf(arrayBuffer);
    
    resultTextarea.value = extractedText;
    updateItemCount(extractedText);
    
    btnCopy.disabled = false;
    btnDownloadTxt.disabled = false;
    
    if (currentDocxFile) {
      btnInjectDocx.disabled = false;
    }
    
    await saveSession();
    showToast("Ruang lingkup berhasil diekstrak! ✅");
  } catch (err) {
    console.error(err);
    resultTextarea.value = `Gagal mengekstrak PDF: ${err.message}`;
    itemCountBadge.textContent = "Error";
  }
});

// Auto-save edited text in textarea
resultTextarea.addEventListener('input', () => {
  const text = resultTextarea.value.trim();
  updateItemCount(resultTextarea.value);
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
    resultTextarea.value = '';
    currentPdfName = '';
    pdfFilenameLabel.innerHTML = 'Pilih atau Drag & Drop file <strong>T.4.1 (PDF)</strong>';
    filePdfInput.value = '';
    itemCountBadge.textContent = '0 item';
    btnCopy.disabled = true;
    btnDownloadTxt.disabled = true;
    btnInjectDocx.disabled = true;

    await storage.remove(['savedResultText', 'savedPdfName', 'savedItemCount']);
    showToast("Sesi berhasil dibersihkan! 🔄");
  });
}

// Extract text from PDF buffer using PDF.js and Katalis / Anchor Logic
async function extractScopeFromPdf(arrayBuffer) {
  const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
  const pdfDoc = await loadingTask.promise;
  
  const allTokens = [];
  
  for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const textContent = await page.getTextContent();
    
    for (const item of textContent.items) {
      const str = item.str.trim();
      if (str.length > 0) {
        allTokens.push(str);
      }
    }
  }

  // Parse scope from tokens using Start & Stop anchors
  return parseScopeTokens(allTokens);
}

function parseScopeTokens(tokens) {
  let startIdx = -1;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (/1\.1\.?/i.test(t) || /Fungsi\s*-\s*fungsi\s+yang\s+di\s*uji/i.test(t)) {
      startIdx = i;
      break;
    }
  }
  if (startIdx === -1) {
    for (let i = 0; i < tokens.length; i++) {
      if (/1\.\s*Cakupan/i.test(tokens[i])) {
        startIdx = i;
        break;
      }
    }
  }
  if (startIdx === -1) startIdx = 0;

  const selected = [];
  let hasStarted = false;

  for (let i = startIdx; i < tokens.length; i++) {
    const t = tokens[i].trim();
    if (!t) continue;

    // Stop Anchor
    if (hasStarted && (
      /1\.2\.?/i.test(t) ||
      /Fungsi\s*-\s*fungsi\s+yang\s+tidak\s+di\s*uji/i.test(t) ||
      /^[2-9]\.\s+(Aspek|Lingkungan|Tugas)/i.test(t)
    )) {
      break;
    }

    // Start detect
    if (/^\d+\./.test(t) && !t.startsWith('1.1')) {
      hasStarted = true;
    } else if (['Input', 'Input:', 'Proses', 'Proses:', 'Output', 'Output:'].includes(t) || t.startsWith('Modul ')) {
      hasStarted = true;
    }

    // Skip headers and page numbers
    if (/^T\.4\.\d+$/i.test(t)) continue;
    if (/^\d+$/.test(t) && !/^\d+\./.test(t)) continue;

    if (!hasStarted) continue;
    selected.push(t);
  }

  // Combine standalone ':' with previous token
  const fixed = [];
  for (const t of selected) {
    if (t === ':' && fixed.length > 0) {
      fixed[fixed.length - 1] += ':';
    } else {
      fixed.push(t);
    }
  }

  // Group tokens into items
  const categoryHeaders = ['Input:', 'Proses:', 'Output:', 'General:', 'Fitur:'];
  const rawItems = [];
  let currentTokens = [];

  for (const t of fixed) {
    const isCategory = categoryHeaders.includes(t) || (t.endsWith(':') && t.split(' ').length <= 2 && !t.startsWith('1.1'));
    const isTop = /^\d+\./.test(t) && !t.startsWith('1.1');
    const isSub = /^[a-z]\./.test(t);

    if (isCategory || isTop || isSub) {
      if (currentTokens.length > 0) {
        rawItems.push(currentTokens);
        currentTokens = [];
      }
    }
    currentTokens.push(t);
  }
  if (currentTokens.length > 0) {
    rawItems.push(currentTokens);
  }

  // Polish each item
  const polished = [];
  for (const itemTokens of rawItems) {
    let merged = "";
    for (let idx = 0; idx < itemTokens.length; idx++) {
      const tok = itemTokens[idx];
      if (idx === 0) {
        merged = tok;
      } else {
        const prevTok = itemTokens[idx - 1];
        if ((prevTok.length <= 2 && /^[a-zA-Z]+$/.test(prevTok)) || ['Penyemp', 'Ko', 'Li', 'M'].includes(prevTok)) {
          merged += tok;
        } else if (tok.startsWith('-') || merged.endsWith('-')) {
          merged += tok;
        } else {
          merged += " " + tok;
        }
      }
    }

    let text = merged.replace(/(\w+)\s*-\s*(\w+)/g, '$1-$2');
    text = text.replace(/\\\(|\(/g, '(').replace(/\\\)|\)/g, ')');
    text = text.replace(/\(CIU\\/g, '(CIU)').replace(/\(TKT\\/g, '(TKT)').replace(/\(SD\\/g, '(SD)').replace(/\(SMP\\/g, '(SMP)').replace(/\(lima\\/g, '(lima)');
    text = text.replace(/M\s+edia/g, 'Media');
    text = text.replace(/\s+([,.:;])/g, '$1');
    text = text.replace(/\s{2,}/g, ' ').trim();

    if (text && !text.startsWith('1.1.')) {
      polished.push(text);
    }
  }

  // Format final plain text
  const finalLines = [];
  for (const it of polished) {
    if (it.endsWith(':') && !/^\d/.test(it)) {
      if (finalLines.length > 0) finalLines.push("");
      finalLines.push(it);
    } else if (/^[a-z]\./.test(it)) {
      finalLines.push("   " + it);
    } else {
      finalLines.push(it);
    }
  }

  return finalLines.join('\n');
}

// Copy to Clipboard
btnCopy.addEventListener('click', async () => {
  const text = resultTextarea.value.trim();
  if (!text) return;
  try {
    await navigator.clipboard.writeText(text);
    showToast("Tersalin ke Clipboard! Siap dipaste (Ctrl+V) ✔");
  } catch (err) {
    resultTextarea.select();
    document.execCommand('copy');
    showToast("Tersalin ke Clipboard! ✔");
  }
});

// Download TXT
btnDownloadTxt.addEventListener('click', () => {
  const text = resultTextarea.value.trim();
  if (!text) return;

  const baseName = currentPdfName ? currentPdfName.replace(/\.pdf$/i, '') : 'T41_Ruang_Lingkup';
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

// DOCX File Selection Handler
fileDocxInput.addEventListener('change', (e) => {
  if (!e.target.files || e.target.files.length === 0) return;
  currentDocxFile = e.target.files[0];
  docxFilenameLabel.innerHTML = `Template T.4.2: <strong>${currentDocxFile.name}</strong>`;
  
  if (resultTextarea.value.trim().length > 0 && !resultTextarea.value.startsWith("Sedang membaca")) {
    btnInjectDocx.disabled = false;
  }
  showToast("Template T.4.2 siap untuk diinjeksi! 📄");
});

// Inject into DOCX & Download
btnInjectDocx.addEventListener('click', async () => {
  if (!currentDocxFile) {
    alert("Silakan upload file template T.4.2 (.docx) terlebih dahulu.");
    return;
  }
  const textToInject = resultTextarea.value.trim();
  if (!textToInject) {
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

    // Add new paragraphs from textarea
    const lines = textToInject.split('\n');
    for (const line of lines) {
      const p = xmlDoc.createElementNS(W_NS, "w:p");
      const pPr = xmlDoc.createElementNS(W_NS, "w:pPr");
      const pStyle = xmlDoc.createElementNS(W_NS, "w:pStyle");
      pStyle.setAttributeNS(W_NS, "w:val", "TableParagraph");
      pPr.appendChild(pStyle);
      p.appendChild(pPr);

      const r = xmlDoc.createElementNS(W_NS, "w:r");
      const t = xmlDoc.createElementNS(W_NS, "w:t");
      t.textContent = line;
      if (line.startsWith(' ') || line.endsWith(' ')) {
        t.setAttribute("xml:space", "preserve");
      }
      r.appendChild(t);
      p.appendChild(r);
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

    showToast("Dokumen T.4.2 berhasil diinjeksi dan didownload! 🎉");
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

// Initialize and restore previous session if any
restoreSession();

