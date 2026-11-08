"""
Tool Ekstraksi Ruang Lingkup (Fungsi-Fungsi yang Diuji) dari Dokumen T.4.1 ke T.4.2
==================================================================================
Mengekstrak poin 1.1 (Fungsi-fungsi yang diuji) dari T.4.1 (PDF).
Hasil teks diekstrak dan dimasukkan apa adanya (plain text, tanpa bold buatan),
persis seperti hasil copy-paste langsung yang bersih dan rapi.
"""

import os
import sys
import re
import zlib
import json
import shutil
import argparse
import subprocess

def parse_cmap(cmap_str):
    """Membaca CMap ToUnicode dari stream PDF untuk mendekode font custom/subset."""
    mapping = {}
    for block in re.finditer(r'beginbfchar(.*?)endbfchar', cmap_str, re.DOTALL):
        for m in re.finditer(r'<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>', block.group(1)):
            mapping[int(m.group(1), 16)] = chr(int(m.group(2), 16))
    for block in re.finditer(r'beginbfrange(.*?)endbfrange', cmap_str, re.DOTALL):
        for m in re.finditer(r'<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>', block.group(1)):
            start = int(m.group(1), 16)
            end = int(m.group(2), 16)
            dst_start = int(m.group(3), 16)
            for i in range(end - start + 1):
                mapping[start + i] = chr(dst_start + i)
    return mapping

def decode_token_fallback(token):
    """
    Mendekode token yang menggunakan encoding 16-bit subset tanpa CMap resmi
    (pola umum subset font BookmanOldStyle/Calibri pada PDF hasil ekspor Office).
    """
    if '\x00' in token:
        chars = []
        for i in range(0, len(token), 2):
            c = ord(token[i+1])
            if c == 0xb4:
                chars.append('“')
            elif c == 0xb5:
                chars.append('”')
            elif c == 0x0f:
                chars.append(',')
            elif c == 0x11:
                chars.append('.')
            elif c == 0x1d:
                chars.append(':')
            elif c == 0x03:
                chars.append(' ')
            else:
                chars.append(chr(c + 29))
        return ''.join(chars)
    return token

def extract_pdf_tokens(pdf_path):
    """Mengekstrak urutan teks dari PDF dengan resolusi font encoding & CMap."""
    with open(pdf_path, 'rb') as f:
        data = f.read()

    objs = {}
    streams = {}
    for m in re.finditer(rb'(\d+)\s+0\s+obj(.*?)endobj', data, re.DOTALL):
        oid = int(m.group(1))
        body = m.group(2)
        objs[oid] = body
        sm = re.search(rb'stream[\r\n]+(.*?)[\r\n]+endstream', body, re.DOTALL)
        if sm:
            streams[oid] = sm.group(1)

    cmaps = {}
    for oid, s_bytes in streams.items():
        try:
            decomp = zlib.decompress(s_bytes).decode('latin1', errors='ignore')
            if 'begincmap' in decomp:
                cmaps[oid] = parse_cmap(decomp)
        except Exception:
            pass

    font_tounicode = {}
    for oid, body_bytes in objs.items():
        body = body_bytes.decode('latin1', errors='ignore')
        if '/Type/Font' in body or '/Type /Font' in body:
            tu_m = re.search(r'\/ToUnicode\s+(\d+)\s+0\s+R', body)
            if tu_m:
                font_tounicode[oid] = int(tu_m.group(1))

    pages = []
    for oid, body_bytes in objs.items():
        body = body_bytes.decode('latin1', errors='ignore')
        if ('/Type/Page' in body or '/Type /Page' in body) and ('/Type/Pages' not in body and '/Type /Pages' not in body):
            fonts_dict = {}
            fonts_m = re.search(r'\/Font\s*<<([^>]*)>>', body)
            if fonts_m:
                for fm in re.finditer(r'\/([a-zA-Z0-9]+)\s+(\d+)\s+0\s+R', fonts_m.group(1)):
                    fname = fm.group(1)
                    fid = int(fm.group(2))
                    tu_id = font_tounicode.get(fid)
                    if tu_id and tu_id in cmaps:
                        fonts_dict[fname] = cmaps[tu_id]
            
            c_m = re.search(r'\/Contents\s+(\d+)\s+0\s+R', body)
            if c_m:
                pages.append({'page_id': oid, 'content_ids': [int(c_m.group(1))], 'fonts': fonts_dict})

    all_tokens = []
    for page in pages:
        page_decomp = b''
        for cid in page.get('content_ids', []):
            if cid in streams:
                try:
                    page_decomp += zlib.decompress(streams[cid]) + b'\n'
                except Exception:
                    pass
        
        p_str = page_decomp.decode('latin1', errors='ignore')
        bts = re.findall(r'BT\s*(.*?)\s*ET', p_str, re.DOTALL)
        for bt in bts:
            current_font = None
            tokens = re.split(r'(\/[a-zA-Z0-9]+\s+[\d\.]+\s+Tf|\[.*?\]\s*TJ|\(.*?\)\s*Tj|\(.*?\)\s*[\'|\"])', bt)
            block_text = []
            for token in tokens:
                token = token.strip()
                if not token:
                    continue
                fm = re.match(r'\/([a-zA-Z0-9]+)\s+[\d\.]+\s+Tf', token)
                if fm:
                    current_font = fm.group(1)
                    continue
                cmap = page['fonts'].get(current_font)
                if token.startswith('[') and token.endswith('TJ'):
                    inner = token[1:-2]
                    parts = re.findall(r'(<[0-9a-fA-F]+>|\(.*?\))', inner)
                    line = ''
                    for p in parts:
                        if p.startswith('<') and p.endswith('>'):
                            hex_str = p[1:-1]
                            if cmap:
                                for i in range(0, len(hex_str), 4):
                                    code = int(hex_str[i:i+4], 16)
                                    line += cmap.get(code, '')
                            else:
                                line += bytes.fromhex(hex_str).decode('latin1', errors='ignore')
                        elif p.startswith('(') and p.endswith(')'):
                            raw = p[1:-1].replace(r'\(', '(').replace(r'\)', ')')
                            line += raw
                    block_text.append(line)
                elif token.startswith('(') and (token.endswith('Tj') or token.endswith("'") or token.endswith('"')):
                    raw = token[1:token.rfind(')')].replace(r'\(', '(').replace(r'\)', ')')
                    block_text.append(raw)
            bt_res = ''.join(block_text).strip()
            if bt_res:
                decoded = decode_token_fallback(bt_res)
                all_tokens.append(decoded)
    return all_tokens

def parse_scope_from_tokens(tokens):
    """
    Mengidentifikasi batasan ruang lingkup (katalis start & stop) dan
    menyusun kembali teks apa adanya tanpa modifikasi atau bolding buatan.
    """
    start_idx = -1
    for i, t in enumerate(tokens):
        if re.search(r'1\.1\.?', t) or re.search(r'Fungsi\s*-\s*fungsi\s+yang\s+di\s*uji', t, re.IGNORECASE):
            start_idx = i
            break
    if start_idx == -1:
        for i, t in enumerate(tokens):
            if re.search(r'1\.\s*Cakupan', t, re.IGNORECASE):
                start_idx = i
                break
    if start_idx == -1:
        start_idx = 0

    selected = []
    has_started = False
    for t in tokens[start_idx:]:
        clean_t = t.strip()
        if not clean_t:
            continue

        # Katalis Berhenti
        if has_started and (
            re.search(r'1\.2\.?', clean_t) or 
            re.search(r'Fungsi\s*-\s*fungsi\s+yang\s+tidak\s+di\s*uji', clean_t, re.IGNORECASE) or
            re.search(r'^[2-9]\.\s+(Aspek|Lingkungan|Tugas)', clean_t, re.IGNORECASE)
        ):
            break

        # Deteksi awal item poin (misal 1.) atau header kategori (Input:, Modul...)
        if re.match(r'^\d+\.', clean_t) and not clean_t.startswith('1.1'):
            has_started = True
        elif clean_t in ['Input', 'Input:', 'Proses', 'Proses:', 'Output', 'Output:'] or clean_t.startswith('Modul '):
            has_started = True

        if re.match(r'^T\.4\.\d+$', clean_t):
            continue
        if re.match(r'^\d+$', clean_t) and not re.match(r'^\d+\.', clean_t):
            continue

        if not has_started:
            continue

        selected.append(clean_t)

    # Gabungkan colon ':' yang terpisah baris
    fixed_selected = []
    for t in selected:
        if t == ':' and fixed_selected:
            fixed_selected[-1] = fixed_selected[-1] + ':'
        else:
            fixed_selected.append(t)

    raw_items = []
    current_tokens = []
    category_headers = ['Input:', 'Proses:', 'Output:', 'General:', 'Fitur:']

    for t in fixed_selected:
        is_category = t in category_headers or (t.endswith(':') and len(t.split()) <= 2 and not t.startswith('1.1'))
        is_top_level = bool(re.match(r'^\d+\.', t)) and not t.startswith('1.1')
        is_sub_level = bool(re.match(r'^[a-z]\.', t))

        if is_category or is_top_level or is_sub_level:
            if current_tokens:
                raw_items.append(current_tokens)
                current_tokens = []
        current_tokens.append(t)

    if current_tokens:
        raw_items.append(current_tokens)

    polished_items = []
    for item_tokens in raw_items:
        merged = ""
        for idx, tok in enumerate(item_tokens):
            if idx == 0:
                merged = tok
            else:
                prev_tok = item_tokens[idx - 1]
                if (len(prev_tok) <= 2 and prev_tok.isalpha()) or prev_tok in ['Penyemp', 'Ko', 'Li', 'M']:
                    merged += tok
                elif tok.startswith('-') or merged.endswith('-'):
                    merged += tok
                else:
                    merged += " " + tok

        # Pembersihan escape standar tanpa manipulasi kurung buatan
        text = re.sub(r'(\w+)\s*-\s*(\w+)', r'\1-\2', merged)
        text = text.replace(r'\(', '(').replace(r'\)', ')').replace(r'\,', ',').replace(r'\.', '.')
        text = text.replace('(SD\\', '(SD)').replace('(SMP\\', '(SMP)').replace('(SMP.', '(SMP).')
        text = text.replace('(CIU\\', '(CIU)').replace('(TKT\\', '(TKT)')
        text = text.replace('(lima\\', '(lima)')
        text = text.replace('M edia', 'Media')
        text = re.sub(r'\s+([,.:;])', r'\1', text)
        text = re.sub(r'\s{2,}', ' ', text)

        cleaned_text = text.strip()
        if cleaned_text and not cleaned_text.startswith('1.1.'):
            polished_items.append(cleaned_text)

    return polished_items

def inject_to_t42_docx(docx_path, scope_items):
    """
    Memasukkan hasil ekstraksi ke tabel target di dokumen T.4.2 Word (.docx)
    apa adanya (plain text, tanpa bold) persis seperti copy-paste.
    """
    abs_docx = os.path.abspath(docx_path)
    json_path = os.path.join(os.path.dirname(abs_docx), "_temp_scope_items.json")
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(scope_items, f, ensure_ascii=False, indent=2)

    ps_script = f"""
$ErrorActionPreference = 'Stop'
$docxPath = "{abs_docx}"
$jsonPath = "{json_path}"

$items = Get-Content -Path $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0

try {{
    $doc = $w.Documents.Open($docxPath, $false, $false)
    
    $targetCell = $null
    for ($t = 1; $t -le $doc.Tables.Count; $t++) {{
        $table = $doc.Tables.Item($t)
        for ($r = 1; $r -le $table.Rows.Count; $r++) {{
            try {{
                $cellText = $table.Cell($r, 1).Range.Text
                if ($cellText -match "Fungsi-fungsi yang diuji") {{
                    $targetCell = $table.Cell($r, 2)
                    break
                }}
            }} catch {{}}
        }}
        if ($targetCell -ne $null) {{ break }}
    }}

    if ($targetCell -eq $null) {{
        throw "Tidak menemukan sel 'Fungsi-fungsi yang diuji' pada tabel dokumen T.4.2"
    }}

    # Masukkan teks apa adanya seperti copas biasa (plain text tanpa bold)
    $plainText = $items -join "`r"
    $targetCell.Range.Text = $plainText
    $targetCell.Range.Bold = $false
    $targetCell.Range.ParagraphFormat.LeftIndent = 0
    $targetCell.Range.ParagraphFormat.SpaceBefore = 0
    $targetCell.Range.ParagraphFormat.SpaceAfter = 2

    $doc.Save()
    Write-Host "BERHASIL_UPDATE"
}} finally {{
    if ($doc) {{ $doc.Close([ref]$false) }}
    if ($w) {{ $w.Quit() }}
}}
"""
    ps_file = os.path.join(os.path.dirname(abs_docx), "_temp_injector.ps1")
    with open(ps_file, 'w', encoding='utf-8') as f:
        f.write(ps_script)

    try:
        proc = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file],
            capture_output=True,
            text=True,
            timeout=60
        )
        if "BERHASIL_UPDATE" in proc.stdout:
            return True, "Sukses mengupdate dokumen T.4.2"
        else:
            return False, f"PowerShell Error: {proc.stderr or proc.stdout}"
    finally:
        if os.path.exists(json_path):
            os.remove(json_path)
        if os.path.exists(ps_file):
            os.remove(ps_file)

def export_to_standalone_docx(output_path, source_name, items):
    """
    Membuat file Word DOCX mandiri baru apa adanya (plain text, tanpa bold).
    """
    abs_out = os.path.abspath(output_path)
    json_path = os.path.join(os.path.dirname(abs_out), "_temp_export_items.json")
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump({'source': source_name, 'items': items}, f, ensure_ascii=False, indent=2)

    ps_script = f"""
$ErrorActionPreference = 'Stop'
$outputPath = "{abs_out}"
$jsonPath = "{json_path}"

$data = Get-Content -Path $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
$source = $data.source
$items = $data.items

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0

try {{
    $doc = $w.Documents.Add()

    # Plain text assignment
    $plainText = $items -join "`r"
    $doc.Content.Text = $plainText
    $doc.Content.Bold = $false
    $doc.Content.Font.Name = "Calibri"
    $doc.Content.Font.Size = 11

    $doc.SaveAs2($outputPath)
    Write-Host "BERHASIL_EXPORT"
}} finally {{
    if ($doc) {{ $doc.Close([ref]$false) }}
    if ($w) {{ $w.Quit() }}
}}
"""
    ps_file = os.path.join(os.path.dirname(abs_out), "_temp_export.ps1")
    with open(ps_file, 'w', encoding='utf-8') as f:
        f.write(ps_script)

    try:
        proc = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file],
            capture_output=True,
            text=True,
            timeout=60
        )
        if "BERHASIL_EXPORT" in proc.stdout:
            return True, "Sukses membuat dokumen ekspor"
        else:
            return False, f"PowerShell Error: {proc.stderr or proc.stdout}"
    finally:
        if os.path.exists(json_path):
            os.remove(json_path)
        if os.path.exists(ps_file):
            os.remove(ps_file)

def main():
    parser = argparse.ArgumentParser(description="Tool Ekstrak Ruang Lingkup T.4.1 ke T.4.2")
    parser.add_argument("--t41", help="Path file T.4.1 (PDF)", default=None)
    parser.add_argument("--t42", help="Path file T.4.2 (DOCX)", default=None)
    parser.add_argument("--preview-only", action="store_true", help="Hanya tampilkan preview di console")
    parser.add_argument("--no-export", action="store_true", help="Jangan buat file export mandiri")
    args = parser.parse_args()

    curr_dir = os.path.dirname(os.path.abspath(__file__))

    # Auto-detect file T.4.1 jika tidak dispesifikasikan
    t41_file = args.t41
    if not t41_file:
        for f in os.listdir(curr_dir):
            if f.lower().startswith("t.4.1") and f.lower().endswith(".pdf"):
                t41_file = os.path.join(curr_dir, f)
                break

    # Auto-detect file T.4.2 jika tidak dispesifikasikan
    t42_file = args.t42
    if not t42_file and not args.preview_only:
        for f in os.listdir(curr_dir):
            if ("t.4.2" in f.lower() or "t42" in f.lower()) and f.lower().endswith(".docx") and not f.endswith(".bak"):
                t42_file = os.path.join(curr_dir, f)
                break

    if not t41_file or not os.path.exists(t41_file):
        print(f"[ERROR] File T.4.1 PDF tidak ditemukan di: {t41_file}")
        sys.exit(1)

    t41_base = os.path.splitext(os.path.basename(t41_file))[0]
    print(f"[INFO] Mengekstrak ruang lingkup dari: {os.path.basename(t41_file)}")
    tokens = extract_pdf_tokens(t41_file)
    items = parse_scope_from_tokens(tokens)

    print(f"[INFO] Berhasil mengekstrak {len(items)} item ruang lingkup:\n")
    for it in items:
        if it.endswith(':') and not it[0].isdigit():
            print(f"\n  {it}")
        elif re.match(r'^[a-z]\.', it):
            print(f"      {it}")
        else:
            print(f"    • {it}")

    if args.preview_only:
        print("\n[INFO] Mode --preview-only aktif. Selesai.")
        return

    # AUTO-EXPORT: Buat file .docx mandiri & file .txt
    if not args.no_export:
        export_docx = os.path.join(curr_dir, f"Ruang_Lingkup_{t41_base}.docx")
        export_txt = os.path.join(curr_dir, f"Ruang_Lingkup_{t41_base}.txt")

        print(f"\n[INFO] Meng-export dokumen mandiri ke: {os.path.basename(export_docx)}...")
        if os.path.exists(export_docx):
            try:
                os.remove(export_docx)
            except Exception:
                pass
        exp_ok, exp_msg = export_to_standalone_docx(export_docx, os.path.basename(t41_file), items)
        if exp_ok:
            print(f"[SUKSES] Dokumen export DOCX berhasil dibuat: {os.path.basename(export_docx)}")
        else:
            print(f"[WARNING] Gagal membuat export DOCX: {exp_msg}")

        with open(export_txt, 'w', encoding='utf-8') as f_txt:
            for it in items:
                if it.endswith(':') and not it[0].isdigit():
                    f_txt.write(f"\n{it}\n")
                elif re.match(r'^[a-z]\.', it):
                    f_txt.write(f"   {it}\n")
                else:
                    f_txt.write(f"{it}\n")
        print(f"[SUKSES] Dokumen export TXT berhasil dibuat: {os.path.basename(export_txt)}")

    # JIKA ADA FILE T.4.2 -> Lakukan Inject langsung
    if t42_file and os.path.exists(t42_file):
        backup_file = t42_file + ".bak"
        if not os.path.exists(backup_file):
            shutil.copy2(t42_file, backup_file)
            print(f"\n[INFO] Backup T.4.2 dibuat di: {os.path.basename(backup_file)}")

        print(f"\n[INFO] Memasukkan hasil ekstraksi ke T.4.2: {os.path.basename(t42_file)}...")
        success, msg = inject_to_t42_docx(t42_file, items)
        if success:
            print(f"[SUKSES] {msg}")
        else:
            print(f"[GAGAL] {msg}")
    else:
        print(f"\n[INFO] Selesai. Hasil ekstraksi tersimpan di file export DOCX dan TXT.")

if __name__ == '__main__':
    main()
