"""
Tool Ekstraksi Ruang Lingkup (Fungsi-Fungsi yang Diuji) dari Dokumen T.4.1 ke T.4.2
==================================================================================
Mendukung input T.4.1 format DOCX (Disarankan, 100% presisi formatting) dan PDF.
Mengekstrak poin 1.1 (Fungsi-fungsi yang diuji) dari T.4.1 dan menginjeksi ke T.4.2.
"""

import os
import sys
import re
import zlib
import json
import shutil
import argparse
import subprocess
import zipfile
import xml.etree.ElementTree as ET

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

def to_letters(num):
    s = ''
    while num > 0:
        rem = (num - 1) % 26
        s = chr(97 + rem) + s
        num = (num - 1) // 26
    return s

def to_roman(num):
    val_map = [
        (1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'),
        (100, 'c'), (90, 'xc'), (50, 'l'), (40, 'xl'),
        (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')
    ]
    res = ''
    for val, roman in val_map:
        while num >= val:
            res += roman
            num -= val
    return res

def format_level_number(val, fmt):
    if fmt == 'lowerLetter':
        return to_letters(val)
    elif fmt == 'upperLetter':
        return to_letters(val).upper()
    elif fmt == 'lowerRoman':
        return to_roman(val)
    elif fmt == 'upperRoman':
        return to_roman(val).upper()
    elif fmt == 'bullet':
        return '•'
    return str(val)

def parse_docx_numbering(num_xml_bytes):
    if not num_xml_bytes:
        return {}, {}
    root = ET.fromstring(num_xml_bytes)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    abstract_num_map = {}
    for ab in root.findall('.//w:abstractNum', ns):
        ab_id = ab.get(f'{{{ns["w"]}}}abstractNumId')
        levels = {}
        for lvl in ab.findall('w:lvl', ns):
            ilvl = lvl.get(f'{{{ns["w"]}}}ilvl', '0')
            num_fmt = lvl.find('w:numFmt', ns)
            lvl_text = lvl.find('w:lvlText', ns)
            start = lvl.find('w:start', ns)
            levels[ilvl] = {
                'fmt': num_fmt.get(f'{{{ns["w"]}}}val', 'decimal') if num_fmt is not None else 'decimal',
                'lvlText': lvl_text.get(f'{{{ns["w"]}}}val', f'%{int(ilvl)+1}.') if lvl_text is not None else f'%{int(ilvl)+1}.',
                'start': int(start.get(f'{{{ns["w"]}}}val', '1')) if start is not None else 1
            }
        abstract_num_map[ab_id] = levels

    num_map = {}
    for n in root.findall('.//w:num', ns):
        num_id = n.get(f'{{{ns["w"]}}}numId')
        ab_ref = n.find('w:abstractNumId', ns)
        if num_id and ab_ref is not None:
            num_map[num_id] = ab_ref.get(f'{{{ns["w"]}}}val')

    return num_map, abstract_num_map

def extract_scope_from_docx(docx_path):
    """
    Mengekstrak ruang lingkup langsung dari file Word T.4.1 (.docx)
    mempertahankan 100% format teks, spasi, Word numbering (numPr), dan struktur aslinya.
    """
    with zipfile.ZipFile(docx_path) as z:
        xml_content = z.read('word/document.xml')
        num_xml_bytes = z.read('word/numbering.xml') if 'word/numbering.xml' in z.namelist() else None
    
    num_map, abstract_num_map = parse_docx_numbering(num_xml_bytes)
    list_counters = {}

    root = ET.fromstring(xml_content)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # Ambil semua paragraf
    all_p = []
    for p in root.findall('.//w:p', ns):
        prefix_text = ''
        pPr = p.find('w:pPr', ns)
        if pPr is not None:
            numPr = pPr.find('w:numPr', ns)
            if numPr is not None:
                num_id_el = numPr.find('w:numId', ns)
                ilvl_el = numPr.find('w:ilvl', ns)
                num_id = num_id_el.get(f'{{{ns["w"]}}}val') if num_id_el is not None else None
                ilvl = ilvl_el.get(f'{{{ns["w"]}}}val', '0') if ilvl_el is not None else '0'

                if num_id and num_id in num_map:
                    ab_id = num_map[num_id]
                    ab_def = abstract_num_map.get(ab_id, {})
                    lvl_def = ab_def.get(ilvl)
                    if lvl_def:
                        counter_key = f"{num_id}_{ilvl}"
                        if counter_key not in list_counters:
                            list_counters[counter_key] = lvl_def.get('start', 1)
                        else:
                            list_counters[counter_key] += 1
                        
                        curr_val = list_counters[counter_key]
                        fmt_str = lvl_def.get('lvlText', '%1.')

                        if lvl_def.get('fmt') == 'bullet':
                            prefix_text = '• '
                        else:
                            def repl(m):
                                t_lvl = str(int(m.group(1)) - 1)
                                v = curr_val if t_lvl == ilvl else list_counters.get(f"{num_id}_{t_lvl}", 1)
                                t_fmt = ab_def[t_lvl]['fmt'] if t_lvl in ab_def else lvl_def['fmt']
                                return format_level_number(v, t_fmt)

                            formatted_num = re.sub(r'%(\d+)', repl, fmt_str)
                            indent = '   ' if int(ilvl) > 0 else ''
                            prefix_text = f"{indent}{formatted_num} "
                    else:
                        prefix_text = '   • ' if int(ilvl) > 0 else '• '
                else:
                    prefix_text = '   • ' if int(ilvl) > 0 else '• '

        raw_text = ''.join(p.itertext()).strip()
        if raw_text:
            already_numbered = bool(re.match(r'^((\d+|[a-zA-Z])[\.\)]|•|\-)\s+', raw_text))
            if prefix_text and already_numbered:
                prefix_text = ''
            all_p.append((prefix_text + raw_text).strip())

    # Temukan batas start 1.1 dan stop 1.2
    start_idx = -1
    for i, t in enumerate(all_p):
        if re.search(r'Fungsi\s*-\s*fungsi\s+yang\s+di\s*uji', t, re.IGNORECASE) or re.search(r'(^|\b)1\.1\.?', t):
            start_idx = i
            break
    if start_idx == -1:
        for i, t in enumerate(all_p):
            if re.search(r'(^|\b)1\.\s*Cakupan', t, re.IGNORECASE):
                start_idx = i
                break
    if start_idx == -1:
        start_idx = 0

    extracted = []
    has_started = False
    for t in all_p[start_idx:]:
        clean_t = t.strip()
        if not clean_t:
            continue

        if has_started and (
            re.search(r'Fungsi\s*-\s*fungsi\s+yang\s+tidak\s+di\s*uji', clean_t, re.IGNORECASE) or
            re.search(r'(^|\b)\d+\.2\.?', clean_t) or
            re.search(r'^[2-9]\.\s+(Aspek|Lingkungan|Tugas|Hasil|Jadwal)', clean_t, re.IGNORECASE)
        ):
            break

        # Lewati header subbab sendiri dan kalimat pengantar naratif sebelum item dimulai
        if re.search(r'Fungsi\s*-\s*fungsi\s+yang\s+di\s*uji', clean_t, re.IGNORECASE) or \
           re.search(r'(^|\b)\d+\.\d+\.?\s*(Ruang\s+Lingkup)?$', clean_t, re.IGNORECASE) or \
           re.search(r'^(Ruang\s+Lingkup)$', clean_t, re.IGNORECASE):
            continue

        if not has_started:
            if re.search(r'sebagai\s+berikut\s*:?$', clean_t, re.IGNORECASE) or \
               re.search(r'adalah\s*:?$', clean_t, re.IGNORECASE) or \
               re.search(r'diuji\s+untuk', clean_t, re.IGNORECASE) or \
               re.search(r'telah\s+diuji', clean_t, re.IGNORECASE) or \
               re.search(r'akan\s+dibangun', clean_t, re.IGNORECASE) or \
               re.search(r'(adapun|berikut)\s+ini', clean_t, re.IGNORECASE):
                continue

            is_actual_first_item = bool(
                re.match(r'^\d+\.\s+', clean_t) or
                re.match(r'^(Fungsi\s+)?(Input|Proses|Output|General|Fitur):?', clean_t, re.IGNORECASE) or
                (clean_t.startswith('Modul ') and ':' in clean_t)
            )
            if not is_actual_first_item:
                continue

        has_started = True

        if re.match(r'^T\.4\.\d+$', clean_t):
            continue

        extracted.append(clean_t)

    return extracted if extracted else all_p

def inject_to_t42_docx(docx_path, scope_items):
    """
    Memasukkan hasil ekstraksi ke tabel target di dokumen T.4.2 Word (.docx)
    menggunakan manipulasi OpenXML langsung (atau fallback PowerShell COM).
    """
    try:
        # Gunakan manipulasi OpenXML murni (Cepat, OS-independent)
        with zipfile.ZipFile(docx_path, 'r') as zin:
            file_map = {name: zin.read(name) for name in zin.namelist()}

        doc_xml = file_map['word/document.xml']
        root = ET.fromstring(doc_xml)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        ET.register_namespace('w', ns['w'])

        target_cell = None
        for tbl in root.findall('.//w:tbl', ns):
            for row in tbl.findall('.//w:tr', ns):
                cells = row.findall('.//w:tc', ns)
                if len(cells) >= 2:
                    c1_text = ''.join(cells[0].itertext())
                    if 'Fungsi-fungsi yang diuji' in c1_text:
                        target_cell = cells[1]
                        break
            if target_cell is not None:
                break

        if target_cell is not None:
            # Simpan tcPr jika ada
            tcPr = target_cell.find('w:tcPr', ns)
            target_cell.clear()
            if tcPr is not None:
                target_cell.append(tcPr)

            for line in scope_items:
                p = ET.Element(f"{{{ns['w']}}}p")
                pPr = ET.SubElement(p, f"{{{ns['w']}}}pPr")
                pStyle = ET.SubElement(pPr, f"{{{ns['w']}}}pStyle")
                pStyle.set(f"{{{ns['w']}}}val", "TableParagraph")

                # Tambahkan indentasi menjorok jika baris adalah sub-poin (misal 'a.', 'b.', atau diawali spasi)
                clean_l = line.strip()
                is_sub = bool(re.match(r'^[a-zA-Z][\.\)]\s+', clean_l) or re.match(r'^\([a-zA-Z0-9]+\)\s+', clean_l) or line.startswith('   '))
                if is_sub:
                    ind = ET.SubElement(pPr, f"{{{ns['w']}}}ind")
                    ind.set(f"{{{ns['w']}}}left", "400")

                r = ET.SubElement(p, f"{{{ns['w']}}}r")
                t = ET.SubElement(r, f"{{{ns['w']}}}t")
                t.text = clean_l
                if clean_l.startswith(' ') or clean_l.endswith(' '):
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

                target_cell.append(p)

            # Tulis ulang zip
            with zipfile.ZipFile(docx_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for name, content in file_map.items():
                    if name == 'word/document.xml':
                        zout.writestr(name, ET.tostring(root, encoding='utf-8', xml_declaration=True))
                    else:
                        zout.writestr(name, content)

            return True, "Sukses mengupdate dokumen T.4.2 via OpenXML Engine"
    except Exception as e:
        print(f"[DEBUG] OpenXML direct inject error, falling back to COM: {e}")

    # Fallback ke PowerShell COM jika ada Word terinstall di Windows
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
    Membuat file Word DOCX mandiri baru.
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
    $range = $doc.Range()
    $range.Text = ($items -join "`r")
    $doc.SaveAs([ref]$outputPath)
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
    parser = argparse.ArgumentParser(description="Tool Ekstrak Ruang Lingkup T.4.1 (DOCX / PDF) ke T.4.2")
    parser.add_argument("--t41", help="Path file T.4.1 (.docx diutamakan, atau .pdf)", default=None)
    parser.add_argument("--t42", help="Path file T.4.2 (DOCX)", default=None)
    parser.add_argument("--preview-only", action="store_true", help="Hanya tampilkan preview di console")
    parser.add_argument("--no-export", action="store_true", help="Jangan buat file export mandiri")
    args = parser.parse_args()

    curr_dir = os.path.dirname(os.path.abspath(__file__))

    # Peringatan utama
    print("=" * 70)
    print("[PERHATIAN / REKOMENDASI UTAMA]")
    print("Disarankan mengutamakan file T.4.1 berformat DOCX (.docx).")
    print("Format Word (.docx) menjaga 100% presisi teks, bold, italic, dan spasi.")
    print("Jika hanya ada file PDF (.pdf), sistem akan melakukan rekonstruksi teks cerdas.")
    print("=" * 70 + "\n")

    # Auto-detect file T.4.1 (Utamakan DOCX jika ada, lalu PDF)
    t41_file = args.t41
    if not t41_file:
        # Cek docx terlebih dahulu (abaikan lock file Word ~$ dan file hasil ekspor)
        for f in os.listdir(curr_dir):
            if f.startswith("~$") or f.startswith("."):
                continue
            if ("t.4.1" in f.lower() or "t41" in f.lower()) and f.lower().endswith(".docx") and not f.startswith("Ruang_Lingkup_") and not f.endswith(".bak"):
                t41_file = os.path.join(curr_dir, f)
                break
        # Jika tidak ada docx, cari PDF
        if not t41_file:
            for f in os.listdir(curr_dir):
                if f.startswith("~$") or f.startswith("."):
                    continue
                if ("t.4.1" in f.lower() or "t41" in f.lower()) and f.lower().endswith(".pdf"):
                    t41_file = os.path.join(curr_dir, f)
                    break

    # Auto-detect file T.4.2 jika tidak dispesifikasikan
    t42_file = args.t42
    if not t42_file and not args.preview_only:
        for f in os.listdir(curr_dir):
            if f.startswith("~$") or f.startswith("."):
                continue
            if ("t.4.2" in f.lower() or "t42" in f.lower()) and f.lower().endswith(".docx") and not f.endswith(".bak"):
                t42_file = os.path.join(curr_dir, f)
                break

    if not t41_file or not os.path.exists(t41_file):
        print(f"[ERROR] File T.4.1 (DOCX/PDF) tidak ditemukan!")
        sys.exit(1)

    is_docx = t41_file.lower().endswith(".docx")
    t41_base = os.path.splitext(os.path.basename(t41_file))[0]

    print(f"[INFO] Sumber Dokumen : {os.path.basename(t41_file)}")
    print(f"[INFO] Tipe Format    : {'DOCX (Akurasi Format 100%)' if is_docx else 'PDF (Ekstraksi Cerdas)'}")

    if is_docx:
        items = extract_scope_from_docx(t41_file)
    else:
        tokens = extract_pdf_tokens(t41_file)
        items = parse_scope_from_tokens(tokens)

    print(f"\n[INFO] Berhasil mengekstrak {len(items)} item ruang lingkup:\n")
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
