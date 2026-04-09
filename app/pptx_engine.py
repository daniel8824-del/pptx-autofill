"""
PPTX Engine — XML 수준 텍스트 분석 + 교체 엔진
"""
import zipfile
import os
import re
from xml.etree import ElementTree as ET

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def unpack(pptx_path: str, dest_dir: str):
    os.makedirs(dest_dir, exist_ok=True)
    with zipfile.ZipFile(pptx_path) as z:
        z.extractall(dest_dir)


def repack(unpacked_dir: str, original_pptx: str, output_pptx: str):
    with zipfile.ZipFile(original_pptx, 'r') as orig, \
         zipfile.ZipFile(output_pptx, 'w', zipfile.ZIP_DEFLATED) as out:
        for item in orig.infolist():
            fpath = os.path.join(unpacked_dir, item.filename)
            if os.path.isfile(fpath):
                with open(fpath, 'rb') as f:
                    out.writestr(item, f.read())
            elif not item.is_dir():
                out.writestr(item, orig.read(item.filename))


def get_slide_files(pptx_path: str) -> list[str]:
    with zipfile.ZipFile(pptx_path) as z:
        return sorted(
            [f for f in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml', f)],
            key=lambda x: int(re.search(r'(\d+)', x).group())
        )


def analyze_template(pptx_path: str) -> dict:
    """템플릿 구조 분석 — 슬라이드별 도형/표/텍스트 매핑"""
    slides = []
    with zipfile.ZipFile(pptx_path) as z:
        slide_files = get_slide_files(pptx_path)
        for slide_name in slide_files:
            num = int(re.search(r'(\d+)', slide_name).group())
            tree = ET.fromstring(z.read(slide_name))
            slide_info = {"number": num, "shapes": [], "tables": [], "images": []}

            # 도형 분석
            for sp in tree.findall('.//p:sp', NS):
                cNvPr = sp.find('p:nvSpPr/p:cNvPr', NS)
                if cNvPr is None:
                    continue
                sp_id = cNvPr.get('id', '?')
                sp_name = cNvPr.get('name', '?')
                paragraphs = []
                for p in sp.findall('.//a:p', NS):
                    texts = [t.text for t in p.findall('.//a:t', NS) if t.text]
                    if texts:
                        paragraphs.append(''.join(texts))
                if paragraphs:
                    slide_info["shapes"].append({
                        "id": sp_id,
                        "name": sp_name,
                        "text": paragraphs,
                    })

            # 표 분석
            for gf in tree.findall('.//p:graphicFrame', NS):
                tbl = gf.find('.//a:tbl', NS)
                if tbl is None:
                    continue
                gf_cNvPr = gf.find('.//p:cNvPr', NS)
                tbl_name = gf_cNvPr.get('name', '표') if gf_cNvPr is not None else '표'
                rows_data = []
                for tr in tbl.findall('a:tr', NS):
                    row = []
                    for tc in tr.findall('a:tc', NS):
                        texts = [t.text for t in tc.findall('.//a:t', NS) if t.text]
                        row.append(''.join(texts) if texts else '')
                    rows_data.append(row)
                slide_info["tables"].append({
                    "name": tbl_name,
                    "rows": rows_data,
                    "row_count": len(rows_data),
                    "col_count": len(rows_data[0]) if rows_data else 0,
                })

            # 이미지 개수
            pics = tree.findall('.//p:pic', NS)
            slide_info["images"] = len(pics)

            slides.append(slide_info)

    return {
        "slide_count": len(slides),
        "slides": slides,
    }


def analyze_template_summary(pptx_path: str) -> str:
    """분석 결과를 Claude API에 전달할 텍스트 요약으로 변환"""
    info = analyze_template(pptx_path)
    lines = [f"총 {info['slide_count']}장 슬라이드\n"]
    for s in info["slides"]:
        lines.append(f"--- Slide {s['number']} ---")
        for sh in s["shapes"]:
            text_preview = ' | '.join(t[:60] for t in sh["text"][:3])
            lines.append(f"  [shape id={sh['id']}] {sh['name']}: {text_preview}")
        for tb in s["tables"]:
            lines.append(f"  [table] {tb['name']}: {tb['row_count']}행 x {tb['col_count']}열")
            for i, row in enumerate(tb["rows"][:2]):
                lines.append(f"    행{i+1}: {' | '.join(c[:30] for c in row)}")
            if tb["row_count"] > 2:
                lines.append(f"    ... (+{tb['row_count']-2}행)")
        if s["images"]:
            lines.append(f"  [이미지 {s['images']}개 — 수정 금지]")
    return '\n'.join(lines)


def find_shape_by_id(root, shape_id: str):
    for sp in root.findall('.//p:sp', NS):
        cNvPr = sp.find('p:nvSpPr/p:cNvPr', NS)
        if cNvPr is not None and cNvPr.get('id') == str(shape_id):
            return sp
    for grp in root.findall('.//p:grpSp', NS):
        for sp in grp.findall('.//p:sp', NS):
            cNvPr = sp.find('p:nvSpPr/p:cNvPr', NS)
            if cNvPr is not None and cNvPr.get('id') == str(shape_id):
                return sp
    return None


def enable_autofit(sp):
    """도형에 자동 축소(normAutofit) 활성화 — 텍스트가 넘칠 때 자동 축소"""
    body_pr = sp.find('.//a:bodyPr', NS)
    if body_pr is None:
        return
    # 기존 autofit 설정 제거
    for child in list(body_pr):
        tag = child.tag.split('}')[-1]
        if tag in ('noAutofit', 'spAutoFit', 'normAutofit'):
            body_pr.remove(child)
    # normAutofit 추가 (fontScale 최소 50%까지 축소 허용)
    ET.SubElement(body_pr, f'{{{NS["a"]}}}normAutofit', attrib={'fontScale': '50000'})


def replace_shape_texts(sp, new_texts: list[str]):
    """도형 내 <a:t> 텍스트만 교체 — 서식 유지 + 자동 축소 활성화"""
    if isinstance(new_texts, str):
        new_texts = [new_texts]
    paragraphs = sp.findall('.//a:p', NS)
    text_paras = [(p, p.findall('.//a:t', NS)) for p in paragraphs]
    text_paras = [(p, runs) for p, runs in text_paras if runs]

    # 원본 글자수 계산
    original_len = sum(len(t.text or '') for _, runs in text_paras for t in runs)

    for i, new_text in enumerate(new_texts):
        if i < len(text_paras):
            p, runs = text_paras[i]
            runs[0].text = new_text
            for r in runs[1:]:
                r.text = ''

    # 새 글자수가 원본 대비 20% 이상 초과하면 autofit 활성화
    new_len = sum(len(t) for t in new_texts)
    if original_len > 0 and new_len > original_len * 1.2:
        enable_autofit(sp)


def replace_table_cell(tbl, row_idx: int, col_idx: int, new_text: str):
    rows = tbl.findall('a:tr', NS)
    if row_idx < len(rows):
        cells = rows[row_idx].findall('a:tc', NS)
        if col_idx < len(cells):
            tc = cells[col_idx]
            runs = tc.findall('.//a:t', NS)
            if runs:
                original_len = sum(len(r.text or '') for r in runs)
                runs[0].text = new_text
                for r in runs[1:]:
                    r.text = ''
                # 셀 텍스트가 20% 이상 초과하면 autofit
                if original_len > 0 and len(new_text) > original_len * 1.2:
                    body_pr = tc.find('.//a:bodyPr', NS)
                    if body_pr is not None:
                        for child in list(body_pr):
                            tag = child.tag.split('}')[-1]
                            if tag in ('noAutofit', 'spAutoFit', 'normAutofit'):
                                body_pr.remove(child)
                        ET.SubElement(body_pr, f'{{{NS["a"]}}}normAutofit',
                                      attrib={'fontScale': '60000'})


def apply_replacements(unpacked_dir: str, replacements: dict):
    """
    replacements 형식:
    {
        1: {  # slide number
            "shapes": { "22": ["새 텍스트1", "새 텍스트2"], ... },
            "tables": { "표 56": [[row0], [row1], ...] }
        },
        2: { ... }
    }
    """
    for slide_num, slide_data in replacements.items():
        path = os.path.join(unpacked_dir, f'ppt/slides/slide{slide_num}.xml')
        if not os.path.exists(path):
            continue
        tree = ET.parse(path)
        root = tree.getroot()

        # 도형 텍스트 교체
        for shape_id, texts in slide_data.get("shapes", {}).items():
            sp = find_shape_by_id(root, shape_id)
            if sp is not None:
                replace_shape_texts(sp, texts)

        # 표 텍스트 교체
        for tbl_name, rows in slide_data.get("tables", {}).items():
            for gf in root.findall('.//p:graphicFrame', NS):
                tbl = gf.find('.//a:tbl', NS)
                if tbl is None:
                    continue
                for r_i, row in enumerate(rows):
                    for c_i, cell_text in enumerate(row):
                        if cell_text is not None:
                            replace_table_cell(tbl, r_i, c_i, cell_text)

        tree.write(path, xml_declaration=True, encoding='UTF-8')


def get_markitdown_text(pptx_path: str) -> str:
    """markitdown으로 텍스트 추출"""
    try:
        from markitdown import MarkItDown
        md = MarkItDown()
        result = md.convert(pptx_path)
        return result.text_content
    except Exception:
        return analyze_template_summary(pptx_path)
