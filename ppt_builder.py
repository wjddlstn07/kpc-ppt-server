"""
ppt_builder.py - python-pptx 기반 템플릿 분석 및 PPT 생성 모듈
"""
from pptx import Presentation
from pptx.util import Pt


# ──────────────────────────────────────────────
# 1. 템플릿 분석
# ──────────────────────────────────────────────

def analyze_template(pptx_path: str) -> dict:
    """
    .pptx 템플릿을 분석하여 슬라이드 구조, 레이아웃, 색상, 폰트 정보를 반환합니다.
    """
    prs = Presentation(pptx_path)

    return {
        "slide_width_inches": round(prs.slide_width.inches, 2),
        "slide_height_inches": round(prs.slide_height.inches, 2),
        "layouts": _get_layouts(prs),
        "fonts": _get_fonts(prs),
        "existing_slides": _get_existing_slides(prs),
    }


def _get_layouts(prs: Presentation) -> list:
    layouts = []
    for idx, layout in enumerate(prs.slide_layouts):
        placeholders = []
        for ph in layout.placeholders:
            placeholders.append({
                "idx": ph.placeholder_format.idx,
                "type": str(ph.placeholder_format.type).split(".")[-1],
                "name": ph.name,
            })
        layouts.append({
            "index": idx,
            "name": layout.name,
            "placeholders": placeholders,
        })
    return layouts


def _get_fonts(prs: Presentation) -> list:
    seen = set()
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.name:
                        seen.add(run.font.name)
    return sorted(seen)


def _get_existing_slides(prs: Presentation) -> list:
    result = []
    for i, slide in enumerate(prs.slides):
        shapes_info = [
            {"name": s.name, "text_preview": s.text[:80].replace("\n", " ")}
            for s in slide.shapes
            if s.has_text_frame
        ]
        result.append({
            "index": i,
            "layout_name": slide.slide_layout.name,
            "shapes": shapes_info,
        })
    return result


# ──────────────────────────────────────────────
# 2. PPT 생성
# ──────────────────────────────────────────────

def build_ppt(template_path: str, slides_json: dict, output_path: str) -> str:
    """
    템플릿 스타일을 유지하면서 slides_json 내용으로 새 PPT를 생성합니다.
    """
    prs = Presentation(template_path)
    _remove_all_slides(prs)

    slides_data = slides_json.get("slides", [])
    if not slides_data:
        raise ValueError("slides_json에 슬라이드 데이터가 없습니다.")

    for slide_idx, slide_data in enumerate(slides_data):
        layout_name = slide_data.get("layout", "")
        print(f"\n[슬라이드 {slide_idx + 1}] 요청 layout_name: '{layout_name}'")

        layout = _find_layout(prs, layout_name)
        print(f"[슬라이드 {slide_idx + 1}] 선택된 레이아웃: '{layout.name}'")

        slide = prs.slides.add_slide(layout)

        ph_idxs = [ph.placeholder_format.idx for ph in slide.placeholders]
        print(f"[슬라이드 {slide_idx + 1}] 플레이스홀더 idx 목록: {ph_idxs}")

        _fill_slide(slide, slide_data)

        filled = {}
        for ph in slide.placeholders:
            idx = ph.placeholder_format.idx
            try:
                text = ph.text_frame.text.strip()
                filled[idx] = bool(text)
            except Exception:
                filled[idx] = False
        print(f"[슬라이드 {slide_idx + 1}] 텍스트 채움 여부 (idx: 채워짐): {filled}")

    prs.save(output_path)
    return output_path


def _remove_all_slides(prs: Presentation) -> None:
    """프레젠테이션의 모든 슬라이드를 제거합니다 (레이아웃·마스터는 유지)."""
    R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    sld_id_lst = prs.slides._sldIdLst

    before_count = len(prs.slides)
    print(f"[_remove_all_slides] 제거 전 슬라이드 수: {before_count}")

    for sld_id_elem in list(sld_id_lst):
        rId = sld_id_elem.get(R_ID)
        try:
            prs.part.drop_rel(rId)
        except Exception:
            pass
        sld_id_lst.remove(sld_id_elem)

    after_count = len(prs.slides)
    print(f"[_remove_all_slides] 제거 후 슬라이드 수: {after_count}")


def _find_layout(prs: Presentation, layout_name: str):
    """레이아웃 이름으로 슬라이드 레이아웃을 찾습니다. 없으면 두 번째 레이아웃 반환."""
    name_lower = (layout_name or "").lower()

    # 1) 이름이 포함된 레이아웃 우선 탐색
    for layout in prs.slide_layouts:
        if name_lower and name_lower in layout.name.lower():
            return layout

    # 2) 한/영 키워드 → 인덱스 매핑
    KEYWORD_MAP = [
        ({"title slide", "제목 슬라이드", "표지"}, 0),
        ({"title and content", "제목 및 내용", "내용"}, 1),
        ({"section header", "구역 머리글", "섹션"}, 2),
        ({"two content", "두 콘텐츠", "two column", "두 열"}, 3),
        ({"blank", "빈 화면", "빈"}, 6),
    ]
    for keywords, fallback_idx in KEYWORD_MAP:
        if any(kw in name_lower for kw in keywords):
            if fallback_idx < len(prs.slide_layouts):
                return prs.slide_layouts[fallback_idx]

    # 3) 기본값: 플레이스홀더가 하나 이상 있는 첫 번째 레이아웃
    for layout in prs.slide_layouts:
        if layout.placeholders:
            return layout

    # 4) 플레이스홀더가 있는 레이아웃조차 없으면 첫 번째 레이아웃 반환
    return prs.slide_layouts[0]


def _fill_slide(slide, slide_data: dict) -> None:
    """슬라이드 플레이스홀더를 slide_data로 채웁니다."""
    title_text = slide_data.get("title", "")
    bullets = slide_data.get("bullets", [])
    content = slide_data.get("content", "")
    subtitle = slide_data.get("subtitle", "")
    notes_text = slide_data.get("notes", "")

    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx

        if idx == 0:  # 제목
            _safe_set_text(ph, title_text)

        elif idx == 1:  # 본문
            tf = ph.text_frame
            tf.clear()
            if bullets:
                for i, bullet in enumerate(bullets):
                    if i == 0:
                        tf.paragraphs[0].text = bullet
                    else:
                        para = tf.add_paragraph()
                        para.text = bullet
                        para.level = 0
            elif content:
                tf.paragraphs[0].text = content

        elif idx in (12, 13) or "subtitle" in ph.name.lower():
            _safe_set_text(ph, subtitle or content)

    if notes_text:
        try:
            slide.notes_slide.notes_text_frame.text = notes_text
        except Exception:
            pass


def _safe_set_text(shape, text: str) -> None:
    if not text:
        return
    try:
        shape.text = text
    except Exception:
        try:
            shape.text_frame.paragraphs[0].text = text
        except Exception:
            pass
