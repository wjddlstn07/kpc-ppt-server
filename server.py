"""
server.py - PPT 자동 생성 Flask 서버
"""
import io
import json
import os
import shutil
import tempfile
import uuid

import anthropic
from flask import Flask, after_this_request, jsonify, request, send_file
from werkzeug.utils import secure_filename

from ppt_builder import analyze_template, build_ppt

app = Flask(__name__)

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
CLAUDE_MODEL = "claude-sonnet-4-20250514"
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50 MB
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


# ──────────────────────────────────────────────
# 라우트
# ──────────────────────────────────────────────

@app.route("/", methods=["GET"])
def health_check():
    return jsonify({"status": "ok", "message": "PPT 생성 서버 가동 중"})


@app.route("/analyze", methods=["POST"])
def analyze():
    if "template" not in request.files:
        return jsonify({"error": "template 파일이 필요합니다 (multipart/form-data)"}), 400

    template_file = request.files["template"]

    if not template_file.filename or not template_file.filename.lower().endswith(".pptx"):
        return jsonify({"error": "템플릿 파일은 .pptx 형식이어야 합니다"}), 400

    tmp_dir = tempfile.mkdtemp(prefix="ppt_analyze_")
    template_path = os.path.join(tmp_dir, "template.pptx")

    try:
        template_file.save(template_path)
        result = analyze_template(template_path)
    except Exception as e:
        return jsonify({"error": f"템플릿 분석 실패: {e}"}), 500
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return jsonify(result)


@app.route("/generate-ppt", methods=["POST"])
def generate_ppt():
    # ── 입력 검증 ──────────────────────────────
    if "template" not in request.files:
        return jsonify({"error": "template 파일이 필요합니다 (multipart/form-data)"}), 400

    template_file = request.files["template"]
    content = request.form.get("content", "").strip()
    title = request.form.get("title", "").strip()

    if not template_file.filename or not template_file.filename.lower().endswith(".pptx"):
        return jsonify({"error": "템플릿 파일은 .pptx 형식이어야 합니다"}), 400

    if not content:
        return jsonify({"error": "content(프로젝트 내용 텍스트)가 필요합니다"}), 400

    if not ANTHROPIC_API_KEY:
        return jsonify({"error": "서버에 ANTHROPIC_API_KEY가 설정되지 않았습니다"}), 500

    # ── 임시 디렉토리 설정 ─────────────────────
    tmp_dir = tempfile.mkdtemp(prefix="ppt_gen_")
    template_path = os.path.join(tmp_dir, "template.pptx")

    output_name = secure_filename(title) if title else f"presentation_{uuid.uuid4().hex[:8]}"
    if not output_name.lower().endswith(".pptx"):
        output_name += ".pptx"
    output_path = os.path.join(tmp_dir, output_name)

    try:
        # 1) 템플릿 저장
        template_file.save(template_path)

        # 2) 템플릿 분석
        template_analysis = analyze_template(template_path)

        # 3) Claude API → 슬라이드 JSON 생성
        slides_json = generate_slides_json(content, template_analysis)

        # 4) python-pptx로 PPT 생성
        build_ppt(template_path, slides_json, output_path)

        # 5) 파일을 메모리로 읽어 반환 (tmp 파일 즉시 정리 가능하도록)
        with open(output_path, "rb") as f:
            ppt_bytes = io.BytesIO(f.read())

    except json.JSONDecodeError as e:
        return jsonify({"error": f"Claude 응답을 JSON으로 파싱하지 못했습니다: {e}"}), 500
    except anthropic.APIError as e:
        return jsonify({"error": f"Claude API 오류: {e}"}), 502
    except Exception as e:
        return jsonify({"error": f"PPT 생성 실패: {e}"}), 500
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    ppt_bytes.seek(0)
    return send_file(
        ppt_bytes,
        as_attachment=True,
        download_name=output_name,
        mimetype=(
            "application/vnd.openxmlformats-officedocument"
            ".presentationml.presentation"
        ),
    )


# ──────────────────────────────────────────────
# Claude API 호출
# ──────────────────────────────────────────────

SYSTEM_PROMPT = """\
당신은 프레젠테이션 전문가입니다. 주어진 프로젝트 내용을 분석하여 논리적이고 시각적으로 효과적인 PPT 슬라이드 구조를 생성합니다.

반드시 아래 JSON 형식으로만 응답하세요. 마크다운 코드블록(```) 없이 순수 JSON만 출력합니다.

{
  "slides": [
    {
      "layout": "레이아웃명",
      "title": "슬라이드 제목",
      "subtitle": "부제목 (표지 슬라이드에만 사용)",
      "bullets": ["핵심 항목1", "핵심 항목2", "핵심 항목3"],
      "content": "단락형 내용 (bullets 대신 사용할 때만)",
      "notes": "발표자 노트 (선택)"
    }
  ]
}

규칙:
- 첫 번째 슬라이드는 반드시 표지 (아래 "실제 템플릿 슬라이드 구조"에서 index=0인 슬라이드의 layout_name 사용)
- 마지막 슬라이드는 마무리 / Q&A 슬라이드
- bullets와 content는 하나만 사용 (bullets 우선)
- 슬라이드당 bullets는 최대 5개
- layout 값은 반드시 "실제 템플릿 슬라이드 구조"의 layout_name 중 하나를 철자 그대로 사용할 것
  - "사용 가능한 레이아웃 목록"은 전체 후보이며, 실제 템플릿에서 쓰인 레이아웃을 우선적으로 선택할 것
  - 적합한 레이아웃이 없을 경우에만 "사용 가능한 레이아웃 목록"에서 선택
"""


def generate_slides_json(content: str, template_analysis: dict) -> dict:
    """Claude API를 호출하여 content 텍스트를 슬라이드 JSON으로 변환합니다."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    layout_names = [l["name"] for l in template_analysis.get("layouts", [])]
    fonts = template_analysis.get("fonts", [])
    slide_size = (
        f"{template_analysis.get('slide_width_inches', 13.33):.2f}\" × "
        f"{template_analysis.get('slide_height_inches', 7.5):.2f}\""
    )
    existing_slides = template_analysis.get("existing_slides", [])

    existing_slides_lines = []
    for s in existing_slides:
        shapes_preview = ", ".join(
            f'"{sh["name"]}: {sh["text_preview"]}"' for sh in s["shapes"]
        ) if s["shapes"] else "없음"
        existing_slides_lines.append(
            f'  - index={s["index"]}, layout_name="{s["layout_name"]}", shapes=[{shapes_preview}]'
        )
    existing_slides_str = "\n".join(existing_slides_lines) if existing_slides_lines else "  (없음)"

    user_message = f"""\
다음 정보를 바탕으로 PPT 슬라이드 JSON을 생성해주세요.

## 템플릿 정보
- 슬라이드 크기: {slide_size}
- 사용 가능한 레이아웃 목록: {json.dumps(layout_names, ensure_ascii=False)}
- 템플릿 폰트: {', '.join(fonts) if fonts else '기본 폰트'}

## 실제 템플릿 슬라이드 구조
(템플릿에 원래 포함된 슬라이드 — layout_name을 layout 값으로 그대로 사용하세요)
{existing_slides_str}

## 프로젝트 내용
{content}

위 내용을 효과적인 프레젠테이션 구조로 변환한 JSON을 출력하세요."""

    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )

    raw = message.content[0].text.strip()

    # 마크다운 코드블록 방어 처리
    if raw.startswith("```"):
        lines = raw.splitlines()
        raw = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])

    return json.loads(raw)


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG", "0") == "1")
