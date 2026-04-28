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
from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename

from ppt_builder import analyze_template, build_ppt

app = Flask(__name__)

MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50 MB
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


# ──────────────────────────────────────────────
# Claude API 시스템 프롬프트
# ──────────────────────────────────────────────

PPT_SYSTEM_PROMPT = """당신은 컨설팅 프로젝트 PPT 구조화 전문가입니다.

사용자가 입력한 프로젝트 내용(줄글)과 템플릿 분석 결과를 바탕으로,
슬라이드 구성 JSON을 생성해주세요.

## 출력 규칙
- 반드시 JSON만 출력 (설명 텍스트, 마크다운 코드블록 없음)
- 한국어로 작성
- 슬라이드 수는 입력 내용 기준으로 자동 결정 (보통 5~8장)
- 각 슬라이드는 title, summary, bullets, pageNum, notes 구성

## JSON 스키마
{
  "slides": [
    {
      "title": "01 슬라이드 제목",
      "summary": "이 슬라이드의 핵심 한 문장 요약",
      "bullets": [
        "첫 번째 핵심 항목",
        "두 번째 핵심 항목",
        "세 번째 핵심 항목"
      ],
      "pageNum": 1,
      "notes": "발표자 노트 (선택사항)"
    }
  ]
}

## 슬라이드 구성 원칙
- title: "번호 제목" 형식 (예: "01 프로젝트 배경")
- summary: 슬라이드 전체를 한 문장으로 요약
- bullets: 핵심 항목 3~5개, 각 항목은 간결하게 (1~2줄)
- pageNum: 순서대로 자동 부여
- 제안서 구조: 배경 → 목적 → 범위 → 로드맵 → 수행실적 → 추진일정
- 보고서 구조: 현황 → 분석 → 문제점 → 전략 → 실행방안 → 기대효과

## 중요
- bullets 항목은 템플릿의 본문 텍스트 영역에 들어갈 내용
- 너무 길면 잘리므로 각 항목은 40자 이내로 작성
- 내용이 많으면 슬라이드를 나눠서 구성"""


# ──────────────────────────────────────────────
# Claude API 호출 함수
# ──────────────────────────────────────────────

def call_claude_api(content: str, template_info: dict = None) -> dict:
    """
    프로젝트 내용(줄글)을 slides_json으로 변환합니다.
    template_info가 있으면 템플릿 구조를 참고하여 슬라이드 수를 조정합니다.
    """
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    # 템플릿 정보가 있으면 슬라이드 수 힌트 추가
    user_message = content
    if template_info:
        slide_count = len(template_info.get("existing_slides", []))
        if slide_count > 0:
            user_message = (
                f"[템플릿 정보]\n"
                f"- 템플릿 슬라이드 수: {slide_count}장\n"
                f"- 폰트: {', '.join(template_info.get('fonts', []))}\n\n"
                f"[프로젝트 내용]\n{content}"
            )

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        system=PPT_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )

    raw_text = message.content[0].text.strip()

    # JSON 파싱 (코드블록 제거 후)
    clean_text = raw_text.replace("```json", "").replace("```", "").strip()
    return json.loads(clean_text)


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
    """
    기존 엔드포인트: slides_json을 직접 받아서 PPT 생성
    (Make.com에서 직접 JSON을 전달할 때 사용)
    """
    if "template" not in request.files:
        return jsonify({"error": "template 파일이 필요합니다 (multipart/form-data)"}), 400

    slides_json_str = request.form.get("slides_json", "").strip()
    if not slides_json_str:
        return jsonify({"error": "slides_json이 필요합니다"}), 400

    template_file = request.files["template"]
    title = request.form.get("title", "").strip()

    if not template_file.filename or not template_file.filename.lower().endswith(".pptx"):
        return jsonify({"error": "템플릿 파일은 .pptx 형식이어야 합니다"}), 400

    try:
        slides_json = json.loads(slides_json_str)
    except json.JSONDecodeError as e:
        return jsonify({"error": f"slides_json 파싱 실패: {e}"}), 400

    if not slides_json.get("slides"):
        return jsonify({"error": "slides_json에 slides 배열이 없습니다"}), 400

    tmp_dir = tempfile.mkdtemp(prefix="ppt_gen_")
    template_path = os.path.join(tmp_dir, "template.pptx")

    output_name = secure_filename(title) if title else f"presentation_{uuid.uuid4().hex[:8]}"
    if not output_name.lower().endswith(".pptx"):
        output_name += ".pptx"
    output_path = os.path.join(tmp_dir, output_name)

    try:
        template_file.save(template_path)
        build_ppt(template_path, slides_json, output_path)

        with open(output_path, "rb") as f:
            ppt_bytes = io.BytesIO(f.read())

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


@app.route("/generate-ppt-auto", methods=["POST"])
def generate_ppt_auto():
    """
    핵심 엔드포인트: 줄글 입력 → Claude API → slides_json → PPTX 생성

    입력 (multipart/form-data):
      - template: .pptx 파일
      - content:  프로젝트 내용 줄글 (필수)
      - title:    파일명 (선택)

    출력:
      - .pptx 파일
    """
    # ── 입력 검증 ──────────────────────────────
    if "template" not in request.files:
        return jsonify({"error": "template 파일이 필요합니다 (multipart/form-data)"}), 400

    content = request.form.get("content", "").strip()
    if not content:
        return jsonify({"error": "content(프로젝트 내용)가 필요합니다"}), 400

    template_file = request.files["template"]
    title = request.form.get("title", "").strip()

    if not template_file.filename or not template_file.filename.lower().endswith(".pptx"):
        return jsonify({"error": "템플릿 파일은 .pptx 형식이어야 합니다"}), 400

    if not os.environ.get("ANTHROPIC_API_KEY"):
        return jsonify({"error": "ANTHROPIC_API_KEY 환경변수가 설정되지 않았습니다"}), 500

    # ── 임시 디렉토리 설정 ─────────────────────
    tmp_dir = tempfile.mkdtemp(prefix="ppt_auto_")
    template_path = os.path.join(tmp_dir, "template.pptx")

    output_name = secure_filename(title) if title else f"presentation_{uuid.uuid4().hex[:8]}"
    if not output_name.lower().endswith(".pptx"):
        output_name += ".pptx"
    output_path = os.path.join(tmp_dir, output_name)

    try:
        # 1) 템플릿 저장
        template_file.save(template_path)

        # 2) 템플릿 분석 (슬라이드 수 파악 → Claude 힌트 제공)
        template_info = analyze_template(template_path)

        # 3) Claude API → slides_json 생성
        slides_json = call_claude_api(content, template_info)

        # 4) ppt_builder → PPTX 생성
        build_ppt(template_path, slides_json, output_path)

        # 5) 메모리로 읽어 반환
        with open(output_path, "rb") as f:
            ppt_bytes = io.BytesIO(f.read())

    except json.JSONDecodeError as e:
        return jsonify({"error": f"Claude API 응답 JSON 파싱 실패: {e}"}), 500
    except anthropic.APIError as e:
        return jsonify({"error": f"Claude API 오류: {e}"}), 500
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
# 진입점
# ──────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG", "0") == "1")
