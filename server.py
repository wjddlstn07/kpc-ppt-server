"""
server.py - PPT 자동 생성 Flask 서버
"""
import io
import json
import os
import shutil
import tempfile
import uuid

from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename

from ppt_builder import analyze_template, build_ppt

app = Flask(__name__)

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

        # 2) python-pptx로 PPT 생성
        build_ppt(template_path, slides_json, output_path)

        # 3) 파일을 메모리로 읽어 반환 (tmp 파일 즉시 정리 가능하도록)
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


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG", "0") == "1")
