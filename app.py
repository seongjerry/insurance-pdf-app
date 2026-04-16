"""
보험 실적입력동의서 PDF 자동 기입 - 웹 애플리케이션
Flask 로컬 서버: http://localhost:5000
"""

from __future__ import annotations

import io
import logging
import os
import sys
from datetime import date
from functools import wraps
from pathlib import Path

from flask import (
    Flask,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

# ---------------------------------------------------------------------------
# 환경 변수 로드 (.env 파일 지원)
# ---------------------------------------------------------------------------
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # python-dotenv 없어도 동작

# ---------------------------------------------------------------------------
# 경로 감지: 개발 환경 vs PyInstaller EXE 번들
#   BASE_DIR   = 사용자 파일 위치 (template.pdf, config.json, output/)
#   BUNDLE_DIR = 코드·템플릿 번들 위치 (templates/, static/, pdf_filler.py)
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    # PyInstaller로 빌드된 EXE 실행 시
    BASE_DIR   = Path(sys.executable).parent   # EXE가 있는 폴더
    BUNDLE_DIR = Path(sys._MEIPASS)            # 임시 추출 번들 폴더
else:
    BASE_DIR   = Path(__file__).parent
    BUNDLE_DIR = BASE_DIR

sys.path.insert(0, str(BUNDLE_DIR))
from pdf_filler import AppConfig, ConfigLoader, PDFFiller, load_data, BatchProcessor

CONFIG_PATH  = BASE_DIR / "config.json"
TEMPLATE_PDF = BASE_DIR / "template.pdf"
OUTPUT_DIR   = BASE_DIR / "output"

app = Flask(
    __name__,
    template_folder=str(BUNDLE_DIR / "templates"),
    static_folder=str(BUNDLE_DIR / "static") if (BUNDLE_DIR / "static").exists() else None,
)
# SECRET_KEY: .env 또는 환경 변수에서 로드. 없으면 매 재시작마다 세션 초기화됨.
app.secret_key = os.environ.get("SECRET_KEY") or os.urandom(24)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# 인증 설정
#   ADMIN_PASSWORD 가 설정되어 있으면 로그인이 필요합니다.
#   설정되어 있지 않으면 인증 없이 접근 가능합니다 (개발/로컬 환경).
# ---------------------------------------------------------------------------
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if ADMIN_PASSWORD and not session.get("logged_in"):
            return redirect(url_for("login", next=request.path))
        return f(*args, **kwargs)
    return decorated


# 미처리 예외를 모두 로그에 기록 (EXE 환경 디버깅용)
@app.errorhandler(Exception)
def handle_all_errors(exc):
    import traceback
    logger.error("미처리 예외:\n%s", traceback.format_exc())
    return f"서버 오류: {exc}", 500


def load_config() -> AppConfig:
    return ConfigLoader.load(CONFIG_PATH)


# ---------------------------------------------------------------------------
# 인증 라우트
# ---------------------------------------------------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    """로그인 페이지"""
    # 이미 로그인된 경우 메인으로
    if session.get("logged_in"):
        return redirect(url_for("index"))

    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            session["logged_in"] = True
            next_url = request.args.get("next") or url_for("index")
            return redirect(next_url)
        error = "아이디 또는 비밀번호가 올바르지 않습니다."

    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# 라우트
# ---------------------------------------------------------------------------

@app.route("/")
@login_required
def index():
    """메인 폼 페이지"""
    today = date.today()
    return render_template(
        "index.html",
        today_year=today.year,
        today_month=f"{today.month:02d}",
        today_day=f"{today.day:02d}",
        error=request.args.get("error"),
    )


@app.route("/generate", methods=["POST"])
@login_required
def generate():
    """단건 폼 제출 → PDF 즉시 다운로드"""
    try:
        config = load_config()
    except FileNotFoundError as e:
        return redirect(url_for("index", error=f"설정 파일 오류: {e}"))

    if not TEMPLATE_PDF.exists():
        return redirect(url_for("index", error="template.pdf 파일이 없습니다."))

    data = {
        "보험회사명": request.form.get("보험회사명", "").strip(),
        "상품명":    request.form.get("상품명", "").strip(),
        "증권번호":  request.form.get("증권번호", "").strip(),
        "계약자명":  request.form.get("계약자명", "").strip(),
        "피보험자명": request.form.get("피보험자명", "").strip(),
        "계약_년도": request.form.get("계약_년도", "").strip(),
        "계약_월":   request.form.get("계약_월", "").strip(),
        "계약_일":   request.form.get("계약_일", "").strip(),
    }

    # 필수값 검증
    missing = [k for k in ["보험회사명", "상품명", "증권번호", "계약자명"] if not data[k]]
    if missing:
        return redirect(url_for("index", error=f"필수 항목 누락: {', '.join(missing)}"))

    try:
        filler = PDFFiller(TEMPLATE_PDF, config)
        filler.fill(data)

        buf = io.BytesIO()
        import fitz
        assert filler._doc is not None
        filler._doc.save(buf, garbage=4, deflate=True)
        filler._doc.close()
        filler._doc = None
        buf.seek(0)

        filename = f"동의서_{data['계약자명']}_{data['증권번호']}.pdf"
        logger.info("단건 생성: %s", filename)

        return send_file(
            buf,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=filename,
        )

    except Exception as exc:
        logger.error("PDF 생성 실패: %s", exc, exc_info=True)
        return redirect(url_for("index", error=f"PDF 생성 실패: {exc}"))


@app.route("/batch", methods=["POST"])
@login_required
def batch():
    """CSV/Excel 업로드 → 일괄 처리 후 합본 PDF 다운로드"""
    try:
        config = load_config()
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 500

    if not TEMPLATE_PDF.exists():
        return jsonify({"error": "template.pdf 파일이 없습니다."}), 500

    file = request.files.get("datafile")
    if not file or file.filename == "":
        return redirect(url_for("index", error="파일을 선택하세요."))

    suffix = Path(file.filename).suffix.lower()
    if suffix not in {".csv", ".xlsx", ".xls"}:
        return redirect(url_for("index", error="CSV 또는 Excel 파일만 지원합니다."))

    # 업로드 파일을 임시 저장
    import tempfile, pandas as pd
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp_path = Path(tmp.name)
        file.save(tmp_path)

    try:
        df = load_data(tmp_path)
    except Exception as exc:
        tmp_path.unlink(missing_ok=True)
        return redirect(url_for("index", error=f"파일 읽기 실패: {exc}"))
    finally:
        tmp_path.unlink(missing_ok=True)

    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        config.output.directory = str(OUTPUT_DIR)
        processor = BatchProcessor(TEMPLATE_PDF, config)
        saved = processor.run(df, merge=True)  # 항상 합본

        merge_pdf_path = saved[0]
        filename = merge_pdf_path.name
        logger.info("일괄 생성 완료: %s (%d건)", filename, len(df))

        return send_file(
            merge_pdf_path,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=filename,
        )

    except Exception as exc:
        logger.error("일괄 처리 실패: %s", exc, exc_info=True)
        return redirect(url_for("index", error=f"일괄 처리 실패: {exc}"))


@app.route("/download-template")
@login_required
def download_template():
    """샘플 CSV 파일 다운로드"""
    csv_content = (
        "보험회사명,상품명,증권번호,계약자명,피보험자명,계약_년도,계약_월,계약_일\n"
        "삼성생명,무배당종신보험(2024),20240001234,홍길동,홍길동,2026,04,16\n"
        "한화생명,변액유니버셜보험,20240005678,김철수,이순희,2026,04,17\n"
    )
    buf = io.BytesIO(csv_content.encode("utf-8-sig"))  # BOM 포함 (Excel 한글 호환)
    return send_file(
        buf,
        mimetype="text/csv",
        as_attachment=True,
        download_name="동의서_샘플데이터.csv",
    )


@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "template_exists": TEMPLATE_PDF.exists(),
        "config_exists": CONFIG_PATH.exists(),
    })


# ---------------------------------------------------------------------------
# 실행
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("=" * 55)
    print("  인카금융서비스 동의서 자동 기입 시스템")
    print("  브라우저에서 http://localhost:5000 으로 접속하세요")
    print("=" * 55)
    app.run(host="0.0.0.0", port=5000, debug=False)
