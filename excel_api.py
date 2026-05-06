from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from openpyxl import load_workbook
import io
import json
import os
from copy import deepcopy

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://kawatsu624.hiho.jp",
        "http://kawatsu624.hiho.jp",
        "http://localhost:5173",
        "http://127.0.0.1:5173",
    ],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_PATH = os.path.join(
    BASE_DIR,
    "templates",
    "08_ﾓﾆﾀﾘﾝｸﾞ結果印刷_ACE既定（モニタリング）_202404.xlsx"
)


class ExcelRequest(BaseModel):
    summary_json: str | None = None
    text: str | None = None
    transcript: str | None = None


DEFAULT_REPORT = {
    "実施日": "",
    "前回実施日": "",
    "次回予定日": "",
    "利用者名": "",
    "利用者名カナ": "",
    "出力氏名": "",
    "住所": "",
    "住所1": "",
    "住所2": "",
    "住所3": "",
    "電話番号": "",
    "電話番号1": "",
    "担当名": "",
    "ケアマネ姓名": "",
    "お話伺った人1": "",
    "お話伺った人2": "",
    "お話伺った人3": "",
    "お話伺った人その他": "",
    "確認方法1": "",
    "確認方法2": "",
    "専門相談員による結果": "",
    "福祉用具利用目標": [],
    "商品一覧": [],
    "身体状況の変化1": "",
    "身体状況の変化2": "",
    "身体状況の変化備考": "",
    "ご家族状況の変化1": "",
    "ご家族状況の変化2": "",
    "ご家族状況の変化備考": "",
    "お気持ちの変化1": "",
    "お気持ちの変化2": "",
    "お気持ちの変化備考": "",
    "生活状況の変化1": "",
    "生活状況の変化2": "",
    "生活状況の変化備考": "",
    "見直しの必要性1": "",
    "見直しの必要性2": ""
}

DEFAULT_GOAL = {
    "目標": "",
    "達成度1": "",
    "達成度2": "",
    "達成度3": "",
    "備考": ""
}

DEFAULT_PRODUCT = {
    "サービス名": "",
    "利用開始日": "",
    "商品名": "",
    "使用状況の問題1": "",
    "点検結果1": "",
    "今後の方針1": "",
    "使用状況の問題2": "",
    "点検結果2": "",
    "今後の方針2": "",
    "モニタリング備考": ""
}


def merge_dict(defaults: dict, actual: dict) -> dict:
    result = deepcopy(defaults)

    if not isinstance(actual, dict):
        return result

    for key, value in actual.items():
        result[key] = value

    return result


def normalize_report(raw: dict) -> dict:
    report = merge_dict(DEFAULT_REPORT, raw if isinstance(raw, dict) else {})

    goals = report.get("福祉用具利用目標")
    if not isinstance(goals, list):
        goals = []

    normalized_goals = []
    for g in goals[:4]:
        normalized_goals.append(
            merge_dict(DEFAULT_GOAL, g if isinstance(g, dict) else {})
        )
    report["福祉用具利用目標"] = normalized_goals

    products = report.get("商品一覧")
    if not isinstance(products, list):
        products = []

    normalized_products = []
    for p in products[:8]:
        normalized_products.append(
            merge_dict(DEFAULT_PRODUCT, p if isinstance(p, dict) else {})
        )
    report["商品一覧"] = normalized_products

    for key, value in list(report.items()):
        if value is None:
            report[key] = ""

    return report


def safe_json(text: str) -> dict:
    try:
        raw = json.loads(text)
        return normalize_report(raw)
    except Exception:
        fallback = deepcopy(DEFAULT_REPORT)
        fallback["専門相談員による結果"] = text
        return fallback


def set_if_exists(ws, cell_ref: str, value):
    ws[cell_ref] = "" if value is None else value


def fill_monitoring_sheet(wb, data: dict):
    ws = wb["レイアウト_モニタリング"]

    mapping = {
        "AC3": data.get("実施日", ""),
        "AC4": data.get("前回実施日", ""),
        "AP3": data.get("利用者名", ""),
        "AP2": data.get("利用者名カナ", ""),
        "AP4": data.get("出力氏名", ""),
        "AC5": data.get("お話伺った人1", ""),
        "AG5": data.get("お話伺った人2", ""),
        "AJ5": data.get("お話伺った人3", ""),
        "AL5": data.get("お話伺った人その他", ""),
        "AC6": data.get("確認方法1", ""),
        "AG6": data.get("確認方法2", ""),
        "AC8": data.get("担当名", ""),
        "AP5": data.get("住所", ""),
        "AR15": data.get("住所1", ""),
        "AR16": data.get("住所2", ""),
        "AR17": data.get("住所3", ""),
        "AP6": data.get("電話番号", ""),
        "AI15": data.get("電話番号1", ""),
        "AI16": data.get("ケアマネ姓名", ""),
        "J91": data.get("専門相談員による結果", ""),
        "AA98": data.get("次回予定日", ""),
        "F85": data.get("身体状況の変化1", ""),
        "F86": data.get("身体状況の変化2", ""),
        "J85": data.get("身体状況の変化備考", ""),
        "Z85": data.get("ご家族状況の変化1", ""),
        "Z86": data.get("ご家族状況の変化2", ""),
        "AD85": data.get("ご家族状況の変化備考", ""),
        "F87": data.get("お気持ちの変化1", ""),
        "F88": data.get("お気持ちの変化2", ""),
        "J87": data.get("お気持ちの変化備考", ""),
        "Z87": data.get("生活状況の変化1", ""),
        "Z88": data.get("生活状況の変化2", ""),
        "AD87": data.get("生活状況の変化備考", ""),
        "F91": data.get("見直しの必要性1", ""),
        "F94": data.get("見直しの必要性2", ""),
    }

    for cell_ref, value in mapping.items():
        set_if_exists(ws, cell_ref, value)

    goal_rows = [20, 23, 26, 29]
    goals = data.get("福祉用具利用目標", [])

    for i, row in enumerate(goal_rows):
        g = goals[i] if i < len(goals) else DEFAULT_GOAL

        set_if_exists(ws, f"C{row}", g.get("目標", ""))
        set_if_exists(ws, f"V{row}", g.get("達成度1", ""))
        set_if_exists(ws, f"V{row + 1}", g.get("達成度2", ""))
        set_if_exists(ws, f"V{row + 2}", g.get("達成度3", ""))
        set_if_exists(ws, f"Z{row}", g.get("備考", ""))

    product_rows = [35, 41, 47, 53, 59, 65, 71, 77]
    products = data.get("商品一覧", [])

    for i, row in enumerate(product_rows):
        p = products[i] if i < len(products) else DEFAULT_PRODUCT

        set_if_exists(ws, f"C{row}", p.get("サービス名", ""))
        set_if_exists(ws, f"O{row}", p.get("利用開始日", ""))
        set_if_exists(ws, f"AB{row}", p.get("モニタリング備考", ""))

        set_if_exists(ws, f"R{row}", p.get("使用状況の問題1", ""))
        set_if_exists(ws, f"U{row}", p.get("点検結果1", ""))
        set_if_exists(ws, f"Y{row}", p.get("今後の方針1", ""))

        item_row = row + 3

        set_if_exists(ws, f"C{item_row}", p.get("商品名", ""))
        set_if_exists(ws, f"R{item_row}", p.get("使用状況の問題2", ""))
        set_if_exists(ws, f"U{item_row}", p.get("点検結果2", ""))
        set_if_exists(ws, f"Y{item_row}", p.get("今後の方針2", ""))

    return wb


@app.get("/")
def root():
    return {
        "ok": True,
        "service": "monitoring-excel-api",
    }


@app.get("/health")
def health():
    return {
        "ok": True,
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "template_path": TEMPLATE_PATH,
        "version": "excel_from_summary_json_text_or_transcript_v1",
    }


@app.post("/api/report-excel")
def report_excel(req: ExcelRequest):
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail="template file not found")

    raw_text = (
        req.summary_json
        or req.text
        or req.transcript
        or ""
    ).strip()

    if not raw_text:
        raise HTTPException(
            status_code=400,
            detail="summary_json or text or transcript empty"
        )

    try:
        report = safe_json(raw_text)

        wb = load_workbook(TEMPLATE_PATH)
        wb = fill_monitoring_sheet(wb, report)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": 'attachment; filename="monitoring_report.xlsx"'
            },
        )

    except HTTPException:
        raise

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
