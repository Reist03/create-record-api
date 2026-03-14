from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import io
import json
import os
from copy import deepcopy

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(
    BASE_DIR,
    "08_ﾓﾆﾀﾘﾝｸﾞ結果印刷_ACE既定（モニタリング）_202404.xlsx"
)


class TranscriptRequest(BaseModel):
    transcript: str


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


def build_prompt(transcript: str) -> str:
    return f"""
"# 指示
あなたは福祉用具レンタル事業所のモニタリング担当です。
次の「福祉用具レンタル事務所のモニタリング担当者」と「福祉施設利用者」による対話記録から、モニタリング帳票入力用の以下のJSON形式で出力してください。"

必ずJSONのみを返してください。
すべてのキーを必ず出力してください。省略は禁止です。
値が不明な場合は空文字、配列項目は空配列で返してください。

JSON形式:
{json.dumps(DEFAULT_REPORT, ensure_ascii=False, indent=2)}

配列要素の形式:
福祉用具利用目標 の1件:
{json.dumps(DEFAULT_GOAL, ensure_ascii=False, indent=2)}

商品一覧 の1件:
{json.dumps(DEFAULT_PRODUCT, ensure_ascii=False, indent=2)}

補足:
- 福祉用具利用目標 は最大4件でよい
- 商品一覧 は最大8件でよい
- 達成度や変化項目は「☑」「✓」「有」「該当」など短い表現で可
- 帳票向けに短く簡潔にしてください

文字起こし:
{transcript}
""".strip()


def extract_text_from_response(res) -> str:
    if hasattr(res, "output_text") and res.output_text:
        return res.output_text

    try:
        parts = []
        for out in getattr(res, "output", []):
            for content in getattr(out, "content", []):
                text_value = getattr(content, "text", None)
                if text_value:
                    parts.append(text_value)
        joined = "\n".join(parts).strip()
        if joined:
            return joined
    except Exception:
        pass

    return str(res)


def merge_dict(defaults: dict, actual: dict) -> dict:
    result = deepcopy(defaults)
    if not isinstance(actual, dict):
        return result

    for key, value in actual.items():
        if key in result:
            result[key] = value
        else:
            result[key] = value
    return result


def normalize_report(raw: dict) -> dict:
    report = merge_dict(DEFAULT_REPORT, raw if isinstance(raw, dict) else {})

    goals = report.get("福祉用具利用目標")
    if not isinstance(goals, list):
        goals = []
    normalized_goals = []
    for g in goals[:4]:
        normalized_goals.append(merge_dict(DEFAULT_GOAL, g if isinstance(g, dict) else {}))
    report["福祉用具利用目標"] = normalized_goals

    products = report.get("商品一覧")
    if not isinstance(products, list):
        products = []
    normalized_products = []
    for p in products[:8]:
        normalized_products.append(merge_dict(DEFAULT_PRODUCT, p if isinstance(p, dict) else {}))
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


# 追加：文字起こしを「。」ごとに分割して1行ずつにする
def split_transcript(text: str):
    if not text:
        return []

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("。", "。\n")
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    return lines


def fill_monitoring_sheet(wb, data: dict, transcript_text: str = ""):
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
        set_if_exists(ws, f"V{row+1}", g.get("達成度2", ""))
        set_if_exists(ws, f"V{row+2}", g.get("達成度3", ""))
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

    # 追加：文字起こしした内容を BJ1, BJ2, BJ3... に表示
    transcript_lines = split_transcript(transcript_text)
    start_row = 1

    for i, line in enumerate(transcript_lines):
        cell = f"BJ{start_row + i}"
        ws[cell] = line
        # 折り返しを使わない
        ws[cell].alignment = Alignment(wrap_text=False)

    return wb


@app.get("/health")
def health():
    return {
        "ok": True,
        "has_api_key": bool(os.getenv("OPENAI_API_KEY")),
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "template_path": TEMPLATE_PATH,
    }


@app.post("/api/report-excel")
def report_excel(req: TranscriptRequest):
    transcript = (req.transcript or "").strip()

    if not transcript:
        raise HTTPException(status_code=400, detail="transcript empty")

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail="template file not found")

    try:
        prompt = build_prompt(transcript)

        res = client.responses.create(
            model="gpt-5-mini",
            input=prompt,
        )

        text = extract_text_from_response(res)
        report = safe_json(text)

        wb = load_workbook(TEMPLATE_PATH)
        wb = fill_monitoring_sheet(wb, report, transcript)

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

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))