from __future__ import annotations

import io
import json
import re
import unicodedata
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests

AS_OF = "2026-09-16"
JPX_PAGE = "https://www.jpx.co.jp/markets/statistics-equities/misc/01.html"
OUT_DIR = Path(__file__).resolve().parent
TEMPLATE = OUT_DIR / "template.html"

TARGET_MARKETS = {
    "プライム（内国株式）",
    "スタンダード（内国株式）",
    "グロース（内国株式）",
    "PRO Market",
}

PATCH_REMOVE = {"3480", "1909", "2180"}
PATCH_ADD = [
    {
        "code": "618A",
        "name": "ＫＯＭＰＥＩＴＯ",
        "market": "グロース（内国株式）",
        "industry33": "サービス業",
        "industry17": "情報通信・サービスその他",
        "size": "",
        "source": "JPX new listing 2026-09-11",
    },
    {
        "code": "619A",
        "name": "オリバー",
        "market": "スタンダード（内国株式）",
        "industry33": "その他製品",
        "industry17": "情報通信・サービスその他",
        "size": "",
        "source": "JPX new listing 2026-09-16",
    },
    {
        "code": "621A",
        "name": "オーディオストック",
        "market": "グロース（内国株式）",
        "industry33": "サービス業",
        "industry17": "情報通信・サービスその他",
        "size": "",
        "source": "JPX new listing 2026-09-16",
    },
]

HEADERS = {"User-Agent": "Mozilla/5.0 industry-atlas-builder/1.0"}


def normalize_code(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return ""
    if isinstance(value, int):
        return str(value).zfill(4)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value)).zfill(4)
        return ""
    text = unicodedata.normalize("NFKC", str(value)).strip().upper()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    if text.isdigit() and len(text) < 4:
        text = text.zfill(4)
    return text


def find_excel_url() -> str:
    r = requests.get(JPX_PAGE, headers=HEADERS, timeout=45)
    r.raise_for_status()
    html = r.text
    candidates = re.findall(r'href=["\']([^"\']*data_j\.(?:xlsx|xls))["\']', html, flags=re.I)
    for href in candidates:
        return urljoin(JPX_PAGE, href)
    return "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xlsx"


def download_excel() -> tuple[bytes, str]:
    urls = [
        find_excel_url(),
        "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xlsx",
        "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq00000030ne-att/data_j.xlsx",
        "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls",
    ]
    last = None
    for url in dict.fromkeys(urls):
        try:
            r = requests.get(url, headers=HEADERS, timeout=60)
            r.raise_for_status()
            data = r.content
            if data.startswith(b"PK\x03\x04") or data.startswith(bytes.fromhex("d0cf11e0")):
                return data, url
        except Exception as exc:
            last = exc
    raise RuntimeError(f"JPX Excel download failed: {last}")


def read_jpx(data: bytes) -> pd.DataFrame:
    if data.startswith(b"PK\x03\x04"):
        return pd.read_excel(io.BytesIO(data), engine="openpyxl")
    return pd.read_excel(io.BytesIO(data), engine="xlrd")


def clean_text(v: object) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def build_records(df: pd.DataFrame) -> tuple[list[dict], str]:
    required = [
        "日付",
        "コード",
        "銘柄名",
        "市場・商品区分",
        "33業種区分",
        "17業種区分",
        "規模区分",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise RuntimeError(f"JPX columns missing: {missing}")

    base_dates = []
    records: dict[str, dict] = {}
    for _, row in df.iterrows():
        market = clean_text(row["市場・商品区分"])
        if market not in TARGET_MARKETS:
            continue
        code = normalize_code(row["コード"])
        if not code:
            continue
        raw_date = row["日付"]
        if not pd.isna(raw_date):
            s = re.sub(r"[^0-9]", "", str(raw_date))
            if len(s) >= 8:
                base_dates.append(s[:8])
        records[code] = {
            "code": code,
            "name": clean_text(row["銘柄名"]),
            "market": market,
            "industry33": clean_text(row["33業種区分"]),
            "industry17": clean_text(row["17業種区分"]),
            "size": clean_text(row["規模区分"]),
            "source": "JPX listed issues",
        }

    for code in PATCH_REMOVE:
        records.pop(code, None)
    for item in PATCH_ADD:
        records[item["code"]] = dict(item)

    base_date = max(base_dates) if base_dates else "20260831"
    return sorted(records.values(), key=lambda x: x["code"]), base_date


def main() -> None:
    data, source_url = download_excel()
    df = read_jpx(data)
    companies, base_date = build_records(df)

    payload = {
        "meta": {
            "title": "日本企業 業界図鑑 2026",
            "asOf": AS_OF,
            "baseDate": f"{base_date[:4]}-{base_date[4:6]}-{base_date[6:8]}",
            "source": "JPX 東証上場銘柄一覧",
            "sourceUrl": source_url,
            "companyCount": len(companies),
        },
        "companies": companies,
    }

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "companies.json").write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    template = TEMPLATE.read_text(encoding="utf-8")
    embedded = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</script>", "<\\/script>")
    html = template.replace("__EMBEDDED_DATA__", embedded)
    (OUT_DIR / "index.html").write_text(html, encoding="utf-8")
    (OUT_DIR / "offline.html").write_text(html, encoding="utf-8")

    print(f"built {len(companies)} companies; base={base_date}; as-of={AS_OF}")


if __name__ == "__main__":
    main()
