"""입찰공고 첨부파일(엑셀 등) 조회·다운로드·내용 추출 도구."""

import io
import logging
import re
import zipfile
from typing import Any, Dict, List, Optional
from urllib.parse import unquote

import requests

from ..core.config import get_service_key

logger = logging.getLogger("naramarket.attachments")

BID_BASE = "https://apis.data.go.kr/1230000/ad/BidPublicInfoService"
BID_OPERATIONS = {
    "물품": "getBidPblancListInfoThng",
    "용역": "getBidPblancListInfoServc",
    "공사": "getBidPblancListInfoCnstwk",
    "외자": "getBidPblancListInfoFrgcpt",
}
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; naramarket-mcp/0.2)",
    "Referer": "https://www.g2b.go.kr/",
}
MAX_BYTES = 20 * 1024 * 1024


def _items(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    body = data.get("response", {}).get("body", {}) or {}
    items = body.get("items") or []
    if isinstance(items, dict):
        items = items.get("item", [])
    if isinstance(items, dict):
        items = [items]
    return items


def _query_bid(bid_ntce_no: str, bid_ntce_ord: Optional[str], business_type: Optional[str]) -> List[Dict[str, Any]]:
    types = [business_type] if business_type in BID_OPERATIONS else list(BID_OPERATIONS)
    found: List[Dict[str, Any]] = []
    for t in types:
        params = {
            "serviceKey": get_service_key(),
            "pageNo": 1,
            "numOfRows": 10,
            "type": "json",
            "inqryDiv": 2,
            "bidNtceNo": bid_ntce_no,
        }
        try:
            r = requests.get(f"{BID_BASE}/{BID_OPERATIONS[t]}", params=params, timeout=20)
            r.raise_for_status()
            items = _items(r.json())
        except Exception as e:  # noqa: BLE001
            logger.warning("bid query failed (%s): %s", t, e)
            continue
        for it in items:
            if bid_ntce_ord and str(it.get("bidNtceOrd", "")).lstrip("0") != str(bid_ntce_ord).lstrip("0"):
                continue
            it["_business_type"] = t
            found.append(it)
        if found:
            break
    return found


def _extract_attachments(item: Dict[str, Any]) -> List[Dict[str, str]]:
    files: List[Dict[str, str]] = []
    for i in range(1, 11):
        url = item.get(f"ntceSpecDocUrl{i}")
        if url:
            files.append({"name": item.get(f"ntceSpecFileNm{i}") or f"첨부{i}", "url": url, "kind": "공고규격서"})
    for key, label in (("stdNtceDocUrl", "표준공고서"), ("sptDscrptDocUrl1", "현장설명서")):
        if item.get(key):
            files.append({"name": label, "url": item[key], "kind": label})
    # 기타 *DocUrl / *FileUrl 필드도 수집
    seen = {f["url"] for f in files}
    for k, v in item.items():
        if isinstance(v, str) and v.startswith("http") and re.search(r"(Doc|File)Url", k) and v not in seen:
            files.append({"name": k, "url": v, "kind": "기타"})
            seen.add(v)
    return files


def _filename(resp: requests.Response, fallback: str) -> str:
    cd = resp.headers.get("Content-Disposition", "")
    m = re.search(r"filename\*=(?:UTF-8'')?([^;]+)", cd) or re.search(r'filename="?([^";]+)"?', cd)
    if m:
        raw = unquote(m.group(1).strip())
        try:
            raw = raw.encode("latin-1").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            try:
                raw = raw.encode("latin-1").decode("euc-kr")
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass
        return raw
    return fallback


def _excel_to_text(content: bytes, name: str, max_rows: int) -> Dict[str, Any]:
    import pandas as pd

    engine = None
    lower = name.lower()
    if lower.endswith(".xls") and not content[:4] == b"PK\x03\x04":
        engine = "xlrd"
    sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None, engine=engine, dtype=str)
    out = {}
    for sheet, df in sheets.items():
        df = df.dropna(how="all").dropna(axis=1, how="all").fillna("")
        total = len(df)
        rows = df.head(max_rows).values.tolist()
        out[str(sheet)] = {
            "total_rows": total,
            "truncated": total > max_rows,
            "rows": [" | ".join(str(c).strip() for c in row) for row in rows],
        }
    return out


def _sniff_kind(content: bytes, name: str) -> str:
    lower = name.lower()
    if lower.endswith((".xlsx", ".xlsm", ".xls")):
        return "excel"
    if content[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" and lower.endswith(".xls"):
        return "excel"
    if content[:4] == b"PK\x03\x04":
        try:
            names = zipfile.ZipFile(io.BytesIO(content)).namelist()
            if any(n.startswith("xl/") for n in names):
                return "excel"
            return "zip"
        except zipfile.BadZipFile:
            pass
    if content[:5] == b"%PDF-" or lower.endswith(".pdf"):
        return "pdf"
    if lower.endswith((".hwp", ".hwpx")):
        return "hwp"
    if content.lstrip()[:15].lower().startswith((b"<!doctype html", b"<html")):
        return "html"
    return "other"


def _read_file(url: str, name: str, max_rows: int) -> Dict[str, Any]:
    result: Dict[str, Any] = {"name": name, "url": url}
    try:
        r = requests.get(url, headers=HEADERS, timeout=40, stream=True, allow_redirects=True)
        r.raise_for_status()
        content = r.raw.read(MAX_BYTES + 1, decode_content=True)
    except Exception as e:  # noqa: BLE001
        result.update(status="download_failed", error=str(e))
        return result
    if len(content) > MAX_BYTES:
        result.update(status="too_large")
        return result
    fname = _filename(r, name)
    result["file_name"] = fname
    result["size_bytes"] = len(content)
    kind = _sniff_kind(content, fname)
    result["type"] = kind
    try:
        if kind == "excel":
            result["sheets"] = _excel_to_text(content, fname, max_rows)
            result["status"] = "ok"
        elif kind == "zip":
            zf = zipfile.ZipFile(io.BytesIO(content))
            inner = []
            for n in zf.namelist():
                if n.lower().endswith((".xlsx", ".xls", ".xlsm")):
                    inner.append({"file_name": n, "sheets": _excel_to_text(zf.read(n), n, max_rows)})
            result.update(status="ok", zip_contents=zf.namelist(), excel_files=inner)
        elif kind == "pdf":
            try:
                from pypdf import PdfReader

                reader = PdfReader(io.BytesIO(content))
                text = "\n".join((p.extract_text() or "") for p in reader.pages[:30])
                result.update(status="ok", text=text[:30000])
            except Exception as e:  # noqa: BLE001
                result.update(status="unsupported", note=f"PDF 텍스트 추출 실패: {e}")
        elif kind == "html":
            result.update(status="login_or_page_required",
                          note="파일 대신 웹페이지가 반환됨 (로그인 필요 또는 링크 만료 가능)")
        else:
            result.update(status="unsupported", note="엑셀/PDF/ZIP 외 형식은 내용 추출을 지원하지 않음 (HWP 등)")
    except Exception as e:  # noqa: BLE001
        result.update(status="parse_failed", error=str(e))
    return result


def get_bid_attachments(
    bid_ntce_no: str,
    bid_ntce_ord: Optional[str] = None,
    business_type: Optional[str] = None,
    read_contents: bool = True,
    excel_only: bool = False,
    max_rows_per_sheet: int = 200,
) -> Dict[str, Any]:
    items = _query_bid(bid_ntce_no.strip(), bid_ntce_ord, business_type)
    if not items:
        return {"success": False, "error": f"공고번호 {bid_ntce_no}를 찾지 못했습니다. 공고번호/업무구분을 확인하세요."}
    item = items[-1]
    attachments = _extract_attachments(item)
    out: Dict[str, Any] = {
        "success": True,
        "bid_ntce_no": item.get("bidNtceNo"),
        "bid_ntce_ord": item.get("bidNtceOrd"),
        "bid_ntce_nm": item.get("bidNtceNm"),
        "business_type": item.get("_business_type"),
        "detail_url": item.get("bidNtceDtlUrl") or item.get("bidNtceUrl"),
        "attachment_count": len(attachments),
        "attachments": attachments,
    }
    if read_contents:
        files = []
        for a in attachments:
            if excel_only and not re.search(r"\.(xlsx?|xlsm|zip)$", a["name"], re.I):
                continue
            files.append(_read_file(a["url"], a["name"], max_rows_per_sheet))
        out["files"] = files
    return out


def read_attachment_url(url: str, file_name: str = "첨부파일", max_rows_per_sheet: int = 200) -> Dict[str, Any]:
    return _read_file(url, file_name, max_rows_per_sheet)


# ── 조건 검색: 공사 기초금액 + 첨부파일 키워드 ─────────────────────────────


def _to_num(v: Any) -> Optional[float]:
    try:
        s = str(v).replace(",", "").strip()
        return float(s) if s else None
    except (TypeError, ValueError):
        return None


def _paged(operation: str, params: Dict[str, Any], max_pages: int) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for page in range(1, max_pages + 1):
        p = dict(params, serviceKey=get_service_key(), pageNo=page, numOfRows=999, type="json")
        try:
            r = requests.get(f"{BID_BASE}/{operation}", params=p, timeout=30)
            r.raise_for_status()
            data = r.json()
        except Exception as e:  # noqa: BLE001
            logger.warning("%s page %s failed: %s", operation, page, e)
            break
        items = _items(data)
        out.extend(items)
        total = _to_num(data.get("response", {}).get("body", {}).get("totalCount")) or 0
        if len(items) < 999 or page * 999 >= total:
            break
    return out


def search_bids_with_attachment(
    business_type: str = "공사",
    min_base_amount: float = 5_000_000_000,
    days_back: int = 14,
    attachment_keywords: Optional[List[str]] = None,
    use_estimated_price_if_no_base: bool = True,
    max_pages: int = 10,
    max_results: int = 50,
) -> Dict[str, Any]:
    from datetime import datetime, timedelta

    if business_type not in BID_OPERATIONS:
        return {"success": False, "error": "business_type은 물품/용역/공사/외자 중 하나"}
    keywords = attachment_keywords if attachment_keywords is not None else ["공내역", "내역서"]
    end = datetime.now()
    start = end - timedelta(days=min(days_back, 31))
    rng = {"inqryDiv": 1, "inqryBgnDt": start.strftime("%Y%m%d0000"), "inqryEndDt": end.strftime("%Y%m%d2359")}

    op = BID_OPERATIONS[business_type]
    bids = _paged(op, rng, max_pages)

    # 기초금액 조회 (공고와 별도 오퍼레이션)
    base_map: Dict[str, float] = {}
    for b in _paged(f"{op}BsisAmount", rng, max_pages):
        amt = _to_num(b.get("bssamt"))
        if amt is not None:
            base_map[f"{b.get('bidNtceNo')}-{b.get('bidNtceOrd')}"] = amt

    latest: Dict[str, Dict[str, Any]] = {}
    for b in bids:  # 같은 공고는 최신 차수만
        no = b.get("bidNtceNo")
        if no and (no not in latest or str(b.get("bidNtceOrd", "")) > str(latest[no].get("bidNtceOrd", ""))):
            latest[no] = b

    results = []
    for b in latest.values():
        key = f"{b.get('bidNtceNo')}-{b.get('bidNtceOrd')}"
        base = base_map.get(key)
        basis = "기초금액"
        if base is None and use_estimated_price_if_no_base:
            base = _to_num(b.get("presmptPrce"))
            basis = "추정가격(기초금액 미공개)"
        if base is None or base < min_base_amount:
            continue
        atts = _extract_attachments(b)
        matched = [a for a in atts if any(k in a["name"] for k in keywords)] if keywords else atts
        if keywords and not matched:
            continue
        results.append({
            "bid_ntce_no": b.get("bidNtceNo"),
            "bid_ntce_ord": b.get("bidNtceOrd"),
            "bid_ntce_nm": b.get("bidNtceNm"),
            "ntce_instt": b.get("ntceInsttNm"),
            "dminstt": b.get("dminsttNm"),
            "amount": base,
            "amount_basis": basis,
            "amount_eok": round(base / 1e8, 2),
            "bid_ntce_dt": b.get("bidNtceDt"),
            "bid_close_dt": b.get("bidClseDt"),
            "openg_dt": b.get("opengDt"),
            "region_limit": b.get("prtcptLmtRgnNm") or b.get("rgnLmtYn"),
            "detail_url": b.get("bidNtceDtlUrl") or b.get("bidNtceUrl"),
            "matched_attachments": matched,
        })
    results.sort(key=lambda x: x["amount"], reverse=True)
    return {
        "success": True,
        "period": f"{start:%Y-%m-%d} ~ {end:%Y-%m-%d}",
        "scanned_bids": len(latest),
        "matched_count": len(results),
        "results": results[:max_results],
        "note": "첨부 엑셀 내용은 get_bid_attachments(공고번호)로 확인",
    }
