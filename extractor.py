from __future__ import annotations

import io
import re
from datetime import datetime
from typing import Any

import pdfplumber


CURRENCY_SYMBOLS = {"USD": "$", "EUR": "€", "GBP": "£", "JPY": "¥", "INR": "₹"}


def _number(value: str | None) -> float | None:
    if not value:
        return None
    cleaned = re.sub(r"[^0-9.\-]", "", value.replace(",", ""))
    try:
        return float(cleaned)
    except ValueError:
        return None


def _date(value: str | None) -> str:
    if not value:
        return ""
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(value.strip(), fmt).date().isoformat()
        except ValueError:
            pass
    return value.strip()


def _compact(value: str | None) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def _looks_blank(value: str | None) -> bool:
    text = _compact(value)
    if not text:
        return True
    return re.sub(r"[\s:;,.|/\\\-]+", "", text) == ""


def _words_in(page: Any, x0: float, top: float, x1: float, bottom: float) -> str:
    sx, sy = page.width / 612.0, page.height / 792.0
    words = page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=False)
    selected = [
        w for w in words
        if x0 * sx <= (w["x0"] + w["x1"]) / 2 <= x1 * sx
        and top * sy <= (w["top"] + w["bottom"]) / 2 <= bottom * sy
    ]
    selected.sort(key=lambda w: (round(w["top"], 1), w["x0"]))
    return " ".join(w["text"] for w in selected)


def _lines_in(page: Any, x0: float, top: float, x1: float, bottom: float) -> list[str]:
    sx, sy = page.width / 612.0, page.height / 792.0
    words = page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=False)
    selected = [
        w for w in words
        if x0 * sx <= (w["x0"] + w["x1"]) / 2 <= x1 * sx
        and top * sy <= (w["top"] + w["bottom"]) / 2 <= bottom * sy
    ]
    selected.sort(key=lambda w: (round(w["top"], 1), w["x0"]))

    lines: list[tuple[float, list[str]]] = []
    for word in selected:
        if not lines or abs(word["top"] - lines[-1][0]) > 3:
            lines.append((word["top"], [word["text"]]))
        else:
            lines[-1][1].append(word["text"])

    return [_compact(" ".join(parts)) for _, parts in lines]


def _particulars(lines: list[str], max_lines: int = 2) -> str:
    cleaned: list[str] = []
    for line in lines:
        line = _compact(re.sub(r"\bY\b", "", line)).strip(" ,;:-")
        if not line or _looks_blank(line):
            continue
        cleaned.append(line)
        if len(cleaned) >= max_lines:
            break
    return ", ".join(cleaned)


def _first(pattern: str, text: str, flags: int = 0, group: int = 1) -> str:
    match = re.search(pattern, text, flags)
    return match.group(group).strip() if match else ""


def extract_pdf(pdf_bytes: bytes, filename: str) -> dict[str, Any]:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        texts = [(p.extract_text(x_tolerance=2, y_tolerance=2) or "") for p in pdf.pages]
        full_text = "\n".join(texts)
        page1 = pdf.pages[0] if pdf.pages else None
        page2 = pdf.pages[1] if len(pdf.pages) > 1 else None

        port = sb_no = sb_date = ""
        header_match = re.search(
            r"INDIAN CUSTOMS EDI SYSTEM\s+([A-Z0-9]+)\s+(\d{6,9})\s+(\d{2}-[A-Z]{3}-\d{2,4})",
            texts[0] if texts else "",
            re.I,
        )
        if header_match:
            port, sb_no, sb_date = header_match.groups()

        invoice_no = invoice_date = ""
        invoice_match = re.search(
            r"(?:^|\n)\s*1\s+(\d{1,12})\s+(\d{2}/\d{2}/\d{2,4})\s+(?:CF|CIF|FOB|CFR)\b",
            texts[1] if len(texts) > 1 else full_text,
            re.I,
        )
        if invoice_match:
            invoice_no, invoice_date = invoice_match.groups()
        elif page2:
            invoice_region = _words_in(page2, 75, 145, 165, 170)
            invoice_match = re.search(r"(\d{1,12})\s+(\d{2}/\d{2}/\d{2,4})", invoice_region)
            if invoice_match:
                invoice_no, invoice_date = invoice_match.groups()

        buyer = _first(
            r"7\.CONSIGNEE NAME\s*&\s*ADDRESS\s+FOR THE ORDER OF\s+([^\n]+)",
            texts[0] if texts else "",
            re.I,
        )
        if _looks_blank(buyer):
            buyer = _particulars(_lines_in(page1, 305, 210, 520, 230)) if page1 else ""
        if _looks_blank(buyer) and page2:
            buyer = _particulars(_lines_in(page2, 305, 176, 505, 212))

        amount = None
        currency = ""
        summary_region = _words_in(page1, 395, 312, 575, 338) if page1 else ""
        summary_match = re.search(r"(?:\b\d+\s+)?(\d{1,12})\s+([\d,.]+)\s+([A-Z]{3})\b", summary_region)
        if summary_match:
            invoice_no = invoice_no or summary_match.group(1)
            amount = _number(summary_match.group(2))
            currency = summary_match.group(3).upper()
        if amount is None:
            details_match = re.search(
                r"H\.INVOICE DETAILS.*?\n\s*1\s+(\d{1,12})\s+([\d,.]+)\s+([A-Z]{3})\b",
                full_text,
                re.I | re.S,
            )
            if details_match:
                invoice_no = invoice_no or details_match.group(1)
                amount = _number(details_match.group(2))
                currency = details_match.group(3).upper()

        exchange_rate = None
        if page2:
            exchange_rate = _number(_words_in(page2, 535, 285, 580, 310))
        if exchange_rate is None:
            exchange_rate = _number(_first(r"\bINR\s+([\d.]+)", texts[1] if len(texts) > 1 else full_text, re.I))

        fob_region = _words_in(page1, 72, 276, 128, 300) if page1 else ""
        drawback_region = _words_in(page1, 355, 276, 405, 298) if page1 else ""
        igst_region = _words_in(page1, 425, 276, 482, 298) if page1 else ""
        taxable_region = _words_in(page1, 345, 297, 407, 316) if page1 else ""
        rodtep_region = _words_in(page1, 438, 297, 474, 316) if page1 else ""
        fob = _number(_first(r"FOB VALUE\s+([\d,.]+)", fob_region, re.I))
        drawback = _number(_first(r"DBK CLAIM\s+([\d,.]+)", drawback_region, re.I))
        igst = _number(_first(r"IGST AMT\s+([\d,.]+)", igst_region, re.I))
        taxable = _number(_first(r"IGST VALUE\s+([\d,.]+)", taxable_region, re.I))
        rodtep = _number(_first(r"RODTEP AMT\s+([\d,.]+)", rodtep_region, re.I))

        if drawback is None:
            drawback = _number(_first(r"1\.DBK CLAIM.*?\n\s*([\d,.]+)", texts[0] if texts else "", re.I | re.S))
        if taxable is None:
            taxable = _number(_first(r"4\.IGST VALUE\s*\n\s*([\d,.]+)", texts[0] if texts else "", re.I))
        if rodtep is None:
            rodtep = _number(_first(r"5\.RODTEP AMT.*?\n\s*([\d,.]+)", texts[0] if texts else "", re.I | re.S))

    values = {
        "particulars": buyer,
        "bill_no": invoice_no,
        "invoice_date": _date(invoice_date),
        "sb_no": sb_no,
        "sb_date": _date(sb_date),
        "port_code": port.upper(),
        "foreign_amount": amount,
        "currency": currency or "USD",
        "exchange_rate": exchange_rate,
        "inr": round(amount * exchange_rate, 2) if amount is not None and exchange_rate is not None else None,
        "taxable_value": taxable,
        "igst": igst,
        "fob": fob,
        "drawback": drawback,
        "rodtep": rodtep,
        "source_file": filename,
    }
    required = [
        "particulars", "bill_no", "invoice_date", "sb_no", "sb_date", "port_code",
        "foreign_amount", "exchange_rate", "taxable_value", "igst", "fob", "drawback", "rodtep",
    ]
    def _present(key: str) -> bool:
        value = values.get(key)
        if isinstance(value, str):
            return not _looks_blank(value)
        return value is not None

    found = sum(_present(key) for key in required)
    missing = [key for key in required if not _present(key)]
    values["confidence"] = round(found / len(required) * 100)
    values["warnings"] = [f"Could not confidently read {key.replace('_', ' ')}" for key in missing]
    values["currency_symbol"] = CURRENCY_SYMBOLS.get(values["currency"], values["currency"])
    return values
