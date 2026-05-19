#!/usr/bin/env python3
"""Локальная проверка парсера ESAD (без GAS). Запуск из корня репозитория."""
import re
import sys
from pathlib import Path

SAMPLE_DIR = Path(r"c:\Users\Сан Саныч\Desktop\Gremlin\Поступления\Мега 8 ГТД")


def extract_gtd(text: str) -> str:
    m = re.search(r"<!--\s*ND=([^>\s]+)\s*-->", text, re.I)
    if m:
        return m.group(1).strip()
    m = re.search(
        r"<GTDID>[\s\S]*?<(?:[\w]+:)?CustomsCode>(\d+)</(?:[\w]+:)?CustomsCode>"
        r"[\s\S]*?<(?:[\w]+:)?RegistrationDate>(\d{4}-\d{2}-\d{2})</(?:[\w]+:)?RegistrationDate>"
        r"[\s\S]*?<(?:[\w]+:)?GTDNumber>(\d+)</(?:[\w]+:)?GTDNumber>",
        text,
        re.I,
    )
    if not m:
        return ""
    iso = m.group(2)
    ddmmyy = iso[8:10] + iso[5:7] + iso[2:4]
    return f"{m.group(1)}/{ddmmyy}/{m.group(3)}"


def parse_payments(block: str) -> dict:
    out = {"duty": 0.0, "vat": 0.0, "fee": 0.0}
    for chunk in re.findall(
        r"<(?:[\w]+:)?ESADout_CUCustomsPaymentCalculation>([\s\S]*?)</(?:[\w]+:)?ESADout_CUCustomsPaymentCalculation>",
        block,
        re.I,
    ):
        code_m = re.search(r"<(?:[\w]+:)?PaymentModeCode>(\d+)</", chunk, re.I)
        amt_m = re.search(r"<(?:[\w]+:)?PaymentAmount>([\d.]+)</", chunk, re.I)
        if not code_m or not amt_m:
            continue
        code, amt = code_m.group(1), float(amt_m.group(1))
        if code == "1010":
            out["fee"] += amt
        elif code == "2010":
            out["duty"] += amt
        elif code == "5010":
            out["vat"] += amt
    return out


def parse_file(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    gtd = extract_gtd(text)
    blocks = re.findall(r"<ESADout_CUGoods>([\s\S]*?)</ESADout_CUGoods>", text, re.I)
    print(f"\n=== {path.name} ===")
    print(f"GTD: {gtd or '—'}")
    print(f"Goods blocks: {len(blocks)}")
    for i, block in enumerate(blocks, 1):
        tnved = re.search(r"<(?:[\w]+:)?GoodsTNVEDCode>(\d+)</", block, re.I)
        qty = re.search(
            r"<(?:[\w]+:)?GoodsGroupQuantity>[\s\S]*?<(?:[\w]+:)?GoodsQuantity>([\d.]+)</",
            block,
            re.I,
        )
        pay = parse_payments(block)
        print(
            f"  #{i} TNVED={tnved.group(1) if tnved else '?'} "
            f"qty={qty.group(1) if qty else '?'} "
            f"duty={pay['duty']:.2f} vat={pay['vat']:.2f} fee={pay['fee']:.2f}"
        )


def main() -> None:
    files = sorted(SAMPLE_DIR.glob("ESAD*.XML"))
    if not files:
        print("No ESAD*.XML in", SAMPLE_DIR, file=sys.stderr)
        sys.exit(1)
    for f in files:
        parse_file(f)


if __name__ == "__main__":
    main()
