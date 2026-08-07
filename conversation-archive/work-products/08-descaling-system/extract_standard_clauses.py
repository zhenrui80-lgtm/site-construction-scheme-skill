from __future__ import annotations

import argparse
from pathlib import Path

from pypdf import PdfReader


DEFAULT_TERMS = (
    "焊接",
    "无损检测",
    "耐压试验",
    "压力试验",
    "管道元件",
    "三级配电",
    "剩余电流",
    "TN-S",
    "电缆支架",
    "电缆桥架",
    "弯曲半径",
    "防火封堵",
    "防腐涂料",
    "涂层缺陷",
    "持证",
    "考试项目",
)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf", type=Path)
    parser.add_argument("terms", nargs="*", default=DEFAULT_TERMS)
    parser.add_argument("--context", action="store_true")
    args = parser.parse_args()

    reader = PdfReader(str(args.pdf))
    texts = [page.extract_text() or "" for page in reader.pages]
    print(f"FILE\t{args.pdf.name}\tPAGES\t{len(texts)}")

    for term in args.terms:
        hits = [index + 1 for index, text in enumerate(texts) if term in text]
        if not hits:
            continue
        print(f"TERM\t{term}\tPAGES\t{','.join(map(str, hits))}")
        if args.context:
            for page_number in hits[:3]:
                text = texts[page_number - 1]
                position = text.find(term)
                start = max(0, position - 180)
                end = min(len(text), position + len(term) + 420)
                excerpt = " ".join(text[start:end].split())
                print(f"CONTEXT\t{page_number}\t{excerpt}")


if __name__ == "__main__":
    main()
