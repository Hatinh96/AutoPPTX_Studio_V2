"""Tóm tắt layout THE NELSON → 16:9 cho tích hợp."""
from __future__ import annotations

import json
import sys
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

SX, SY = 13.33 / 20.0, 7.5 / 11.25


def sc(v, axis):
    return round(v * (SX if axis == "x" else SY), 3)


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "THE NELSON.pptx"
    prs = Presentation(path)
    out = {"file": path, "slides": len(prs.slides), "content_slides": []}
    for si in range(len(prs.slides)):
        slide = prs.slides[si]
        items = []
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                continue
            l = shape.left / 914400 if shape.left else 0
            t = shape.top / 914400 if shape.top else 0
            w = shape.width / 914400 if shape.width else 0
            h = shape.height / 914400 if shape.height else 0
            item = {
                "name": shape.name,
                "type": int(shape.shape_type),
                "x": sc(l, "x"),
                "y": sc(t, "y"),
                "w": sc(w, "x"),
                "h": sc(h, "y"),
            }
            if shape.has_text_frame and shape.text_frame.text.strip():
                item["text"] = shape.text_frame.text.strip()[:120]
                r = shape.text_frame.paragraphs[0].runs[0] if shape.text_frame.paragraphs[0].runs else None
                if r and r.font.size:
                    item["font_pt"] = round(r.font.size.pt * SY, 1)
            if shape.has_table:
                item["table"] = f"{len(shape.table.rows)}x{len(shape.table.columns)}"
                item["cells"] = [
                    [shape.table.cell(ri, ci).text[:30] for ci in range(len(shape.table.columns))]
                    for ri in range(min(4, len(shape.table.rows)))
                ]
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                item["picture"] = True
            items.append(item)
        if items:
            out["content_slides"].append({"index": si + 1, "items": items})
    with open(sys.argv[2] if len(sys.argv) > 2 else "tools/nelson_summary.json", "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=2)
    # print slide 4 summary (typical content)
    for s in out["content_slides"]:
        if s["index"] in (4, 8):
            print(f"\n=== SLIDE {s['index']} ===")
            for it in s["items"]:
                if it.get("text") or it.get("table"):
                    print(it)


if __name__ == "__main__":
    main()
