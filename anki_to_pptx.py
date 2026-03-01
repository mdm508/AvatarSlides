#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import html
import io
import re
import tempfile
from pathlib import Path
from typing import Dict, List

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches, Pt

DEFAULT_MEDIA_DIR = Path("/Users/m/Library/Application Support/Anki2/test/collection.media")

# =========================
# CONSTANTS
# =========================

# Left/right deck slide size
LEFT_RIGHT_SLIDE_WIDTH_IN = 13.333
LEFT_RIGHT_SLIDE_HEIGHT_IN = 5

# Top/down deck slide size
TOP_DOWN_SLIDE_WIDTH_IN = 8.3
TOP_DOWN_SLIDE_HEIGHT_IN = 7.5

# Shared spacing
MARGIN_IN = 0.35
GUTTER_IN = 0.20
SIDE_GAP_IN = 0.18
TEXT_AFTER_EN_PT = 10

# Left/right layout fonts
LEFT_RIGHT_EN_FONT_PT = 24
LEFT_RIGHT_ZH_FONT_PT = LEFT_RIGHT_EN_FONT_PT + 4

# Top/down layout fonts
TOP_DOWN_EN_FONT_PT = 20
TOP_DOWN_ZH_FONT_PT = TOP_DOWN_EN_FONT_PT + 3

# Top/down image size
TOP_DOWN_IMAGE_WIDTH_IN = 8
TOP_DOWN_IMAGE_HEIGHT_IN = 6
TOP_DOWN_TEXT_GAP_IN = 0.00

# Image trimming
CROP_LEFT_PX = 12
CROP_RIGHT_PX = 12
CROP_TOP_PX = 0
CROP_BOTTOM_PX = 0


def clean_text(text: str) -> str:
    text = html.unescape(text or "")
    text = text.replace("<br>", " ").replace("<br/>", " ").replace("<br />", " ")
    text = re.sub(r"<[^>]+>", "", text)
    text = text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_anki_export(input_path: Path) -> List[Dict[str, str]]:
    raw = input_path.read_text(encoding="utf-8-sig")
    lines = raw.splitlines()

    while lines and lines[0].startswith("#"):
        lines.pop(0)

    body = "\n".join(lines)
    reader = csv.reader(io.StringIO(body), delimiter="\t", quotechar='"')
    parsed_rows: List[Dict[str, str]] = []

    for row in reader:
        if not row:
            continue

        cell = next((c for c in row if c.strip()), "")
        if not cell:
            continue

        img_match = re.search(r'<img\s+src="([^"]+)"', cell, re.IGNORECASE)
        img_name = img_match.group(1).strip() if img_match else ""

        div_matches = list(re.finditer(r"<div[^>]*>(.*?)</div>", cell, re.DOTALL | re.IGNORECASE))
        zh_raw = ""
        audio_name = ""
        en_section = cell

        if div_matches:
            first_end = div_matches[0].end()
            en_section = cell[first_end:]

            if len(div_matches) >= 2:
                zh_raw = div_matches[1].group(1)
                en_section = cell[first_end:div_matches[1].start()]

            if len(div_matches) >= 3:
                audio_content = div_matches[2].group(1)
                audio_match = re.search(r"\[sound:([^\]]+)\]", audio_content)
                if audio_match:
                    audio_name = audio_match.group(1).strip()

        parsed_rows.append(
            {
                "En": clean_text(en_section),
                "Zh": clean_text(zh_raw),
                "ImgBaseName": img_name,
                "AudioBaseName": audio_name,
            }
        )

    return parsed_rows


def export_sanitized_csv(rows: List[Dict[str, str]], output_path: Path) -> None:
    with output_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["En", "Zh", "ImgBaseName", "AudioBaseName"])
        writer.writeheader()
        writer.writerows(rows)


def make_trimmed_image(image_path: Path) -> Path:
    """
    Create a temporary cropped copy of the image and return its path.
    Caller does not need to clean it up manually; OS temp area is used.
    """
    with Image.open(image_path) as img:
        width, height = img.size

        left = CROP_LEFT_PX
        top = CROP_TOP_PX
        right = width - CROP_RIGHT_PX
        bottom = height - CROP_BOTTOM_PX

        # Safety clamp
        left = max(0, min(left, width - 1))
        top = max(0, min(top, height - 1))
        right = max(left + 1, min(right, width))
        bottom = max(top + 1, min(bottom, height))

        cropped = img.crop((left, top, right, bottom))

        suffix = image_path.suffix if image_path.suffix else ".png"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp_path = Path(tmp.name)
        tmp.close()

        save_format = img.format if img.format else None
        cropped.save(tmp_path, format=save_format)
        return tmp_path


def get_prepared_image_path(image_path: Path | None) -> Path | None:
    if not image_path or not image_path.exists():
        return None
    return make_trimmed_image(image_path)


def add_slide_left_right(prs: Presentation, en_text: str, zh_text: str, image_path: Path | None) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    margin = Inches(MARGIN_IN)
    gutter = Inches(GUTTER_IN)

    left_x = margin
    left_y = margin
    left_w = slide_width / 2 - margin - gutter / 2
    left_h = slide_height - 2 * margin

    right_x = slide_width / 2 + gutter / 2
    right_y = margin
    right_w = slide_width / 2 - margin - gutter / 2
    right_h = slide_height - 2 * margin

    prepared_image = get_prepared_image_path(image_path)

    if prepared_image:
        pic = slide.shapes.add_picture(str(prepared_image), 0, 0)
        iw = pic.width
        ih = pic.height
        scale = min(left_w / iw, left_h / ih)
        new_w = int(iw * scale)
        new_h = int(ih * scale)
        pic.width = new_w
        pic.height = new_h
        pic.left = int(left_x + (left_w - new_w) / 2)
        pic.top = int(left_y + (left_h - new_h) / 2)
        content_top = pic.top
    else:
        box = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            int(left_x),
            int(left_y),
            int(left_w),
            int(left_h),
        )
        box.text_frame.text = "Image not found"
        box.fill.background()
        box.line.fill.background()
        content_top = left_y

    text_box = slide.shapes.add_textbox(
        int(right_x),
        int(content_top),
        int(right_w),
        int(right_h - (content_top - right_y)),
    )
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.clear()
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    p1 = tf.paragraphs[0]
    p1.text = en_text
    p1.font.size = Pt(LEFT_RIGHT_EN_FONT_PT)
    p1.space_after = Pt(TEXT_AFTER_EN_PT)

    p2 = tf.add_paragraph()
    p2.text = zh_text
    p2.font.size = Pt(LEFT_RIGHT_ZH_FONT_PT)


def add_slide_top_image_text_below(prs: Presentation, en_text: str, zh_text: str, image_path: Path | None) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    slide_width = prs.slide_width
    slide_height = prs.slide_height

    bottom_margin = Inches(MARGIN_IN)
    text_gap = Inches(TOP_DOWN_TEXT_GAP_IN)

    img_w = Inches(TOP_DOWN_IMAGE_WIDTH_IN)
    img_h = Inches(TOP_DOWN_IMAGE_HEIGHT_IN)

    img_left = int((slide_width - img_w) / 2)
    img_top = 0

    prepared_image = get_prepared_image_path(image_path)

    if prepared_image:
        slide.shapes.add_picture(str(prepared_image), img_left, img_top, width=img_w, height=img_h)
    else:
        box = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            img_left,
            img_top,
            int(img_w),
            int(img_h),
        )
        box.text_frame.text = "Image not found"
        box.fill.background()
        box.line.fill.background()

    text_left = img_left
    text_top = int(img_top + img_h + text_gap)
    text_width = int(img_w)
    text_height = int(slide_height - text_top - bottom_margin)

    text_box = slide.shapes.add_textbox(text_left, text_top, text_width, text_height)
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.clear()
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    p1 = tf.paragraphs[0]
    p1.text = en_text
    p1.font.size = Pt(TOP_DOWN_EN_FONT_PT)
    p1.space_after = Pt(6)

    p2 = tf.add_paragraph()
    p2.text = zh_text
    p2.font.size = Pt(TOP_DOWN_ZH_FONT_PT)


def build_presentation_left_right(rows: List[Dict[str, str]], media_dir: Path, output_path: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(LEFT_RIGHT_SLIDE_WIDTH_IN)
    prs.slide_height = Inches(LEFT_RIGHT_SLIDE_HEIGHT_IN)

    for row in rows:
        img_name = row["ImgBaseName"].strip()
        image_path = media_dir / img_name if img_name else None
        add_slide_left_right(prs, row["En"], row["Zh"], image_path)

    prs.save(output_path)


def build_presentation_top_image_text_below(rows: List[Dict[str, str]], media_dir: Path, output_path: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(TOP_DOWN_SLIDE_WIDTH_IN)
    prs.slide_height = Inches(TOP_DOWN_SLIDE_HEIGHT_IN)

    for row in rows:
        img_name = row["ImgBaseName"].strip()
        image_path = media_dir / img_name if img_name else None
        add_slide_top_image_text_below(prs, row["En"], row["Zh"], image_path)

    prs.save(output_path)


def test_run() -> None:
    input_file = Path("02.txt")
    media_dir = DEFAULT_MEDIA_DIR

    rows = parse_anki_export(input_file)
    export_sanitized_csv(rows, input_file.with_name("02_sanitized.csv"))

    left_right_output = input_file.with_suffix(".pptx")
    top_down_output = input_file.with_name(input_file.stem + "_top_image.pptx")

    build_presentation_left_right(rows, media_dir, left_right_output)
    build_presentation_top_image_text_below(rows, media_dir, top_down_output)

    print(f"Rows parsed: {len(rows)}")
    print(f"CSV created: {input_file.with_name('02_sanitized.csv')}")
    print(f"PPTX created: {left_right_output}")
    print(f"PPTX created: {top_down_output}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_file", help="Anki export text file")
    parser.add_argument("--media-dir", default=str(DEFAULT_MEDIA_DIR))
    parser.add_argument("--output-left-right", help="Output path for left/right pptx")
    parser.add_argument("--output-top-down", help="Output path for top/down pptx")
    parser.add_argument("--export-csv", action="store_true", help="Also export sanitized CSV")
    args = parser.parse_args()

    input_path = Path(args.input_file).expanduser()
    media_dir = Path(args.media_dir).expanduser()

    left_right_output = (
        Path(args.output_left_right).expanduser()
        if args.output_left_right
        else input_path.with_suffix(".pptx")
    )
    top_down_output = (
        Path(args.output_top_down).expanduser()
        if args.output_top_down
        else input_path.with_name(input_path.stem + "_top_image.pptx")
    )

    rows = parse_anki_export(input_path)

    if args.export_csv:
        export_sanitized_csv(rows, input_path.with_name(input_path.stem + "_sanitized.csv"))

    build_presentation_left_right(rows, media_dir, left_right_output)
    build_presentation_top_image_text_below(rows, media_dir, top_down_output)

    print(f"Rows parsed: {len(rows)}")
    print(f"PPTX created: {left_right_output}")
    print(f"PPTX created: {top_down_output}")


if __name__ == "__main__":
    main()