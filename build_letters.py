#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os

import openpyxl


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(BASE_DIR, "字母图纸.xlsx")
OUT_PATH = os.path.join(BASE_DIR, "letters.js")

# (rmin, rmax, cmin, cmax, baseline_row)
# baseline_row is the Excel row where the letter's baseline sits.
LOWER_REGIONS = {
    "a": (3, 9, 4, 8, 9),
    "b": (3, 9, 12, 16, 9),
    "c": (3, 9, 19, 23, 9),
    "d": (3, 9, 26, 30, 9),
    "e": (3, 9, 34, 38, 9),
    "f": (3, 9, 41, 45, 9),
    "g": (3, 12, 49, 53, 9),
    "h": (3, 9, 57, 61, 9),
    "i": (3, 9, 65, 65, 9),
    "j": (3, 10, 68, 70, 9),
    "k": (3, 9, 74, 77, 9),
    "l": (3, 9, 81, 81, 9),
    "m": (14, 18, 2, 8, 18),
    "n": (14, 18, 12, 16, 18),
    "o": (14, 18, 19, 23, 18),
    "p": (14, 21, 26, 30, 18),
    "q": (14, 21, 33, 37, 18),
    "r": (14, 18, 41, 44, 18),
    "s": (14, 18, 49, 53, 18),
    "t": (14, 18, 56, 60, 18),
    "u": (14, 18, 63, 67, 18),
    "v": (14, 18, 69, 73, 18),
    "w": (22, 25, 4, 8, 25),
    "x": (22, 26, 12, 16, 26),
    "y": (22, 29, 19, 23, 26),
    "z": (22, 26, 28, 32, 26),
}

UPPER_REGIONS = {
    "A": (33, 39, 4, 8, 39),
    "B": (33, 39, 12, 16, 39),
    "C": (33, 39, 20, 24, 39),
    "D": (33, 39, 29, 33, 39),
    "E": (33, 39, 37, 41, 39),
    "F": (33, 39, 45, 49, 39),
    "G": (33, 39, 53, 57, 39),
    "H": (33, 39, 61, 65, 39),
    "I": (33, 39, 69, 73, 39),
    "J": (33, 39, 77, 81, 39),
    "K": (33, 39, 85, 89, 39),
    "L": (33, 39, 93, 97, 39),
    "M": (43, 49, 2, 8, 49),
    "N": (43, 49, 12, 16, 49),
    "O": (43, 49, 20, 24, 49),
    "P": (43, 49, 29, 33, 49),
    "Q": (43, 49, 37, 41, 49),
    "R": (43, 49, 45, 49, 49),
    "S": (43, 49, 52, 56, 49),
    "T": (43, 49, 59, 63, 49),
    "U": (43, 49, 66, 70, 49),
    "V": (43, 49, 74, 78, 49),
    "W": (43, 49, 81, 85, 49),
    "X": (43, 49, 88, 92, 49),
    "Y": (53, 59, 4, 8, 59),
    "Z": (53, 59, 11, 15, 59),
}

# Manual corrections supplied by the user. These win over the workbook extraction.
PATTERN_OVERRIDES = {
    "w": {
        "baseline": 4,
        "pattern": [
            [1, 0, 1, 0, 1],
            [1, 0, 1, 0, 1],
            [1, 0, 1, 0, 1],
            [1, 0, 1, 0, 1],
            [0, 1, 0, 1, 0],
        ],
    },
    "t": {
        "baseline": 5,
        "pattern": [
            [0, 0, 1, 0, 0],
            [0, 0, 1, 0, 0],
            [1, 1, 1, 1, 1],
            [0, 0, 1, 0, 0],
            [0, 0, 1, 0, 0],
            [0, 0, 1, 1, 1],
        ],
    },
    "Q": {
        "baseline": 6,
        "pattern": [
            [0, 1, 1, 1, 0, 0],
            [1, 0, 0, 0, 1, 0],
            [1, 0, 0, 0, 1, 0],
            [1, 0, 0, 0, 1, 0],
            [1, 0, 0, 1, 1, 0],
            [1, 0, 0, 0, 1, 0],
            [0, 1, 1, 1, 0, 1],
        ],
    },
    "y": {
        "baseline": 3,
        "pattern": [
            [1, 0, 0, 0, 1],
            [1, 0, 0, 0, 1],
            [1, 0, 0, 1, 1],
            [0, 1, 1, 0, 1],
            [0, 0, 0, 0, 1],
            [1, 0, 0, 0, 1],
            [0, 1, 1, 1, 0],
        ],
    },
}


def extract_letter(ws, region):
    rmin, rmax, cmin, cmax, baseline = region
    cells = {
        (r, c)
        for r in range(rmin, rmax + 1)
        for c in range(cmin, cmax + 1)
        if ws.cell(row=r, column=c).fill.fill_type not in (None, "none")
    }
    if not cells:
        raise ValueError("empty region")

    rows = [r for r, _ in cells]
    cols = [c for _, c in cells]
    rrmin, rrmax = min(rows), max(rows)
    ccmin, ccmax = min(cols), max(cols)

    pattern = []
    for r in range(rrmin, rrmax + 1):
        pattern.append([
            1 if (r, c) in cells else 0
            for c in range(ccmin, ccmax + 1)
        ])

    return {
        "pattern": pattern,
        "width": len(pattern[0]),
        "height": len(pattern),
        "baseline": baseline - rrmin,
    }


def trim_pattern(pattern, baseline):
    rows = [r for r, row in enumerate(pattern) if any(row)]
    cols = [c for c in range(len(pattern[0])) if any(pattern[r][c] for r in range(len(pattern)))]
    if not rows or not cols:
        return pattern, baseline
    rmin, rmax = min(rows), max(rows)
    cmin, cmax = min(cols), max(cols)
    trimmed = [row[cmin:cmax + 1] for row in pattern[rmin:rmax + 1]]
    return trimmed, baseline - rmin


def bold_pattern(pattern):
    h = len(pattern)
    w = len(pattern[0])
    result = [[0] * w for _ in range(h)]
    for r in range(h):
        for c in range(w):
            if not pattern[r][c]:
                continue
            result[r][c] = 1
            if c > 0:
                result[r][c - 1] = 1
            if c < w - 1:
                result[r][c + 1] = 1
    return result


def upscale_pattern(pattern, baseline, target_width, target_height, target_baseline):
    h = len(pattern)
    w = len(pattern[0])

    def src_y(row):
        if row <= target_baseline:
            if target_baseline == 0:
                return 0.0
            return row * baseline / target_baseline
        denom = target_height - 1 - target_baseline
        if denom <= 0:
            return h - 1.0
        return baseline + (row - target_baseline) * (h - 1 - baseline) / denom

    def src_x(col):
        return col * (w - 1) / (target_width - 1)

    result = []
    for row in range(target_height):
        src_row = src_y(row)
        y0 = int(src_row)
        y1 = min(y0 + 1, h - 1)
        dy = src_row - y0
        line = []
        for col in range(target_width):
            src_col = src_x(col)
            x0 = int(src_col)
            x1 = min(x0 + 1, w - 1)
            dx = src_col - x0
            value = (
                pattern[y0][x0] * (1 - dx) + pattern[y0][x1] * dx
            ) * (1 - dy) + (
                pattern[y1][x0] * (1 - dx) + pattern[y1][x1] * dx
            ) * dy
            line.append(1 if value >= 0.45 else 0)
        result.append(line)
    return result


def build_large_bold(letter, key):
    pattern = letter["pattern"]
    baseline = letter["baseline"]
    height = letter["height"]
    width = letter["width"]

    target_width = max(3, min(9, int(width * 7 / 5 + 0.5)))
    if key.isupper():
        target_height = 9
        target_baseline = 8
    elif baseline < height - 1:
        target_height = 9
        target_baseline = 6
    elif height >= 6:
        target_height = 9
        target_baseline = 8
    else:
        target_height = 7
        target_baseline = 6

    pattern = upscale_pattern(
        pattern, baseline, target_width, target_height, target_baseline
    )
    pattern, baseline = trim_pattern(pattern, target_baseline)
    return {
        "pattern": pattern,
        "width": len(pattern[0]),
        "height": len(pattern),
        "baseline": baseline,
    }


def main():
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb["Sheet1"]

    letters = {}
    for key, region in list(LOWER_REGIONS.items()) + list(UPPER_REGIONS.items()):
        letters[key] = extract_letter(ws, region)
    for key, override in PATTERN_OVERRIDES.items():
        pattern = override["pattern"]
        letters[key] = {
            "pattern": pattern,
            "width": len(pattern[0]),
            "height": len(pattern),
            "baseline": override["baseline"],
        }

    styles = {
        "classic": letters,
        "bold": {key: build_large_bold(letter, key) for key, letter in letters.items()},
    }
    payload = {
        "letters": letters,
        "styles": styles,
        "generated_from": os.path.basename(XLSX_PATH),
    }
    with open(OUT_PATH, "w", encoding="utf-8") as fh:
        fh.write("// Generated by build_letters.py. Do not edit by hand.\n")
        fh.write("window.LETTERS = ")
        fh.write(json.dumps(payload, ensure_ascii=False, separators=(",", ":")))
        fh.write(";\n")

    print("wrote", OUT_PATH, "letters:", len(letters))


if __name__ == "__main__":
    main()
