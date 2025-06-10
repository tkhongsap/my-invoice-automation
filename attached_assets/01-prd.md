# PRD — Zoomed Invoice Screenshot Generator

*Version 0.1 — 2025-06-10*

---

## 1  Overview

We need clearer, tighter screenshots of American-Express invoice PDFs.
Instead of a full-page capture the script must **crop to the transaction block** and **zoom it (≈ 150 – 200 %)** so reviewers can read the image at 100 % scale in Excel or any viewer.

---

## 2  Goals

| ID  | Goal                                            | Metric / Target                                      |
| --- | ----------------------------------------------- | ---------------------------------------------------- |
| G-1 | Produce legible screenshots without manual zoom | 100 % of samples pass visual check                   |
| G-2 | Reduce image clutter                            | Only the transaction table appears (no page margins) |
| G-3 | Keep file sizes moderate                        | ≤ 1 MB per PNG at ≤ 1800 px wide                     |

---

## 3  Scope

### In-Scope

* Update or extend **`pdf_screenshot_generator.py`** so it

  1. Renders page at \~120 DPI.
  2. Crops a **fixed bounding box** (configurable JSON for other layouts).
  3. Upscales the crop ×2 (simulate 200 % zoom).
  4. Saves PNG to **`/output/screenshot_zoomed/`** with the original filename.
* Add idempotency — skip PNG if it already exists.
* Log progress (`[OK] filename`) and errors.

### Out-of-Scope

* OCR, dynamic table detection, or multi-page invoices.
* Changing the Excel organizer (handles any image size already).

---

## 4  Functional Requirements

| Ref  | Requirement                                                             |
| ---- | ----------------------------------------------------------------------- |
| FR-1 | Batch-process every PDF in `/invoices/`.                                |
| FR-2 | Use **fixed crop coords** `(x1,y1,x2,y2)` (tuned once for AmEx layout). |
| FR-3 | Upscale the crop by a factor of 2 (configurable).                       |
| FR-4 | Save PNG with identical base name in `output/screenshot_zoomed/`.       |
| FR-5 | Provide clear console output with success / failure per file.           |

---

## 5  Non-Functional Requirements

| NFR | Requirement                                                 |
| --- | ----------------------------------------------------------- |
| N-1 | Runtime ≤ 2 s per PDF on a modern laptop.                   |
| N-2 | Code ≤ 100 LOC; follow existing project style.              |
| N-3 | Uses only **pdf2image** + **Pillow** (no extra heavy deps). |
| N-4 | Works cross-platform (Windows + macOS/Linux with Poppler).  |

---

## 6  Suggested Implementation Snippet

```python
from pdf2image import convert_from_path
from pathlib import Path
from PIL import Image

CROP_BOX = (100, 300, 1300, 700)   # tune once
ZOOM = 2                           # 2× scale
DPI = 120

def process(pdf_path: Path, out_dir: Path):
    page = convert_from_path(pdf_path, dpi=DPI, first_page=1, last_page=1)[0]
    crop = page.crop(CROP_BOX)
    zoomed = crop.resize((crop.width * ZOOM, crop.height * ZOOM), Image.LANCZOS)
    out_dir.mkdir(parents=True, exist_ok=True)
    zoomed.save(out_dir / (pdf_path.stem + ".png"))
```

---

## 7  Acceptance Criteria

* **AC-1** — Screenshots show only the “DATE | DESCRIPTION | AMOUNT” block (see sample).
* **AC-2** — Text is sharp when viewed at 100 % in Excel; no pixelation.
* **AC-3** — Output folder populated for all PDFs; filenames match.
* **AC-4** — Script rerun does not overwrite unchanged images (idempotent).

---

## 8  Deliverables

| Item                                         | Location                                  |
| -------------------------------------------- | ----------------------------------------- |
| Updated script `pdf_screenshot_generator.py` | `/scr/`                                   |
| New config `crop_coords.json` (optional)     | `/scr/`                                   |
| Sample output PNGs                           | `/output/screenshot_zoomed/`              |
| README addendum                              | root `README.md` (update “Usage” section) |

---

## 9  Timeline

| Task                                                 | Owner    | ETA        |
| ---------------------------------------------------- | -------- | ---------- |
| T-1 Coordinate tuning of `CROP_BOX` on 3 sample PDFs | Dev + QA | **12 Jun** |
| T-2 Implement & commit script changes                | Dev      | **13 Jun** |
| T-3 Peer review & merge                              | Lead     | **14 Jun** |
| T-4 Release to staging & QA sign-off                 | QA       | **15 Jun** |

---

**Contact:**
*Dev Lead — `dev@example.com`*
