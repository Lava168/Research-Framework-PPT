# Research Framework PPT

Structure-only PowerPoint generator for a four-slide research framework template (16:9 widescreen).

Generates layout primitives — panels, cards, arrows, chapter flows — **without domain text**. Fill content separately after export.

## Requirements

```bash
pip install python-pptx
```

## Usage

```bash
python scripts/build_research_framework_ppt.py
```

Output: `Research Framework.pptx` in the project root (4 slides, 13.333 × 7.5 in).

## Typography

Matches the original `Research Framework.pptx`:

| Slides | Font | Notes |
|--------|------|-------|
| 1, 3, 4 | **Arial** | Panel titles 12 pt; group headers 9 pt; chapter tags 8.5 pt / names 11 pt |
| 2 | **Times New Roman** | Panel/section headers 12 / 10.5 pt; chapter tags 11 pt |

Font sizes are centralized in `FS` inside `scripts/ppt_framework_common.py`.

## Slides

| Slide | Layout |
|-------|--------|
| 1 | Left: 5 stacked evaluation blocks · Right: 4 chapter flow rows |
| 2 | Left: data grid (A/B/C) · Right: chapter task rows + outputs |
| 3 | Future plan stage-1 template (upgrade + graph + 5 route cards) |
| 4 | Future plan stage-2 template (dynamics + 4 decision cards) |

## Files

- `scripts/ppt_framework_common.py` — shared layout helpers (colors, rects, arrows, groups)
- `scripts/build_research_framework_ppt.py` — builds all four slides
