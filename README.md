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
