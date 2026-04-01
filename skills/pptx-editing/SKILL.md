---
name: pptx-editing
description: Use when editing, creating, or inspecting PowerPoint (.pptx) files programmatically
---

# PowerPoint Editing

## Overview

Edit PowerPoint presentations using `rtk pptx`. It handles reading, writing, and modifying slides in Rust. Falls back to python-pptx only for operations rtk does not support yet.

**Core principle:** Read fast, edit fast, always verify. Try rtk first, python second.

## When to Use

- User asks to edit, update, or create slides in a .pptx file
- User asks to inspect slide contents, list shapes, read text
- User asks to add, remove, or reorder slides
- User asks to modify text, colors, fonts, or layout on slides
- User shares a .pptx file and asks about its contents

## Available Commands

### Reading

```bash
rtk pptx info file.pptx                    # Slide count, dimensions, title, author
rtk pptx slides file.pptx                  # List all slides with titles
rtk pptx read file.pptx 3                  # Read slide 3 (shapes, text, colors, fonts)
rtk pptx read file.pptx 3-5                # Read slides 3 through 5
rtk pptx find file.pptx "Status"           # Find shapes containing "Status" across all slides
```

### Modifying Existing Content

```bash
rtk pptx set-text file.pptx 3 "Shape Name" "New text"     # Replace text in a named shape
rtk pptx set-fill file.pptx 3 "Shape Name" "#4472C4"      # Set shape fill color
```

Shape names come from `rtk pptx read` output. Use the XML name shown in brackets.

### Adding New Content

```bash
# Add a textbox
rtk pptx add-textbox file.pptx 3 --left 500000 --top 2000000 --width 5000000 --height 500000 --text "Hello"

# Add a shape (rectangle, oval, diamond, arrow, chevron, cloud, star, heart, etc.)
rtk pptx add-shape file.pptx 3 rectangle --left 500000 --top 2000000 --width 2000000 --height 800000 --fill "#4472C4" --text "Label"

# Add a table
rtk pptx add-table file.pptx 3 5 3 --left 500000 --top 1500000 --width 8000000 --height 3000000
```

Positions and sizes are in EMU (English Metric Units). Quick reference:
- 1 inch = 914400 EMU
- 1 cm = 360000 EMU
- Typical slide: 9144000 x 5143500 EMU (10" x 5.63")

### Slide Management

```bash
rtk pptx delete-slide file.pptx 5          # Delete slide 5
rtk pptx move-slide file.pptx 5 2          # Move slide 5 to position 2
```

### Supported Shape Types for add-shape

rectangle, oval, diamond, triangle, right-triangle, parallelogram, trapezoid, pentagon, hexagon, octagon, arrow-right, arrow-left, arrow-up, arrow-down, chevron, star-5, heart, cloud, callout, plus

## Workflow

### Step 1: Inspect First

Always read the presentation before making changes.

```bash
rtk pptx slides file.pptx                  # Overview of all slides
rtk pptx read file.pptx 15                 # Detailed view of slide 15
```

### Step 2: Plan Changes

Tell the user what you plan to change. Get confirmation before destructive operations (deleting slides, replacing content).

### Step 3: Edit

Use rtk pptx commands. For each edit:

```bash
# Edit
rtk pptx set-text file.pptx 15 "Google Shape;245;p30" "New Status"

# Or add new content
rtk pptx add-textbox file.pptx 15 --left 500000 --top 3000000 --width 8000000 --height 400000 --text "Added note"
```

### Step 4: Verify

After every edit, re-read the slide:

```bash
rtk pptx read file.pptx 15
```

Never claim an edit is done without verifying.

## When to Fall Back to python-pptx

rtk pptx does not support these operations yet. Use python-pptx for:

- Modifying font properties on existing text (size, bold, italic, color)
- Editing table cell content and formatting
- Adding images to slides
- Adding or modifying charts
- Working with animations and transitions
- Cloning slides
- Modifying SmartArt
- Group shape manipulation
- Setting text alignment on existing shapes

For these cases:

```python
python3 << 'PYEOF'
from pptx import Presentation
from pptx.util import Pt, Emu, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation('file.pptx')
slide = prs.slides[14]  # 0-indexed

for shape in slide.shapes:
    if shape.has_text_frame and "target text" in shape.text:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                run.font.size = Pt(24)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

prs.save('file.pptx')
PYEOF
```

After any python-pptx edit, verify with `rtk pptx read`.

## Red Flags

| Thought | Reality |
|---------|---------|
| "Let me rewrite the whole slide" | Modify in-place. Rebuilding loses formatting. |
| "I'll use python-pptx for everything" | Try rtk first. It's faster and uses fewer tokens. |
| "I don't need to verify" | Always read the slide after editing. |
| "I'll save to a new file" | Save to the same file unless asked otherwise. |
| "rtk failed so skip to python" | Check the error. If it's a shape name issue, fix the name. |

## Fallback Order

1. `rtk pptx` commands (Rust-native, fast, both read and write)
2. `python3 -c "from pptx import Presentation; ..."` (for unsupported operations)
3. If python-pptx not installed: `pip install python-pptx` then retry
