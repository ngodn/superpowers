---
name: pptx-editing
description: Use when editing, creating, or inspecting PowerPoint (.pptx) files programmatically
---

# PowerPoint Editing

## Overview

Edit PowerPoint presentations programmatically using `rtk pptx` for fast read/inspect operations, with python-pptx as fallback for complex edits.

**Core principle:** Read fast, edit carefully, always verify.

## When to Use

- User asks to edit, update, or create slides in a .pptx file
- User asks to inspect slide contents, list shapes, read text
- User asks to add, remove, or reorder slides
- User asks to modify text, colors, fonts, or layout on slides
- User shares a .pptx file and asks about its contents

## Workflow

### Step 1: Inspect First

Before making any changes, read the presentation structure.

**Try `rtk pptx` first (fast, no Python startup):**

```bash
rtk pptx info presentation.pptx           # Slide count, dimensions, metadata
rtk pptx slides presentation.pptx         # List all slides with titles
rtk pptx read presentation.pptx 3         # Read slide 3 in detail (shapes, text, positions)
rtk pptx read presentation.pptx 3-5       # Read slides 3 through 5
```

**If `rtk pptx` is not available or fails, fall back to python-pptx:**

```python
python3 -c "
from pptx import Presentation
prs = Presentation('presentation.pptx')
for i, slide in enumerate(prs.slides):
    title = ''
    for shape in slide.shapes:
        if shape.has_text_frame:
            title = shape.text_frame.text[:80]
            break
    print(f'Slide {i+1}: {title}')
"
```

### Step 2: Plan Changes

Before editing, tell the user what you plan to change:
- Which slides will be modified
- What specific shapes/text will change
- Whether new slides will be added or existing ones removed

Get confirmation before proceeding with destructive changes (removing slides, replacing content).

### Step 3: Make Edits

**Try `rtk pptx` for supported operations:**

```bash
rtk pptx text presentation.pptx 3 "Shape Name" "New text content"
rtk pptx export presentation.pptx 3 slide3.png    # Export slide as image for review
```

**For complex edits (shapes, colors, positions, new elements), use python-pptx:**

```python
python3 << 'PYEOF'
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor

prs = Presentation('presentation.pptx')
slide = prs.slides[2]  # Slide 3 (0-indexed)

# Edit existing text
for shape in slide.shapes:
    if shape.has_text_frame and "old text" in shape.text:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                run.text = run.text.replace("old text", "new text")

prs.save('presentation.pptx')
print("Saved successfully")
PYEOF
```

### Step 4: Verify

After every edit, re-read the modified slide to confirm the change took effect:

```bash
rtk pptx read presentation.pptx 3
```

Or with python-pptx if rtk is not available.

Never claim an edit is done without verifying the output.

## Common Patterns

### Reading all text from a slide

```bash
rtk pptx read presentation.pptx 5
```

### Finding a shape by name or content

```bash
rtk pptx find presentation.pptx "Status"    # Find shapes containing "Status"
```

### Cloning a slide

python-pptx does not have a built-in clone. Use the lxml approach:

```python
python3 << 'PYEOF'
from pptx import Presentation
from copy import deepcopy
from lxml import etree

prs = Presentation('presentation.pptx')
source = prs.slides[2]  # Clone slide 3

# Deep copy the slide XML
slide_layout = source.slide_layout
new_slide = prs.slides.add_slide(slide_layout)

# Copy all shapes from source to new slide
for shape in source.shapes:
    el = deepcopy(shape._element)
    new_slide.shapes._spTree.append(el)

prs.save('presentation.pptx')
print("Slide cloned")
PYEOF
```

### Updating a diagram or flowchart

For complex visual elements (flowcharts, lifecycle diagrams):
1. Read the slide to understand existing shape positions and connections
2. Modify text and colors in-place rather than rebuilding from scratch
3. Preserve the original layout structure

## Red Flags

| Thought | Reality |
|---------|---------|
| "Let me just rewrite the whole slide" | Modify in-place. Rebuilding loses formatting, animations, positioning. |
| "I'll use a template" | Unless the user asked for a new slide, edit the existing one. |
| "I don't need to verify" | Always read the slide after editing. python-pptx silently corrupts sometimes. |
| "This shape doesn't have text" | Check shape.has_text_frame. Group shapes contain child shapes. |
| "I'll save to a new file" | Save to the same file unless user asks otherwise. Avoids file proliferation. |

## Fallback Order

1. `rtk pptx` (fast, Rust-native, no Python startup)
2. `python3 -c "from pptx import Presentation; ..."` (full python-pptx, slower but complete)
3. If python-pptx not installed: `pip install python-pptx` then retry
