---
name: pptx-editing
description: Use when editing, creating, or inspecting PowerPoint (.pptx) files programmatically
---

# PowerPoint Editing

## Overview

Edit PowerPoint presentations using `rtk pptx` for both reading and writing. Falls back to python-pptx for operations rtk cannot handle yet.

**Core principle:** Read fast, edit fast, always verify. Try Rust first, Python second.

## When to Use

- User asks to edit, update, or create slides in a .pptx file
- User asks to inspect slide contents, list shapes, read text
- User asks to add, remove, or reorder slides
- User asks to modify text, colors, fonts, or layout on slides
- User shares a .pptx file and asks about its contents

## Workflow

### Step 1: Inspect First

Before making any changes, understand the presentation structure.

```bash
rtk pptx info presentation.pptx           # Slide count, dimensions, metadata
rtk pptx slides presentation.pptx         # List all slides with titles
rtk pptx read presentation.pptx 3         # Read slide 3 (shapes, text, colors, fonts)
rtk pptx read presentation.pptx 3-5       # Read slides 3 through 5
rtk pptx find presentation.pptx "Status"  # Find shapes containing "Status" across all slides
```

If `rtk pptx` is not installed, fall back to python-pptx:

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
- Whether new slides will be added or removed

Get confirmation before destructive changes (removing slides, replacing content).

### Step 3: Edit

**Try `rtk pptx` write commands first:**

```bash
# Modify text in a shape
rtk pptx set-text presentation.pptx 3 "Shape Name" "New text content"

# Set shape fill color
rtk pptx set-fill presentation.pptx 3 "Shape Name" "#4472C4"

# Set font properties
rtk pptx set-font presentation.pptx 3 "Shape Name" --size 24 --bold --color "#FFFFFF"

# Add a new text box
rtk pptx add-textbox presentation.pptx 3 "My Label" --left 1in --top 2in --width 4in --height 1in --text "Hello"

# Delete a slide
rtk pptx delete-slide presentation.pptx 5

# Move a slide
rtk pptx move-slide presentation.pptx 5 2    # Move slide 5 to position 2
```

**If `rtk pptx` does not support the operation or fails, fall back to python-pptx:**

```python
python3 << 'PYEOF'
from pptx import Presentation
from pptx.util import Pt, Emu, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation('presentation.pptx')
slide = prs.slides[2]  # Slide 3 (0-indexed)

for shape in slide.shapes:
    if shape.has_text_frame and "old text" in shape.text:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                run.text = run.text.replace("old text", "new text")

prs.save('presentation.pptx')
print("Saved")
PYEOF
```

Common python-pptx operations that may not be in rtk yet:
- Adding images to slides
- Creating charts
- Modifying animations and transitions
- Working with SmartArt
- Cloning slides (use lxml deepcopy approach)
- Group shape manipulation
- Table cell editing

### Step 4: Verify

After every edit, re-read the modified slide:

```bash
rtk pptx read presentation.pptx 3
```

Never claim an edit is done without verifying the output. python-pptx silently corrupts files sometimes. rtk pptx read will show you the actual state.

## Common Patterns

### Replacing text across all slides

```bash
rtk pptx find presentation.pptx "Old Company"
# Shows which slides/shapes contain it
# Then for each:
rtk pptx set-text presentation.pptx 3 "Title 1" "New Company"
```

### Updating colors in a diagram

```bash
rtk pptx read presentation.pptx 15    # See all shapes with fill colors
rtk pptx set-fill presentation.pptx 15 "Ordered" "#2563EB"
rtk pptx set-fill presentation.pptx 15 "In Progress" "#059669"
```

### Cloning a slide

python-pptx does not have built-in clone. Use lxml:

```python
python3 << 'PYEOF'
from pptx import Presentation
from copy import deepcopy

prs = Presentation('presentation.pptx')
source = prs.slides[2]

slide_layout = source.slide_layout
new_slide = prs.slides.add_slide(slide_layout)

for shape in source.shapes:
    el = deepcopy(shape._element)
    new_slide.shapes._spTree.append(el)

prs.save('presentation.pptx')
print("Slide cloned")
PYEOF
```

### Updating a flowchart or diagram

1. Read the slide to understand existing shapes and positions
2. Modify text and colors in-place, do not rebuild from scratch
3. Preserve the original layout structure
4. Verify after each change

## Red Flags

| Thought | Reality |
|---------|---------|
| "Let me rewrite the whole slide" | Modify in-place. Rebuilding loses formatting, animations, positioning. |
| "I'll use a template" | Unless the user asked for a new slide, edit the existing one. |
| "I don't need to verify" | Always read the slide after editing. |
| "This shape doesn't have text" | Check all shapes. Group shapes contain child shapes. |
| "I'll save to a new file" | Save to the same file unless asked otherwise. |
| "rtk pptx failed, skip to python" | Check the error. If it's a missing subcommand, python is fine. If it's a file error, python will fail too. |

## Fallback Order

1. `rtk pptx` commands (fast, Rust-native, no Python startup, both read and write)
2. `python3 -c "from pptx import Presentation; ..."` (full python-pptx, slower but handles everything)
3. If python-pptx not installed: `pip install python-pptx` then retry
