# PowerPoint-FormatToolkit

A collection of PowerPoint VBA tools for fast slide formatting:

- Multi-slot **shape format painter** (CopyA/B/C + Paste)
- **TextBox format painter** (font, highlight, fill, border)
- **Paging tools**: apply formats to each slide / every N slides
- **Shape paging**: paste a template shape to following slides
- **Link helpers**: quickly create internal slide links
- **Crop** and crop-duplicate helpers (optional, WIP)

---

## Folder structure

- `src/vba/` — exported VBA modules (`*.bas`) you can import to your own PPTM/PPAM
- `dist/` — compiled add-in (`macros.ppam`)
- `examples/` — demo PPT files before/after using the toolkit
- `docs/images/` — screenshots used in this README

---

## Installation

1. Download `dist/macros.ppam`.
2. In PowerPoint: **File → Options → Add-ins**  
3. Manage: `PowerPoint Add-ins` → `Go...`  
4. Click **Add New...**, choose `macros.ppam`.
5. Restart PowerPoint (if needed). A new tab **macros** should appear.

---

## Main features

### Shape format painter (CopyA/B/C + Paste)

1. Select a picture or shape.
2. Click **CopyA** (or CopyB/CopyC) to store its format:
   - position / size
   - rotation / flip
   - crop
   - fill color & transparency
   - line style & weight
3. Select one or multiple shapes and click **Paste**:
   - 1 stored format → 1 output shape
   - 2–3 stored formats → the target shape is duplicated and each duplicate gets one format

### Text format painter (CP_Text / CV_Text)

- `CP_Text`: copy font name, size, bold/italic, underline, highlight color, fill & border of the textbox.
- `CV_Text`: apply to one or multiple textboxes.

### Paging tools

- `Paging`: starting from current slide, apply shape formats to **each slide’s first picture**.
- `PagingII`: starting from current slide, apply formats **every N slides** (N is asked via InputBox).

(Then add your link tools / anti-crop 简短介绍…)

---

## Development

- Edit macros inside PowerPoint (`macros.pptm`).
- Export modules to `src/vba/*.bas` to keep them under version control.
- Build a new `macros.ppam` when you want to release a new version.
