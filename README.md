# MoDocs

**Mo**bile **Doc**ument**s** — a lightweight Android app for viewing and working with documents on your phone.

If you're like me, you sometimes use your phone to look at docs, fill out forms, or open up a quick Excel file — but you don't do any heavy document editing on your phone. There was no single lightweight app on the Play Store that supported opening PDF, DOCX, XLSX, and PPTX without showing you a bunch of ads, so I decided to vibe code one.

## Features

- **PDF** — View, search, fill supported form fields, place text/signatures/checkmarks/dates, reposition marks, save filled copies, and print while preserving original page content
- **DOCX** — View Word documents, edit simple paragraphs with undo, save copies without rebuilding the whole document, export PDF, and print
- **XLSX** — View merged cells and frozen headings, jump to cells, search, edit input cells with undo, and save copies while preserving worksheet metadata
- **PPTX** — View presentations with text, shapes, images, and backgrounds

Recent files can be searched and filtered by format. Viewer save actions create a copy;
after saving, sharing sends that copy. Pending edits trigger a save-and-share prompt,
and Back offers to save, discard, or keep editing.

## Settings

- Theme: system, light, or dark; optional Android wallpaper colors
- Keep the screen awake while reading
- Remember PDF/Word page and spreadsheet row position
- Default spreadsheet zoom: 75%, 100%, 125%, or 150%
- Remember recent files; clear recent-file and reading-position history
- Check for, download, and install app updates

Document colors remain unchanged by the app theme. Documents are processed locally.
Turning off recent files clears the recent-file list; turning off position memory clears
saved positions. These metadata stores are excluded from backup and device transfer.

## Installation

Grab the latest APK from [Releases](https://github.com/llmassisted/modocs/releases) and sideload it on your Android device.

**Requires Android 8.0 (API 26) or higher.**

## Known Limitations

### DOCX
- Documents won't show up exactly as in Word, especially with nested tables or unusual positioning. Headers/footers support basic default text; first/even-page variants and complex positioning remain limited.
- Editing supports simple body paragraphs. Complex paragraphs remain read-only; unsupported XML is preserved in saved DOCX copies. PDF export/printing uses the lightweight renderer and can omit unsupported layout features.

### XLSX
- Formulas are not evaluated: displayed results are cached and can be stale after input edits. Saved workbooks request recalculation in a spreadsheet app. Formula cells are read-only; saves reject edits within shared/array formula ranges.
- Editing is limited to viewing and modifying cell values (no formatting changes, no inserting rows/columns)

### PPTX
- This is the least baked feature (because it's very rare I open a PPT on my phone)
- Complex graphics, SmartArt, charts, and animations will not render
- No editing support for presentations — just viewing
- If you need to do a lot of presentation editing on mobile, just get Microsoft Office

### PDF
- Fill & sign appends transparent image overlays for new marks; original page streams, text, links, crop boxes, and dimensions are preserved. Added marks are not searchable text or cryptographic signatures.
- Form support covers editable text fields, checkboxes, and single-choice fields. Radio groups, XFA, signature fields, and multi-select choices are not edited. Form changes appear in the saved copy; the field editor shows pending values.
- Very large PDFs may be slow to save
- Text search is unavailable on PDFs over 32 MB — viewing and fill & sign still work; the cap keeps a huge file from exhausting memory
- Overlay resolution is bounded on extreme-aspect-ratio pages to limit bitmap memory; original page content stays intact.

## Roadmap

PRs welcome for any of these.

### DOCX
- [x] Table rendering (borders, merged cells, column widths)
- [x] Basic default header/footer text in page view, PDF export, and printing
- [ ] Text wrapping around images
- [x] List indentation, numbering continuation, level text, and start overrides
- [x] Basic paragraph editing, undo, and preservation-first copy saving

### XLSX
- [ ] Formula evaluation (SUM, AVERAGE, VLOOKUP, etc.)
- [ ] Insert/delete rows and columns
- [ ] Cell formatting (bold, color, borders) from the editor
- [x] Merged cell display, frozen headings, and jump-to-cell
- [ ] Charts and graphs rendering

### PPTX
- [ ] SmartArt rendering
- [ ] Charts and graphs
- [ ] Slide transitions and animation previews
- [ ] Speaker notes viewer
- [ ] Video/audio placeholder indicators

### PDF
- [x] Preserve original PDF content when adding fill-and-sign marks
- [x] Draggable/resizable annotations after placement
- [ ] Highlight and strikethrough text tools
- [x] Detect and edit supported AcroForm fields
- [ ] Optimize save performance for large documents

### General
- [x] Searchable recent files with format filters
- [x] Persistent reader, appearance, privacy, and update settings
- [ ] Dark mode support for document viewers
- [x] PDF and DOCX printing; worksheet/slide printing remains future work
- [x] Share saved documents and save-and-share edited copies

## Reliability checks

A synthetic regression corpus lives in `app/src/androidTest/assets/`. Tests verify
DOCX XML preservation, XLSX metadata/formula preservation, PDF dimensions/rotation/text/forms,
saved-copy tracking, draft commits, undo, XML rejection, and preference persistence.

```bash
./gradlew :app:testDebugUnitTest :app:connectedDebugAndroidTest :app:lintDebug
```

The connected tests require an Android emulator or device. Copy saving stages serialization
before writing to the selected destination. Android document providers do not all support
atomic replacement; the original is protected by creating a new destination.

## Building from Source

```bash
# Clone the repo
git clone https://github.com/llmassisted/modocs.git
cd modocs

# Build debug APK
./gradlew assembleDebug

# APK will be at app/build/outputs/apk/debug/app-debug.apk
```

## License

[MIT](LICENSE)
