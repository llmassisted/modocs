# Making MoDocs more complete

MoDocs should make short document tasks dependable: open, read, make a small change,
and send back. Keep full office-suite editing out of the near-term scope.

## Next priorities

1. **Recover unfinished work.** Restore Word edits, spreadsheet drafts, PDF marks,
   and form values after process death. Add redo and a clear saved/unsaved indicator.
2. **Broader compatibility tests.** Expand the synthetic corpus with public,
   redistributable files from Word, Excel, PowerPoint, Google exports, and LibreOffice.
   Add visual comparisons and release-APK smoke tests on Android 8 and current Android.
3. **Better document navigation.** Page thumbnails, outlines/bookmarks, go-to-page,
   persistent spreadsheet column position and presentation slide position.
4. **Accessibility and phone layouts.** Individual cell accessibility, dependable
   text selection, larger touch targets, large-font checks, landscape/tablet layouts,
   and an explicit reflow reading mode.
5. **PDF improvements.** Search scanned documents with on-device OCR; use native
   text/vector marks so additions remain searchable and small; show form changes
   immediately on the page; support radio groups and richer choice fields.
6. **Word fidelity.** First/even-page headers, per-section layouts, floating images,
   nested tables, footnotes, accurate fields/page numbers, and complex list overrides.
7. **Spreadsheet usefulness.** Dates/currency/percent formats, filters and sort in a
   temporary view, copy ranges, selectable column widths, and worksheet printing.
   Add formula evaluation only with clear support limits and recalculation tests.
8. **Presentation basics.** Slide thumbnails, notes, placeholder indicators for
   unsupported content, and slide printing. Defer animations and SmartArt editing.
9. **File management.** Pin favorites, remove individual recents, display file size
   and location, and help users reselect documents whose access expired.
10. **Performance budgets.** Measure first-page latency, large-file memory, scrolling,
    export time and APK size; load ZIP parts and images on demand.

## Useful future settings

- Default reading mode and fit-width/fit-page behavior by document type
- Reflow text size, line spacing and reading background
- Remember zoom, spreadsheet column and slide position
- Default export filename suffix and open/share after saving
- Recent-history retention and per-document forgetting
- Saved-signature management, if reusable signatures are introduced
- Language, date format and measurement units
- Orientation preference and document-specific defaults

Every setting should change visible behavior. Keep rare technical options out of the
main settings screen, and preserve document colors by default.

## Included in v1.77

Preservation-first Word/spreadsheet copy saving; undo and protected navigation;
save-and-share using the saved copy; PDF overlays that retain original page content;
supported form-field editing; PDF/Word printing; merged cells, frozen headings,
cell navigation and cached-formula notices; partial-load notices; basic Word
header/footer text and consistent list labels; searchable recents; reader/theme/
privacy settings; and automated document regression tests.
