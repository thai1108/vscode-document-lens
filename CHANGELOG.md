# Changelog

## 0.1.0

- First release. One read-only custom editor per format (`.pdf`, `.docx`, `.doc`, `.xlsx`, `.xls`, `.pptx`, `.ppt`)
  that checks the file on the host (size cap, file signature, ZIP limits for `.docx`/`.xlsx`/`.pptx`) and hands its
  bytes to a sandboxed webview viewer. Shared viewer frame with file name, zoom (buttons, `+`/`-`/`0`, Ctrl/Cmd +
  scroll), error and "simplified preview" banners; reloads when the file changes on disk. Links in documents open in
  the browser only when they are http(s).
- PDF viewer (pdf.js): pages rendered near the view with selectable text, find with match count, page navigation,
  zoom and fit width on the current page, outline, links; remembers page, zoom and outline. Japanese and other CJK
  text, standard fonts and JBIG2/CCITT/JPEG 2000 images work offline; password-protected and damaged files show an
  error.
- `.docx` viewer: pages rendered with docx-preview (styles, headers and footers, footnotes, pictures, tables, lists,
  embedded fonts), in-document links, fit width and find. The document's HTML and CSS are sanitized, pictures linked
  from the web are not loaded, and oversized, malformed or forged parts are refused with a message.
- `.doc` viewer: simplified text preview of Word 97–2003 documents (paragraphs, tables, page breaks, links, footnotes
  and endnotes, comments, headers/footers, text boxes; tracked deletions and hidden text left out), with clear
  messages for password-protected and Word 6.0/95 files.
- Spreadsheet viewer for `.xlsx` and `.xls`: sheet tabs, merged cells, column widths and row heights, frozen panes,
  fills, fonts, borders, number and date formats (including Japanese eras), links, a cell bar with formulas, and
  smooth scrolling through sheets with hundreds of thousands of rows. Pictures and charts are not shown yet.
- `.pptx` viewer: slide list with thumbnails and the current slide fitted to the window, with shapes, text, pictures
  and tables in their theme colors and master/layout styles; charts, embedded objects and media appear as labelled
  placeholders. Navigate with ←/→, PageUp/PageDown, Home/End.
- `.ppt` viewer: simplified preview with each slide's title, text and pictures (PNG, JPEG, BMP; metafiles as labelled
  placeholders), hidden slides marked, find with Ctrl/Cmd+F.
- The ZIP check on the host refuses archives that another reader would unpack differently (prepended data,
  non-canonical or duplicate names, mismatched headers, overlapping entries).
