# Changelog

## Unreleased

- Foundation: one read-only custom editor per format (`.pdf`, `.docx`, `.doc`, `.xlsx`, `.xls`, `.pptx`, `.ppt`)
  that checks the file on the host (size cap, file signature, ZIP limits for `.docx`/`.xlsx`/`.pptx`) and hands its
  bytes to a sandboxed webview viewer. Shared viewer frame with file name, zoom (buttons, `+`/`-`/`0`, Ctrl/Cmd +
  scroll), error and "simplified preview" banners; reloads when the file changes on disk. Links in documents open in
  the browser only when they are http(s). The viewers themselves are placeholders for now.
