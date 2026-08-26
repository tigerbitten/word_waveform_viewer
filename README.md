# Waveform Viewer/Editor

Word add-in: author WaveJSON timing diagrams (via WaveDrom), insert them as
images, and edit them again later. The diagram's JSON source is stored in
the image's alt-text, so no companion files and no server needed.

## Files

- `taskpane.html` — the whole add-in. Editor, WaveDrom render, Office.js
  calls, all in one file.
- `manifest.xml` — points Word at the hosted taskpane, sideload this file.
- `prototype.html` — standalone WaveJSON-to-SVG renderer, no Word/Office.js.
  Useful for iterating on rendering without touching Word at all.
- `vendor/wavedrom.min.js`, `vendor/default.js` — WaveDrom itself, vendored
  into the repo instead of pulled from jsDelivr. `office.js` is still loaded
  from Microsoft's CDN (required — can't be self-hosted).

## How it works

1. Edit WaveJSON in the taskpane, live preview renders via WaveDrom (SVG).
2. **Insert new diagram** / **Replace selected diagram** rasterizes the SVG
   to PNG (Word rejects SVG directly) and inserts it, writing the JSON
   into the image's `altTextDescription`, prefixed with a sentinel
   (`WAVEJSON:v1:`) so corrupted/foreign alt-text fails loudly instead of
   silently.
3. **Load selected diagram** reads that alt-text back into the editor.

## Testing (sideload, no store/IT approval needed)

1. Host `taskpane.html` + `manifest.xml` somewhere HTTPS-reachable (this
   repo uses GitHub Pages).
2. In Word for the web: Insert → Add-ins → **Upload My Add-in**, choose
   `manifest.xml`.
3. **Whenever `taskpane.html` changes**, bump the `?v=N` in `manifest.xml`'s
   `SourceLocation` and the `build vN` marker at the top of the taskpane —
   Word caches aggressively and silently serves stale versions otherwise.
   If bumping still doesn't work, sideload in a fresh private/incognito
   window.

## LLM-readability

A timing diagram inserted by this add-in is a PNG — an LLM reading the
document has to interpret pixels, which it will do poorly or not at all,
even with vision. The WaveJSON source is right there in the image's
alt-text, but only some export formats keep it accessible:

- **.docx** — alt-text is a real XML attribute (`wp:docPr descr`), extractable
  by any tool that parses the docx directly (e.g. `python-docx`, or just
  unzipping and reading `document.xml`).
- **HTML** — Word's HTML export keeps alt-text verbatim as `<img alt="...">`.
- **Markdown** (via Pandoc, from .docx or HTML) — alt-text carries into
  `![alt](path)` syntax.
- **PDF** — alt-text is dropped on export. This is the common case for
  "send this doc out," and the one place the JSON is genuinely lost.

Two things help an LLM actually use this:

1. **Export as .docx, HTML, or Markdown instead of PDF** when the document
   will be read by an LLM. No code change needed, just an export choice.
2. **Tell the LLM where to look.** Timing diagrams read as opaque images by
   default — a prompt instruction like *"timing diagrams in this document
   are images; the exact WaveJSON source for each one is in that image's
   alt-text — read it from there instead of trying to interpret the image"*
   measurably improves how well the model understands them, since it's
   reading structured data instead of guessing from a picture.

## Known limitations

- PNG only, not SVG — Word's `insertInlinePictureFromBase64` rejects SVG.
  No resolution independence.
- Tested on Word for the web only. Desktop Word untested so far.
- Alt-text capacity tested up to 200,000 characters with no truncation —
  not expected to be a real constraint.
- **No offline mode.** Needs internet on every load: the taskpane itself
  (GitHub Pages) and `office.js` (Microsoft's CDN, required to be loaded
  live) are fetched live. WaveDrom itself is now vendored locally (see
  `vendor/`), so it no longer depends on jsDelivr being reachable. Would
  still fail entirely on a machine with no reach to GitHub Pages / the
  Microsoft CDN. Possible future work: host the taskpane internally too.
