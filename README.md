# Waveform Viewer/Editor

Word add-in: author WaveJSON timing diagrams (via WaveDrom), insert them as
images, and edit them again later. The diagram's JSON source is stored in
the image's alt-text, so no companion files and no server needed.

## Desktop Word setup (per-person, not centralized)

This is a one-time setup per machine; each person
who wants to use this add-in on desktop Word repeats it themselves.

### Windows

1. Make a local folder, e.g. `C:\AddinCatalog`.
2. Download `manifest.xml` from this repo into that folder. MAKE SURE TO USE THE GITHUB 'Download Raw File' BUTTON. The manifest file should be ~2KB.
3. In Word: **File → Options → Trust Center → Trust Center Settings →
   Trusted Add-in Catalogs**. As the Catalog Url, use
   `\\localhost\c$\AddinCatalog` (swap in your actual folder path), check
   **Show in Menu**, OK, OK, then fully restart Word.
4. Inside a Document, **Home → Add-ins → Advanced Settings → Shared Folder** — select the
   add-in there.

### Mac

1. Download `manifest.xml` from this repo. MAKE SURE TO USE THE GITHUB 'Download Raw File' BUTTON. The manifest file should be ~2KB.
2. Run `mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
3. Move `manifest.xml` into that `wef` folder.
4. Quit Word completely, then reopen it.
5. Inside a Document, **Home → Add-ins → Developer Add-ins** — select the add-in there. It may have a GitHub logo.

## How it works

### The core loop

1. Type or paste WaveJSON in the taskpane. The preview updates as you type.
2. **Insert new diagram** — puts the diagram in your document at the cursor.
3. Click a diagram, then **Load selected diagram** — pulls it back into the
   editor to change it.
4. **Replace selected diagram** — swaps the clicked diagram for the edited one.

Every diagram carries its own source, so any of them can be reopened and
edited later. No side files to keep track of.

### The other editor buttons

- **New (blank)** — resets the editor to a single empty signal.
- **Add signal** — appends an empty signal to the `signal` array. Needs the
  current JSON to parse first; it edits the parsed object, not the text.
- **Copy JSON** — selects the editor text so you can press Ctrl+C. It does
  not copy for you: `navigator.clipboard.writeText` is blocked by Word's
  iframe permissions policy.

Two editor conveniences that aren't buttons: drag the bar under the textarea
to resize it, and Tab inserts an indent instead of jumping focus out.

### Templates and cheat sheet

Collapsed behind **Templates & wave-char cheat sheet**. Four templates —
Clock, Data bus, Handshake (req/ack), and Edges (arrows between signals) —
each of which *replaces* the whole editor contents, so they're a starting
point, not something to click mid-edit.

The cheat sheet table covers the wave characters (`p n P N 0 1 h l x z 2-9
= . |`) and how bus labels pull from the `data` array. Arrows need a `node`
string per signal plus a top-level `edge` array; the Edges template is the
working example.

## Files

- `taskpane.html` — the whole add-in. Editor, WaveDrom render, Office.js
  calls, all in one file.
- `manifest.xml` — points Word at the hosted taskpane, sideload this file.
- `prototype.html` — standalone WaveJSON-to-SVG renderer, no Word/Office.js.
  Useful for iterating on rendering without touching Word at all.
- `vendor/wavedrom.min.js`, `vendor/default.js` — WaveDrom itself, vendored
  into the repo instead of pulled from jsDelivr. `office.js` is still loaded
  from Microsoft's CDN (required — can't be self-hosted).

## LLM-readability

A timing diagram inserted by this add-in is a PNG, so an LLM reading the
document has to interpret pixels, which it will do poorly or not at all,
even with vision. The WaveJSON source is right there in the image's
alt-text, but only some export formats keep it accessible:

- **.docx** — no special export step, the native save/download-as-.docx
  already keeps it: alt-text is a real XML attribute (`wp:docPr descr`)
  inside `document.xml`. The catch is that naive text extraction (copy-paste,
  `strings`, most "convert docx to text" tools) skips it, since it isn't
  part of the visible text. An LLM (or a script feeding one) needs to parse
  the docx directly with something like `python-docx`
  (`inline_shape._inline.docPr.get("descr")`) or by unzipping the .docx and
  reading the `descr` attribute out of `document.xml` — not by reading
  the document as plain text.
- **HTML** — Word's HTML export keeps alt-text verbatim as `<img alt="...">`.
- **PDF** — alt-text is dropped on export.

Two things help an LLM actually use this:

1. **Keep it as .docx or HTML instead of exporting to PDF** when the
   document will be read by an LLM. No code change needed, just a format
   choice.
2. **Tell the LLM where to look, and how to get it.** Timing diagrams read
   as opaque images by default — a prompt instruction like *"timing diagrams
   in this document are images; the exact WaveJSON source for each one is in
   that image's alt-text. If reading a .docx, parse it directly (e.g.
   `python-docx`) to get the alt-text — don't just extract the visible text,
   it won't be there"* measurably improves how well the model understands
   them, since it's reading structured data instead of guessing from a
   picture.
