# Coding style for this project

## What this is

A Word add-in that renders WaveJSON timing diagrams to SVG inside a Word
document, storing the JSON source in the image's alt-text so diagrams stay
editable later.

## Style

Goal: minimalistic, short, code that works. Not clever, not complete, not
future-proof — just correct and as small as it can be.

Optimize for one thing: a reader can go top to bottom and understand the whole
pipeline without jumping through indirection. This is a prototype, not a
platform — code like it.

- Prefer one flat script over a package with modules, until a file is
  genuinely too long to hold in your head (>300-400 lines is the rough
  trigger, not a hard rule).
- No class where a function will do. No framework where a function will do.
  No config system (YAML/JSON/env-driven settings) until there's a second
  real use case that needs it — hardcode the value and leave a comment.
- No abstraction for a single call site. If something is called once,
  inline it. Don't build for imagined future flexibility.
- Prefer explicit, boring code over clever code. If you have to explain a
  trick, don't use the trick.
- Write it so it fails loudly. Assertions and explicit checks over silent
  fallbacks or broad try/except. A crash with a clear message beats a
  quietly wrong result.
- Comments explain *why*, not *what*. If a line needs a "what" comment, the
  line should be rewritten to not need it.
- Delete code aggressively. Dead branches, unused params, speculative
  hooks — remove them the moment they're unused, don't leave them "just in
  case."
- No premature error handling for inputs that can't occur here. Validate
  only at real boundaries (user input, file I/O, network, the Word API).
- Print/log state at the boundaries you're least sure about (WaveJSON
  parse, SVG generation, Office.js calls) rather than wrapping everything
  in logging.

When in doubt: fewer files, fewer layers, fewer knobs.

## Word caches the taskpane aggressively

Every time `taskpane.html` changes and gets pushed, bump the `?v=N` query
param on `SourceLocation` in `manifest.xml` to the next number. Without a
new URL, Word keeps serving a stale cached copy of the taskpane even after
a fresh GitHub Pages deploy, and changes silently don't show up. This one
line has repeatedly cost more debugging time than everything else in this
project — always bump it, no exceptions.
