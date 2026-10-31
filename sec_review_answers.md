## 1. Project description, business problem and benefits

The Wavedrom-in-Word add-in is a Microsoft Word and PowerPoint add-in that lets engineers write digital timing diagrams (clocks, buses, handshakes) as short text, using the open-source WaveJSON/WaveDrom format, and insert them into a document as images. Today, timing diagrams are usually drawn in various separate tools and pasted in as static pictures, so editing one later means finding the original source file or redrawing it. This add-in stores each diagram's source inside the document, next to the image, so anyone can click a diagram, edit it and replace it in seconds. The benefits are faster spec revisions, consistent-looking diagrams, no lost source files, and diagrams that search tools and AI tools can read as text rather than just pixels.

## 2. Architecture diagram

See `sec_review_architecture.png`.

The add-in is about 12 MB of static files (HTML, JavaScript and icons, almost all of it Microsoft's `office.js` and its per-application and per-language files) that Word or PowerPoint downloads over HTTPS and runs inside Office's sandboxed add-in web view. It has no server-side code, database, user accounts or credentials. All diagram rendering happens locally in memory, and the only place data is stored is inside the user's own document, which is saved through Office's normal mechanisms. The only network traffic is outbound HTTPS downloads of the add-in files from the hosting location. Microsoft's standard `office.js` library is bundled with the add-in rather than loaded from Microsoft's CDN, so there is exactly one destination. The add-in never sends document or user data anywhere, and the page's Content Security Policy enforces this: it may load scripts only from the hosting location, cannot run inline or dynamically evaluated script, and cannot open network connections, load external images or embed frames. Diagram source read from a document is treated as untrusted input, so a document crafted to inject script into the add-in cannot run code or send data out.

## 3. Technical and security documentation

There is no formal security assessment yet, but the full source is small (a ~100-line HTML page and a ~450-line JavaScript file, plus a manifest) and can be handed to IT for review. The only third-party code is WaveDrom, an open-source MIT-licensed library, and Microsoft's own `office.js`, both bundled with the add-in. `office.js` is copied from Microsoft's CDN by a short script in the repository; Microsoft's guidance for locked-down environments is to bundle it this way. The add-in requests only the standard `ReadWriteDocument` permission, which limits it to the currently open document (read and write access to the whole of that document, which is why the Content Security Policy above matters).

## 4. AI

The Add-In does not interact with any AI services.

**References**

- WaveDrom (the open-source diagram renderer bundled with the add-in, v3.6.2): [wavedrom.com](https://wavedrom.com), [tutorial / WaveJSON format](https://wavedrom.com/tutorial.html), [source code](https://github.com/wavedrom/wavedrom), [MIT license](https://github.com/wavedrom/wavedrom/blob/trunk/LICENSE)
- Microsoft: [Privacy and security for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/privacy-and-security), which describes the sandbox, permissions and HTTPS requirements this add-in runs under
