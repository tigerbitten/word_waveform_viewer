# Waveform Viewer/Editor: likely follow-up questions

## 1. Can the add-in read or leak the contents of our documents?

It has the standard `ReadWriteDocument` permission, which gives it read and write access to the whole of the document that is open, but no other document or file. In practice the code reads only the alt-text (Word) or shape tag (PowerPoint) of the picture the user selected, and makes no network requests of its own: no analytics or API calls. Microsoft's `office.js` includes telemetry code, but the Content Security Policy blocks what it tries to send. Because that alt-text can come from a document someone else wrote, it is treated as untrusted input: the add-in refuses diagram source that contains script hooks, and the page's Content Security Policy blocks inline or dynamically evaluated script and all outbound connections, external images and frames. So even a document crafted to attack the add-in cannot make it run code or send data out. IT can confirm this by reading the source (about 550 lines), the policy at the top of `taskpane.html`, or by watching the network traffic while it runs.

## 2. What do you need from the hosting location?

Static HTTPS file hosting for about 12 MB of files (`taskpane.html`, the add-in's JavaScript, the bundled WaveDrom and Microsoft `office.js` files, and two icons), with no server-side code, database or special configuration. The certificate must be one that users' machines already trust, because Office refuses to load add-ins from self-signed or invalid-certificate URLs. An internal-only host is fine as long as everyone who uses the add-in can reach it over HTTPS.

## 3. How do updates work, and who can change the code?

To update, copy the new files to the host and bump the version number in the manifest's taskpane URL (Office caches the add-in aggressively, so the URL change forces a reload). Today the code runs from GitHub Pages, so anyone with push access to that repository could change what runs inside users' Office. Moving it to a company host puts every change under IT's control, which is a main reason for this request. There is no build step and nothing is fetched from package registries at runtime: WaveDrom (v3.6.2, MIT license) and Microsoft's `office.js` are fixed copies in the repository, so users' machines contact only the hosting location. `office.js` should be refreshed from Microsoft every few months by re-running `vendor/fetch-office-js.sh`.

## 4. How do users install it?

Each user downloads `manifest.xml` and adds it through a trusted add-in catalog folder in Word's Trust Center settings, which takes a few minutes and is described in the README. The manifest is a small XML file that names the add-in, its icon and the hosted URL; it contains no code. Removing the catalog entry uninstalls the add-in.

## 5. What happens if the host goes down or you leave?

Existing documents are unaffected. Every diagram is a normal PNG picture, so documents still open, print and share as normal; only editing diagrams through the add-in is unavailable until the host is back. Because the source of each diagram is stored in the document itself, nothing is lost, and the WaveJSON can be copied out of the alt-text by hand. The code is two small files with a README, so another engineer could take it over, and the repository can be moved into a company GitHub organization.

## 6. Does the diagram source create hidden data in our documents?

Yes, by design: each diagram's WaveJSON source sits in the picture's alt-text, which is not visible on the page. It contains the same signal names and labels the picture shows, but it survives cropping, covering or blurring the picture, and Word's HTML export keeps it. Anyone redacting a document for external release should delete or regenerate the alt-text of any diagram they changed visually. PDF export drops alt-text.
