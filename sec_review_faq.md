# Waveform Viewer/Editor: likely follow-up questions

## 1. Can the add-in read or leak the contents of our documents?

It has the standard `ReadWriteDocument` permission, which covers only the document that is open when the user clicks a button, and in practice it reads only the alt-text (Word) or shape tag (PowerPoint) of the picture the user selected. The code makes no network requests of its own: no telemetry, analytics or API calls. The only traffic is the initial download of its own files and Microsoft's `office.js`. IT can confirm this by reading the ~500-line source or by watching the network traffic while it runs.

## 2. What do you need from the hosting location?

Static HTTPS file hosting for about 120 KB of files (`taskpane.html`, two JavaScript files and two icons), with no server-side code, database or special configuration. The certificate must be one that users' machines already trust, because Office refuses to load add-ins from self-signed or invalid-certificate URLs. An internal-only host is fine as long as everyone who uses the add-in can reach it over HTTPS.

## 3. How do updates work, and who can change the code?

To update, copy the new files to the host and bump the version number in the manifest's taskpane URL (Office caches the add-in aggressively, so the URL change forces a reload). Today the code runs from GitHub Pages, so anyone with push access to that repository could change what runs inside users' Office. Moving it to a company host puts every change under IT's control, which is a main reason for this request. There is no build step and nothing is fetched from package registries at runtime: WaveDrom (v3.6.2, MIT license) is a fixed copy in the repository, and the only other external code is Microsoft's `office.js`.

## 4. How do users install it?

Each user downloads `manifest.xml` and adds it through a trusted add-in catalog folder in Word's Trust Center settings, which takes a few minutes and is described in the README. The manifest is a small XML file that names the add-in, its icon and the hosted URL; it contains no code. Removing the catalog entry uninstalls the add-in.

## 5. What happens if the host goes down or you leave?

Existing documents are unaffected. Every diagram is a normal PNG picture, so documents still open, print and share as normal; only editing diagrams through the add-in is unavailable until the host is back. Because the source of each diagram is stored in the document itself, nothing is lost, and the WaveJSON can be copied out of the alt-text by hand. The code is a single small file with a README, so another engineer could take it over, and the repository can be moved into a company GitHub organization.
