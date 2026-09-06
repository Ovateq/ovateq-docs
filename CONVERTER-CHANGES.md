# Document Converter — added to this build

This is your `ovateq-docs-app-dist-v37-fixed` build with the document
converter **built into the same app**, not a separate deploy. Nothing in
`assets/` (the compiled app bundle) was touched, so the app you already have
is unaffected — this only adds a floating **"Convert"** button that opens an
overlay panel on top of it, on every screen.

## What changed vs. the original zip

- **Added** `converter/converter.js` and `converter/converter.css` — a
  small, dependency-free vanilla JS module (no React, no build step).
- **Edited** `index.html` — two lines added: one `<link>` for the CSS, one
  `<script defer>` for the JS, right before `</body>`. Nothing else touched.
- **Edited** `sw.js` — added the two new files to the service worker's
  precache list, and updated `index.html`'s cache revision so people who
  already have the app installed actually receive this update on their next
  visit (otherwise the old cached `index.html` would keep being served).

No Flask/Python backend is included or required — everything runs in the
browser, which is what makes this deployable as-is to GitHub Pages
(a backend server can't run there).

## Why not the Flask backend from `files.zip`?

The `files.zip` package (`app.py` + `DocumentConverter.jsx`) assumed a
Python server running alongside the app. Your GitHub Pages site
(`nogafa-missions.github.io/ovateq-docs`) only serves static files — there's
nowhere for a Flask process to run. So the five conversion features were
rebuilt to run entirely client-side instead:

| Feature | How it works now |
|---|---|
| PDF → Word (.docx) | PDF.js extracts the text in the browser; a small built-in writer packages it as a real `.docx` |
| PDF → Text (.txt) | Same text extraction, saved as plain text |
| PDF → Images (.png) | Each page is rendered to a canvas and saved as PNG (zipped if there's more than one page) |
| Images → PDF | Selected images are combined into one PDF with jsPDF |
| Scan to Document | Tesseract.js OCRs the photo in-browser, then the same `.docx` writer builds the Word file |

**Trade-off to know about:** PDF.js, jsPDF, JSZip and Tesseract.js are
loaded from a CDN the first time someone opens the converter panel (they're
too large to hand-roll reliably). This means that specific action needs
internet access the first time; the app shell itself still works fully
offline as before. After that first load, the browser caches those library
files normally.

**Also know:** the PDF → Word conversion extracts text only (no table/image
reconstruction). Building real layout-aware table detection client-side
without a PDF-analysis library like `pdfplumber` isn't something that can be
done reliably in the time available, so the honest choice was a converter
that always works over one that sometimes silently mangles a table.

## Where the button lives

A small pill-shaped **"Convert"** button is fixed to the bottom-left of the
screen, raised above the safe-area inset so it clears a bottom tab bar. It's
outside the app's own `#root` element, so it can never conflict with the
app's own state or re-renders. If it visually clashes with something in your
actual app once you see it live, the position is one CSS rule to change:
`#ovq-conv-fab` in `converter/converter.css` (`left`/`bottom`/`right`).

## Deploying

This zip's inner folder is a drop-in replacement for whatever you currently
publish to GitHub Pages for `nogafa-missions/ovateq-docs`:

1. Unzip it.
2. Copy everything **inside** `ovateq-docs-app-dist-v36-final/` into the
   branch/folder your Pages site serves from (e.g. the `gh-pages` branch, or
   `/docs` on `main` — whichever this repo already uses).
3. Commit and push. GitHub Pages will redeploy automatically.
4. Because the service worker's `index.html` revision changed, people who
   already have the app open will pick up the update automatically the next
   time they load it (the existing "Ovateq Docs couldn't start" recovery
   screen is unaffected — that logic wasn't touched).

## Limitation worth knowing

This whole patch was applied on top of the **compiled** `dist` output — the
zip you provided doesn't contain the app's source code, only its build
artifacts. That's why this was done as a bolt-on overlay rather than a true
integrated in-app screen/route: the minified bundle in `assets/` can't be
safely hand-edited. If you have the source repo, the cleaner long-term move
is to port `converter.js`'s logic into a real component/route in the app and
rebuild — happy to help with that if you share the source.
