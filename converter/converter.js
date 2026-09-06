/*!
 * Ovateq Docs — Document Converter
 * Self-contained, framework-free feature bolted onto the built app.
 * Everything runs client-side in the browser. No backend server involved:
 * PDF parsing/rendering and OCR run third-party libraries loaded on demand
 * from a CDN the first time this panel is opened (then cached by the
 * browser like any other asset).
 *
 * Features:
 *  - PDF -> Word (.docx)      (text extraction, one paragraph per line)
 *  - PDF -> Text (.txt)
 *  - PDF -> Images (.png)     (zipped if the PDF has more than one page)
 *  - Images -> PDF            (combine one or more images into one .pdf)
 *  - Scan to Document (.docx) (camera/photo -> OCR -> Word, all in-browser)
 */
(function () {
  'use strict';

  if (window.__ovqConverterMounted) return;
  window.__ovqConverterMounted = true;

  // ---------------------------------------------------------------------
  // CDN libraries, loaded lazily (only once, only when actually needed)
  // ---------------------------------------------------------------------
  var CDN = {
    pdfjs: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
    pdfjsWorker: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
    jszip: 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
    jspdf: 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
    tesseract: 'https://cdn.jsdelivr.net/npm/tesseract.js@5.0.4/dist/tesseract.min.js'
  };
  var scriptPromises = {};
  function loadScript(url) {
    if (scriptPromises[url]) return scriptPromises[url];
    scriptPromises[url] = new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.src = url;
      s.async = true;
      s.onload = function () { resolve(); };
      s.onerror = function () {
        delete scriptPromises[url];
        reject(new Error('Could not load a required library. Check your internet connection and try again.'));
      };
      document.head.appendChild(s);
    });
    return scriptPromises[url];
  }
  function ensurePdfJs() {
    return loadScript(CDN.pdfjs).then(function () {
      if (window.pdfjsLib && !window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = CDN.pdfjsWorker;
      }
      return window.pdfjsLib;
    });
  }
  function ensureJsZip() { return loadScript(CDN.jszip).then(function () { return window.JSZip; }); }
  function ensureJsPdf() { return loadScript(CDN.jspdf).then(function () { return window.jspdf.jsPDF; }); }
  function ensureTesseract() { return loadScript(CDN.tesseract).then(function () { return window.Tesseract; }); }

  // ---------------------------------------------------------------------
  // Minimal .docx (OOXML) writer — no dependency beyond JSZip for the
  // zip container itself. Produces a valid Word document with headings
  // and plain paragraphs.
  // ---------------------------------------------------------------------
  function escapeXml(str) {
    return String(str == null ? '' : str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  // blocks: array of { type: 'heading'|'para'|'pageBreak', text }
  function buildDocx(blocks) {
    return ensureJsZip().then(function (JSZip) {
      var bodyXml = blocks.map(function (b) {
        if (b.type === 'pageBreak') {
          return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
        }
        var runProps = b.type === 'heading' ? '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>' : '';
        var lines = String(b.text || '').split('\n');
        var runs = lines.map(function (line, i) {
          var br = i < lines.length - 1 ? '<w:br/>' : '';
          return '<w:r>' + runProps + '<w:t xml:space="preserve">' + escapeXml(line) + '</w:t>' + br + '</w:r>';
        }).join('');
        return '<w:p>' + runs + '</w:p>';
      }).join('');

      var documentXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body>' + bodyXml +
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
        '</w:sectPr></w:body></w:document>';

      var contentTypes =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
        '</Types>';

      var rootRels =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
        '</Relationships>';

      var nowIso = new Date().toISOString();
      var coreXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        '<dc:title>Converted document</dc:title>' +
        '<dc:creator>Ovateq Docs</dc:creator>' +
        '<dcterms:created xsi:type="dcterms:W3CDTF">' + nowIso + '</dcterms:created>' +
        '<dcterms:modified xsi:type="dcterms:W3CDTF">' + nowIso + '</dcterms:modified>' +
        '</cp:coreProperties>';

      var appXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
        '<Application>Ovateq Docs Converter</Application></Properties>';

      var zip = new JSZip();
      zip.file('[Content_Types].xml', contentTypes);
      zip.folder('_rels').file('.rels', rootRels);
      zip.folder('word').file('document.xml', documentXml);
      zip.folder('docProps').file('core.xml', coreXml);
      zip.folder('docProps').file('app.xml', appXml);

      return zip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
    });
  }

  // ---------------------------------------------------------------------
  // PDF helpers
  // ---------------------------------------------------------------------
  function readFileAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(new Error('Could not read the selected file.')); };
      fr.readAsArrayBuffer(file);
    });
  }

  // Group a page's text items into reading-order lines using their
  // vertical position, since pdf.js gives flat text items, not lines.
  function textContentToLines(textContent) {
    var items = textContent.items.filter(function (it) { return it.str !== undefined; });
    var lines = [];
    var current = null;
    var lastY = null;
    items.forEach(function (it) {
      var y = it.transform[5];
      if (current === null || lastY === null || Math.abs(y - lastY) > 4) {
        current = [];
        lines.push(current);
      }
      current.push(it.str);
      lastY = y;
    });
    return lines.map(function (parts) { return parts.join(' ').replace(/\s+/g, ' ').trim(); })
      .filter(function (l) { return l.length > 0; });
  }

  function loadPdf(arrayBuffer) {
    return ensurePdfJs().then(function (pdfjsLib) {
      return pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    });
  }

  function pdfToWord(file, onProgress) {
    return readFileAsArrayBuffer(file).then(loadPdf).then(function (pdf) {
      var blocks = [];
      var chain = Promise.resolve();
      for (var i = 1; i <= pdf.numPages; i++) {
        (function (pageNum) {
          chain = chain.then(function () {
            onProgress && onProgress(Math.round((pageNum - 1) / pdf.numPages * 90));
            return pdf.getPage(pageNum).then(function (page) {
              return page.getTextContent().then(function (tc) {
                var lines = textContentToLines(tc);
                if (lines.length === 0) lines = ['(No selectable text found on this page.)'];
                lines.forEach(function (line) { blocks.push({ type: 'para', text: line }); });
                if (pageNum < pdf.numPages) blocks.push({ type: 'pageBreak' });
              });
            });
          });
        })(i);
      }
      return chain.then(function () {
        onProgress && onProgress(95);
        return buildDocx(blocks);
      });
    });
  }

  function pdfToText(file, onProgress) {
    return readFileAsArrayBuffer(file).then(loadPdf).then(function (pdf) {
      var pages = [];
      var chain = Promise.resolve();
      for (var i = 1; i <= pdf.numPages; i++) {
        (function (pageNum) {
          chain = chain.then(function () {
            onProgress && onProgress(Math.round((pageNum - 1) / pdf.numPages * 95));
            return pdf.getPage(pageNum).then(function (page) {
              return page.getTextContent().then(function (tc) {
                pages.push(textContentToLines(tc).join('\n'));
              });
            });
          });
        })(i);
      }
      return chain.then(function () {
        var text = pages.join('\n\n--- PAGE BREAK ---\n\n');
        return new Blob([text], { type: 'text/plain' });
      });
    });
  }

  function pdfToImages(file, onProgress) {
    return readFileAsArrayBuffer(file).then(loadPdf).then(function (pdf) {
      var pageBlobs = [];
      var chain = Promise.resolve();
      for (var i = 1; i <= pdf.numPages; i++) {
        (function (pageNum) {
          chain = chain.then(function () {
            onProgress && onProgress(Math.round((pageNum - 1) / pdf.numPages * 90));
            return pdf.getPage(pageNum).then(function (page) {
              var viewport = page.getViewport({ scale: 2 });
              var canvas = document.createElement('canvas');
              canvas.width = viewport.width;
              canvas.height = viewport.height;
              var ctx = canvas.getContext('2d');
              return page.render({ canvasContext: ctx, viewport: viewport }).promise.then(function () {
                return new Promise(function (resolve) {
                  canvas.toBlob(function (blob) {
                    pageBlobs.push({ name: 'page_' + pageNum + '.png', blob: blob });
                    resolve();
                  }, 'image/png');
                });
              });
            });
          });
        })(i);
      }
      return chain.then(function () {
        onProgress && onProgress(95);
        if (pageBlobs.length === 1) {
          return { blob: pageBlobs[0].blob, filename: pageBlobs[0].name, mime: 'image/png' };
        }
        return ensureJsZip().then(function (JSZip) {
          var zip = new JSZip();
          pageBlobs.forEach(function (p) { zip.file(p.name, p.blob); });
          return zip.generateAsync({ type: 'blob' }).then(function (blob) {
            return { blob: blob, filename: 'pages.zip', mime: 'application/zip' };
          });
        });
      });
    });
  }

  function loadImageElement(file) {
    return new Promise(function (resolve, reject) {
      var img = new Image();
      var url = URL.createObjectURL(file);
      img.onload = function () { resolve({ img: img, url: url }); };
      img.onerror = function () { URL.revokeObjectURL(url); reject(new Error('Could not read image: ' + file.name)); };
      img.src = url;
    });
  }

  function imagesToPdf(files, onProgress) {
    return ensureJsPdf().then(function (jsPDF) {
      var doc = null;
      var chain = Promise.resolve();
      files.forEach(function (file, idx) {
        chain = chain.then(function () {
          onProgress && onProgress(Math.round(idx / files.length * 90));
          return loadImageElement(file).then(function (res) {
            var w = res.img.naturalWidth, h = res.img.naturalHeight;
            var orientation = w >= h ? 'l' : 'p';
            var format = [w, h];
            if (!doc) {
              doc = new jsPDF({ orientation: orientation, unit: 'px', format: format, compress: true });
            } else {
              doc.addPage(format, orientation);
            }
            var canvas = document.createElement('canvas');
            canvas.width = w; canvas.height = h;
            canvas.getContext('2d').drawImage(res.img, 0, 0, w, h);
            var dataUrl = canvas.toDataURL('image/jpeg', 0.92);
            doc.addImage(dataUrl, 'JPEG', 0, 0, w, h);
            URL.revokeObjectURL(res.url);
          });
        });
      });
      return chain.then(function () {
        onProgress && onProgress(95);
        return doc.output('blob');
      });
    });
  }

  function scanToWord(file, title, onOcrProgress) {
    return ensureTesseract().then(function (Tesseract) {
      return Tesseract.recognize(file, 'eng', {
        logger: function (m) {
          if (m.status === 'recognizing text' && typeof m.progress === 'number') {
            onOcrProgress && onOcrProgress(Math.round(m.progress * 100));
          }
        }
      }).then(function (result) {
        var text = (result && result.data && result.data.text) || '';
        var blocks = [{ type: 'heading', text: title || 'Scanned Document' }];
        text.split('\n').forEach(function (line) {
          if (line.trim()) blocks.push({ type: 'para', text: line });
        });
        return buildDocx(blocks);
      });
    });
  }

  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  }

  // ---------------------------------------------------------------------
  // UI
  // ---------------------------------------------------------------------
  var FORMATS = {
    'pdf-to-word': {
      icon: '📄', title: 'PDF to Word', accept: '.pdf', multiple: false,
      desc: 'Turn a PDF into an editable Word document (.docx).'
    },
    'pdf-to-text': {
      icon: '📋', title: 'PDF to Text', accept: '.pdf', multiple: false,
      desc: 'Pull all the text out of a PDF as a plain .txt file.'
    },
    'pdf-to-images': {
      icon: '🖼️', title: 'PDF to Images', accept: '.pdf', multiple: false,
      desc: 'Save each PDF page as a PNG image.'
    },
    'images-to-pdf': {
      icon: '📑', title: 'Images to PDF', accept: '.png,.jpg,.jpeg,.gif,.webp,.bmp', multiple: true,
      desc: 'Combine one or more photos or images into a single PDF.'
    },
    'image-ocr-to-word': {
      icon: '📸', title: 'Scan to Document', accept: '.png,.jpg,.jpeg,.gif,.bmp,.webp', multiple: false,
      desc: 'Photo of a hard copy → text (OCR) → editable Word doc.'
    }
  };

  function buildMarkup() {
    var cards = Object.keys(FORMATS).map(function (key) {
      var f = FORMATS[key];
      return '' +
        '<button type="button" class="ovq-conv-card" data-format="' + key + '">' +
        '<div class="ovq-icon">' + f.icon + '</div>' +
        '<div class="ovq-t">' + f.title + '</div>' +
        '<div class="ovq-d">' + f.desc + '</div>' +
        '</button>';
    }).join('');

    var root = document.createElement('div');
    root.id = 'ovq-conv-root';
    root.innerHTML =
      '<div class="ovq-conv-backdrop" data-close="1"></div>' +
      '<div class="ovq-conv-panel" role="dialog" aria-modal="true" aria-label="Document Converter">' +
      '<div class="ovq-conv-head">' +
      '<div><h2>Document Converter</h2><p>Convert files right in your browser — nothing is uploaded.</p></div>' +
      '<button type="button" class="ovq-conv-close" data-close="1" aria-label="Close">✕</button>' +
      '</div>' +
      '<div class="ovq-conv-body">' +
      '<div class="ovq-conv-grid">' + cards + '</div>' +
      '<div class="ovq-conv-details" style="display:none;">' +
      '<h3 class="ovq-details-title"></h3>' +
      '<div class="ovq-file-list" style="display:none;"><p class="ovq-label">Selected:</p><div class="ovq-file-items"></div></div>' +
      '<button type="button" class="ovq-btn ovq-btn-outline ovq-select-btn">Select file(s)</button>' +
      '<div class="ovq-progress-wrap" style="display:none;"><div class="ovq-progress-bar"><div class="ovq-progress-fill" style="width:0%"></div></div><p class="ovq-progress-text">0%</p></div>' +
      '<div class="ovq-status-ok" style="display:none;"></div>' +
      '<div class="ovq-status-err" style="display:none;"></div>' +
      '<button type="button" class="ovq-btn ovq-btn-primary ovq-convert-btn" disabled>Convert now</button>' +
      '</div>' +
      '<p class="ovq-hint">PDF and image tools load a small open-source library the first time you use them. Everything then runs on this device — your files are never sent anywhere.</p>' +
      '</div>' +
      '</div>' +
      '<input type="file" class="ovq-file-input" style="display:none">';
    return root;
  }

  function mount() {
    var fab = document.createElement('button');
    fab.id = 'ovq-conv-fab';
    fab.type = 'button';
    fab.setAttribute('aria-label', 'Open Document Converter');
    fab.innerHTML =
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
      '<path d="M9 12h6M9 16h6M9 8h1"/><path d="M7 3h7l5 5v13a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"/>' +
      '</svg>' +
      '<span>Convert</span><span class="ovq-fab-dot"></span>';
    document.body.appendChild(fab);

    var root = buildMarkup();
    document.body.appendChild(root);

    var state = { format: null, files: [] };

    var detailsEl = root.querySelector('.ovq-conv-details');
    var detailsTitle = root.querySelector('.ovq-details-title');
    var fileListEl = root.querySelector('.ovq-file-list');
    var fileItemsEl = root.querySelector('.ovq-file-items');
    var selectBtn = root.querySelector('.ovq-select-btn');
    var convertBtn = root.querySelector('.ovq-convert-btn');
    var fileInput = root.querySelector('.ovq-file-input');
    var progWrap = root.querySelector('.ovq-progress-wrap');
    var progFill = root.querySelector('.ovq-progress-fill');
    var progText = root.querySelector('.ovq-progress-text');
    var okEl = root.querySelector('.ovq-status-ok');
    var errEl = root.querySelector('.ovq-status-err');

    function openModal() { root.classList.add('ovq-open'); document.documentElement.style.overflow = 'hidden'; }
    function closeModal() { root.classList.remove('ovq-open'); document.documentElement.style.overflow = ''; }

    function resetDetails() {
      state.files = [];
      fileListEl.style.display = 'none';
      fileItemsEl.innerHTML = '';
      progWrap.style.display = 'none';
      progFill.style.width = '0%';
      progText.textContent = '0%';
      okEl.style.display = 'none';
      errEl.style.display = 'none';
      convertBtn.disabled = true;
      convertBtn.textContent = 'Convert now';
    }

    function selectFormat(key) {
      state.format = key;
      resetDetails();
      var f = FORMATS[key];
      detailsTitle.textContent = f.title;
      detailsEl.style.display = 'block';
      root.querySelectorAll('.ovq-conv-card').forEach(function (c) {
        c.classList.toggle('ovq-active', c.getAttribute('data-format') === key);
      });
      fileInput.accept = f.accept;
      fileInput.multiple = !!f.multiple;
      detailsEl.scrollIntoView({ block: 'nearest' });
    }

    root.querySelectorAll('.ovq-conv-card').forEach(function (card) {
      card.addEventListener('click', function () {
        selectFormat(card.getAttribute('data-format'));
        fileInput.value = '';
        fileInput.click();
      });
    });

    selectBtn.addEventListener('click', function () { fileInput.click(); });

    fileInput.addEventListener('change', function () {
      var files = Array.prototype.slice.call(fileInput.files || []);
      if (!files.length) return;
      state.files = files;
      fileListEl.style.display = 'block';
      fileItemsEl.innerHTML = files.map(function (f) { return '<div class="ovq-file-item">📎 ' + escapeHtml(f.name) + '</div>'; }).join('');
      okEl.style.display = 'none';
      errEl.style.display = 'none';
      convertBtn.disabled = false;
    });

    function escapeHtml(s) {
      return String(s).replace(/[&<>"']/g, function (c) {
        return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
      });
    }

    function setProgress(pct) {
      progWrap.style.display = 'block';
      progFill.style.width = Math.max(0, Math.min(100, pct)) + '%';
      progText.textContent = Math.max(0, Math.min(100, pct)) + '%';
    }

    function showError(msg) {
      errEl.textContent = msg;
      errEl.style.display = 'block';
      okEl.style.display = 'none';
    }
    function showOk(msg) {
      okEl.textContent = msg;
      okEl.style.display = 'block';
      errEl.style.display = 'none';
    }

    convertBtn.addEventListener('click', function () {
      if (!state.format || state.files.length === 0) { showError('Please select a file first.'); return; }
      convertBtn.disabled = true;
      convertBtn.textContent = 'Converting…';
      okEl.style.display = 'none';
      errEl.style.display = 'none';
      setProgress(0);

      var done = function (blob, filename) {
        setProgress(100);
        downloadBlob(blob, filename);
        showOk('✅ Done — your download has started.');
        convertBtn.disabled = false;
        convertBtn.textContent = 'Convert now';
      };
      var fail = function (err) {
        showError((err && err.message) || 'Conversion failed. Please try again.');
        convertBtn.disabled = false;
        convertBtn.textContent = 'Convert now';
        progWrap.style.display = 'none';
      };

      try {
        if (state.format === 'pdf-to-word') {
          pdfToWord(state.files[0], setProgress).then(function (blob) { done(blob, 'converted.docx'); }, fail);
        } else if (state.format === 'pdf-to-text') {
          pdfToText(state.files[0], setProgress).then(function (blob) { done(blob, 'extracted.txt'); }, fail);
        } else if (state.format === 'pdf-to-images') {
          pdfToImages(state.files[0], setProgress).then(function (res) { done(res.blob, res.filename); }, fail);
        } else if (state.format === 'images-to-pdf') {
          imagesToPdf(state.files, setProgress).then(function (blob) { done(blob, 'combined.pdf'); }, fail);
        } else if (state.format === 'image-ocr-to-word') {
          var title = state.files[0].name.replace(/\.[^/.]+$/, '');
          scanToWord(state.files[0], title, setProgress).then(function (blob) { done(blob, 'scanned_document.docx'); }, fail);
        }
      } catch (e) {
        fail(e);
      }
    });

    fab.addEventListener('click', openModal);
    root.addEventListener('click', function (e) {
      if (e.target && e.target.getAttribute && e.target.getAttribute('data-close')) closeModal();
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && root.classList.contains('ovq-open')) closeModal();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', mount);
  } else {
    mount();
  }
})();
