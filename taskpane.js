let hostApp = '';
const src = document.getElementById('src');
const target = document.getElementById('target');
const status = document.getElementById('status');
const ALT_TEXT_SENTINEL = 'WAVEJSON:v1:';

function setStatus(text, kind) {
  status.textContent = text;
  status.className = kind ? 'status-' + kind : '';
}

function checkWaveLengths(json) {
  if (!json || !Array.isArray(json.signal)) return null;
  const lengths = json.signal
    .filter((s) => s && typeof s.wave === 'string')
    .map((s) => ({ name: s.name || '(unnamed)', len: s.wave.length }));
  if (lengths.length < 2) return null;
  const first = lengths[0].len;
  if (lengths.every((l) => l.len === first)) return null;
  return 'wave length mismatch: ' + lengths.map((l) => l.name + '=' + l.len).join(', ');
}

async function withButtonsDisabled(fn) {
  const buttons = document.querySelectorAll('#actions button');
  buttons.forEach((b) => { b.disabled = true; });
  try {
    await fn();
  } finally {
    buttons.forEach((b) => { b.disabled = false; });
  }
}

// WaveDrom turns arrays in name/text fields into raw SVG elements with whatever
// attributes they list, so a diagram from someone else's document can carry
// markup. on* handlers: the CSP already blocks them, this just makes it loud.
// The rest the CSP can't stop, because it doesn't govern navigation: href and
// attributeName (via <set>) make links, xmlns lets in HTML such as
// <meta http-equiv="refresh">, which redirects the taskpane on render.
function assertNoActiveMarkup(value) {
  if (value === null || typeof value !== 'object') return;
  for (const [key, child] of Object.entries(value)) {
    if (/^on|href$|^xmlns|^attributename$/i.test(key)) {
      throw new Error('"' + key + '" is not allowed in a diagram (it could run script or navigate the taskpane)');
    }
    assertNoActiveMarkup(child);
  }
}

function parseWaveJSON(text) {
  let json;
  // Strict JSON first: the loose rewrite below mangles apostrophes inside
  // double-quoted strings (e.g. "don't care"), and loaded diagrams are strict JSON.
  try {
    json = JSON.parse(text);
  } catch (e) {
    json = JSON.parse(text
      .replace(/'([^']*)'/g, '"$1"')
      .replace(/([{,]\s*)([A-Za-z_$][A-Za-z0-9_$]*)\s*:/g, '$1"$2":'));
  }
  if (json === null || typeof json !== 'object' || Array.isArray(json)) {
    throw new Error('top level must be an object, like { signal: [...] }');
  }
  assertNoActiveMarkup(json);
  return json;
}

function render() {
  let json;
  try {
    json = parseWaveJSON(src.value);
    const warning = checkWaveLengths(json);
    setStatus(warning || '', warning ? 'warning' : '');
  } catch (e) {
    setStatus('parse error: ' + e.message, 'error');
    return;
  }
  // Hand WaveDrom the object directly. ProcessAll() would need the source spliced
  // into a <script> tag via innerHTML, where a "</script>" inside any string breaks
  // out into live HTML, and it eval()s the text, which the CSP forbids.
  WaveDrom.RenderWaveForm(0, json, 'WaveDrom_Display_');
}

function describeError(e) {
  // Both: common-API errors (setSelectedDataAsync) carry a bare numeric code
  // with the useful text in message.
  return (e && [e.code, e.message].filter(Boolean).join(': ')) + ' | ' + JSON.stringify(e && e.debugInfo);
}

// Rasterize at 2x so the diagram stays sharp when enlarged or projected. The PNG
// is then twice its intended size, so callers must set the picture's size
// explicitly from the returned width/height (in points: 1 CSS px = 0.75 pt).
const RENDER_SCALE = 2;
function svgToPng() {
  const svg = target.querySelector('svg');
  if (!svg) return Promise.resolve(null);
  const svgString = new XMLSerializer().serializeToString(svg);
  const width = svg.width.baseVal.value || svg.getBoundingClientRect().width;
  const height = svg.height.baseVal.value || svg.getBoundingClientRect().height;

  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width = width * RENDER_SCALE;
      canvas.height = height * RENDER_SCALE;
      canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
      resolve({
        base64: canvas.toDataURL('image/png').split(',')[1],
        width: width * 0.75,
        height: height * 0.75,
      });
    };
    img.onerror = reject;
    img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgString)));
  });
}

const TEMPLATES = {
  tplClock: { signal: [ { name: 'clk', wave: 'p.......' } ]},
  tplBus: { signal: [
    { name: 'clk',  wave: 'p.......' },
    { name: 'data', wave: 'x.345x..', data: ['a', 'b', 'c'] }
  ]},
  tplHandshake: { signal: [
    { name: 'clk', wave: 'p.......' },
    { name: 'req', wave: '0.1..0..' },
    { name: 'ack', wave: '0...1.0.' }
  ]},
  tplEdges: {
    signal: [
      { name: 'clk', wave: 'p......' },
      { name: 'req', wave: '0.1...0', node: '..a....' },
      { name: 'ack', wave: '0....10', node: '.....b.' }
    ],
    edge: ['a~>b']
  }
};

// Word keeps in-line-with-text pictures and floating (wrapped) pictures in separate
// collections. Floating ones are only reachable via WordApiDesktop 1.2, i.e. desktop
// Word, not Word on the web. Returns { inline } or { floating } or {} if neither.
async function getSelectedWordPicture(context, sel) {
  const inline = sel.inlinePictures;
  inline.load('items');
  await context.sync();
  if (inline.items.length > 0) return { inline: inline.items[0] };
  if (!Office.context.requirements.isSetSupported('WordApiDesktop', '1.2')) return {};
  const shapes = sel.shapes;
  shapes.load('items/type');
  await context.sync();
  console.log('word selection: floating shape types', shapes.items.map((s) => s.type));
  return { floating: shapes.items.find((s) => s.type === Word.ShapeType.picture) };
}

async function insertOrReplace(mode) {
  await withButtonsDisabled(async () => {
    let png;
    try {
      png = await svgToPng();
    } catch (e) {
      setStatus('rasterize failed: ' + describeError(e), 'error');
      return;
    }
    if (!png) { setStatus('nothing rendered to insert', 'error'); return; }
    const base64 = png.base64;
    const aspect = png.height / png.width;
    
    let json;
    try {
      json = parseWaveJSON(src.value);
    } catch (e) {
      setStatus('parse error, not inserting: ' + e.message, 'error');
      return;
    }
    
    try {
      if (hostApp === Office.HostType.Word) {
        let noSelection = false;
        await Word.run(async (context) => {
          const sel = context.document.getSelection();
          let pic;
          if (mode === 'replace') {
            const { inline, floating } = await getSelectedWordPicture(context, sel);
            // Keep the user's width; height follows the new diagram's aspect ratio
            // so an added signal doesn't get squashed.
            if (inline) {
              inline.load('width');
              await context.sync();
              // Replace the picture itself, not the whole selection, so any text
              // selected along with it isn't wiped out.
              pic = inline.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
              pic.width = inline.width;
              pic.height = inline.width * aspect;
            } else if (floating) {
              floating.load('left,top,width,relativeHorizontalPosition,relativeVerticalPosition,textWrap/type');
              await context.sync();
              // There's no "swap the image" call for a floating picture, so insert a
              // new one anchored at the selection and copy the old one's layout over.
              // Wrap and the relative-to frame go first, since left/top are measured from it.
              pic = sel.insertPictureFromBase64(base64, { width: floating.width, height: floating.width * aspect });
              pic.textWrap.type = floating.textWrap.type;
              pic.relativeHorizontalPosition = floating.relativeHorizontalPosition;
              pic.relativeVerticalPosition = floating.relativeVerticalPosition;
              pic.left = floating.left;
              pic.top = floating.top;
              floating.delete();
            } else {
              noSelection = true;
              return;
            }
          } else {
            pic = sel.insertInlinePictureFromBase64(base64, Word.InsertLocation.after);
            pic.width = png.width;
            pic.height = png.height;
          }
          pic.altTextDescription = ALT_TEXT_SENTINEL + JSON.stringify(json);
          await context.sync();
        });
        setStatus(
          noSelection ? 'nothing selected to replace -- select a diagram first' : (mode === 'replace' ? 'replaced' : 'inserted'),
          noSelection ? 'error' : 'success'
        );

      } else if (hostApp === Office.HostType.PowerPoint) {
        // PowerPoint.run has no insert-picture call, so we go through the common
        // setSelectedDataAsync API. The inserted picture is NOT reliably selected
        // afterwards (tagging "the selection" hit the old shape or nothing), so we
        // find it by diffing the slide's shape ids before and after the insert.
        let beforeIds;
        let old = null;
        await PowerPoint.run(async (context) => {
          const slideShapes = context.presentation.getSelectedSlides().getItemAt(0).shapes;
          slideShapes.load('items/id');
          const selected = context.presentation.getSelectedShapes();
          selected.load('items/id,items/type,items/left,items/top,items/width');
          await context.sync();
          beforeIds = new Set(slideShapes.items.map((s) => s.id));
          const s = selected.items[0];
          if (mode === 'replace' && s) {
            old = { id: s.id, left: s.left, top: s.top, width: s.width };
          }
        });
        if (mode === 'replace' && old === null) {
          setStatus('nothing selected to replace -- select a diagram first', 'error');
          return;
        }

        await new Promise((resolve, reject) => {
          Office.context.document.setSelectedDataAsync(base64, { coercionType: Office.CoercionType.Image }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) reject(result.error);
            else resolve();
          });
        });

        await PowerPoint.run(async (context) => {
          const slideShapes = context.presentation.getSelectedSlides().getItemAt(0).shapes;
          slideShapes.load('items/id');
          await context.sync();
          const added = slideShapes.items.filter((s) => !beforeIds.has(s.id));
          console.log('ppt insert: new shape ids', added.map((s) => s.id));
          if (added.length !== 1) {
            throw new Error('expected exactly 1 new shape on the slide after insert, found ' + added.length);
          }
          const pic = added[0];
          // Pictures have no writable alt-text here, so the source lives in a shape tag
          // (PowerPoint upper-cases tag keys, hence the all-caps name).
          const tagValue = ALT_TEXT_SENTINEL + JSON.stringify(json);
          pic.tags.add('WAVEJSON', tagValue);
          pic.load('left,top,width,height');
          await context.sync();

          // Tag value length limits aren't documented, so read it back to catch truncation.
          const tag = pic.tags.getItem('WAVEJSON');
          tag.load('value');
          await context.sync();
          if (tag.value !== tagValue) {
            throw new Error('tag readback mismatch: wrote ' + tagValue.length + ' chars, read ' + tag.value.length);
          }

          // Geometry is set directly rather than via setSelectedDataAsync's imageLeft/
          // imageTop/... options, which PowerPoint ignored (new picture landed centered).
          // Replace: old position and width, height from the new aspect ratio.
          // Insert: intended size (the 2x PNG came in double size), re-centered where
          // PowerPoint put it.
          const width = old ? old.width : png.width;
          const height = width * aspect;
          pic.left = old ? old.left : pic.left + (pic.width - width) / 2;
          pic.top = old ? old.top : pic.top + (pic.height - height) / 2;
          pic.width = width;
          pic.height = height;
          // setSelectedDataAsync adds a picture rather than replacing the selected one,
          // so delete the old diagram ourselves (if PowerPoint already did, it's gone).
          const oldShape = old && slideShapes.items.find((s) => s.id === old.id);
          if (oldShape) oldShape.delete();
          await context.sync();
        });
        setStatus(mode === 'replace' ? 'replaced' : 'inserted', 'success');
      } else {
        throw new Error('unsupported host: ' + hostApp);
      }
    } catch (e) {
      setStatus('insert failed: ' + describeError(e), 'error');
    }
  });
}

(function setupResizeHandle() {
  const handle = document.getElementById('resizeHandle');
  let startY = 0;
  let startHeight = 0;

  function onMove(e) {
    src.style.height = Math.max(80, startHeight + (e.clientY - startY)) + 'px';
  }
  function onUp() {
    window.removeEventListener('mousemove', onMove);
    window.removeEventListener('mouseup', onUp);
  }
  handle.addEventListener('mousedown', (e) => {
    startY = e.clientY;
    startHeight = src.offsetHeight;
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    e.preventDefault();
  });
})();

Office.onReady((info) => {
  hostApp = info.host;
  render();

  document.getElementById('insertBtn').addEventListener('click', () => insertOrReplace('insert'));
  document.getElementById('replaceBtn').addEventListener('click', () => insertOrReplace('replace'));

  for (const id of Object.keys(TEMPLATES)) {
    document.getElementById(id).addEventListener('click', () => {
      src.value = JSON.stringify(TEMPLATES[id], null, 2);
      render();
    });
  }

  document.getElementById('loadBtn').addEventListener('click', () => {
    withButtonsDisabled(async () => {
      
      function processLoadedData(raw) {
        if (!raw.startsWith(ALT_TEXT_SENTINEL)) {
          setStatus('selected image has no recognizable waveform source (not inserted by this tool)', 'error');
          return;
        }
        let json;
        try {
          json = JSON.parse(raw.slice(ALT_TEXT_SENTINEL.length));
        } catch (e) {
          setStatus('diagram source is corrupted, refusing to load: ' + e.message, 'error');
          return;
        }
        // Checked before it reaches the editor, since a document from someone else
        // is exactly where a hostile diagram would come from.
        try {
          assertNoActiveMarkup(json);
        } catch (e) {
          setStatus('diagram source is unsafe, refusing to load: ' + e.message, 'error');
          return;
        }
        src.value = JSON.stringify(json, null, 2);
        render();
        setStatus('loaded from selection', 'success');
      }

      try {
        if (hostApp === Office.HostType.Word) {
          await Word.run(async (context) => {
            const { inline, floating } = await getSelectedWordPicture(context, context.document.getSelection());
            const pic = inline || floating;
            if (!pic) {
              setStatus('select a diagram image first', 'error');
              return;
            }
            pic.load('altTextDescription');
            await context.sync();

            const raw = pic.altTextDescription || '';
            processLoadedData(raw);
          });
          
        } else if (hostApp === Office.HostType.PowerPoint) {
          await PowerPoint.run(async (context) => {
            const shapes = context.presentation.getSelectedShapes();
            shapes.load('items');
            await context.sync();
            
            if (shapes.items.length === 0) {
              setStatus('select a diagram image first', 'error');
              return;
            }
            
            const shape = shapes.items[0];
            const tag = shape.tags.getItemOrNullObject('WAVEJSON');
            tag.load('value');
            await context.sync();
            if (tag.isNullObject) {
              // Show what PowerPoint actually has on this shape, so a failed load
              // (wrong shape selected, unsaved deck, older build) is diagnosable.
              shape.load('id,type,name');
              shape.tags.load('key');
              await context.sync();
              const keys = shape.tags.items.map((t) => t.key);
              console.log('ppt load: no WAVEJSON tag', shape.id, shape.type, shape.name, keys);
              setStatus('selected shape has no WAVEJSON tag -- type=' + shape.type +
                ', name="' + shape.name + '", tags=[' + keys.join(', ') + ']', 'error');
              return;
            }
            processLoadedData(tag.value);
          });
        } else {
          throw new Error('unsupported host: ' + hostApp);
        }
      } catch (e) {
        setStatus('load failed: ' + describeError(e), 'error');
      }
    });
  });

  document.getElementById('clearBtn').addEventListener('click', () => {
    src.value = JSON.stringify({ signal: [{ name: '', wave: '' }] }, null, 2);
    render();
    setStatus('cleared', 'success');
  });

  document.getElementById('addSignalBtn').addEventListener('click', () => {
    let json;
    try {
      json = parseWaveJSON(src.value);
    } catch (e) {
      setStatus('fix the parse error before adding a signal: ' + e.message, 'error');
      return;
    }
    if (!Array.isArray(json.signal)) json.signal = [];
    json.signal.push({ name: '', wave: '' });
    src.value = JSON.stringify(json, null, 2);
    render();
  });

  document.getElementById('copyBtn').addEventListener('click', () => {
    src.focus();
    src.select();
    setStatus('JSON selected above -- press Ctrl+C to copy', 'success');
  });
});

src.addEventListener('input', render);

src.addEventListener('keydown', (e) => {
  if (e.key === 'Tab') {
    e.preventDefault();
    e.stopPropagation();
    const start = src.selectionStart;
    const end = src.selectionEnd;
    src.value = src.value.slice(0, start) + '  ' + src.value.slice(end);
    src.selectionStart = src.selectionEnd = start + 2;
    render();
  }
});
