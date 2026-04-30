(function () {
  'use strict';

  function text(value) {
    return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
  }

  function rowsFromTable(table) {
    if (!table) return [];
    var out = [];
    var trs = table.querySelectorAll('tr');
    Array.prototype.forEach.call(trs, function (tr) {
      var cells = tr.querySelectorAll('th,td');
      if (!cells.length) return;
      out.push(Array.prototype.map.call(cells, function (cell) {
        return text(cell.innerText || cell.textContent || '');
      }));
    });
    return out;
  }

  function rowsFromObjects(rows) {
    if (!Array.isArray(rows) || !rows.length) return [];
    var keys = [];
    var seen = {};
    rows.slice(0, 60).forEach(function (row) {
      Object.keys(row || {}).forEach(function (key) {
        if (!seen[key] && typeof row[key] !== 'object') {
          seen[key] = 1;
          keys.push(key);
        }
      });
    });
    if (!keys.length) return [];
    var out = [keys];
    rows.slice(0, 60).forEach(function (row) {
      out.push(keys.map(function (key) { return text(row && row[key]); }));
    });
    if (rows.length > 60) out.push(['Showing first 60 of ' + rows.length + ' rows']);
    return out;
  }

  function roundedRect(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
  }

  function drawWrapped(ctx, value, x, y, maxWidth, lineHeight) {
    var words = text(value).split(' ');
    var line = '';
    var lines = [];
    words.forEach(function (word) {
      var next = line ? line + ' ' + word : word;
      if (ctx.measureText(next).width > maxWidth && line) {
        lines.push(line);
        line = word;
      } else {
        line = next;
      }
    });
    if (line) lines.push(line);
    lines.slice(0, 2).forEach(function (ln, i) {
      ctx.fillText(ln, x, y + i * lineHeight);
    });
  }

  function canvasFromMatrix(matrix) {
    var rows = (matrix || []).filter(function (r) { return Array.isArray(r) && r.length; });
    if (!rows.length) throw new Error('no_table_rows');

    var maxCols = Math.min(Math.max.apply(null, rows.map(function (r) { return r.length; })), 12);
    rows = rows.map(function (r) { return r.slice(0, maxCols); });

    var probe = document.createElement('canvas').getContext('2d');
    probe.font = '600 13px Arial, sans-serif';
    var widths = [];
    for (var c = 0; c < maxCols; c++) {
      var width = 84;
      rows.slice(0, 61).forEach(function (row) {
        probe.font = row === rows[0] ? '700 13px Arial, sans-serif' : '500 13px Arial, sans-serif';
        width = Math.max(width, Math.min(230, probe.measureText(text(row[c])).width + 28));
      });
      widths.push(Math.ceil(width));
    }

    var rowHeight = 42;
    var titleHeight = 44;
    var pad = 18;
    var tableWidth = widths.reduce(function (sum, w) { return sum + w; }, 0);
    var canvas = document.createElement('canvas');
    canvas.width = Math.min(4096, tableWidth + pad * 2);
    canvas.height = Math.min(4096, titleHeight + rows.length * rowHeight + pad * 2);
    var ctx = canvas.getContext('2d');

    ctx.fillStyle = '#F8FAFC';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#FFFFFF';
    roundedRect(ctx, pad, pad, tableWidth, canvas.height - pad * 2, 8);
    ctx.fill();

    ctx.fillStyle = '#8B1A2D';
    roundedRect(ctx, pad, pad, tableWidth, titleHeight, 8);
    ctx.fill();
    ctx.fillStyle = '#FFFFFF';
    ctx.font = '700 16px Arial, sans-serif';
    ctx.fillText('NLPL AI Table', pad + 14, pad + 28);

    var y = pad + titleHeight;
    rows.forEach(function (row, ri) {
      var x = pad;
      ctx.fillStyle = ri === 0 ? '#F1F5F9' : (ri % 2 ? '#FFFFFF' : '#FAFAFA');
      ctx.fillRect(pad, y, tableWidth, rowHeight);
      widths.forEach(function (w, ci) {
        ctx.strokeStyle = '#E2E8F0';
        ctx.strokeRect(x, y, w, rowHeight);
        ctx.fillStyle = ri === 0 ? '#111827' : '#1F2937';
        ctx.font = ri === 0 ? '700 12px Arial, sans-serif' : '500 12px Arial, sans-serif';
        drawWrapped(ctx, row[ci], x + 10, y + 17, w - 20, 15);
        x += w;
      });
      y += rowHeight;
    });
    return canvas;
  }

  function writeCanvas(canvas, fileName) {
    return new Promise(function (resolve, reject) {
      canvas.toBlob(function (blob) {
        if (!blob) return reject(new Error('blob_failed'));
        if (navigator.clipboard && window.ClipboardItem) {
          navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]).then(resolve).catch(function () {
            downloadBlob(blob, fileName || 'nlpl_ai_table.png');
            resolve();
          });
        } else {
          downloadBlob(blob, fileName || 'nlpl_ai_table.png');
          resolve();
        }
      }, 'image/png');
    });
  }

  function downloadBlob(blob, fileName) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fileName || 'nlpl_ai_table.png';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 300);
  }

  window._aiCopyTableAsImage = function (target, options) {
    options = options || {};
    var table = target && target.matches && target.matches('table')
      ? target
      : target && target.querySelector
        ? target.querySelector('table')
        : null;
    var matrix = table ? rowsFromTable(table) : rowsFromObjects(options.rows);
    var canvas = canvasFromMatrix(matrix);
    return writeCanvas(canvas, options.fileName);
  };

  document.addEventListener('click', function (event) {
    var btn = event.target && event.target.closest && event.target.closest('.ai-row-bot [data-act="copy"]');
    if (!btn) return;
    var row = btn.closest('.ai-row-bot');
    if (!row || !row.querySelector('.ai-table')) return;
    btn.setAttribute('title', 'Copy table as image');
    btn.setAttribute('aria-label', 'Copy table image');
  }, true);
})();
