/* Удаление кодов из Excel по PDF. Работает полностью в браузере.
 * Чистые функции (parsePdf, extractCodeAndNumber, parseRange, ...) изолированы
 * от DOM, чтобы их можно было переиспользовать в Node-тестах. */
(function () {
  'use strict';

  /* ============================ PDF ============================ */

  // Собирает элементы pdf.js в строки: кластер по координате Y, внутри — сортировка по X.
  function groupIntoLines(items) {
    var prepared = [];
    items.forEach(function (it) {
      var s = (it.str || '').replace(/\u00a0/g, ' ');
      if (!s.trim()) return;
      prepared.push({ str: s, x: it.transform[4], y: it.transform[5] });
    });
    prepared.sort(function (a, b) { return b.y - a.y || a.x - b.x; });

    var clusters = [];
    prepared.forEach(function (p) {
      var last = clusters[clusters.length - 1];
      if (last && Math.abs(p.y - last.y) <= 3) {
        last.parts.push(p);
        last.y = (last.y * (last.parts.length - 1) + p.y) / last.parts.length;
      } else {
        clusters.push({ y: p.y, parts: [p] });
      }
    });

    return clusters
      .map(function (c) {
        return c.parts
          .sort(function (a, b) { return a.x - b.x; })
          .map(function (p) { return p.str; })
          .join('')
          .trim();
      })
      .filter(Boolean);
  }

  // Ищет на странице номер этикетки (№N) и код под QR-кодом:
  // строка из 16 цифр, начинающаяся с 01 (AI 01 + GTIN), за которой следует
  // строка, начинающаяся с 21 (AI 21 + серийный номер); либо весь код одной строкой.
  function extractCodeAndNumber(lines) {
    var num = null;
    for (var i = 0; i < lines.length; i++) {
      var m = lines[i].match(/№\s*([0-9]+)\s*$/);
      if (m) { num = parseInt(m[1], 10); break; }
    }

    var code = null;
    for (var j = 0; j < lines.length; j++) {
      var l = lines[j];
      if (/^01[0-9]{14}$/.test(l)) {
        var nxt = lines[j + 1];
        if (nxt && /^21/.test(nxt)) { code = l + nxt.replace(/\s+/g, ''); break; }
      }
      var one = l.match(/^(01[0-9]{14}21\S+)$/);
      if (one) { code = one[1]; break; }
    }

    return code && num !== null ? { num: num, code: code } : null;
  }

  function parsePdf(arrayBuffer) {
    var data = new Uint8Array(arrayBuffer);
    return pdfjsLib.getDocument({ data: data }).promise.then(function (doc) {
      var entries = [];
      var failedPages = [];
      var chain = Promise.resolve();

      for (var p = 1; p <= doc.numPages; p++) {
        (function (pageNum) {
          chain = chain.then(function () {
            return doc.getPage(pageNum).then(function (page) {
              return page.getTextContent();
            }).then(function (tc) {
              var found = extractCodeAndNumber(groupIntoLines(tc.items));
              if (found) entries.push(found);
              else failedPages.push(pageNum);
            });
          });
        })(p);
      }

      return chain.then(function () {
        entries.sort(function (a, b) { return a.num - b.num; });
        return { numPages: doc.numPages, entries: entries, failedPages: failedPages };
      });
    });
  }

  /* ============================ Диапазоны ============================ */

  // Принимает: "" (все), "485", "485-500", "485-", а также "с 485" / "от 485".
  function parseRange(str, min, max) {
    var s = String(str == null ? '' : str).trim().toLowerCase();
    s = s.replace(/^(с|от)\s+/, '').replace(/[–—]/g, '-');
    if (!s) return { from: min, to: max };
    var m = s.match(/^([0-9]+)\s*-\s*([0-9]+)$/);
    if (m) return { from: parseInt(m[1], 10), to: parseInt(m[2], 10) };
    m = s.match(/^([0-9]+)\s*-\s*$/);
    if (m) return { from: parseInt(m[1], 10), to: max };
    m = s.match(/^([0-9]+)$/);
    if (m) return { from: parseInt(m[1], 10), to: max };
    return null;
  }

  /* ============================ Excel ============================ */

  function cellText(cell) {
    var v = cell.value;
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(function (r) { return r.text; }).join('');
      if (v.result !== undefined) return String(v.result);
      return String(cell.text != null ? cell.text : '');
    }
    return String(v);
  }

  // Лист, у которого в A1 заголовок «Код», иначе первый лист.
  function pickWorksheet(wb) {
    for (var i = 0; i < wb.worksheets.length; i++) {
      var ws = wb.worksheets[i];
      if (String(cellText(ws.getRow(1).getCell(1))).trim().toLowerCase() === 'код') return ws;
    }
    return wb.worksheets[0];
  }

  // Столбец «Код» в строке заголовков, по умолчанию — первый.
  function findCodeColumn(ws) {
    var header = ws.getRow(1);
    for (var c = 1; c <= header.cellCount; c++) {
      if (String(cellText(header.getCell(c))).trim().toLowerCase() === 'код') return c;
    }
    return 1;
  }

  // code -> список номеров строк с этим кодом
  function collectExcelCodes(ws, col) {
    var map = new Map();
    for (var r = 2; r <= ws.rowCount; r++) {
      var t = cellText(ws.getRow(r).getCell(col)).trim();
      if (!t) continue;
      if (!map.has(t)) map.set(t, []);
      map.get(t).push(r);
    }
    return map;
  }

  function loadWorkbook(buffer) {
    var wb = new ExcelJS.Workbook();
    return wb.xlsx.load(buffer).then(function () { return wb; });
  }

  /* ============================ Состояние ============================ */

  var state = {
    excel: null, // { name, buffer, sheetName, dataRows, codeCol }
    excelCodes: null, // Map code -> [rows] из текущего Excel
    pdfs: [], // { id, name, numPages, entries, failedPages, min, max }
    nextId: 1,
  };

  /* ============================ DOM ============================ */

  function $(id) { return document.getElementById(id); }

  function setupDropzone(dropId, inputId, handler) {
    var drop = $(dropId);
    var input = $(inputId);
    drop.addEventListener('click', function () { input.click(); });
    drop.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
    });
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('dragover'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('dragover'); });
    drop.addEventListener('drop', function (e) {
      e.preventDefault();
      drop.classList.remove('dragover');
      if (e.dataTransfer.files.length) handler(Array.prototype.slice.call(e.dataTransfer.files));
    });
    input.addEventListener('change', function () {
      if (input.files.length) handler(Array.prototype.slice.call(input.files));
      input.value = '';
    });  }

  function showError(msg) {
    var el = $('error');
    el.textContent = msg;
    el.hidden = !msg;
  }

  function setExcelFile(files) {
    var file = files && files[0];
    if (!file) return;
    if (!/\.xlsx$/i.test(file.name)) {
      showError('Файл «' + file.name + '» не является .xlsx');
      return;
    }
    showError('');
    file.arrayBuffer().then(function (buf) {
      return loadWorkbook(buf).then(function (wb) {
        var ws = pickWorksheet(wb);
        var col = findCodeColumn(ws);
        state.excel = {
          name: file.name,
          buffer: buf,
          sheetName: ws.name,
          dataRows: Math.max(ws.rowCount - 1, 0),
          codeCol: col,
        };
        state.excelCodes = collectExcelCodes(ws, col);
        var info = $('excel-info');
        info.hidden = false;
        info.textContent = '✓ ' + file.name + ' — лист «' + ws.name + '», строк с кодами: ' +
          state.excel.dataRows + ', столбец «Код»: ' + colLetter(col);
        refreshAll();
      });
    }).catch(function (e) {
      state.excel = null;
      state.excelCodes = null;
      $('excel-info').hidden = true;
      showError('Не удалось прочитать Excel: ' + (e && e.message ? e.message : e));
      refreshAll();
    });
  }

  function colLetter(n) {
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function addPdfFiles(files) {
    files.forEach(function (file) {
      if (!/\.pdf$/i.test(file.name)) {
        showError('Файл «' + file.name + '» не является PDF');
        return;
      }
      showError('');
      var item = { id: state.nextId++, name: file.name, loading: true };
      state.pdfs.push(item);
      renderPdfList();

      file.arrayBuffer().then(function (buf) {
        return parsePdf(buf).then(function (res) {
          item.loading = false;
          item.numPages = res.numPages;
          item.entries = res.entries;
          item.failedPages = res.failedPages;
          item.min = res.entries.length ? res.entries[0].num : null;
          item.max = res.entries.length ? res.entries[res.entries.length - 1].num : null;
          item.buf = buf;
        });
      }).catch(function (e) {
        item.loading = false;
        item.error = 'Ошибка чтения PDF: ' + (e && e.message ? e.message : e);
      }).then(function () {
        renderPdfList();
        refreshAll();
      });
    });
  }

  function removePdf(id) {
    state.pdfs = state.pdfs.filter(function (p) { return p.id !== id; });
    renderPdfList();
    refreshAll();
  }

  // Коды выбранного диапазона для одного PDF
  function selectedCodes(pdf) {
    var input = document.querySelector('input[data-range-for="' + pdf.id + '"]');
    var range = parseRange(input ? input.value : '', pdf.min, pdf.max);
    if (!range) return { range: null, codes: [] };
    var codes = pdf.entries
      .filter(function (e) { return e.num >= range.from && e.num <= range.to; })
      .map(function (e) { return e.code; });
    return { range: range, codes: codes };
  }

  function renderPdfList() {
    var list = $('pdf-list');
    list.innerHTML = '';
    state.pdfs.forEach(function (pdf) {
      var div = document.createElement('div');
      div.className = 'pdf-item';

      var head = document.createElement('div');
      head.className = 'pdf-head';
      var name = document.createElement('span');
      name.className = 'pdf-name';
      name.textContent = pdf.name;
      var del = document.createElement('button');
      del.className = 'pdf-remove';
      del.title = 'Убрать файл';
      del.textContent = '✕';
      del.addEventListener('click', function () { removePdf(pdf.id); });
      head.appendChild(name);
      head.appendChild(del);
      div.appendChild(head);

      var meta = document.createElement('div');
      meta.className = 'pdf-meta';
      if (pdf.loading) meta.textContent = 'Чтение PDF…';
      else if (pdf.error) meta.innerHTML = '<span class="bad">✕ ' + escapeHtml(pdf.error) + '</span>';
      else {
        meta.textContent = 'Страниц: ' + pdf.numPages + ' · найдено кодов: ' + pdf.entries.length +
          ' · номера №' + pdf.min + '–№' + pdf.max;
        if (pdf.failedPages.length) {
          meta.innerHTML += ' · <span class="bad">без кода: стр. ' + pdf.failedPages.join(', ') + '</span>';
        }
      }
      div.appendChild(meta);

      if (!pdf.loading && !pdf.error && pdf.entries.length) {
        var row = document.createElement('div');
        row.className = 'range-row';
        var label = document.createElement('label');
        label.htmlFor = 'range-' + pdf.id;
        label.textContent = 'Удалить коды с номерами:';
        var input = document.createElement('input');
        input.type = 'text';
        input.id = 'range-' + pdf.id;
        input.placeholder = 'все (примеры: 485 или 456-500)';
        input.setAttribute('data-range-for', pdf.id);
        input.addEventListener('input', refreshAll);
        row.appendChild(label);
        row.appendChild(input);
        div.appendChild(row);

        var count = document.createElement('div');
        count.className = 'pdf-count';
        count.setAttribute('data-count-for', pdf.id);
        div.appendChild(count);
      }
      list.appendChild(div);
    });
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }

  function refreshAll() {
    var totalSelected = 0, totalFound = 0, totalMissing = 0;

    state.pdfs.forEach(function (pdf) {
      var countEl = document.querySelector('[data-count-for="' + pdf.id + '"]');
      if (!countEl || pdf.loading || pdf.error || !pdf.entries.length) return;
      var sel = selectedCodes(pdf);
      if (!sel.range) {
        countEl.className = 'pdf-count err';
        countEl.textContent = 'Непонятный диапазон. Примеры: 485 · 485-500 · 485-';
        return;
      }
      var found = sel.codes.filter(function (c) { return state.excelCodes && state.excelCodes.has(c); });
      var missing = sel.codes.length - found.length;
      countEl.className = 'pdf-count ' + (missing ? 'warn' : 'ok');
      countEl.textContent = 'Будет удалено кодов: ' + found.length +
        (sel.range.from !== sel.range.to
          ? ' (№' + sel.range.from + '–№' + sel.range.to + ')'
          : ' (№' + sel.range.from + ')') +
        (missing ? ' · не найдено в Excel: ' + missing : '');
      totalSelected += sel.codes.length;
      totalFound += found.length;
      totalMissing += missing;
    });

    $('sum-selected').textContent = totalSelected;
    $('sum-found').textContent = totalFound;
    $('sum-missing').textContent = totalMissing;

    var ready = state.excel && state.pdfs.some(function (p) { return !p.loading && p.entries && p.entries.length; });
    $('process-btn').disabled = !ready;
    $('pdf-card').hidden = false;
  }

  /* ============================ Обработка ============================ */

  function processAndDownload() {
    var btn = $('process-btn');
    btn.classList.add('busy');
    btn.textContent = 'Обработка…';
    showError('');

    Promise.resolve().then(function () {
      var deleteCodes = new Set();
      state.pdfs.forEach(function (pdf) {
        if (pdf.loading || pdf.error || !pdf.entries.length) return;
        selectedCodes(pdf).codes.forEach(function (c) { deleteCodes.add(c); });
      });

      return loadWorkbook(state.excel.buffer).then(function (wb) {
        var ws = pickWorksheet(wb);
        var col = findCodeColumn(ws);
        var codeMap = collectExcelCodes(ws, col);

        var rowsToDelete = [];
        var notFound = [];
        deleteCodes.forEach(function (c) {
          if (codeMap.has(c)) codeMap.get(c).forEach(function (r) { rowsToDelete.push(r); });
          else notFound.push(c);
        });
        rowsToDelete.sort(function (a, b) { return b - a; });
        rowsToDelete.forEach(function (r) { ws.spliceRows(r, 1); });

        return wb.xlsx.writeBuffer().then(function (out) {
          var stem = state.excel.name.replace(/\.xlsx$/i, '');
          var outName = stem + '_обновленный.xlsx';
          download(out, outName);

          var msg = 'Готово! Удалено строк: ' + rowsToDelete.length +
            ' из ' + (state.excel.dataRows) + '. Осталось строк: ' + (state.excel.dataRows - rowsToDelete.length) + '.';
          if (notFound.length) {
            msg += '\nВнимание: ' + notFound.length + ' код(ов) из PDF не найдены в Excel и не удалены.';
          }
          var res = $('result');
          res.textContent = msg;
          res.className = 'result' + (notFound.length ? ' warn' : '');
          res.hidden = false;
        });
      });
    }).catch(function (e) {
      showError('Ошибка при обработке: ' + (e && e.message ? e.message : e));
    }).then(function () {
      btn.classList.remove('busy');
      btn.textContent = 'Удалить строки и скачать Excel';
    });
  }

  function download(buffer, filename) {
    var blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 5000);
  }

  /* ============================ Инициализация ============================ */

  function init() {
    if (window.pdfjsLib) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
    }
    setupDropzone('excel-drop', 'excel-input', setExcelFile);
    setupDropzone('pdf-drop', 'pdf-input', addPdfFiles);
    $('process-btn').addEventListener('click', processAndDownload);
  }

  if (typeof document !== 'undefined') init();

  /* Экспорт чистых функций для Node-тестов */
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
      groupIntoLines: groupIntoLines,
      extractCodeAndNumber: extractCodeAndNumber,
      parsePdf: parsePdf,
      parseRange: parseRange,
      pickWorksheet: pickWorksheet,
      findCodeColumn: findCodeColumn,
      collectExcelCodes: collectExcelCodes,
      loadWorkbook: loadWorkbook,
      cellText: cellText,
    };
  }
})();
