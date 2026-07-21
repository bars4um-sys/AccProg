/* ============================================================
 * Функция: Авансовый отчёт vs Банк
 * Шаг 1 — проверка дублей в Авансовом отчёте (Дата + Сумма).
 * Шаг 2 — строки Банка без соответствия в АО (по Сумме).
 * Шаг 3 — строки АО без соответствия в Банке (по Сумме).  ← НОВЫЙ
 * Все результаты можно скачать в Excel.
 * ============================================================ */
AccProg.register({
  id: 'advance-vs-bank',
  title: 'Авансовый отчёт vs Банк',
  icon: '🏦',
  description: '<span class="step-tag">Авансовый отчёт</span> проверяется на дубли, ' +
    'затем сравнивается с <span class="step-tag">Банком</span> по столбцу «Сумма» в обоих направлениях.',

  /* --- Панель 1: ввод --- */
  renderInput: function (ctx) {
    ctx.inputEl.innerHTML =
      '<div class="upload-grid">' +
      '  <div class="upload-box">' +
      '    <h3>Авансовый отчёт</h3>' +
      '    <input type="file" id="avb-file1" accept=".xlsx,.xls,.xlsm,.xlsb">' +
      '    <div class="file-info" id="avb-info1">Файл не выбран</div>' +
      '  </div>' +
      '  <div class="upload-box">' +
      '    <h3>Банковская выписка</h3>' +
      '    <input type="file" id="avb-file2" accept=".xlsx,.xls,.xlsm,.xlsb">' +
      '    <div class="file-info" id="avb-info2">Файл не выбран</div>' +
      '  </div>' +
      '</div>';

    var u = ctx.utils;

    ctx.inputEl.querySelector('#avb-file1').addEventListener('change', function (e) {
      var file = e.target.files[0]; if (!file) return;
      var info = ctx.inputEl.querySelector('#avb-info1');
      info.textContent = 'Чтение файла…'; info.className = 'file-info';
      u.readSheetMatrix(file).then(function (r) {
        ctx.state.aoData = r.rows;
        var h = u.detectHeaderRow(r.rows);
        info.textContent = file.name + ' | лист: ' + r.sheetName +
          ' | строк данных: ' + (r.rows.length - 1 - h) +
          ' | «Дата»: 2-й столбец | «Сумма»: 5-й столбец';
        info.className = 'file-info ok';
        resetWorkflow(ctx);
        updateButtons(ctx);
      }).catch(function (err) {
        ctx.state.aoData = null;
        info.textContent = 'Ошибка чтения: ' + err.message; info.className = 'file-info error';
        updateButtons(ctx);
      });
    });

    ctx.inputEl.querySelector('#avb-file2').addEventListener('change', function (e) {
      var file = e.target.files[0]; if (!file) return;
      var info = ctx.inputEl.querySelector('#avb-info2');
      info.textContent = 'Чтение файла…'; info.className = 'file-info';
      u.readSheetMatrix(file).then(function (r) {
        ctx.state.bankData = r.rows;
        var h = u.detectHeaderRow(r.rows);
        info.textContent = file.name + ' | лист: ' + r.sheetName +
          ' | строк данных: ' + (r.rows.length - 1 - h) + ' | «Сумма»: 2-й столбец';
        info.className = 'file-info ok';
        resetWorkflow(ctx);
        updateButtons(ctx);
      }).catch(function (err) {
        ctx.state.bankData = null;
        info.textContent = 'Ошибка чтения: ' + err.message; info.className = 'file-info error';
        updateButtons(ctx);
      });
    });

    function updateButtons(ctx) {
      var both = !!(ctx.state.aoData && ctx.state.bankData);
      ctx.enableAction('check-dups', both);
    }
    function resetWorkflow(ctx) {
      ctx.state.dupsState      = null;
      ctx.state.resultRows     = null;   // строки Банка без АО
      ctx.state.resultAORows   = null;   // строки АО без Банка
      ctx.enableAction('compare',          false);
      ctx.enableAction('compare-ao',       false);
      ctx.enableAction('download-dups',    false);
      ctx.enableAction('download-result',  false);
      ctx.enableAction('download-result-ao', false);
      ctx.clearOutput();
    }

    // сразу выставим состояние кнопок
    ctx.enableAction('check-dups', !!(ctx.state.aoData && ctx.state.bankData));
  },

  /* --- Панель 2: операции --- */
  actions: [
    {
      id: 'check-dups',
      label: 'Шаг 1: Дубли в Авансовом отчёте',
      disabled: true,
      run: function (ctx) { runCheckDups(ctx); }
    },
    {
      id: 'compare',
      label: 'Шаг 2: Банк → не найдено в АО',
      variant: 'secondary', disabled: true,
      run: function (ctx) { runCompareBank(ctx); }
    },
    {
      id: 'compare-ao',
      label: 'Шаг 3: АО → не найдено в Банке',
      variant: 'secondary', disabled: true,
      run: function (ctx) { runCompareAO(ctx); }
    },
    {
      id: 'download-dups',
      label: 'Скачать дубли (Excel)',
      variant: 'secondary', disabled: true,
      run: function (ctx) { downloadDups(ctx); }
    },
    {
      id: 'download-result',
      label: 'Скачать Банк без АО (Excel)',
      variant: 'secondary', disabled: true,
      run: function (ctx) { downloadResultBank(ctx); }
    },
    {
      id: 'download-result-ao',
      label: 'Скачать АО без Банка (Excel)',
      variant: 'secondary', disabled: true,
      run: function (ctx) { downloadResultAO(ctx); }
    }
  ]
});

/* ---------- Константы столбцов ---------- */
var AVB = { DATE_COLUMN_AO: 1, SUM_COLUMN_AO: 4, SUM_COLUMN_BANK: 1 };

/* ======================================================
 * ШАГ 1 — Дубли в Авансовом отчёте (Дата + Сумма)
 * ====================================================== */
function runCheckDups(ctx) {
  var u = ctx.utils;
  if (!ctx.state.aoData) {
    ctx.setOutput(ctx.message('Сначала загрузите Авансовый отчёт.', 'error')); return;
  }

  var headerRowIdx = u.detectHeaderRow(ctx.state.aoData);
  var dataStart    = headerRowIdx + 1;
  var headers      = (ctx.state.aoData[headerRowIdx] || []).slice();
  var colsCount    = headers.length;

  var groups = new Map();
  for (var i = dataStart; i < ctx.state.aoData.length; i++) {
    var row     = ctx.state.aoData[i];
    var dateKey = u.normalizeDate(row[AVB.DATE_COLUMN_AO]);
    var sumKey  = u.normalizeSum(row[AVB.SUM_COLUMN_AO]);
    if (dateKey === null || sumKey === null) continue;
    var key = dateKey + '|' + sumKey;
    if (!groups.has(key)) groups.set(key, { date: dateKey, sum: sumKey, rows: [] });
    groups.get(key).rows.push({ idx: i, row: row });
  }

  var dupGroups = [];
  groups.forEach(function (g) {
    if (g.rows.length > 1) {
      g.rows.sort(function (a, b) { return a.idx - b.idx; });
      dupGroups.push(g);
    }
  });

  ctx.state.dupsState = { headers: headers, colsCount: colsCount, groups: dupGroups };
  renderDups(ctx);
}

function renderDups(ctx) {
  var u         = ctx.utils;
  var st        = ctx.state.dupsState;
  var headerIdx = u.detectHeaderRow(ctx.state.aoData);
  var totalRows = ctx.state.aoData.length - 1 - headerIdx;
  var dupCount  = 0;
  st.groups.forEach(function (g) { dupCount += g.rows.length; });

  var html = '<h3 style="margin-bottom:10px">Шаг 1. Повторяющиеся записи в Авансовом отчёте</h3>';

  if (st.groups.length === 0) {
    html += '<div class="summary">Повторов не найдено. <b class="ok">Можно переходить к сравнению.</b></div>';
    html += '<div class="empty-msg">Дублей по паре «Дата + Сумма» нет.</div>';
    ctx.setOutput(html);
    ctx.enableAction('download-dups', false);
    ctx.enableAction('compare',    !!ctx.state.bankData);
    ctx.enableAction('compare-ao', !!ctx.state.bankData);
    return;
  }

  html += '<div class="summary">Групп дублей: <b class="warn">' + st.groups.length +
    '</b> · строк-дублей: <b class="warn">' + dupCount +
    '</b> · из общего числа: ' + totalRows +
    '.<br>Поиск по совпадению <b>Даты</b> (2-й столбец) и <b>Суммы</b> (5-й столбец).</div>';

  html += '<div class="table-wrapper"><table><thead><tr>';
  st.headers.forEach(function (h, idx) {
    var marker = idx === AVB.DATE_COLUMN_AO ? ' <span class="col-marker">★</span>' :
                 idx === AVB.SUM_COLUMN_AO  ? ' <span class="col-marker-blue">✓</span>' : '';
    html += '<th>' + u.escapeHtml(String(h == null ? '' : h)) + marker + '</th>';
  });
  html += '</tr></thead><tbody>';
  st.groups.forEach(function (g) {
    html += '<tr class="group-separator"><td colspan="' + st.colsCount + '">Группа: дата «' +
      u.escapeHtml(g.date) + '» + сумма «' + u.escapeHtml(String(g.sum)) + '» — ' +
      g.rows.length + ' строк</td></tr>';
    g.rows.forEach(function (entry) {
      html += '<tr class="dup-row">';
      for (var c = 0; c < st.colsCount; c++) {
        var v = entry.row[c];
        html += '<td>' + u.escapeHtml(v === null || v === undefined ? '' : String(v)) + '</td>';
      }
      html += '</tr>';
    });
  });
  html += '</tbody></table></div>';

  ctx.setOutput(html);
  ctx.enableAction('download-dups', true);
  ctx.enableAction('compare',    !!ctx.state.bankData);
  ctx.enableAction('compare-ao', !!ctx.state.bankData);
}

/* ======================================================
 * ШАГ 2 — Строки Банка без соответствия в АО
 * ====================================================== */
function runCompareBank(ctx) {
  var u = ctx.utils;
  if (!ctx.state.aoData)   { ctx.setOutput(ctx.message('Загрузите Авансовый отчёт.', 'error')); return; }
  if (!ctx.state.bankData) { ctx.setOutput(ctx.message('Загрузите Банковскую выписку.', 'error')); return; }

  // Собираем множество сумм из АО
  var sumsAO = new Set();
  for (var i = 1; i < ctx.state.aoData.length; i++) {
    var s = u.normalizeSum(ctx.state.aoData[i][AVB.SUM_COLUMN_AO]);
    if (s !== null) sumsAO.add(s);
  }

  var headerIdx    = u.detectHeaderRow(ctx.state.bankData);
  var resultHeaders = (ctx.state.bankData[headerIdx] || []).slice();
  var colsCount    = resultHeaders.length;
  var resultRows   = [];

  for (var j = headerIdx + 1; j < ctx.state.bankData.length; j++) {
    var row = ctx.state.bankData[j];
    var sb  = u.normalizeSum(row[AVB.SUM_COLUMN_BANK]);
    if (sb === null || !sumsAO.has(sb)) {
      var padded = new Array(colsCount);
      for (var c = 0; c < colsCount; c++) padded[c] = row[c] !== undefined ? row[c] : '';
      resultRows.push(padded);
    }
  }

  ctx.state.resultHeaders = resultHeaders;
  ctx.state.resultRows    = resultRows;
  renderCompareBank(ctx);
}

function renderCompareBank(ctx) {
  var u         = ctx.utils;
  var headerIdx = u.detectHeaderRow(ctx.state.bankData);
  var totalRows = ctx.state.bankData.length - 1 - headerIdx;
  var rows      = ctx.state.resultRows;
  var matched   = totalRows - rows.length;

  var html = '<h3 style="margin-bottom:10px">Шаг 2. Строки Банка без соответствия в Авансовом отчёте</h3>';
  html += '<div class="summary">Строк в Банке: <b>' + totalRows +
    '</b> · совпало с АО: <b class="ok">' + matched +
    '</b> · без соответствия: <b class="warn">' + rows.length + '</b></div>';

  if (rows.length === 0) {
    html += '<div class="empty-msg">Все строки Банка имеют соответствие в Авансовом отчёте.</div>';
    ctx.setOutput(html);
    ctx.enableAction('download-result', false);
    return;
  }

  html += u.renderTable(ctx.state.resultHeaders, rows, {
    markCols: (function () { var m = {}; m[AVB.SUM_COLUMN_BANK] = { mark: '★', cls: 'col-marker' }; return m; })()
  });
  ctx.setOutput(html);
  ctx.enableAction('download-result', true);
}

/* ======================================================
 * ШАГ 3 — Строки АО без соответствия в Банке  ← НОВЫЙ
 * ====================================================== */
function runCompareAO(ctx) {
  var u = ctx.utils;
  if (!ctx.state.aoData)   { ctx.setOutput(ctx.message('Загрузите Авансовый отчёт.', 'error')); return; }
  if (!ctx.state.bankData) { ctx.setOutput(ctx.message('Загрузите Банковскую выписку.', 'error')); return; }

  // Собираем множество сумм из Банка
  var sumsBank = new Set();
  var bankHeaderIdx = u.detectHeaderRow(ctx.state.bankData);
  for (var i = bankHeaderIdx + 1; i < ctx.state.bankData.length; i++) {
    var s = u.normalizeSum(ctx.state.bankData[i][AVB.SUM_COLUMN_BANK]);
    if (s !== null) sumsBank.add(s);
  }

  // Ищем строки АО, суммы которых нет в Банке
  var aoHeaderIdx    = u.detectHeaderRow(ctx.state.aoData);
  var resultAOHeaders = (ctx.state.aoData[aoHeaderIdx] || []).slice();
  var colsCount      = resultAOHeaders.length;
  var resultAORows   = [];

  for (var j = aoHeaderIdx + 1; j < ctx.state.aoData.length; j++) {
    var row = ctx.state.aoData[j];
    var sa  = u.normalizeSum(row[AVB.SUM_COLUMN_AO]);
    if (sa === null || !sumsBank.has(sa)) {
      var padded = new Array(colsCount);
      for (var c = 0; c < colsCount; c++) padded[c] = row[c] !== undefined ? row[c] : '';
      resultAORows.push(padded);
    }
  }

  ctx.state.resultAOHeaders = resultAOHeaders;
  ctx.state.resultAORows    = resultAORows;
  renderCompareAO(ctx);
}

function renderCompareAO(ctx) {
  var u         = ctx.utils;
  var aoHeaderIdx = u.detectHeaderRow(ctx.state.aoData);
  var totalRows = ctx.state.aoData.length - 1 - aoHeaderIdx;
  var rows      = ctx.state.resultAORows;
  var matched   = totalRows - rows.length;

  var html = '<h3 style="margin-bottom:10px">Шаг 3. Строки Авансового отчёта без соответствия в Банке</h3>';
  html += '<div class="summary">Строк в Авансовом отчёте: <b>' + totalRows +
    '</b> · совпало с Банком: <b class="ok">' + matched +
    '</b> · без соответствия: <b class="warn">' + rows.length + '</b></div>';

  if (rows.length === 0) {
    html += '<div class="empty-msg">Все строки Авансового отчёта имеют соответствие в Банковской выписке.</div>';
    ctx.setOutput(html);
    ctx.enableAction('download-result-ao', false);
    return;
  }

  html += u.renderTable(ctx.state.resultAOHeaders, rows, {
    markCols: (function () { var m = {}; m[AVB.SUM_COLUMN_AO] = { mark: '★', cls: 'col-marker' }; return m; })()
  });
  ctx.setOutput(html);
  ctx.enableAction('download-result-ao', true);
}

/* ======================================================
 * Скачивание результатов
 * ====================================================== */
function downloadDups(ctx) {
  var st = ctx.state.dupsState;
  if (!st || st.groups.length === 0) return;
  var data = [st.headers];
  st.groups.forEach(function (g) {
    g.rows.forEach(function (entry) {
      var padded = new Array(st.headers.length);
      for (var c = 0; c < st.headers.length; c++)
        padded[c] = entry.row[c] !== undefined ? entry.row[c] : '';
      data.push(padded);
    });
  });
  ctx.utils.exportXlsx(data, 'Дубли АО', 'dubles');
}

function downloadResultBank(ctx) {
  if (!ctx.state.resultRows || !ctx.state.resultRows.length) return;
  var data = [ctx.state.resultHeaders].concat(ctx.state.resultRows);
  ctx.utils.exportXlsx(data, 'Банк без АО', 'bank_bez_ao');
}

function downloadResultAO(ctx) {
  if (!ctx.state.resultAORows || !ctx.state.resultAORows.length) return;
  var data = [ctx.state.resultAOHeaders].concat(ctx.state.resultAORows);
  ctx.utils.exportXlsx(data, 'АО без Банка', 'ao_bez_bank');
}
