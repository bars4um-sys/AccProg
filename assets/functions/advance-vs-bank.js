/* ============================================================
 * Функция: Авансовый отчёт vs Банк (портирована из SearchFall)
 * Шаг 1 — проверка дублей в Авансовом отчёте (Дата + Сумма).
 * Шаг 2 — сравнение с Банком по столбцу «Сумма»,
 *         поиск строк Банка без соответствия.
 * Оба результата можно скачать в Excel.
 * ============================================================ */
AccProg.register({
  id: 'advance-vs-bank',
  title: 'Авансовый отчёт vs Банк',
  icon: '🏦',
  description: '<span class="step-tag">Авансовый отчёт</span> проверяется на дубли, ' +
    'затем сравнивается с <span class="step-tag">Банком</span> по столбцу «Сумма».',

  // Фиксированные индексы столбцов (по визуальному положению)
  DATE_COLUMN_AO: 1,
  SUM_COLUMN_AO: 4,
  SUM_COLUMN_BANK: 1,

  /* --- Панель 1: ввод --- */
  renderInput: function (ctx) {
    var self = this;
    ctx.inputEl.innerHTML =
      '<div class="upload-grid">' +
      '  <div class="upload-box">' +
      '    <h3>Авансовый отчёт</h3>' +
      '    <input type="file" id="avb-file1" accept=".xlsx,.xls,.xlsm,.xlsb">' +
      '    <div class="file-info" id="avb-info1">Файл не выбран</div>' +
      '  </div>' +
      '  <div class="upload-box">' +
      '    <h3>Банк</h3>' +
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
      ctx.enableAction('check-dups', !!(ctx.state.aoData && ctx.state.bankData));
    }
    function resetWorkflow(ctx) {
      ctx.state.dupsState = null;
      ctx.state.resultRows = null;
      ctx.enableAction('compare', false);
      ctx.enableAction('download-dups', false);
      ctx.enableAction('download-result', false);
      ctx.clearOutput();
    }

    // сразу выставим состояние кнопок
    ctx.enableAction('check-dups', !!(ctx.state.aoData && ctx.state.bankData));
  },

  /* --- Панель 2: операции --- */
  actions: [
    {
      id: 'check-dups', label: 'Проверить дубли в Авансовом отчёте', disabled: true,
      run: function (ctx) { runCheckDups(ctx); }
    },
    {
      id: 'compare', label: 'Сравнить с Банком', variant: 'secondary', disabled: true,
      run: function (ctx) { runCompare(ctx); }
    },
    {
      id: 'download-dups', label: 'Скачать дубли (Excel)', variant: 'secondary', disabled: true,
      run: function (ctx) { downloadDups(ctx); }
    },
    {
      id: 'download-result', label: 'Скачать результат (Excel)', variant: 'secondary', disabled: true,
      run: function (ctx) { downloadResult(ctx); }
    }
  ]
});

/* ---------- Логика (как в SearchFall) ---------- */
var AVB = { DATE_COLUMN_AO: 1, SUM_COLUMN_AO: 4, SUM_COLUMN_BANK: 1 };

function runCheckDups(ctx) {
  var u = ctx.utils;
  if (!ctx.state.aoData) { ctx.setOutput(ctx.message('Сначала загрузите Авансовый отчёт.', 'error')); return; }

  var headerRowIdx = u.detectHeaderRow(ctx.state.aoData);
  var dataStart = headerRowIdx + 1;
  var headers = (ctx.state.aoData[headerRowIdx] || []).slice();
  var colsCount = headers.length;

  var groups = new Map();
  for (var i = dataStart; i < ctx.state.aoData.length; i++) {
    var row = ctx.state.aoData[i];
    var dateKey = u.normalizeDate(row[AVB.DATE_COLUMN_AO]);
    var sumKey = u.normalizeSum(row[AVB.SUM_COLUMN_AO]);
    if (dateKey === null || sumKey === null) continue;
    var key = dateKey + '|' + sumKey;
    if (!groups.has(key)) groups.set(key, { date: dateKey, sum: sumKey, rows: [] });
    groups.get(key).rows.push({ idx: i, row: row });
  }

  var dupGroups = [];
  groups.forEach(function (g) {
    if (g.rows.length > 1) { g.rows.sort(function (a, b) { return a.idx - b.idx; }); dupGroups.push(g); }
  });

  ctx.state.dupsState = { headers: headers, colsCount: colsCount, groups: dupGroups };
  renderDups(ctx);
}

function renderDups(ctx) {
  var u = ctx.utils;
  var st = ctx.state.dupsState;
  var headerRowIdx = u.detectHeaderRow(ctx.state.aoData);
  var totalRows = ctx.state.aoData.length - 1 - headerRowIdx;
  var groupsCount = st.groups.length;
  var dupCount = 0; st.groups.forEach(function (g) { dupCount += g.rows.length; });

  var html = '<h3 style="margin-bottom:10px">Шаг 1. Повторяющиеся записи в Авансовом отчёте</h3>';

  if (groupsCount === 0) {
    html += '<div class="summary">Повторов не найдено. <b class="ok">Можно переходить к сравнению с Банком.</b></div>';
    html += '<div class="empty-msg">Дублей по паре «Дата + Сумма» нет.</div>';
    ctx.setOutput(html);
    ctx.enableAction('download-dups', false);
    ctx.enableAction('compare', !!ctx.state.bankData);
    return;
  }

  html += '<div class="summary">Групп дублей: <b class="warn">' + groupsCount +
    '</b> · всего строк-дублей: <b class="warn">' + dupCount +
    '</b> · из общего числа: ' + totalRows +
    '.<br>Поиск по совпадению <b>Даты</b> (2-й столбец) и <b>Суммы</b> (5-й столбец).</div>';

  html += '<div class="table-wrapper"><table><thead><tr>';
  st.headers.forEach(function (h, idx) {
    var marker = idx === AVB.DATE_COLUMN_AO ? ' <span class="col-marker">★</span>' :
      (idx === AVB.SUM_COLUMN_AO ? ' <span class="col-marker-blue">✓</span>' : '');
    html += '<th>' + u.escapeHtml(String(h == null ? '' : h)) + marker + '</th>';
  });
  html += '</tr></thead><tbody>';
  st.groups.forEach(function (g) {
    html += '<tr class="group-separator"><td colspan="' + st.colsCount + '">Группа: дата «' +
      u.escapeHtml(g.date) + '» + сумма «' + u.escapeHtml(String(g.sum)) + '» — ' + g.rows.length + ' строк</td></tr>';
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
  ctx.enableAction('compare', !!ctx.state.bankData);
}

function runCompare(ctx) {
  var u = ctx.utils;
  if (!ctx.state.aoData) { ctx.setOutput(ctx.message('Сначала загрузите Авансовый отчёт.', 'error')); return; }
  if (!ctx.state.bankData) { ctx.setOutput(ctx.message('Сначала загрузите Банк.', 'error')); return; }

  var sumsAO = new Set();
  for (var i = 1; i < ctx.state.aoData.length; i++) {
    var s = u.normalizeSum(ctx.state.aoData[i][AVB.SUM_COLUMN_AO]);
    if (s !== null) sumsAO.add(s);
  }

  var headerRowIdxBank = u.detectHeaderRow(ctx.state.bankData);
  var resultHeaders = (ctx.state.bankData[headerRowIdxBank] || []).slice();
  var colsCount = resultHeaders.length;
  var resultRows = [];

  for (var j = headerRowIdxBank + 1; j < ctx.state.bankData.length; j++) {
    var row = ctx.state.bankData[j];
    var sb = u.normalizeSum(row[AVB.SUM_COLUMN_BANK]);
    if (sb === null || !sumsAO.has(sb)) {
      var padded = new Array(colsCount);
      for (var c = 0; c < colsCount; c++) padded[c] = row[c] !== undefined ? row[c] : '';
      resultRows.push(padded);
    }
  }

  ctx.state.resultHeaders = resultHeaders;
  ctx.state.resultRows = resultRows;
  renderCompare(ctx);
}

function renderCompare(ctx) {
  var u = ctx.utils;
  var headerRowIdxBank = u.detectHeaderRow(ctx.state.bankData);
  var totalRows = ctx.state.bankData.length - 1 - headerRowIdxBank;
  var resultRows = ctx.state.resultRows;
  var matched = totalRows - resultRows.length;

  var html = '<h3 style="margin-bottom:10px">Шаг 2. Строки из Банка без соответствия в Авансовом отчёте</h3>';
  html += '<div class="summary">Строк в Банке: <b>' + totalRows +
    '</b> · совпало с АО: <b class="ok">' + matched +
    '</b> · без соответствия: <b class="warn">' + resultRows.length + '</b></div>';

  if (resultRows.length === 0) {
    html += '<div class="empty-msg">Все строки Банка имеют соответствие в Авансовом отчёте по столбцу «Сумма».</div>';
    ctx.setOutput(html);
    ctx.enableAction('download-result', false);
    return;
  }

  html += u.renderTable(ctx.state.resultHeaders, resultRows, {
    markCols: (function () { var m = {}; m[AVB.SUM_COLUMN_BANK] = { mark: '★', cls: 'col-marker' }; return m; })()
  });
  ctx.setOutput(html);
  ctx.enableAction('download-result', true);
}

function downloadDups(ctx) {
  var st = ctx.state.dupsState;
  if (!st || st.groups.length === 0) return;
  var data = [st.headers];
  st.groups.forEach(function (g) {
    g.rows.forEach(function (entry) {
      var padded = new Array(st.headers.length);
      for (var c = 0; c < st.headers.length; c++) padded[c] = entry.row[c] !== undefined ? entry.row[c] : '';
      data.push(padded);
    });
  });
  ctx.utils.exportXlsx(data, 'Дубли АО', 'dubles');
}

function downloadResult(ctx) {
  if (!ctx.state.resultRows || !ctx.state.resultRows.length) return;
  var data = [ctx.state.resultHeaders].concat(ctx.state.resultRows);
  ctx.utils.exportXlsx(data, 'Результат', 'result');
}
