/* ============================================================
 * Функция: Сравнение таблиц (портирована из SearchDiff)
 * Загружает 2 таблицы, агрегирует суммы по компаниям и
 * выводит расхождения. Добавлен экспорт результата в Excel.
 * ============================================================ */
AccProg.register({
  id: 'compare-tables',
  title: 'Сравнение таблиц',
  icon: '📊',
  description: 'Суммирует значения по компаниям из двух таблиц и показывает расхождения.',

  /* --- Панель 1: ввод данных --- */
  renderInput: function (ctx) {
    ctx.inputEl.innerHTML =
      '<div class="upload-grid">' +
      '  <div class="upload-box">' +
      '    <h3>Таблица 1</h3>' +
      '    <input type="file" id="ct-file1" accept=".xlsx,.xls,.csv">' +
      '    <div class="file-info" id="ct-info1">Файл не выбран</div>' +
      '  </div>' +
      '  <div class="upload-box">' +
      '    <h3>Таблица 2</h3>' +
      '    <input type="file" id="ct-file2" accept=".xlsx,.xls,.csv">' +
      '    <div class="file-info" id="ct-info2">Файл не выбран</div>' +
      '  </div>' +
      '</div>';

    var u = ctx.utils;
    ctx.inputEl.querySelector('#ct-file1').addEventListener('change', function (e) {
      handleUpload(ctx, e, 1);
    });
    ctx.inputEl.querySelector('#ct-file2').addEventListener('change', function (e) {
      handleUpload(ctx, e, 2);
    });

    function handleUpload(ctx, event, num) {
      var file = event.target.files[0];
      if (!file) return;
      var info = ctx.inputEl.querySelector('#ct-info' + num);
      info.textContent = 'Чтение файла…';
      info.className = 'file-info';
      u.readSheetObjects(file).then(function (json) {
        json = u.cleanData(json);
        if (!json.length) {
          info.textContent = 'Таблица пуста или не содержит данных';
          info.className = 'file-info error';
          return;
        }
        ctx.state['table' + num] = json;
        info.textContent = '✓ ' + file.name + ' · строк: ' + json.length;
        info.className = 'file-info ok';
        ctx.enableAction('compare', !!(ctx.state.table1 && ctx.state.table2));
      }).catch(function (err) {
        info.textContent = 'Ошибка: ' + err.message;
        info.className = 'file-info error';
      });
    }
  },

  /* --- Панель 2: операции --- */
  actions: [
    {
      id: 'compare', label: 'Сравнить таблицы', disabled: true,
      run: function (ctx) {
        if (!ctx.state.table1 || !ctx.state.table2) {
          ctx.setOutput(ctx.message('Загрузите обе таблицы', 'error'));
          return;
        }
        var agg1 = aggregate(ctx.state.table1);
        var agg2 = aggregate(ctx.state.table2);
        var diffs = compare(agg1, agg2);
        ctx.state.lastResult = diffs;
        renderResult(ctx, diffs);
        ctx.enableAction('export', diffs.length > 0);
      }
    },
    {
      id: 'export', label: 'Скачать результат', variant: 'secondary', disabled: true,
      run: function (ctx) {
        var diffs = ctx.state.lastResult || [];
        if (!diffs.length) return;
        var data = [['Компания', 'Таблица 1', 'Таблица 2', 'Разница']];
        diffs.forEach(function (d) {
          data.push([d.company, d.amount1, d.amount2, d.difference]);
        });
        ctx.utils.exportXlsx(data, 'Расхождения', 'sravnenie');
      }
    },
    {
      id: 'reset', label: 'Очистить', variant: 'danger',
      run: function (ctx) {
        ctx.state.table1 = null; ctx.state.table2 = null; ctx.state.lastResult = null;
        var f1 = ctx.inputEl.querySelector('#ct-file1');
        var f2 = ctx.inputEl.querySelector('#ct-file2');
        if (f1) f1.value = '';
        if (f2) f2.value = '';
        ctx.inputEl.querySelector('#ct-info1').textContent = 'Файл не выбран';
        ctx.inputEl.querySelector('#ct-info1').className = 'file-info';
        ctx.inputEl.querySelector('#ct-info2').textContent = 'Файл не выбран';
        ctx.inputEl.querySelector('#ct-info2').className = 'file-info';
        ctx.setOutput('<div class="placeholder">Здесь появится результат после выполнения операции.</div>');
        ctx.enableAction('compare', false);
        ctx.enableAction('export', false);
      }
    }
  ]
});

/* ---------- Вспомогательная логика (как в SearchDiff) ---------- */
function aggregate(tableData) {
  var agg = {};
  tableData.forEach(function (row) {
    var company = null, amount = null;
    for (var k in row) {
      if (typeof row[k] === 'string' && !company) company = row[k].trim();
    }
    var keys = Object.keys(row);
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i], v = row[key];
      if (typeof v === 'number' && amount === null) { amount = v; break; }
      if (key.toLowerCase().includes('сумма') || key.toLowerCase().includes('sum')) {
        amount = parseFloat(v) || 0; break;
      }
    }
    if (company && amount !== null) {
      agg[company] = (agg[company] || 0) + amount;
    }
  });
  return agg;
}

function compare(agg1, agg2) {
  var diffs = [];
  var all = new Set(Object.keys(agg1).concat(Object.keys(agg2)));
  all.forEach(function (company) {
    var a1 = agg1[company] || 0, a2 = agg2[company] || 0;
    if (a1 !== a2) {
      diffs.push({ company: company, amount1: a1, amount2: a2, difference: Math.abs(a1 - a2) });
    }
  });
  diffs.sort(function (a, b) { return a.company.localeCompare(b.company, 'ru'); });
  return diffs;
}

function renderResult(ctx, diffs) {
  var u = ctx.utils;
  if (!diffs.length) {
    ctx.setOutput('<div class="summary">Расхождений не найдено ✓</div>' +
      '<div class="no-diff">Обе таблицы содержат одинаковые суммы по всем компаниям</div>');
    return;
  }
  var html = '<div class="summary">Обнаружено <b class="warn">' + diffs.length +
    '</b> компани(й) с расхождениями:</div><ul class="diff-list">';
  diffs.forEach(function (d) {
    html += '<li class="diff-item">' +
      '<div class="diff-company">' + u.escapeHtml(d.company) + '</div>' +
      '<div class="diff-amounts">' +
      '<div class="diff-amount"><span class="lbl">Таблица 1</span><span class="val">' + u.formatNumber(d.amount1) + '</span></div>' +
      '<div class="diff-amount"><span class="lbl">Таблица 2</span><span class="val">' + u.formatNumber(d.amount2) + '</span></div>' +
      '<div class="diff-amount"><span class="lbl">Разница</span><span class="val" style="color:#dc3545">' + u.formatNumber(d.difference) + '</span></div>' +
      '</div></li>';
  });
  html += '</ul>';
  ctx.setOutput(html);
}
