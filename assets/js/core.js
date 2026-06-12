/* ============================================================
 * AccProg — ядро оболочки (core.js)
 * Отвечает за: реестр функций, построение меню, роутинг,
 * предоставление контекста (ctx) модулям-функциям.
 * ============================================================ */
(function () {
  'use strict';

  var registry = {};   // id -> определение функции
  var order = [];      // порядок id (как в реестре)
  var active = null;   // id активной функции
  var states = {};     // id -> состояние функции (сохраняется между переключениями)

  // DOM
  var els = {};
  function $(id) { return document.getElementById(id); }

  /* ---------- Публичный API регистрации модуля ---------- */
  // Каждый файл assets/functions/<id>.js вызывает AccProg.register({...})
  var AccProg = {
    register: function (def) {
      if (!def || !def.id) { console.error('AccProg.register: нет id', def); return; }
      registry[def.id] = def;
    }
  };
  window.AccProg = AccProg;

  /* ---------- Контекст для модуля ---------- */
  function makeCtx(id) {
    if (!states[id]) states[id] = {};
    return {
      id: id,
      inputEl: els.inputBody,
      actionsEl: els.actionsBody,
      outputEl: els.outputBody,
      state: states[id],
      utils: window.AccProgUtils,
      // Хелперы для модуля:
      setOutput: function (html) { els.outputBody.innerHTML = html; },
      clearOutput: function () { els.outputBody.innerHTML = ''; },
      enableAction: function (actionId, enabled) {
        var btn = els.actionsBody.querySelector('[data-action="' + actionId + '"]');
        if (btn) btn.disabled = !enabled;
      },
      message: function (text, kind) {
        return '<div class="msg ' + (kind || '') + '">' + window.AccProgUtils.escapeHtml(text) + '</div>';
      }
    };
  }

  /* ---------- Построение меню из реестра ---------- */
  function buildMenu() {
    els.menu.innerHTML = '';
    order.forEach(function (id) {
      var def = registry[id];
      if (!def) return;
      var btn = document.createElement('button');
      btn.className = 'fn-item';
      btn.dataset.id = id;
      btn.innerHTML = '<span class="fn-icon">' + (def.icon || '•') + '</span>' +
        '<span class="fn-label">' + window.AccProgUtils.escapeHtml(def.title || id) + '</span>';
      btn.addEventListener('click', function () { activate(id); });
      els.menu.appendChild(btn);
    });
  }

  /* ---------- Активация функции ---------- */
  function activate(id) {
    var def = registry[id];
    if (!def) { console.error('Функция не найдена:', id); return; }
    active = id;

    // подсветка меню
    Array.prototype.forEach.call(els.menu.children, function (b) {
      b.classList.toggle('active', b.dataset.id === id);
    });

    // шапка
    els.title.textContent = def.title || id;
    els.description.innerHTML = def.description || '';

    // очистка панелей
    els.inputBody.innerHTML = '';
    els.actionsBody.innerHTML = '';
    els.outputBody.innerHTML = '<div class="placeholder">Здесь появится результат после выполнения операции.</div>';

    var ctx = makeCtx(id);

    // Панель 1 — ввод
    if (typeof def.renderInput === 'function') def.renderInput(ctx);

    // Панель 2 — кнопки операций
    if (Array.isArray(def.actions)) {
      def.actions.forEach(function (act) {
        var btn = document.createElement('button');
        btn.textContent = act.label;
        btn.dataset.action = act.id;
        if (act.variant) btn.className = act.variant; // 'secondary' | 'danger'
        if (act.disabled) btn.disabled = true;
        btn.addEventListener('click', function () {
          if (typeof act.run === 'function') act.run(ctx);
        });
        els.actionsBody.appendChild(btn);
      });
    }

    // location hash для навигации/перезагрузки
    if (location.hash !== '#' + id) {
      history.replaceState(null, '', '#' + id);
    }
  }

  /* ---------- Загрузка реестра и скриптов функций ---------- */
  function loadFunction(entry) {
    return new Promise(function (resolve) {
      var s = document.createElement('script');
      s.src = entry.script;
      s.onload = resolve;
      s.onerror = function () {
        console.error('Не удалось загрузить модуль:', entry.script);
        resolve();
      };
      document.body.appendChild(s);
    });
  }

  function init() {
    els = {
      menu: $('fnMenu'),
      title: $('fnTitle'),
      description: $('fnDescription'),
      inputBody: $('panelInputBody'),
      actionsBody: $('panelActionsBody'),
      outputBody: $('panelOutputBody')
    };

    fetch('functions.json')
      .then(function (r) {
        if (!r.ok) throw new Error('functions.json не найден (' + r.status + ')');
        return r.json();
      })
      .then(function (list) {
        order = list.map(function (e) { return e.id; });
        // последовательно подгружаем модули
        return list.reduce(function (chain, entry) {
          return chain.then(function () { return loadFunction(entry); });
        }, Promise.resolve());
      })
      .then(function () {
        buildMenu();
        // активируем функцию из hash или первую
        var fromHash = location.hash.replace('#', '');
        if (fromHash && registry[fromHash]) activate(fromHash);
        else if (order.length) activate(order[0]);
      })
      .catch(function (err) {
        els.inputBody.innerHTML =
          '<div class="msg error">Ошибка инициализации: ' +
          window.AccProgUtils.escapeHtml(err.message) +
          '<br>Запустите через локальный сервер (например: <code>python3 -m http.server</code>), ' +
          'т.к. fetch не работает с file://</div>';
      });
  }

  document.addEventListener('DOMContentLoaded', init);
})();
