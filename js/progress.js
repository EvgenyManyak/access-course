/* ============================================================
   progress.js — прогресс обучения
   - отметки «посещён» / «пройден» по модулям;
   - результаты тестов (лучший результат по модулю);
   - интерактивные чек-листы с сохранением состояния;
   - индикаторы прогресса (сайдбар + топбар);
   - события ac:progress для перерисовки навигации.
   Хранилище: localStorage, ключ ac.progress.v1.
   ============================================================ */
(function () {
  'use strict';
  var AC = window.AC;
  var KEY = 'ac.progress.v1';
  var MODULE_IDS = ['m01','m02','m03','m04','m05','m06','m07','m08','m09','m10','m11','m12','m13','m14','m15'];

  var data = { visited: {}, done: {}, quizzes: {}, checks: {} };

  function load() {
    var saved = AC.store.get(KEY, null);
    if (saved && typeof saved === 'object') {
      data.visited = saved.visited || {};
      data.done = saved.done || {};
      data.quizzes = saved.quizzes || {};
      data.checks = saved.checks || {};
    }
  }
  function save() {
    AC.store.set(KEY, data);
    document.dispatchEvent(new CustomEvent('ac:progress'));
  }

  AC.progress = {
    data: data,

    percent: function () {
      var done = 0;
      MODULE_IDS.forEach(function (id) { if (data.done[id]) done++; });
      return Math.round(done / MODULE_IDS.length * 100);
    },

    markVisited: function (id) {
      if (!id || data.visited[id]) return;
      data.visited[id] = Date.now();
      save();
    },

    isDone: function (id) { return !!data.done[id]; },
    isVisited: function (id) { return !!data.visited[id]; },

    toggleDone: function (id) {
      if (!id) return;
      if (data.done[id]) { delete data.done[id]; } else { data.done[id] = Date.now(); }
      save();
    },

    quizBest: function (id) {
      return data.quizzes[id] || null; /* {score, total, ts} */
    },

    saveQuiz: function (id, score, total) {
      var prev = data.quizzes[id];
      if (!prev || score > prev.score) {
        data.quizzes[id] = { score: score, total: total, ts: Date.now() };
        save();
      }
    },

    checkKey: function (key) { return !!data.checks[key]; },
    setCheck: function (key, on) {
      if (on) data.checks[key] = true; else delete data.checks[key];
      save();
    },

    checkedCount: function (moduleId) {
      var n = 0;
      Object.keys(data.checks).forEach(function (k) {
        if (k.indexOf(moduleId + ':') === 0 && data.checks[k]) n++;
      });
      return n;
    },

    reset: function () {
      data = { visited: {}, done: {}, quizzes: {}, checks: {} };
      AC.progress.data = data;
      AC.store.remove(KEY);
      save();
    },

    /* ---------- Отрисовка индикаторов ---------- */
    render: function () {
      var pct = this.percent();

      var fill = AC.$('#sidebar-progress-fill');
      if (fill) fill.style.width = pct + '%';
      var val = AC.$('#sidebar-progress-value');
      if (val) val.textContent = pct + '%';
      var topFill = AC.$('#topbar-progress-fill');
      if (topFill) topFill.style.width = pct + '%';

      /* Кнопка «Модуль пройден» */
      var id = AC.getCurrentId();
      var doneBtn = AC.$('#done-row');
      if (doneBtn && MODULE_IDS.indexOf(id) !== -1) {
        var isDone = this.isDone(id);
        doneBtn.innerHTML =
          '<button type="button" class="btn ' + (isDone ? 'btn--ghost' : '') + '" data-action="toggle-done">' +
          (isDone ? '✓ Модуль отмечен пройденным — снять отметку' : '✓ Отметить модуль пройденным') +
          '</button>';
      }
    },

    /* ---------- Чек-листы ---------- */
    bindChecklists: function () {
      var self = this;
      AC.$$('.checklist[data-checklist]').forEach(function (ul) {
        var items = AC.$$('input[type="checkbox"][data-ck]', ul);
        items.forEach(function (input) {
          var key = input.getAttribute('data-ck');
          input.checked = self.checkKey(key);
          input.addEventListener('change', function () {
            self.setCheck(key, input.checked);
          });
        });
      });
    },

    init: function () {
      var self = this;
      load();
      this.bindChecklists();
      this.render();

      AC.registerAction('toggle-done', function () {
        self.toggleDone(AC.getCurrentId());
        self.render();
        var id = AC.getCurrentId();
        AC.toast(self.isDone(id) ? 'Модуль отмечен как пройденный ✓' : 'Отметка о прохождении снята');
      });
      AC.registerAction('reset-progress', function () {
        if (window.confirm('Сбросить весь прогресс: отметки о прохождении, результаты тестов и чек-листы?')) {
          self.reset();
          self.render();
          AC.navigation && AC.navigation.refreshNavStates();
          AC.toast('Прогресс сброшен');
        }
      });

      document.addEventListener('ac:progress', function () { self.render(); });
    }
  };
})();
