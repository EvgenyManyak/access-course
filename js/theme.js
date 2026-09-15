/* ============================================================
   theme.js — светлая и тёмная тема
   - сохранение выбора в localStorage (ключ ac.theme.v1);
   - первый визит: по системной настройке prefers-color-scheme;
   - переключатель в топбаре (data-action="toggle-theme").
   ============================================================ */
(function () {
  'use strict';
  var AC = window.AC;
  var KEY = 'ac.theme.v1';

  AC.theme = {
    get: function () {
      var saved = AC.store.get(KEY, null);
      if (saved === 'light' || saved === 'dark') return saved;
      if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
        return 'dark';
      }
      return 'light';
    },

    apply: function (theme) {
      var root = document.documentElement;
      root.setAttribute('data-theme', theme === 'dark' ? 'dark' : 'light');
      var icon = AC.$('#theme-btn-icon');
      if (icon) icon.textContent = theme === 'dark' ? '☀️' : '🌙';
      var meta = document.querySelector('meta[name="theme-color"]');
      if (meta) meta.setAttribute('content', theme === 'dark' ? '#0f141b' : '#ffffff');
    },

    toggle: function () {
      var next = this.get() === 'dark' ? 'light' : 'dark';
      AC.store.set(KEY, next);
      this.apply(next);
      AC.toast(next === 'dark' ? 'Тёмная тема включена' : 'Светлая тема включена', 1400);
    },

    init: function () {
      this.apply(this.get());
      var self = this;
      AC.registerAction('toggle-theme', function () { self.toggle(); });

      /* Если пользователь не выбирал тему вручную — следим за системной */
      if (window.matchMedia && AC.store.get(KEY, null) === null) {
        var mq = window.matchMedia('(prefers-color-scheme: dark)');
        var onChange = function () { self.apply(self.get()); };
        if (mq.addEventListener) mq.addEventListener('change', onChange);
        else if (mq.addListener) mq.addListener(onChange);
      }
    }
  };
})();
