/* ============================================================
   app.js — ядро учебника
   - манифест модулей курса;
   - общие утилиты ($, store, escapeHtml, toast);
   - подсветка SQL/VBA;
   - сборка оболочки интерфейса (сайдбар, топбар, подвал, тосты);
   - порядок инициализации остальных модулей.
   Загружается ПЕРВЫМ; theme.js, progress.js, export.js,
   navigation.js, quiz.js регистрируют свои init() в AC.
   ============================================================ */
(function () {
  'use strict';

  var AC = window.AC = {};

  /* ---------- 1. Манифест курса ---------- */
  AC.MANIFEST = [
    { id: 'index',    file: 'index.html', num: '',    icon: '🏠', group: 'Обзор',
      title: 'Главная', short: 'Обзор курса' },

    { id: 'start', file: 'getting-started.html', num: '', icon: '🚀', group: 'Обзор',
      title: 'Начало работы', short: 'Настройка рабочей среды' },

    { id: 'm01', file: 'modules/01-access-basics.html', num: '01', icon: '🧭', group: 'Модули курса',
      title: 'Access для аналитика: отличие от Excel и SQL',
      desc: 'Интерфейс Access, объекты базы данных, когда выбирать Access вместо Excel или SQL' },
    { id: 'm02', file: 'modules/02-tables.html', num: '02', icon: '🗄️', group: 'Модули курса',
      title: 'Проектирование таблиц и типы данных',
      desc: 'Конструктор таблиц, типы данных Access, первичные ключи, индексы, нормализация' },
    { id: 'm03', file: 'modules/03-relationships.html', num: '03', icon: '🔗', group: 'Модули курса',
      title: 'Ключи, индексы и связи между таблицами',
      desc: 'Схема данных, ссылочная целостность, каскадные обновления и удаления' },
    { id: 'm04', file: 'modules/04-import.html', num: '04', icon: '📥', group: 'Модули курса',
      title: 'Импорт CSV и Excel',
      desc: 'Мастер импорта, типы и ошибки импорта, сохранённые спецификации импорта' },
    { id: 'm05', file: 'modules/05-cleaning.html', num: '05', icon: '🧹', group: 'Модули курса',
      title: 'Очистка и проверка данных',
      desc: 'Дедупликация, пропуски, проверки на уровне поля и таблицы, маски ввода' },
    { id: 'm06', file: 'modules/06-query-builder.html', num: '06', icon: '🧮', group: 'Модули курса',
      title: 'Конструктор запросов: фильтрация и сортировка',
      desc: 'Условия отбора, вычисляемые поля, сортировка, выборка TOP N' },
    { id: 'm07', file: 'modules/07-access-sql.html', num: '07', icon: '⌨️', group: 'Модули курса',
      title: 'Access SQL и соединения таблиц',
      desc: 'Диалект Access SQL, INNER/LEFT/RIGHT JOIN, декартово произведение, особенности ANSI-89/92' },
    { id: 'm08', file: 'modules/08-analytics.html', num: '08', icon: '📊', group: 'Модули курса',
      title: 'Аналитические запросы: агрегаты, параметры, перекрёстные, action',
      desc: 'GROUP BY и агрегатные функции, параметрические и перекрёстные запросы, запросы на изменение' },
    { id: 'm09', file: 'modules/09-forms.html', num: '09', icon: '🪟', group: 'Модули курса',
      title: 'Формы и подчинённые формы',
      desc: 'Мастер форм, элементы управления, кнопки, подчинённые формы, навигационная форма' },
    { id: 'm10', file: 'modules/10-reports.html', num: '10', icon: '🖨️', group: 'Модули курса',
      title: 'Отчёты',
      desc: 'Мастер отчётов, группировка и итоги, выражения, экспорт отчётов в PDF' },
    { id: 'm11', file: 'modules/11-macros.html', num: '11', icon: '🤖', group: 'Модули курса',
      title: 'Макросы Access',
      desc: 'Конструктор макросов, действия и условия, макросы данных, кнопочные макросы' },
    { id: 'm12', file: 'modules/12-vba.html', num: '12', icon: '💻', group: 'Модули курса',
      title: 'VBA на прикладном уровне',
      desc: 'Чтение готового кода, запуск процедур, события, DoCmd, циклы по Recordset' },
    { id: 'm13', file: 'modules/13-export.html', num: '13', icon: '🔁', group: 'Модули курса',
      title: 'Импорт и экспорт: автоматизация обмена данными',
      desc: 'TransferText/TransferSpreadsheet, OutputTo, спецификации обмена, связанные таблицы' },
    { id: 'm14', file: 'modules/14-optimization.html', num: '14', icon: '🛠️', group: 'Модули курса',
      title: 'Оптимизация, разделение базы и обслуживание',
      desc: 'Сжатие и восстановление, индексы, разделение front-end/back-end, резервные копии' },
    { id: 'm15', file: 'modules/15-final-project.html', num: '15', icon: '🏁', group: 'Модули курса',
      title: 'Итоговый аналитический проект',
      desc: 'Сквозное решение: от CSV до разделённой базы с формами, отчётами и автоматизацией' },

    { id: 'datasets', file: 'datasets/kaggle-links.html', num: '', icon: '🗃️', group: 'Материалы',
      title: 'Датасеты и ссылки Kaggle', short: 'Датасеты курса' }
  ];

  AC.DOCS_LINKS = [
    { file: 'docs/course-outline.md', title: 'Программа курса (course-outline.md)' },
    { file: 'docs/student-checklist.md', title: 'Чек-лист студента (student-checklist.md)' }
  ];

  /* ---------- 2. Утилиты ---------- */
  AC.$ = function (sel, root) { return (root || document).querySelector(sel); };
  AC.$$ = function (sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); };

  AC.escapeHtml = function (s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  };

  /* Безопасный localStorage: file:// и приватные режимы могут бросать исключения */
  (function () {
    var mem = {};
    var available = false;
    try {
      var probe = '__ac_probe__';
      window.localStorage.setItem(probe, '1');
      window.localStorage.removeItem(probe);
      available = true;
    } catch (e) { available = false; }

    AC.store = {
      get: function (key, fallback) {
        try {
          if (available) {
            var raw = window.localStorage.getItem(key);
            return raw === null ? fallback : JSON.parse(raw);
          }
          return (key in mem) ? mem[key] : fallback;
        } catch (e) { return fallback; }
      },
      set: function (key, value) {
        try {
          var raw = JSON.stringify(value);
          if (available) window.localStorage.setItem(key, raw); else mem[key] = value;
        } catch (e) { mem[key] = value; }
      },
      remove: function (key) {
        try { if (available) window.localStorage.removeItem(key); } catch (e) { /* ignore */ }
        delete mem[key];
      }
    };
  })();

  var toastTimer = null;
  AC.toast = function (message, ms) {
    var wrap = AC.$('.toast-wrap');
    if (!wrap) return;
    var t = document.createElement('div');
    t.className = 'toast';
    t.textContent = message;
    wrap.appendChild(t);
    requestAnimationFrame(function () { t.classList.add('is-shown'); });
    setTimeout(function () {
      t.classList.remove('is-shown');
      setTimeout(function () { if (t.parentNode) t.parentNode.removeChild(t); }, 300);
    }, ms || 2600);
  };

  /* Префикс до корня проекта: страницы в modules/ и datasets/ ссылаются на файлы через ../ */
  AC.rootPrefix = /\/(modules|datasets)\//.test(window.location.pathname) ? '../' : '';

  AC.navUrl = function (file) {
    return window.IS_STANDALONE ? null : (AC.rootPrefix + file);
  };

  AC.getCurrentId = function () {
    var b = document.body;
    if (window.IS_STANDALONE) {
      return (AC.route && AC.route.id) ? AC.route.id : 'index';
    }
    return b.getAttribute('data-module') || b.getAttribute('data-page') || 'index';
  };

  AC.moduleById = function (id) {
    for (var i = 0; i < AC.MANIFEST.length; i++) {
      if (AC.MANIFEST[i].id === id) return AC.MANIFEST[i];
    }
    return null;
  };

  AC.neighborModules = function (id) {
    var idx = -1, i;
    for (i = 0; i < AC.MANIFEST.length; i++) { if (AC.MANIFEST[i].id === id) { idx = i; break; } }
    if (idx === -1) return { prev: null, next: null };
    return {
      prev: idx > 0 ? AC.MANIFEST[idx - 1] : null,
      next: idx < AC.MANIFEST.length - 1 ? AC.MANIFEST[idx + 1] : null
    };
  };

  AC.goTo = function (id, anchor) {
    var m = AC.moduleById(id);
    if (!m) return;
    if (window.IS_STANDALONE) {
      window.location.hash = '#/' + id + (anchor ? '/' + anchor : '');
    } else {
      window.location.href = AC.rootPrefix + m.file + (anchor ? '#' + anchor : '');
    }
  };

  /* ---------- 3. Подсветка SQL и VBA ---------- */
  var SQL_KEYWORDS = 'SELECT|FROM|WHERE|INNER|LEFT|RIGHT|FULL|OUTER|JOIN|ON|GROUP BY|ORDER BY|HAVING|AS|AND|OR|NOT|NULL|IS|LIKE|IN|BETWEEN|EXISTS|DISTINCT|DISTINCTROW|TOP|INSERT INTO|VALUES|UPDATE|SET|DELETE|SELECT INTO|ALTER TABLE|CREATE TABLE|DROP TABLE|ADD|COLUMN|CONSTRAINT|PRIMARY KEY|FOREIGN KEY|REFERENCES|UNIQUE|INDEX|PARAMETERS|TRANSFORM|PIVOT|UNION|ALL|ANY|SOME|CASE|WHEN|THEN|ELSE|END|ASC|DESC|COUNT|SUM|AVG|MIN|MAX|FIRST|LAST|STDEV|VAR|IIF|SWITCH|CHOOSE|FORMAT|NZ|DATEADD|DATEDIFF|DATEPART|DATESERIAL|YEAR|MONTH|DAY|WEEKDAY|NOW|DATE|CVAR|CINT|CLNG|CDBL|CSTR|CDATE|CCUR|VAL|LEN|LEFT|RIGHT|LTRIM|RTRIM|TRIM|MID|INSTR|REPLACE|UCASE|LCASE|SPACE|STRING|ABS|INT|FIX|ROUND|SGN|SQR|RND|PARTITION|USER|DOMAIN'.split('|');
  var VBA_KEYWORDS = 'Sub|Function|Property Get|Property Let|Property Set|End Sub|End Function|End If|End With|Dim|Const|Set|Let|As|If|Then|ElseIf|Else|For|Each|In|To|Step|Next|Do|While|Until|Loop|With|Select Case|Case|Call|ByVal|ByRef|Exit|Public|Private|Static|Option Explicit|On Error|GoTo|Resume|Error|True|False|Nothing|Empty|Null|And|Or|Not|Xor|Mod|New|Me|Long|Integer|String|Double|Single|Currency|Date|Boolean|Variant|Object|Byte|Application|DoCmd|Screen|Forms|Reports|CurrentDb|CurrentProject|Access'.split('|');

  function buildRegex(keywords) {
    var kw = keywords.slice().sort(function (a, b) { return b.length - a.length; }).join('|');
    return new RegExp(
      "('(?:[^']|'')*')|" +          // 1: строка SQL
      "(\"[^\"]*\")|" +              // 2: строка в кавычках
      "('(?:[^'\\n])*')|" +          // 3: строка VBA (апостроф)
      "(\\s(?:Rem\\s)[^\\n]*)|" +    // 4: Rem-комментарий
      "(\\b\\d+(?:\\.\\d+)?\\b)|" +  // 5: число
      "(#[^#\\n]+#)|" +              // 6: литерал даты Access #...#
      "(--[^\\n]*)|" +               // 7: SQL-комментарий --
      "(\\b(?:" + kw + ")\\b)",      // 8: ключевое слово
      'gi');
  }
  var RX_SQL = buildRegex(SQL_KEYWORDS);
  var RX_VBA = buildRegex(VBA_KEYWORDS);

  /* VBA-комментарий начинается с апострофа в начале строки или после пробела,
     поэтому для VBA строка-апостроф и комментарий конфликтуют: раздельные проходы */
  function highlightVba(escaped) {
    /* Сначала вырезаем и защищаем комментарии (апостроф до конца строки) */
    var parts = escaped.split(/(\n)/);
    var out = [];
    for (var i = 0; i < parts.length; i++) {
      var chunk = parts[i];
      if (chunk === '\n') { out.push(chunk); continue; }
      var apos = chunk.indexOf("'");
      var code = chunk, comment = null;
      if (apos !== -1) { code = chunk.slice(0, apos); comment = chunk.slice(apos); }
      code = code.replace(RX_VBA, function (m, str1, str2, str3, rem, num, date, sqlc, kw) {
        if (str1 || str2 || str3 || rem || date) return '<span class="tok-str">' + m + '</span>';
        if (sqlc) return '<span class="tok-com">' + m + '</span>';
        if (num) return '<span class="tok-num">' + m + '</span>';
        if (kw) return '<span class="tok-kw">' + m + '</span>';
        return m;
      });
      if (comment) code += '<span class="tok-com">' + comment + '</span>';
      out.push(code);
    }
    return out.join('');
  }

  function highlightSql(escaped) {
    return escaped.replace(RX_SQL, function (m, str1, str2, str3, rem, num, date, com, kw) {
      if (str1 || str2 || str3 || date) return '<span class="tok-str">' + m + '</span>';
      if (com || rem) return '<span class="tok-com">' + m + '</span>';
      if (num) return '<span class="tok-num">' + m + '</span>';
      if (kw) return '<span class="tok-kw">' + m + '</span>';
      return m;
    });
  }

  AC.highlightElement = function (codeEl) {
    var lang = codeEl.className;
    var text = codeEl.textContent;
    var escaped = AC.escapeHtml(text);
    var html;
    try {
      if (lang.indexOf('lang-vba') !== -1 || lang.indexOf('lang-vb') !== -1) {
        html = highlightVba(escaped);
      } else if (lang.indexOf('lang-sql') !== -1 || lang.indexOf('lang-access') !== -1) {
        html = highlightSql(escaped);
      } else {
        html = escaped;
      }
    } catch (e) { html = escaped; }
    codeEl.innerHTML = html;
  };

  AC.highlightAll = function (root) {
    AC.$$('pre.code > code', root || document).forEach(function (c) {
      if (c.getAttribute('data-hl') === '1') return;
      AC.highlightElement(c);
      c.setAttribute('data-hl', '1');
    });
  };

  /* ---------- 4. Сборка оболочки ---------- */
  function buildShell() {
    var content = AC.$('#content');
    if (!content) return;

    var app = document.createElement('div');
    app.className = 'app';

    /* Сайдбар */
    var sidebar = document.createElement('aside');
    sidebar.className = 'sidebar';
    sidebar.id = 'courseMenu';   /* эталонный id мобильного меню (см. menuToggle) */
    sidebar.innerHTML =
      '<div class="sidebar__brand">' +
      '  <div class="sidebar__brand-icon">📗</div>' +
      '  <div>' +
      '    <div class="sidebar__brand-title">Microsoft Access 2007–2016</div>' +
      '    <div class="sidebar__brand-sub">Интерактивный учебник для аналитиков</div>' +
      '  </div>' +
      '</div>' +
      '<div class="sidebar__progress">' +
      '  <div class="progress-row">' +
      '    <span class="progress-label">Прогресс курса</span>' +
      '    <span class="progress-value" id="sidebar-progress-value">0%</span>' +
      '  </div>' +
      '  <div class="progress-track"><div class="progress-fill" id="sidebar-progress-fill"></div></div>' +
      '  <button type="button" class="progress-reset" data-action="reset-progress">Сбросить прогресс</button>' +
      '</div>' +
      '<nav class="sidebar__nav" id="sidebar-nav" aria-label="Навигация по курсу"></nav>' +
      '<div class="sidebar__footer">' +
      '  <div><kbd>←</kbd> <kbd>→</kbd> — модули, <kbd>/</kbd> — поиск, <kbd>Esc</kbd> — закрыть</div>' +
      '  <div style="margin-top:6px"><a href="#" class="js-doc-link" data-file="docs/course-outline.md">Программа курса</a> · ' +
      '  <a href="#" class="js-doc-link" data-file="docs/student-checklist.md">Чек-лист студента</a></div>' +
      '</div>';

    var backdrop = document.createElement('div');
    backdrop.className = 'sidebar-backdrop';
    backdrop.setAttribute('data-action', 'close-nav');
    backdrop.setAttribute('aria-hidden', 'true');

    var main = document.createElement('div');
    main.className = 'app__main';

    /* Топбар */
    var topbar = document.createElement('header');
    topbar.className = 'topbar';
    topbar.innerHTML =
      '<button type="button" id="menuToggle" class="icon-btn topbar__burger" data-action="toggle-nav" aria-label="Открыть меню" aria-expanded="false" aria-controls="courseMenu" title="Меню">☰</button>' +
      '<div class="topbar__title" id="topbar-title"></div>' +
      '<div class="topbar__actions">' +
      '  <button type="button" class="icon-btn" data-action="open-search" aria-label="Поиск по учебнику" title="Поиск ( / )">🔍<span class="btn-text">Поиск</span></button>' +
      '  <button type="button" class="icon-btn" data-action="toggle-theme" aria-label="Светлая или тёмная тема" title="Светлая / тёмная тема"><span id="theme-btn-icon">🌙</span></button>' +
      '  <button type="button" class="icon-btn icon-btn--accent" data-action="open-export" aria-label="Экспорт" title="Экспорт и печать">⬇<span class="btn-text">Экспорт</span></button>' +
      '</div>';

    var topProgress = document.createElement('div');
    topProgress.className = 'topbar-progress';
    topProgress.innerHTML = '<div class="topbar-progress__fill" id="topbar-progress-fill"></div>';

    var page = document.createElement('div');
    page.className = 'page';

    /* Переносим существующий контент внутрь каркаса */
    content.parentNode.insertBefore(app, content);
    page.appendChild(content);
    main.appendChild(topbar);
    main.appendChild(topProgress);
    main.appendChild(page);
    app.appendChild(sidebar);
    app.appendChild(backdrop);
    app.appendChild(main);

    /* Подвал: навигация назад/далее + кнопка пройдено + строка-футер */
    var navFooter = document.createElement('div');
    navFooter.className = 'page__nav';
    navFooter.id = 'page-nav';
    page.appendChild(navFooter);

    var doneRow = document.createElement('div');
    doneRow.className = 'done-btn-row';
    doneRow.id = 'done-row';
    page.appendChild(doneRow);

    var pageFooter = document.createElement('footer');
    pageFooter.className = 'page__footer';
    pageFooter.innerHTML =
      '<span>Учебник «Microsoft Access для аналитика» · версия 1.0 · 2026</span>' +
      '<span>Сквозной датасет: ShopData (Olist Brazilian E-commerce)</span>';
    page.appendChild(pageFooter);

    /* Тосты */
    var toasts = document.createElement('div');
    toasts.className = 'toast-wrap';
    document.body.appendChild(toasts);

    /* Поисковый оверлей (наполняет navigation.js) */
    var search = document.createElement('div');
    search.className = 'overlay';
    search.id = 'search-overlay';
    search.innerHTML =
      '<div class="dialog" role="dialog" aria-modal="true" aria-label="Поиск по учебнику">' +
      '  <div class="dialog__head"><span>Поиск по учебнику</span><button type="button" class="dialog__close" data-action="close-search" aria-label="Закрыть">✕</button></div>' +
      '  <div class="dialog__body">' +
      '    <input type="search" class="search-input" id="search-input" placeholder="Например: внешние ключи, перекрёстный запрос, TransferText…" autocomplete="off">' +
      '    <ul class="search-results" id="search-results"></ul>' +
      '  </div>' +
      '</div>';
    document.body.appendChild(search);

    /* Панель экспорта (наполняет export.js) */
    var exportOverlay = document.createElement('div');
    exportOverlay.className = 'overlay';
    exportOverlay.id = 'export-overlay';
    exportOverlay.innerHTML = '<div class="dialog" id="export-dialog" role="dialog" aria-modal="true" aria-label="Экспорт и печать"></div>';
    document.body.appendChild(exportOverlay);

    /* Титульная страница для печати */
    var cover = document.createElement('div');
    cover.className = 'print-cover';
    cover.id = 'print-cover';
    var first = content.firstChild;
    if (first) content.insertBefore(cover, first); else content.appendChild(cover);

    /* Делегирование действий */
    document.addEventListener('click', function (ev) {
      var el = ev.target.closest ? ev.target.closest('[data-action]') : null;
      if (!el) return;
      var action = el.getAttribute('data-action');
      var handlers = AC.actionHandlers;
      if (handlers && handlers[action]) {
        ev.preventDefault();
        handlers[action](el, ev);
      }
    });
  }

  AC.extractModuleTitle = function (h1) {
    var kicker = h1.querySelector ? h1.querySelector('.module-title__kicker') : null;
    var rest = h1.cloneNode(true);
    if (rest.querySelector) {
      var k = rest.querySelector('.module-title__kicker');
      if (k) k.parentNode.removeChild(k);
      /* <br> внутри заголовка — это перенос строки, а не склейка слов */
      Array.prototype.slice.call(rest.querySelectorAll('br')).forEach(function (br) {
        br.parentNode.replaceChild(document.createTextNode(' '), br);
      });
    }
    var body = rest.textContent.replace(/\s+/g, ' ').trim();
    if (kicker) return kicker.textContent.replace(/\s+/g, ' ').trim() + '. ' + body;
    return body;
  };

  AC.actionHandlers = {};
  AC.registerAction = function (name, fn) { AC.actionHandlers[name] = fn; };

  /* ---------- 5. Инициализация ---------- */
  function init() {
    try { buildShell(); } catch (e) { /* оболочка критична, но не блокируем контент */ }

    /* Тема — до отрисовки остальных частей, чтобы не мигало */
    try { AC.theme.init(); } catch (e) {}

    try { AC.progress.init(); } catch (e) {}
    try { AC.export.init(); } catch (e) {}
    try { AC.navigation.init(); } catch (e) {}
    try { AC.quiz.init(); } catch (e) {}

    try { AC.highlightAll(); } catch (e) {}

    /* Заголовок в топбаре */
    var titleEl = AC.$('#topbar-title');
    if (titleEl) {
      var h1 = AC.$('#content h1');
      titleEl.textContent = h1 ? AC.extractModuleTitle(h1) : 'Microsoft Access 2007–2016';
    }

    /* Перед печатью раскрываем все скрытые решения, после — возвращаем как было */
    var openedForPrint = [];
    window.addEventListener('beforeprint', function () {
      openedForPrint = AC.$$('details:not([open])');
      openedForPrint.forEach(function (d) { d.setAttribute('open', 'open'); });
    });
    window.addEventListener('afterprint', function () {
      openedForPrint.forEach(function (d) { d.removeAttribute('open'); });
      openedForPrint = [];
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
