/* ============================================================
   navigation.js — навигация и поиск
   - сайдбар: группы «Обзор / Модули курса / Материалы», отметки
     посещения и прохождения (из progress.js);
   - кликабельное оглавление текущего модуля (h2) + scrollspy;
   - «Назад / Далее» внизу страницы;
   - поиск по всему учебнику (полный текст из кэша курса);
   - клавиатура: ← → (модули), / или Ctrl+K (поиск), Esc (закрыть);
   - мобильный бургер;
   - роутер standalone-версии (hash #/m03/якорь).
   ============================================================ */
(function () {
  'use strict';
  var AC = window.AC;

  var searchIndex = null;   /* [{modId, modTitle, secId, heading, text}] */
  var selectedHit = 0;

  /* ================= 1. Сайдбар ================= */

  function navHref(m) {
    if (window.IS_STANDALONE) return '#/' + m.id;
    return AC.rootPrefix + m.file;
  }

  function buildSidebarNav() {
    var nav = AC.$('#sidebar-nav');
    if (!nav) return;
    var groups = [];
    AC.MANIFEST.forEach(function (m) {
      var g = m.group;
      if (!groups.length || groups[groups.length - 1].name !== g) groups.push({ name: g, items: [] });
      groups[groups.length - 1].items.push(m);
    });

    var html = '';
    groups.forEach(function (g) {
      html += '<div class="nav-group"><div class="nav-group__title">' + g.name + '</div>';
      g.items.forEach(function (m) {
        var label = m.num ? '<span class="nav-item__num">' + m.num + '</span>' : '<span class="nav-item__icon">' + m.icon + '</span>';
        html += '<a class="nav-item" data-nav="' + m.id + '" href="' + navHref(m) + '">' +
          label +
          '<span class="nav-item__label">' + (m.short || m.title) + '</span>' +
          '<span class="nav-item__state" data-state="' + m.id + '"></span>' +
          '</a>';
      });
      html += '</div>';
    });
    nav.innerHTML = html;

    nav.addEventListener('click', function (ev) {
      var a = ev.target.closest ? ev.target.closest('[data-nav]') : null;
      if (!a) return;
      closeMenu();
      if (window.IS_STANDALONE) {
        ev.preventDefault();
        window.location.hash = '#/' + a.getAttribute('data-nav');
      }
      /* обычный режим — обычный переход по href */
    });

    refreshNavStates();
    setActiveNav(AC.getCurrentId());
  }

  function refreshNavStates() {
    AC.$$('#sidebar-nav [data-state]').forEach(function (span) {
      var id = span.getAttribute('data-state');
      span.className = 'nav-item__state';
      if (AC.progress.isDone(id)) { span.classList.add('is-done'); span.textContent = '✓'; }
      else if (AC.progress.isVisited(id)) { span.classList.add('is-visited'); span.textContent = '◐'; }
      else { span.textContent = ''; }
    });

    /* Обновить карточки на главной (если они есть) */
    AC.$$('[data-module-card]').forEach(function (card) {
      var id = card.getAttribute('data-module-card');
      var state = card.querySelector('.module-card__state');
      if (!state) return;
      state.className = 'module-card__state';
      if (AC.progress.isDone(id)) { state.classList.add('is-done'); state.textContent = '✓ Пройден'; }
      else if (AC.progress.isVisited(id)) { state.classList.add('is-visited'); state.textContent = '◐ Начат'; }
      else { state.textContent = 'Не начат'; }
    });
  }

  function setActiveNav(id) {
    AC.$$('#sidebar-nav .nav-item').forEach(function (a) {
      a.classList.toggle('is-active', a.getAttribute('data-nav') === id);
    });
  }

  /* ================= 2. Оглавление модуля ================= */

  function buildToc() {
    var nav = AC.$('#sidebar-nav');
    if (!nav) return;
    var old = AC.$('#toc-group');
    if (old) old.parentNode.removeChild(old);

    var sections = AC.$$('#content .sec[id], #content section[id]');
    if (!sections.length) return;

    var html = '<div class="nav-group" id="toc-group">' +
      '<div class="nav-group__title">Содержание модуля</div>';
    sections.forEach(function (s) {
      var h2 = s.querySelector('h2');
      var label = h2 ? h2.textContent.trim() : s.id;
      html += '<a class="nav-item nav-item--toc" href="#' + s.id + '" data-toc="' + s.id + '">' +
        '<span class="nav-item__label">' + label + '</span></a>';
    });
    html += '</div>';
    nav.insertAdjacentHTML('beforeend', html);
  }

  function setupScrollSpy() {
    var sections = AC.$$('#content .sec[id], #content section[id]');
    if (!sections.length || !('IntersectionObserver' in window)) return;
    var links = {};
    AC.$$('#sidebar-nav [data-toc]').forEach(function (a) {
      links[a.getAttribute('data-toc')] = a;
    });
    var observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          AC.$$('#sidebar-nav [data-toc]').forEach(function (a) { a.classList.remove('is-active'); });
          var link = links[entry.target.id];
          if (link) {
            link.classList.add('is-active');
            var nav = AC.$('#sidebar-nav');
            if (nav && link.offsetTop > nav.scrollTop + nav.clientHeight - 80) {
              nav.scrollTop = link.offsetTop - nav.clientHeight / 2;
            }
          }
        }
      });
    }, { rootMargin: '-45px 0px -55% 0px', threshold: 0 });
    sections.forEach(function (s) { observer.observe(s); });
  }

  /* ================= 3. Назад / Далее ================= */

  function buildPrevNext() {
    var wrap = AC.$('#page-nav');
    if (!wrap) return;
    var id = AC.getCurrentId();
    var nb = AC.neighborModules(id);
    var html = '';
    if (nb.prev) {
      html += '<a class="page-nav__btn" data-nav="' + nb.prev.id + '" href="' + navHref(nb.prev) + '">' +
        '<span class="page-nav__dir">← Назад</span>' +
        '<span class="page-nav__name">' + (nb.prev.title || nb.prev.short) + '</span></a>';
    } else {
      html += '<span class="page-nav__spacer"></span>';
    }
    if (nb.next) {
      html += '<a class="page-nav__btn page-nav__btn--next" data-nav="' + nb.next.id + '" href="' + navHref(nb.next) + '">' +
        '<span class="page-nav__dir">Далее →</span>' +
        '<span class="page-nav__name">' + (nb.next.title || nb.next.short) + '</span></a>';
    } else {
      html += '<span class="page-nav__spacer"></span>';
    }
    wrap.innerHTML = html;
  }

  /* ================= 4. Поиск ================= */

  function buildSearchIndex() {
    if (searchIndex) return searchIndex;
    var index = [];
    var cache = AC.courseCache || {};

    function addHtml(modId, modTitle, html) {
      var root = document.createElement('div');
      root.innerHTML = html;
      var h1 = root.querySelector('h1');
      var pageTitle = h1 ? h1.textContent.trim() : modTitle;

      var secs = root.querySelectorAll('.sec[id], section[id]');
      if (secs.length) {
        secs.forEach(function (s) {
          var h2 = s.querySelector('h2');
          index.push({
            modId: modId,
            modTitle: pageTitle,
            secId: s.id,
            heading: h2 ? h2.textContent.trim() : s.id,
            text: s.textContent.replace(/\s+/g, ' ').trim()
          });
        });
      }
      index.push({
        modId: modId, modTitle: pageTitle, secId: '',
        heading: 'Введение',
        text: root.textContent.replace(/\s+/g, ' ').trim().slice(0, 500)
      });
    }

    /* Кэш, инъецированный сборкой (работает офлайн, включая file://) */
    AC.MANIFEST.forEach(function (m) {
      if (cache[m.id]) addHtml(m.id, m.title, cache[m.id]);
    });

    /* Резерв: если кэша нет (например, несобранная копия) — индексируем текущую страницу */
    if (!index.length) {
      var content = AC.$('#content');
      if (content) {
        var id = AC.getCurrentId();
        var meta = AC.moduleById(id);
        addHtml(id, meta ? meta.title : '', '<div>' + content.innerHTML + '</div>');
      }
    }
    searchIndex = index;
    return index;
  }

  function snippetAround(text, query) {
    var low = text.toLowerCase();
    var pos = low.indexOf(query);
    if (pos === -1) return text.slice(0, 140) + (text.length > 140 ? '…' : '');
    var start = Math.max(0, pos - 60);
    var end = Math.min(text.length, pos + query.length + 80);
    var snip = (start > 0 ? '…' : '') + text.slice(start, end) + (end < text.length ? '…' : '');
    var rx = new RegExp('(' + query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi');
    return snip.replace(rx, '<mark>$1</mark>');
  }

  function doSearch(query) {
    var resultsUl = AC.$('#search-results');
    if (!resultsUl) return;
    var q = query.trim().toLowerCase();
    if (q.length < 2) {
      resultsUl.innerHTML = '<li class="search-empty">Введите не менее двух символов. Поиск идёт по всем модулям, включая код и задания.</li>';
      selectedHit = 0;
      return;
    }
    var hits = [];
    buildSearchIndex().forEach(function (e) {
      var headingPos = e.heading.toLowerCase().indexOf(q);
      var textPos = e.text.toLowerCase().indexOf(q);
      if (headingPos !== -1 || textPos !== -1) {
        hits.push({ e: e, score: (headingPos !== -1 ? 3 : 0) + (headingPos === 0 ? 2 : 0) });
      }
    });
    hits.sort(function (a, b) { return b.score - a.score; });
    hits = hits.slice(0, 25);

    if (!hits.length) {
      resultsUl.innerHTML = '<li class="search-empty">Ничего не найдено. Попробуйте другой запрос: например, «связи», «нормализация», «перекрёстный», «TransferText».</li>';
      return;
    }

    selectedHit = 0;
    resultsUl.innerHTML = hits.map(function (h, i) {
      var meta = AC.moduleById(h.e.modId);
      var path = meta && meta.num ? 'Модуль ' + meta.num : (meta ? meta.title : '');
      return '<li><button type="button" class="search-hit' + (i === 0 ? ' is-selected' : '') + '" data-hit-mod="' + h.e.modId + '" data-hit-sec="' + h.e.secId + '">' +
        '<div class="search-hit__path">' + path + ' · ' + h.e.heading + '</div>' +
        '<div class="search-hit__snippet">' + snippetAround(h.e.text, q) + '</div>' +
        '</button></li>';
    }).join('');
  }

  function openSearch() {
    var ov = AC.$('#search-overlay');
    if (!ov) return;
    ov.classList.add('is-open');
    var input = AC.$('#search-input');
    if (input) { input.value = ''; doSearch(''); setTimeout(function () { input.focus(); }, 50); }
  }
  function closeSearch() {
    var ov = AC.$('#search-overlay');
    if (ov) ov.classList.remove('is-open');
  }

  function goSearchHit(modId, secId) {
    closeSearch();
    if (!modId) return;
    if (window.IS_STANDALONE) {
      window.location.hash = '#/' + modId + (secId ? '/' + secId : '');
    } else if (modId === AC.getCurrentId()) {
      if (secId) {
        var el = document.getElementById(secId);
        if (el) el.scrollIntoView({ behavior: 'smooth' });
      }
    } else {
      var m = AC.moduleById(modId);
      if (m) window.location.href = AC.rootPrefix + m.file + (secId ? '#' + secId : '');
    }
  }

  /* ================= 5а. Мобильное меню (бургер) =================
     Единая машина состояний для бургера и off-canvas сайдбара.
     Состояние: .is-open на меню + .is-active на кнопке + body.nav-open
     (хук существующего CSS для фона и блокировки прокрутки) + ARIA. */

  var MOBILE_NAV_BREAKPOINT = 1024;  /* тот же порог, что в responsive.css */

  function menuElements() {
    return {
      toggle: document.getElementById('menuToggle'),
      menu: document.getElementById('courseMenu')
    };
  }

  function syncMenuAria(toggle, menu, isOpen) {
    var desktop = window.innerWidth > MOBILE_NAV_BREAKPOINT;
    toggle.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
    toggle.setAttribute('aria-label', isOpen ? 'Закрыть меню' : 'Открыть меню');
    /* На десктопе сайдбар виден всегда — скрывать его от скринридеров нельзя */
    menu.setAttribute('aria-hidden', (isOpen || desktop) ? 'false' : 'true');
  }

  function closeMenu() {
    var els = menuElements();
    if (!els.toggle || !els.menu) return;
    els.menu.classList.remove('is-open');
    els.toggle.classList.remove('is-active');
    document.body.classList.remove('nav-open');
    syncMenuAria(els.toggle, els.menu, false);
  }

  function openMenu() {
    var els = menuElements();
    if (!els.toggle || !els.menu) return;
    els.menu.classList.add('is-open');
    els.toggle.classList.add('is-active');
    document.body.classList.add('nav-open');
    syncMenuAria(els.toggle, els.menu, true);
  }

  function toggleMenu() {
    var els = menuElements();
    if (!els.toggle || !els.menu) return;
    if (els.menu.classList.contains('is-open')) closeMenu(); else openMenu();
  }

  function setupMobileMenu() {
    var els = menuElements();
    if (!els.toggle || !els.menu) {
      console.error('Не найдены menuToggle или courseMenu — мобильное меню не работает');
      return;
    }

    /* Начальное ARIA-состояние с учётом текущей ширины экрана */
    syncMenuAria(els.toggle, els.menu, false);

    /* Прямой обработчик на кнопке (не полагаемся только на делегирование):
       stopPropagation исключает двойное срабатывание вместе с делегированным toggle-nav */
    els.toggle.addEventListener('click', function (ev) {
      ev.stopPropagation();
      toggleMenu();
    });

    /* Клик вне меню и вне кнопки — закрыть */
    document.addEventListener('click', function (ev) {
      if (!els.menu.classList.contains('is-open')) return;
      var t = ev.target;
      if (els.menu.contains(t) || els.toggle.contains(t)) return;
      closeMenu();
    });

    /* Escape — закрыть и вернуть фокус на кнопку (доступность с клавиатуры) */
    document.addEventListener('keydown', function (ev) {
      if (ev.key !== 'Escape') return;
      if (!els.menu.classList.contains('is-open')) return;
      closeMenu();
      els.toggle.focus();
    });

    /* Выбор пункта меню на мобильном — закрыть (в т.ч. якоря оглавления) */
    AC.$$('a', els.menu).forEach(function (link) {
      link.addEventListener('click', function () { closeMenu(); });
    });

    /* Переход к широкой версии — сброс мобильного состояния; обратно — синхронизация ARIA */
    window.addEventListener('resize', function () {
      if (window.innerWidth > MOBILE_NAV_BREAKPOINT) {
        closeMenu();
        els.menu.setAttribute('aria-hidden', 'false');
      } else if (!els.menu.classList.contains('is-open')) {
        els.menu.setAttribute('aria-hidden', 'true');
      }
    });
  }

  /* ================= 5. Клавиатура ================= */

  function setupKeyboard() {
    document.addEventListener('keydown', function (ev) {
      var inInput = /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement.tagName);

      if (ev.key === 'Escape') {
        closeSearch();
        AC.export && AC.export.closeExport();
        document.body.classList.remove('nav-open');
        return;
      }
      if ((ev.key === '/' && !inInput) || ((ev.ctrlKey || ev.metaKey) && ev.key.toLowerCase() === 'k')) {
        ev.preventDefault();
        openSearch();
        return;
      }
      var searchOpen = AC.$('#search-overlay') && AC.$('#search-overlay').classList.contains('is-open');
      if (searchOpen) {
        if (ev.key === 'Enter') {
          var first = AC.$('#search-results .search-hit');
          if (first) { ev.preventDefault(); goSearchHit(first.getAttribute('data-hit-mod'), first.getAttribute('data-hit-sec')); }
        }
        if (ev.key === 'ArrowDown' || ev.key === 'ArrowUp') {
          ev.preventDefault();
          var hits = AC.$$('#search-results .search-hit');
          if (!hits.length) return;
          hits[selectedHit] && hits[selectedHit].classList.remove('is-selected');
          selectedHit = ev.key === 'ArrowDown'
            ? Math.min(hits.length - 1, selectedHit + 1)
            : Math.max(0, selectedHit - 1);
          hits[selectedHit].classList.add('is-selected');
          hits[selectedHit].scrollIntoView({ block: 'nearest' });
        }
        return;
      }
      if (inInput) return;
      if (ev.key === 'ArrowRight' || ev.key === 'ArrowLeft') {
        var nb = AC.neighborModules(AC.getCurrentId());
        var target = ev.key === 'ArrowRight' ? nb.next : nb.prev;
        if (target) { AC.goTo(target.id); }
      }
    });

    /* Клики по результатам поиска */
    var results = AC.$('#search-results');
    if (results) {
      results.addEventListener('click', function (ev) {
        var btn = ev.target.closest ? ev.target.closest('.search-hit') : null;
        if (!btn) return;
        goSearchHit(btn.getAttribute('data-hit-mod'), btn.getAttribute('data-hit-sec'));
      });
    }
    var input = AC.$('#search-input');
    if (input) {
      var timer = null;
      input.addEventListener('input', function () {
        clearTimeout(timer);
        timer = setTimeout(function () { doSearch(input.value); }, 180);
      });
    }
  }

  /* ================= 6. Роутер standalone ================= */

  function renderFromCache(id, anchor) {
    var cache = AC.courseCache || {};
    var meta = AC.moduleById(id);
    var content = AC.$('#content');
    if (!meta || !cache[id] || !content) {
      /* Неизвестный маршрут — главная */
      id = 'index';
      meta = AC.moduleById('index');
      if (!cache[id] || !content) return;
    }

    /* В standalone страницы лежат в корне — убираем ../ у локальных ассетов */
    var html = cache[id].replace(/(src|href)="\.\.\//g, '$1="');
    content.innerHTML = html;

    /* innerHTML затёр титульную страницу печати — пересоздаём её */
    var oldCover = AC.$('#print-cover');
    if (oldCover) oldCover.parentNode.removeChild(oldCover);
    content.insertAdjacentHTML('afterbegin', '<div class="print-cover" id="print-cover"></div>');
    buildPrintCover();

    /* Пересобрать всё, что зависело от содержимого */
    try { AC.highlightAll(content); } catch (e) {}
    try { AC.progress.bindChecklists(); } catch (e) {}
    try { AC.progress.render(); } catch (e) {}
    try { AC.quiz.init(); } catch (e) {}
    try { buildToc(); setupScrollSpy(); } catch (e) {}
    try { buildPrevNext(); } catch (e) {}
    setActiveNav(id);
    refreshNavStates();
    AC.route = { id: id };

    var h1 = content.querySelector('h1');
    var titleEl = AC.$('#topbar-title');
    if (titleEl) titleEl.textContent = h1 ? AC.extractModuleTitle(h1) : (meta.title || '');
    document.title = (h1 ? AC.extractModuleTitle(h1) : meta.title) + ' — Учебник Microsoft Access';

    if (anchor) {
      var el = document.getElementById(anchor);
      if (el) { setTimeout(function () { el.scrollIntoView(); }, 30); return; }
    }
    window.scrollTo(0, 0);
  }

  function setupStandaloneRouter() {
    window.addEventListener('hashchange', function () {
      var parts = window.location.hash.replace(/^#\/?/, '').split('/');
      renderFromCache(parts[0] || 'index', parts[1] || '');
    });
    var parts = window.location.hash.replace(/^#\/?/, '').split('/');
    renderFromCache(parts[0] || 'index', parts[1] || '');
  }

  /* ================= 7. Титульная страница для печати ================= */

  function buildPrintCover() {
    var cover = AC.$('#print-cover');
    if (!cover) return;
    var items = AC.MANIFEST.filter(function (m) { return m.num; }).map(function (m) {
      return '<li><span class="toc-num">' + m.num + '</span>' + m.title + '</li>';
    }).join('');
    var d = new Date();
    cover.innerHTML =
      '<div class="print-cover__title">Microsoft Access 2007–2016<br>для аналитика данных</div>' +
      '<div class="print-cover__subtitle">Интерактивный учебник: от Excel и SQL к реляционной базе данных и прикладному VBA</div>' +
      '<div class="print-cover__toc-title">Содержание курса</div>' +
      '<ul class="print-cover__toc">' + items + '</ul>' +
      '<div class="print-cover__meta">80 академических часов · 15 модулей · 20 практических заданий<br>' +
      'Сквозной датасет: ShopData (Olist Brazilian E-commerce)<br>' +
      'Версия 1.0 · ' + d.getFullYear() + '</div>';
  }

  /* ================= 8. Инициализация ================= */

  AC.navigation = {
    refreshNavStates: refreshNavStates,
    openSearch: openSearch,
    closeSearch: closeSearch,

    init: function () {
      buildSidebarNav();
      buildPrintCover();

      /* Ссылки на документацию — с учётом глубины страницы */
      AC.$$('.js-doc-link').forEach(function (a) {
        a.href = AC.rootPrefix + a.getAttribute('data-file');
      });

      if (window.IS_STANDALONE) {
        setupStandaloneRouter();
      } else {
        buildToc();
        setupScrollSpy();
        buildPrevNext();
      }

      setupKeyboard();

      AC.registerAction('open-search', openSearch);
      AC.registerAction('close-search', closeSearch);
      AC.registerAction('toggle-nav', function () { toggleMenu(); });
      AC.registerAction('close-nav', function () { closeMenu(); });

      setupMobileMenu();

      /* Клик по якорным ссылкам внутри страницы — плавная прокрутка */
      document.addEventListener('click', function (ev) {
        var a = ev.target.closest ? ev.target.closest('a[href^="#"]') : null;
        if (!a || window.IS_STANDALONE) return;
        var id = a.getAttribute('href').slice(1);
        if (!id) return;
        var el = document.getElementById(id);
        if (el) {
          ev.preventDefault();
          el.scrollIntoView({ behavior: 'smooth' });
          if (history.replaceState) history.replaceState(null, '', '#' + id);
        }
      });

      /* Standalone: любые ссылки с data-nav (карточки главной, назад/далее) → hash-маршрут */
      document.addEventListener('click', function (ev) {
        if (!window.IS_STANDALONE) return;
        var a = ev.target.closest ? ev.target.closest('a[data-nav]') : null;
        if (!a) return;
        ev.preventDefault();
        window.location.hash = '#/' + a.getAttribute('data-nav');
      });

      document.addEventListener('ac:progress', refreshNavStates);
    }
  };
})();
