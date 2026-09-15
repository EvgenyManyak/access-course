/* ============================================================
   quiz.js — тесты по модулям
   Тест задаётся декларативно в HTML модуля:
     <section class="quiz" data-quiz="m03">
       <script type="application/json">
         { "title": "Тест: …", "passScore": 80,
           "questions": [
             { "q": "Вопрос?", "multi": false,
               "options": ["a", "b", "c", "d"],
               "correct": [1],
               "explain": "Пояснение…" } ] }
       <\/script>
     </section>
   quiz.js рендерит интерфейс, проверяет ответы, показывает
   пояснения, сохраняет лучший результат в progress.js.
   init() можно вызывать повторно (роутер standalone-версии).
   ВАЖНО: этот файл может встраиваться inline в standalone.html,
   поэтому в комментариях запрещён литерал закрывающего script-тега
   (см. ниже — он записан с экранированной косой чертой).
   ============================================================ */
(function () {
  'use strict';
  var AC = window.AC;

  function esc(s) { return AC.escapeHtml(s); }

  function renderQuiz(section) {
    var script = section.querySelector('script[type="application/json"]');
    if (!script) return;
    var data;
    try {
      data = JSON.parse(script.textContent);
    } catch (e) {
      section.innerHTML = '<div class="callout callout--error"><div class="callout__title">Ошибка теста</div>' +
        '<p>Не удалось прочитать данные теста. Пожалуйста, сообщите авторам учебника.</p></div>';
      return;
    }

    var quizId = section.getAttribute('data-quiz') || AC.getCurrentId();
    var passScore = typeof data.passScore === 'number' ? data.passScore : 80;
    var questions = data.questions || [];

    function renderState(savedMsg) {
      var html = '<div class="quiz__head">' +
        '<div><div class="quiz__title">📝 ' + esc(data.title || 'Тест по модулю') + '</div>' +
        '<div class="quiz__hint">Вопросов: ' + questions.length + ' · Для зачёта нужно ' + passScore + '% · ' +
        'Вопросы могут иметь несколько верных ответов</div></div></div>';
      html += '<form class="quiz__form" novalidate>';

      questions.forEach(function (q, qi) {
        var multi = q.multi || (Array.isArray(q.correct) && q.correct.length > 1);
        html += '<div class="quiz__q" data-qi="' + qi + '">' +
          '<div class="quiz__q-text"><span class="quiz__q-num">' + (qi + 1) + '</span><span>' + esc(q.q) +
          (multi ? ' <span class="muted small">(несколько ответов)</span>' : '') + '</span></div>' +
          '<div class="quiz__opts">';
        (q.options || []).forEach(function (opt, oi) {
          html += '<label class="quiz__opt" data-oi="' + oi + '">' +
            '<input type="' + (multi ? 'checkbox' : 'radio') + '" name="q-' + quizId + '-' + qi + '" value="' + oi + '">' +
            '<span>' + esc(opt) + '</span></label>';
        });
        html += '</div><div class="quiz__explain" hidden></div></div>';
      });

      html += '<div class="quiz__actions">' +
        '<button type="submit" class="btn">Проверить ответы</button>' +
        '<button type="button" class="btn btn--ghost" data-quiz-reset>Пройти заново</button>' +
        '<span class="quiz__result" aria-live="polite">' + (savedMsg || '') + '</span>' +
        '</div></form>';

      section.innerHTML = html;

      var form = section.querySelector('.quiz__form');
      form.addEventListener('submit', function (ev) { ev.preventDefault(); check(); });
      form.querySelector('[data-quiz-reset]').addEventListener('click', function () { renderState(); });
    }

    function check() {
      var total = questions.length;
      var correctCount = 0;

      questions.forEach(function (q, qi) {
        var qEl = section.querySelector('.quiz__q[data-qi="' + qi + '"]');
        var inputs = Array.prototype.slice.call(qEl.querySelectorAll('input'));
        var chosen = inputs.filter(function (i) { return i.checked; }).map(function (i) { return parseInt(i.value, 10); });
        var correct = q.correct || [];
        var isCorrect = chosen.length === correct.length && correct.every(function (c) { return chosen.indexOf(c) !== -1; });
        if (isCorrect) correctCount++;

        qEl.classList.remove('is-answered-correctly', 'is-answered-wrongly');
        qEl.classList.add(isCorrect ? 'is-answered-correctly' : 'is-answered-wrongly');

        qEl.querySelectorAll('.quiz__opt').forEach(function (lab) {
          var oi = parseInt(lab.getAttribute('data-oi'), 10);
          lab.classList.remove('quiz__opt--correct', 'quiz__opt--wrong');
          if (correct.indexOf(oi) !== -1) lab.classList.add('quiz__opt--correct');
          else if (chosen.indexOf(oi) !== -1 && !isCorrect) lab.classList.add('quiz__opt--wrong');
        });

        var explainEl = qEl.querySelector('.quiz__explain');
        var msg = (isCorrect ? '✓ Верно. ' : '✗ Неверно. ') + (q.explain || '');
        explainEl.textContent = msg;
        explainEl.hidden = false;
      });

      var pct = Math.round(correctCount / total * 100);
      var passed = pct >= passScore;
      var resultEl = section.querySelector('.quiz__result');
      resultEl.textContent = 'Результат: ' + correctCount + ' из ' + total + ' (' + pct + '%) — ' +
        (passed ? 'зачёт ✓' : 'не достигнут порог ' + passScore + '%');
      resultEl.className = 'quiz__result ' + (passed ? 'is-pass' : 'is-fail');

      AC.progress.saveQuiz(quizId, correctCount, total);
      AC.toast(passed ? 'Тест сдан: ' + pct + '%' : 'Тест пройден на ' + pct + '% — загляните в пояснения');

      resultEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    var best = AC.progress.quizBest(quizId);
    renderState(best ? ('Лучший результат: ' + best.score + ' из ' + best.total) : '');
  }

  AC.quiz = {
    init: function () {
      AC.$$('.quiz').forEach(renderQuiz);
    }
  };
})();
