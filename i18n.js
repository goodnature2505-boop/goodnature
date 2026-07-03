/* GoodNature i18n — KO/EN language switch.
   Most text uses CSS-visibility twins: <span lang="ko">..</span><span lang="en">..</span>
   (handled purely by CSS via html[lang]). This script additionally:
   - toggles <html lang> on switch + persists the choice (localStorage)
   - swaps <title> / <meta description> (data-title-en / data-desc-en on <html>)
   - swaps textContent for elements that can't hold child spans, e.g. <option>
     (data-en attribute) and input/textarea placeholders (data-ph-en). */
(function () {
  var KEY = 'gn_lang';
  var html = document.documentElement;
  var titleEn = html.getAttribute('data-title-en');
  var titleKo = document.title;
  var descEl = document.querySelector('meta[name="description"]');
  var descKo = descEl ? descEl.getAttribute('content') : null;
  var descEn = html.getAttribute('data-desc-en');

  // Cache originals for attribute/text swaps (elements without CSS twins).
  var textEls = [].slice.call(document.querySelectorAll('[data-en]'));
  textEls.forEach(function (el) { el._ko = el.textContent; });
  var phEls = [].slice.call(document.querySelectorAll('[data-ph-en]'));
  phEls.forEach(function (el) { el._phKo = el.getAttribute('placeholder') || ''; });

  function setLang(lang) {
    if (lang !== 'en') lang = 'ko';
    html.setAttribute('lang', lang);

    if (lang === 'en') {
      if (titleEn) document.title = titleEn;
      if (descEl && descEn) descEl.setAttribute('content', descEn);
      textEls.forEach(function (el) { el.textContent = el.getAttribute('data-en'); });
      phEls.forEach(function (el) { el.setAttribute('placeholder', el.getAttribute('data-ph-en')); });
    } else {
      document.title = titleKo;
      if (descEl && descKo != null) descEl.setAttribute('content', descKo);
      textEls.forEach(function (el) { el.textContent = el._ko; });
      phEls.forEach(function (el) { el.setAttribute('placeholder', el._phKo); });
    }

    var btns = document.querySelectorAll('.lang-switch button');
    for (var i = 0; i < btns.length; i++) {
      btns[i].classList.toggle('active', btns[i].getAttribute('data-lang') === lang);
    }
    try { localStorage.setItem(KEY, lang); } catch (e) {}
  }

  document.addEventListener('click', function (e) {
    var b = e.target.closest ? e.target.closest('.lang-switch button') : null;
    if (b) { e.preventDefault(); setLang(b.getAttribute('data-lang')); }
  });

  var saved = null;
  try { saved = localStorage.getItem(KEY); } catch (e) {}
  setLang(saved === 'en' ? 'en' : 'ko');
})();
