/**
 * MathJax Equation Copy Menu
 *
 * Hover over any MathJax-rendered equation to see copy buttons.
 * Supports two copy formats:
 *   - LaTeX: copies the TeX source wrapped in \(...\) or \[...\] delimiters
 *   - MathML: copies the MathML markup via MathJax's internal serializer
 *
 * Compatible with MathJax 3.x/4.x (native MathML output via startup.js).
 * Automatically skips initialization in SharePoint edit mode.
 *
 * Usage:
 *   Include this script after MathJax in your HTML, or embed inline.
 *   The script waits for MathJax to finish typesetting before attaching handlers.
 */
(function () {
    'use strict';

    var activeMenu = null;
    var hideTimer = null;
    var retryCount = 0;

    /* ------------------------------------------------------------------ */
    /*  MathJax math items accessor (handles different MathJax 3 builds)  */
    /* ------------------------------------------------------------------ */
    function getMathItems() {
        try {
            var math = MathJax.startup.document.math;
            if (typeof math.toArray === 'function') return math.toArray();
            if (typeof Symbol !== 'undefined' && math[Symbol.iterator]) return Array.from(math);
            if (math._list) return math._list;
        } catch (e) {}
        return [];
    }

    /* ------------------------------------------------------------------ */
    /*  Initialization                                                     */
    /* ------------------------------------------------------------------ */
    function init() {
        // Skip in SharePoint edit mode
        var qs = window.location.search.toLowerCase();
        if (qs.indexOf('mode=edit') > -1) return;
        if (document.querySelector('.sp-pageLayout-editMode')) return;
        if (document.querySelector('#spPageCanvasContent [contenteditable="true"]')) return;

        retryCount++;

        // Wait for MathJax to load (retry up to 30 seconds)
        if (!window.MathJax || !MathJax.startup || !MathJax.startup.promise) {
            if (retryCount < 60) setTimeout(init, 500);
            return;
        }

        // Wait for MathJax to finish typesetting
        MathJax.startup.promise
            .then(function () { setup(); })
            .catch(function () { setup(); });
    }

    function setup() {
        // Inject styles for the copy menu
        var s = document.createElement('style');
        s.textContent =
            '.math-copy-btns{position:absolute;z-index:10000;display:flex;direction:rtl;gap:3px;' +
            'background:#fff;border:1px solid #cbd5e0;border-radius:6px;padding:3px;' +
            'box-shadow:0 2px 8px rgba(0,0,0,.12);font-family:"Segoe UI",sans-serif}' +
            '.math-copy-btns button{padding:3px 8px;border:none;border-radius:4px;' +
            'background:#edf2f7;color:#2d3748;font-size:11px;cursor:pointer;transition:all .12s}' +
            '.math-copy-btns button:hover{background:#3182ce;color:#fff}' +
            '.math-copy-btns button.ok{background:#38a169;color:#fff}';
        document.head.appendChild(s);

        // Attach hover handlers to all rendered equation containers
        var containers = document.querySelectorAll('mjx-container');
        for (var i = 0; i < containers.length; i++) {
            containers[i].style.cursor = 'pointer';
            containers[i].addEventListener('mouseenter', onEnter);
            containers[i].addEventListener('mouseleave', onLeave);
        }

        console.log('[MathCopy] ' + containers.length + ' equations ready');
    }

    /* ------------------------------------------------------------------ */
    /*  Find the MathJax math item for a given mjx-container element      */
    /* ------------------------------------------------------------------ */
    function findItem(el) {
        var items = getMathItems();
        for (var i = 0; i < items.length; i++) {
            if (items[i].typesetRoot === el) return items[i];
        }
        return null;
    }

    /* ------------------------------------------------------------------ */
    /*  Hover handlers                                                     */
    /* ------------------------------------------------------------------ */
    function onEnter() {
        var el = this;
        clearTimeout(hideTimer);
        if (activeMenu && activeMenu._target === el) return;
        removeMenu();
        showBtns(el);
    }

    function onLeave() {
        hideTimer = setTimeout(function () {
            if (activeMenu && !activeMenu.matches(':hover')) removeMenu();
        }, 300);
    }

    /* ------------------------------------------------------------------ */
    /*  Copy menu UI                                                       */
    /* ------------------------------------------------------------------ */
    function showBtns(el) {
        var div = document.createElement('div');
        div.className = 'math-copy-btns';
        div._target = el;

        // Arabic labels: "نسخ LaTeX" and "نسخ MathML"
        div.appendChild(makeBtn('\u0646\u0633\u062E LaTeX', function () { doCopy(el, 'latex', this); }));
        div.appendChild(makeBtn('\u0646\u0633\u062E MathML', function () { doCopy(el, 'mathml', this); }));

        // Keep menu visible while hovering over it
        div.addEventListener('mouseenter', function () { clearTimeout(hideTimer); });
        div.addEventListener('mouseleave', function () {
            hideTimer = setTimeout(removeMenu, 300);
        });

        document.body.appendChild(div);
        activeMenu = div;

        // Position above the equation, centered horizontally
        var rect = el.getBoundingClientRect();
        var menuW = div.offsetWidth;
        var menuH = div.offsetHeight;
        var top = rect.top + window.scrollY - menuH - 4;
        var left = rect.left + window.scrollX + (rect.width - menuW) / 2;
        left = Math.max(4, Math.min(left, window.innerWidth - menuW - 4));
        if (top < window.scrollY + 4) top = rect.bottom + window.scrollY + 4;
        div.style.top = top + 'px';
        div.style.left = left + 'px';
    }

    function makeBtn(label, handler) {
        var b = document.createElement('button');
        b.textContent = label;
        b.addEventListener('click', function (e) { e.stopPropagation(); handler.call(this); });
        return b;
    }

    function removeMenu() {
        if (activeMenu) { activeMenu.remove(); activeMenu = null; }
    }

    /* ------------------------------------------------------------------ */
    /*  Copy logic                                                         */
    /* ------------------------------------------------------------------ */
    function doCopy(el, fmt, btnEl) {
        var item = findItem(el);
        var text = '';

        if (fmt === 'latex') {
            var latex = item ? item.math : '';
            if (!latex) {
                // "فارغ!" = Empty!
                showFeedback(btnEl, '\u0641\u0627\u0631\u063A!', '#e53e3e', '\u0646\u0633\u062E LaTeX');
                return;
            }
            text = item.display ? '\\[' + latex + '\\]' : '\\(' + latex + '\\)';
        } else {
            text = toMathML(item);
            if (!text) {
                // "خطأ!" = Error!
                showFeedback(btnEl, '\u062E\u0637\u0623!', '#e53e3e', '\u0646\u0633\u062E MathML');
                return;
            }
        }

        var p = (navigator.clipboard && navigator.clipboard.writeText)
            ? navigator.clipboard.writeText(text)
            : fallbackCopy(text);

        // "تم النسخ!" = Copied!
        p.then(function () {
            showFeedback(btnEl, '\u062A\u0645 \u0627\u0644\u0646\u0633\u062E!', '#38a169',
                fmt === 'latex' ? '\u0646\u0633\u062E LaTeX' : '\u0646\u0633\u062E MathML');
        });
    }

    /* ------------------------------------------------------------------ */
    /*  MathML extraction                                                  */
    /* ------------------------------------------------------------------ */
    function toMathML(item) {
        if (!item) return '';

        // With native MathML output, the <math> element is in the DOM
        if (item.typesetRoot) {
            var mathEl = item.typesetRoot.querySelector('math');
            if (mathEl) return mathEl.outerHTML;
        }

        // Fallback: MathJax serializer
        try {
            var r = MathJax.startup.toMML(item.root);
            if (r) return r;
        } catch (e) {}

        return '';
    }

    /* ------------------------------------------------------------------ */
    /*  Helpers                                                            */
    /* ------------------------------------------------------------------ */
    function showFeedback(btnEl, msg, color, restoreLabel) {
        btnEl.textContent = msg;
        btnEl.style.background = color;
        btnEl.style.color = '#fff';
        setTimeout(function () {
            btnEl.textContent = restoreLabel;
            btnEl.style.background = '';
            btnEl.style.color = '';
        }, 1000);
    }

    function fallbackCopy(text) {
        return new Promise(function (resolve) {
            var ta = document.createElement('textarea');
            ta.value = text;
            ta.style.cssText = 'position:fixed;left:-9999px;opacity:0';
            document.body.appendChild(ta);
            ta.select();
            try { document.execCommand('copy'); } catch (e) {}
            document.body.removeChild(ta);
            resolve();
        });
    }

    /* ------------------------------------------------------------------ */
    /*  Bootstrap                                                          */
    /* ------------------------------------------------------------------ */
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();
