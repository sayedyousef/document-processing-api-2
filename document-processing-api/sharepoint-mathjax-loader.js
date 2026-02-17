(function () {

    // Skip in SharePoint edit mode
    var qs = window.location.search.toLowerCase();
    var h = window.location.hash.toLowerCase();
    if (qs.indexOf('mode=edit') > -1 || h.indexOf('mode=edit') > -1) return;
    if (document.querySelector('.sp-pageLayout-editMode')) return;
    if (document.querySelector('#spPageCanvasContent [contenteditable="true"]')) return;
    var mso = document.querySelector('#MSOLayout_InDesignMode');
    if (mso && mso.value === '1') return;
    if (window._spPageContextInfo && window._spPageContextInfo.isEditMode) return;
    var dm = document.getElementById('MSOSPWebPartManager_DisplayModeName');
    if (dm && dm.value === 'Design') return;

    function go() {
        var wrap = document.getElementById('mathjax-content');
        var target = wrap ? '#mathjax-content' : null;

        // if no wrapper div, look for latex
        if (!target) {
            var html = document.body ? document.body.innerHTML : '';
            if (html.indexOf('\\(') === -1 && html.indexOf('\\[') === -1) return;
        }

        // don't run inside sp content editable
        if (wrap) {
            var p = wrap.parentElement;
            while (p) {
                if (p.getAttribute && p.getAttribute('contenteditable') === 'true') return;
                p = p.parentElement;
            }
        }

        // Inject styles for mjx-container (native MathML output)
        var style = document.createElement('style');
        style.textContent = 'mjx-container{display:inline}mjx-container[display="block"]{display:block;text-align:center;margin:1em 0}';
        document.head.appendChild(style);

        window.MathJax = {
            loader: {load: ['input/tex']},
            tex: {
                inlineMath: [['\\(', '\\)']],
                displayMath: [['\\[', '\\]']]
            },
            options: {
                ignoreHtmlClass: 'sp-.*|od-.*|canvasTextArea',
                processHtmlClass: 'mathjax-content',
                renderActions: {
                    assistiveMml: [],
                    typeset: [150,
                        function(doc) { for (var math of doc.math) MathJax.config.renderMathML(math, doc); },
                        function(math, doc) { MathJax.config.renderMathML(math, doc); }
                    ]
                }
            },
            startup: {
                elements: target ? [target] : null,
                pageReady: function() {
                    return MathJax.startup.document.render();
                }
            },
            renderMathML: function(math, doc) {
                math.typesetRoot = document.createElement('mjx-container');
                var mml = MathJax.startup.toMML(math.root);
                // Strip invisible Unicode operators (entity and literal forms)
                // Strip invisible operators, zero-width chars, bidi marks, BOM
                mml = mml.replace(/[\u2060-\u2064\u200B-\u200F\u061C\u202A-\u202C\u2066-\u2069\uFEFF]/g, '');
                mml = mml.replace(/&#x(206[0-9a-f]|200[b-f]|061c|202[a-c]|feff);/gi, '');
                mml = mml.replace(/<mo[^>]*>\s*<\/mo>/g, '');
                mml = mml.replace(/ data-[a-z-]+="[^"]*"/g, '');
                // Collapse <msup><mi></mi><mo>X</mo></msup> → <mo>X</mo>
                // MathJax empty-base prime pattern causes browser to insert invisible chars
                mml = mml.replace(/<msup>\s*<mi\s*\/?\s*>\s*(<\/mi>)?\s*(<mo[^>]*>[^<]*<\/mo>)\s*<\/msup>/g, '$2');
                math.typesetRoot.innerHTML = mml;
                if (math.display) math.typesetRoot.setAttribute('display', 'block');
            }
        };

        // Load MathJax 4 (native MathML output)
        var s = document.createElement('script');
        s.src = 'https://cdn.jsdelivr.net/npm/mathjax@4/startup.js';
        s.async = true;
        document.head.appendChild(s);
    }

    // Clean clipboard: browser's native MathML renderer adds invisible
    // Unicode operators during rendering; strip them on Ctrl+C / copy
    var INVISIBLE_RE = /[\u2060-\u2064\u200B-\u200F\u061C\u202A-\u202C\u2066-\u2069\uFEFF]/g;
    document.addEventListener('copy', function(e) {
        var sel = window.getSelection();
        if (!sel || !sel.rangeCount || !sel.toString()) return;
        var range = sel.getRangeAt(0);
        var mathEls = document.querySelectorAll('mjx-container, math');
        var touchesMath = false;
        for (var i = 0; i < mathEls.length; i++) {
            if (range.intersectsNode(mathEls[i])) { touchesMath = true; break; }
        }
        if (!touchesMath) return;
        var text = sel.toString().replace(INVISIBLE_RE, '');
        e.clipboardData.setData('text/plain', text);
        var div = document.createElement('div');
        div.appendChild(range.cloneContents());
        var html = div.innerHTML.replace(INVISIBLE_RE, '');
        e.clipboardData.setData('text/html', html);
        e.preventDefault();
    });

    if (document.readyState === 'loading')
        document.addEventListener('DOMContentLoaded', go);
    else
        go();

})();
