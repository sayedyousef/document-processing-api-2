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
                mml = mml.replace(/&#x206[1-3];/gi, '');
                mml = mml.replace(/[\u2061-\u2063]/g, '');
                mml = mml.replace(/<mo[^>]*>\s*<\/mo>/g, '');
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

    if (document.readyState === 'loading')
        document.addEventListener('DOMContentLoaded', go);
    else
        go();

})();
