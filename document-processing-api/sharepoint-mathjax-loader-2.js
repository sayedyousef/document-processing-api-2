/*
 * mathjax loader - handles equation rendering on SP pages
 * put this in Site Assets, add via Script Editor / Content Editor web part
 * content needs to be inside <div id="mathjax-content">...</div>
 */
(function () {

    // bail out if we're in edit mode - mathjax messes up the SP editor
    var qs = window.location.search.toLowerCase();
    var h = window.location.hash.toLowerCase();
    if (qs.indexOf('mode=edit') > -1 || h.indexOf('mode=edit') > -1) return;
    if (document.querySelector('.sp-pageLayout-editMode')) return;
    if (document.querySelector('#spPageCanvasContent [contenteditable="true"]')) return;
    if (document.querySelector('#MSOLayout_InDesignMode')) return;
    if (window._spPageContextInfo && window._spPageContextInfo.isEditMode) return;
    if (document.getElementById('MSOSPWebPartManager_DisplayModeName')) return;

    function go() {
        var wrap = document.getElementById('mathjax-content');
        var target = wrap ? '#mathjax-content' : null;

        // no wrapper div? check if there's even any latex on the page
        if (!target) {
            var html = document.body ? document.body.innerHTML : '';
            if (html.indexOf('\\(') === -1 && html.indexOf('\\[') === -1) return;
        }

        // don't run inside contenteditable (SP puts content there when editing)
        if (wrap) {
            var p = wrap.parentElement;
            while (p) {
                if (p.getAttribute && p.getAttribute('contenteditable') === 'true') return;
                p = p.parentElement;
            }
        }

        window.MathJax = {
            tex: {
                inlineMath: [['\\(', '\\)']],
                displayMath: [['\\[', '\\]']]
            },
            options: {
                elements: target ? [target] : null,
                ignoreHtmlClass: 'sp-.*|ms-.*|od-.*|canvasTextArea',
                processHtmlClass: 'mathjax-content'
            },
            svg: { fontCache: 'global' },
            startup: {
                ready: function () {
                    MathJax.startup.defaultReady();
                }
            }
        };

        var s = document.createElement('script');
        s.src = 'https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js';
        s.async = true;
        document.head.appendChild(s);
    }

    if (document.readyState === 'loading')
        document.addEventListener('DOMContentLoaded', go);
    else
        go();

})();
