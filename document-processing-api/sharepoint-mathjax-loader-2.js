(function () {

    // handed SP edit mode 
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

        window.MathJax = {
            tex: {
                inlineMath: [['\\(', '\\)']],
                displayMath: [['\\[', '\\]']]
            },
            options: {
                ignoreHtmlClass: 'sp-.*|ms-.*|od-.*|canvasTextArea',
                processHtmlClass: 'mathjax-content',
                enableMenu: false
            },
            startup: {
                elements: target ? [target] : null,
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
