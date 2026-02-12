/**
 * MathJax Loader for SharePoint
 *
 * Upload this file to SharePoint (e.g., Site Assets library) and reference it
 * via a Script Editor web part, Content Editor web part, or SPFx extension.
 *
 * What it does:
 *   1. Detects if the page is in edit mode — if so, skips MathJax entirely
 *   2. Scopes MathJax to only scan elements with id="mathjax-content"
 *   3. Extra safety: won't run if content is inside a contenteditable area
 *   4. Renders LaTeX equations: \(...\) for inline, \[...\] for display math
 *
 * Usage:
 *   - Wrap your HTML content in: <div id="mathjax-content">...</div>
 *   - Add this script to the page via Script Editor or Content Editor web part
 */
(function () {
    'use strict';

    // ── 1. Detect SharePoint edit mode ──────────────────────────────
    var search = window.location.search.toLowerCase();
    var hash = window.location.hash.toLowerCase();

    var isEditMode = (
        // SharePoint Online modern pages
        search.indexOf('mode=edit') !== -1 ||
        hash.indexOf('mode=edit') !== -1 ||
        // SharePoint Online modern page edit indicators
        document.querySelector('.sp-pageLayout-editMode') !== null ||
        document.querySelector('#spPageCanvasContent [contenteditable="true"]') !== null ||
        // Classic SharePoint edit mode
        document.querySelector('#MSOLayout_InDesignMode') !== null ||
        (typeof window._spPageContextInfo !== 'undefined' &&
            window._spPageContextInfo.isEditMode === true) ||
        // SharePoint designer
        document.getElementById('MSOSPWebPartManager_DisplayModeName') !== null
    );

    if (isEditMode) {
        console.log('[MathJax Loader] Edit mode detected — MathJax will NOT load.');
        return;
    }

    // ── 2. Wait for content to be ready ─────────────────────────────
    function initMathJax() {
        var contentEl = document.getElementById('mathjax-content');

        // If no mathjax-content wrapper found, try to find any content with LaTeX
        var targetSelector = contentEl ? '#mathjax-content' : null;

        if (!targetSelector) {
            // Fallback: look for LaTeX delimiters anywhere in the page body
            var bodyText = document.body ? document.body.innerHTML : '';
            if (bodyText.indexOf('\\(') === -1 && bodyText.indexOf('\\[') === -1) {
                console.log('[MathJax Loader] No LaTeX content found — skipping.');
                return;
            }
            // If LaTeX exists but no wrapper, use body but log a warning
            console.warn('[MathJax Loader] No #mathjax-content div found. Scanning document body. ' +
                'For best results, wrap content in <div id="mathjax-content">...</div>');
        }

        // ── 3. Extra safety: skip if inside contenteditable ─────────
        if (contentEl) {
            var parent = contentEl.parentElement;
            while (parent) {
                if (parent.getAttribute && parent.getAttribute('contenteditable') === 'true') {
                    console.log('[MathJax Loader] Content is inside contenteditable — skipping.');
                    return;
                }
                parent = parent.parentElement;
            }
        }

        // ── 4. Configure MathJax ────────────────────────────────────
        window.MathJax = {
            tex: {
                inlineMath: [['\\(', '\\)']],
                displayMath: [['\\[', '\\]']]
            },
            options: {
                // Only scan our content div (not SharePoint's UI)
                elements: targetSelector ? [targetSelector] : null,
                // Skip elements that SharePoint uses for editing
                ignoreHtmlClass: 'sp-.*|ms-.*|od-.*|canvasTextArea',
                processHtmlClass: 'mathjax-content'
            },
            svg: { fontCache: 'global' },
            startup: {
                ready: function () {
                    console.log('[MathJax Loader] MathJax ready — rendering equations.');
                    MathJax.startup.defaultReady();
                }
            }
        };

        // ── 5. Load MathJax from CDN ────────────────────────────────
        var script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js';
        script.async = true;
        script.onload = function () {
            console.log('[MathJax Loader] MathJax loaded successfully.');
        };
        script.onerror = function () {
            console.error('[MathJax Loader] Failed to load MathJax from CDN.');
        };
        document.head.appendChild(script);
    }

    // ── 6. Run when DOM is ready ────────────────────────────────────
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initMathJax);
    } else {
        // DOM already loaded (script loaded async or deferred)
        initMathJax();
    }

})();
