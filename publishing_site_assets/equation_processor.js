/**
 * equation_processor.js
 *
 * CRITICAL COMPONENT: Post-processes HTML to convert equation markers to proper HTML elements
 * This enables LaTeX equations to survive Word-to-HTML conversion and render with MathJax/KaTeX
 *
 * Author: Document Processing API Team
 * Version: 1.0.0
 *
 * Usage:
 * 1. Include this script in your HTML after body content
 * 2. It automatically processes markers on page load
 * 3. Markers are converted to span/div elements with appropriate classes
 */

(function() {
    'use strict';

    /**
     * Main function to process all equation markers in the document
     * Converts:
     * - MATHSTARTINLINE...\)MATHENDINLINE → <span class="inlineMath">...</span>
     * - MATHSTARTDISPLAY...\]MATHENDDISPLAY → <div class="Math_box">...</div>
     */
    function processEquationMarkers() {
        console.log('🔧 Processing equation markers...');

        let content = document.body.innerHTML;
        let inlineCount = 0;
        let displayCount = 0;

        // Process inline equations
        content = content.replace(
            /MATHSTARTINLINE(\\?\(.*?\\?\))MATHENDINLINE/g,
            function(match, equation) {
                inlineCount++;
                return `<span class="inlineMath" data-equation-id="inline-${inlineCount}">${equation}</span>`;
            }
        );

        // Process display equations
        content = content.replace(
            /MATHSTARTDISPLAY(\\?\[.*?\\?\])MATHENDDISPLAY/g,
            function(match, equation) {
                displayCount++;
                return `<div class="Math_box" data-equation-id="display-${displayCount}">${equation}</div>`;
            }
        );

        // Update the document
        document.body.innerHTML = content;

        console.log(`✅ Processed ${inlineCount} inline equations`);
        console.log(`✅ Processed ${displayCount} display equations`);
        console.log(`✅ Total equations: ${inlineCount + displayCount}`);

        // Add CSS if not already present
        addEquationStyles();

        // Trigger MathJax if available
        if (window.MathJax && MathJax.typesetPromise) {
            console.log('🎯 Triggering MathJax rendering...');
            MathJax.typesetPromise().then(() => {
                console.log('✅ MathJax rendering complete');
            }).catch((e) => console.error('❌ MathJax error:', e));
        }

        // Trigger KaTeX if available
        if (window.katex && window.renderMathInElement) {
            console.log('🎯 Triggering KaTeX rendering...');
            renderMathInElement(document.body, {
                delimiters: [
                    {left: '\\[', right: '\\]', display: true},
                    {left: '\\(', right: '\\)', display: false}
                ]
            });
            console.log('✅ KaTeX rendering complete');
        }

        // Return statistics
        return {
            inline: inlineCount,
            display: displayCount,
            total: inlineCount + displayCount
        };
    }

    /**
     * Add default styles for equation containers
     */
    function addEquationStyles() {
        if (document.getElementById('equation-processor-styles')) return;

        const styles = `
            <style id="equation-processor-styles">
                /* Inline equations - flow with text */
                .inlineMath {
                    display: inline;
                    margin: 0 0.2em;
                }

                /* Display equations - centered blocks */
                .Math_box {
                    display: block;
                    text-align: center;
                    margin: 1em auto;
                    overflow-x: auto;
                    padding: 0.5em;
                    max-width: 100%;
                }

                /* Optional: Add subtle highlighting */
                .inlineMath:hover, .Math_box:hover {
                    background-color: rgba(255, 255, 0, 0.05);
                    transition: background-color 0.3s ease;
                }

                /* For responsive design */
                @media (max-width: 768px) {
                    .Math_box {
                        font-size: 0.9em;
                        padding: 0.3em;
                    }
                }

                /* Debug mode - shows equation IDs */
                .debug-equations .inlineMath::before,
                .debug-equations .Math_box::before {
                    content: "[" attr(data-equation-id) "] ";
                    color: #ff6b6b;
                    font-size: 0.7em;
                    font-family: monospace;
                }
            </style>
        `;

        document.head.insertAdjacentHTML('beforeend', styles);
    }

    /**
     * Get statistics about equations in the document
     */
    function getEquationStats() {
        const inline = document.querySelectorAll('.inlineMath');
        const display = document.querySelectorAll('.Math_box');

        return {
            inline: inline.length,
            display: display.length,
            total: inline.length + display.length,
            elements: {
                inline: Array.from(inline),
                display: Array.from(display)
            }
        };
    }

    /**
     * Enable debug mode to show equation IDs
     */
    function enableDebugMode() {
        document.body.classList.add('debug-equations');
        console.log('🐛 Debug mode enabled - equation IDs are now visible');
    }

    /**
     * Disable debug mode
     */
    function disableDebugMode() {
        document.body.classList.remove('debug-equations');
        console.log('Debug mode disabled');
    }

    /**
     * Manually trigger equation processing
     * Useful if content is loaded dynamically
     */
    function reprocessEquations() {
        console.log('♻️ Reprocessing equations...');
        return processEquationMarkers();
    }

    // Export functions to global scope
    window.EquationProcessor = {
        process: processEquationMarkers,
        reprocess: reprocessEquations,
        getStats: getEquationStats,
        enableDebug: enableDebugMode,
        disableDebug: disableDebugMode,
        version: '1.0.0'
    };

    // Auto-process on DOMContentLoaded
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function() {
            console.log('📄 Document ready - processing equations');
            processEquationMarkers();
        });
    } else {
        // DOM already loaded
        console.log('📄 Document already loaded - processing equations immediately');
        processEquationMarkers();
    }

    // Handle dynamic content (for SPAs)
    let observer = null;
    window.EquationProcessor.watchForChanges = function() {
        if (observer) return;

        observer = new MutationObserver(function(mutations) {
            let shouldReprocess = false;
            mutations.forEach(function(mutation) {
                if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
                    // Check if new content contains markers
                    mutation.addedNodes.forEach(function(node) {
                        if (node.nodeType === 1 && node.innerHTML &&
                            (node.innerHTML.includes('MATHSTART') ||
                             node.innerHTML.includes('MATHEND'))) {
                            shouldReprocess = true;
                        }
                    });
                }
            });

            if (shouldReprocess) {
                console.log('🔄 New content with markers detected - reprocessing');
                processEquationMarkers();
            }
        });

        observer.observe(document.body, {
            childList: true,
            subtree: true
        });

        console.log('👀 Watching for dynamic content changes');
    };

    // Log initialization
    console.log('✨ Equation Processor v1.0.0 loaded');
    console.log('📖 Usage: window.EquationProcessor.process()');
    console.log('📊 Stats: window.EquationProcessor.getStats()');

})();