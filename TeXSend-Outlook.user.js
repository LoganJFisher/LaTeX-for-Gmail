// ==UserScript==
// @name            TeXSend-Outlook
// @version         1.0.0
// @description     LaTeX compiling for Outlook
// @author          Logan J. Fisher & GTK & MistralMireille
// @license         MIT
// @namespace       https://github.com/LoganJFisher/TeXSend/
// @downloadURL     https://raw.githubusercontent.com/LoganJFisher/TeXSend/refs/heads/main/TeXSend-Outlook.user.js
// @updateURL       https://raw.githubusercontent.com/LoganJFisher/TeXSend/refs/heads/main/TeXSend-Outlook.user.js
// @supportURL      https://github.com/LoganJFisher/TeXSend/issues
// @match           *://outlook.live.com/mail/*
// @noframes
// @grant           GM_registerMenuCommand
// @grant           GM_addStyle
// @grant           GM_getValue
// @grant           GM_setValue
// @grant           GM_getResourceText
// @require         https://cdn.jsdelivr.net/npm/katex@0.16.22/dist/katex.min.js
// @require         https://cdn.jsdelivr.net/npm/katex@0.16.22/dist/contrib/mhchem.min.js
// @require         https://cdn.jsdelivr.net/npm/katex@0.16.22/dist/contrib/copy-tex.min.js
// @resource css    https://gist.github.com/raw/8272456e1ead908609004a31bd4e59c2/katex.inline.woff2.min.css
// ==/UserScript==

/* globals katex */

// ===================================================================================================
// Constants
// ===================================================================================================

const selectors = {
    messageBody: "div[id^=UniqueMessageBody]",
    draftBody: "div[id^=editorParent] > div[role=textbox]",
    discardButton: "button#discardCompose",
    bottomBar: "div#BottomBar",
    newMailButton: "button[aria-label='New mail']",
};

const BASE_DELIMITERS = [
    {left: '$$' , right: '$$' , display: true, includeDelimiter: false},
    {left: '$' , right: '$' , display: false, includeDelimiter: false},
    {left: '[(;' , right: ';)]' , display: true, includeDelimiter: false},
    {left: '\\[' , right: '\\]' , display: true, includeDelimiter: false},
    {left: '[;' , right: ';]' , display: false, includeDelimiter: false},
    {left: '\\(' , right: '\\)' , display: false, includeDelimiter: false},
    {left: '\\begin{displaymath}' , right: '\\end{displaymath}' , display: true, includeDelimiter: false},
    {left: '\\begin{math}' , right: '\\end{math}', display: false, includeDelimiter: false},

    {left: '\\begin{align}' , right: '\\end{align}', display: true, includeDelimiter: true},
    {left: '\\begin{align*}' , right: '\\end{align*}', display: true, includeDelimiter: true},
    {left: '\\begin{aligned}' , right: '\\end{aligned}', display: true, includeDelimiter: true},
    {left: '\\begin{alignat}' , right: '\\end{alignat}', display: true, includeDelimiter: true},
    {left: '\\begin{alignat*}' , right: '\\end{alignat*}', display: true, includeDelimiter: true},
    {left: '\\begin{alignedat}' , right: '\\end{alignedat}', display: true, includeDelimiter: true},
    {left: '\\begin{array}' , right: '\\end{array}', display: true, includeDelimiter: true},
    {left: '\\begin{bmatrix}' , right: '\\end{bmatrix}', display: true, includeDelimiter: true},
    {left: '\\begin{bmatrix*}' , right: '\\end{bmatrix*}', display: true, includeDelimiter: true},
    {left: '\\begin{Bmatrix}' , right: '\\end{Bmatrix}', display: true, includeDelimiter: true},
    {left: '\\begin{Bmatrix*}' , right: '\\end{Bmatrix*}', display: true, includeDelimiter: true},
    {left: '\\begin{cases}' , right: '\\end{cases}', display: true, includeDelimiter: true},
    {left: '\\begin{CD}' , right: '\\end{CD}', display: true, includeDelimiter: true},
    {left: '\\begin{darray}' , right: '\\end{darray}', display: true, includeDelimiter: true},
    {left: '\\begin{drcases}' , right: '\\end{drcases}', display: true, includeDelimiter: true},
    {left: '\\begin{equation}' , right: '\\end{equation}', display: true, includeDelimiter: true},
    {left: '\\begin{equation*}' , right: '\\end{equation*}', display: true, includeDelimiter: true},
    {left: '\\begin{gather}' , right: '\\end{gather}', display: true, includeDelimiter: true},
    {left: '\\begin{gathered}' , right: '\\end{gathered}', display: true, includeDelimiter: true},
    {left: '\\begin{matrix}' , right: '\\end{matrix}', display: true, includeDelimiter: true},
    {left: '\\begin{matrix*}' , right: '\\end{matrix*}', display: true, includeDelimiter: true},
    {left: '\\begin{pmatrix}' , right: '\\end{pmatrix}', display: true, includeDelimiter: true},
    {left: '\\begin{pmatrix*}' , right: '\\end{pmatrix*}', display: true, includeDelimiter: true},
    {left: '\\begin{rcases}' , right: '\\end{rcases}', display: true, includeDelimiter: true},
    {left: '\\begin{smallmatrix}' , right: '\\end{smallmatrix}', display: false, includeDelimiter: true},
    {left: '\\begin{split}' , right: '\\end{split}', display: true, includeDelimiter: true},
    {left: '\\begin{subarray}' , right: '\\end{subarray}', display: true, includeDelimiter: true},
    {left: '\\begin{Vmatrix}' , right: '\\end{Vmatrix}', display: true, includeDelimiter: true},
    {left: '\\begin{Vmatrix*}' , right: '\\end{Vmatrix*}', display: true, includeDelimiter: true},
    {left: '\\begin{vmatrix}' , right: '\\end{vmatrix}', display: true, includeDelimiter: true},
    {left: '\\begin{vmatrix*}' , right: '\\end{vmatrix*}', display: true, includeDelimiter: true},
]

function buildRegex(delims) {
    const escape = string => string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const expressions = delims.map( d => {
        const display = d.display ? '(?<d>)' : '';
        const exp = d.includeDelimiter ? `(?<tex>${escape(d.left)}.+?${escape(d.right)})${display}` : `${escape(d.left)}(?<tex>.+?)${escape(d.right)}${display}`;
        return exp;
    })

    return new RegExp(expressions.join('|'), 'gs');
}

// ===============================================================================================
// PREFERENCES
// ===============================================================================================

let DOLLAR_DELIMITERS_ENABLED = GM_getValue("DOLLAR_DELIMITERS_ENABLED", true);
function getDelimiters() {
    return DOLLAR_DELIMITERS_ENABLED ? BASE_DELIMITERS : BASE_DELIMITERS.slice(2);
}

function toggleDollarDelimiters() {
    DOLLAR_DELIMITERS_ENABLED = !DOLLAR_DELIMITERS_ENABLED;
    GM_setValue("DOLLAR_DELIMITERS_ENABLED", DOLLAR_DELIMITERS_ENABLED);

    REGEX = buildRegex(getDelimiters());
    LATEX_CACHE.clear();
    refreshMessages();
    refreshDrafts();

    console.log(`Dollar sign delimiters ${DOLLAR_DELIMITERS_ENABLED ? 'enabled' : 'disabled'}`);
}

let MESSAGES_TOGGLE_DEFAULT = GM_getValue('MESSAGES_TOGGLE_DEFAULT', false);
function toggleDefaultMessagesState() {
    MESSAGES_TOGGLE_DEFAULT = !MESSAGES_TOGGLE_DEFAULT;
    GM_setValue('MESSAGES_TOGGLE_DEFAULT', MESSAGES_TOGGLE_DEFAULT);
    console.log(`Default message LaTeX rendering ${MESSAGES_TOGGLE_DEFAULT ? 'enabled' : 'disabled'}`);
}


// ===================================================================================================
// LATEX
// ===================================================================================================

let REGEX = buildRegex(getDelimiters());
const LATEX_CACHE = new Map();



function renderLatex(html) {
    let cached = LATEX_CACHE.get(html);
    if (cached) {
        return cached;
    }

    const div = document.createElement('div');
    let result = html.replace(REGEX, function() {
        const groups = arguments[arguments.length - 1];
        const display = groups.d !== undefined;
        div.innerHTML = POLICY.createHTML(groups.tex.replace(/&nbsp;/gs, '').trim());
        return katex.renderToString(div.textContent, { throwOnError: false, displayMode: display, trust: true, strict: false });
    })

    LATEX_CACHE.set(html, result);
    return result;
}


function updateLatex(messageList, state = true) {
    messageList.forEach(message => {
        let html = state ? renderLatex(message.innerHTML) : message.innerHTML;
        if (message.tex_overlay.innerHTML === html) return;

        message.tex_overlay.innerHTML = POLICY.createHTML(html);
    });
}

// ===============================================================================================
// MESSAGES
// ===============================================================================================

function insertTexButton(mailButton) {
    if (document.querySelector(".tex-button-main")) return;

    const btn = createTexButton(mailButton.className + " tex-button-main");
    btn.addEventListener("click", toggleMessages)
    document.querySelector(selectors.bottomBar)?.appendChild(btn);
}


function processMessage(message) {
    if (message.tex_overlay) return;

    let overlay = document.createElement("div");
    overlay.className = message.className;
    overlay.classList.add("tex-overlay");
    message.before(overlay);
    message.tex_overlay = overlay;
}

function refreshMessages(elements) {
    elements = elements || document.querySelectorAll(selectors.messageBody);
    updateLatex(elements, MESSAGES_TOGGLE);
}


let MESSAGES_TOGGLE = MESSAGES_TOGGLE_DEFAULT;
function toggleMessages() {
    MESSAGES_TOGGLE = !MESSAGES_TOGGLE;
    refreshMessages();
}


// ===============================================================================================
// DRAFTS
// ===============================================================================================


function addDraftPreview(draftBody) {
    const preview = document.createElement("div");
    preview.className = "draft-tex-preview";
    preview.style = draftBody.style.cssText

    const render = () => {
        if (preview.classList.contains("show")) {
            preview.style = draftBody.style.cssText;
            preview.innerHTML = POLICY.createHTML(renderLatex(draftBody.innerHTML));
        }
    }

    draftBody.render = render;
    draftBody.toggle = () => {
        preview.classList.toggle("show");
        render();
    }

    draftBody.addEventListener("keyup", debounce(render, 250));
    draftBody.after(preview);
}

function processDraft(draftBody) {
    const discardButton = document.querySelector(selectors.discardButton);

    if (draftBody.processed || !discardButton) return;
    draftBody.processed = true;

    const btn = createTexButton(discardButton.className + " tex-button-draft");
    btn.addEventListener("click", () => draftBody.toggle())
    discardButton.before(btn);

    draftBody.addEventListener('keyup', draftShortcutHandler, true);

    addDraftPreview(draftBody);
}

function refreshDrafts(elements) {
    elements = elements || document.querySelectorAll(selectors.draftBody);
    elements.forEach(elem => elem.render());
}

function draftShortcutHandler(event) {
    if (event.ctrlKey && event.altKey && event.keyCode === 75) {
        event.currentTarget?.toggle();
    }
}


// ===============================================================================================
// UTILS
// ===============================================================================================

function debounce(func, duration) {
    let timeout = null;

    return function() {
        clearTimeout(timeout);
        timeout = setTimeout(func, duration, ...arguments)
    }
}

function createTexButton(className) {
    let button = document.createElement("button");
    button.id = `tex-button-${Math.random().toString(16).slice(2)}`;
    button.className = className;
    button.title = "Toggle LaTeX";
    button.ariaLabel = "TeX";
    button.innerHTML = katex.renderToString('\\footnotesize \\TeX', { throwOnError: false });
    return button;
}

function addShortcuts() {
    const handler = event => {
        if (event.ctrlKey && event.altKey && event.keyCode === 76) {
            toggleMessages();
        }
    };

    document.addEventListener("keyup", handler, true);
}

function injectStyles() {
    GM_addStyle(GM_getResourceText("css"));

    GM_addStyle(`
        .tex-button-main {
            min-width: 40px;
            margin: 4px 12px;
            border-right: unset;
            border-radius: 4px;
            border: none;
        }

        .tex-button-draft {
            min-width: 30px;
        }

        .draft-tex-preview {
            display: none;
            background-color: inherit;
            overflow-wrap: break-word;
        }

        div:has(> .draft-tex-preview.show)::before {
            content: "";
            position: absolute;
            display: block;
            background: currentColor;
            width: 2px;
            height: 100%;
            top: 0;
            left: 50%;
            translate: -50% 0%;
        }

        .draft-tex-preview.show {
            display: block;
        }

        div:has(> .draft-tex-preview.show) {
            display: flex;
            flex-direction: row;
            --gap: 2%;
            gap: var(--gap);
            position: relative;
        }

        div:has(> .draft-tex-preview.show) > div {
            --width: calc(50% - var(--gap) / 2);
            flex: 0 0 var(--width);
            width: var(--width);
        }

        .tex-overlay, .draft-tex-preview {
            counter-reset: katexEqnNo mmlEqnNo;
        }

        .katex-display {
            max-width: 99%;
        }

        .tex-overlay {
            display: block;
        }

        ${selectors.messageBody} {
            display: none;
        }
    `)
}

// ===============================================================================================
// MAIN
// ===============================================================================================

const watch = (function() {
    const TARGETS = new Map();
    function _watch(selector, callback) {
        let callbacks = TARGETS.get(selector);
        if (!callbacks) {
            callbacks = new Set();
            TARGETS.set(selector, callbacks);
        }

        callbacks.add(callback);
        getAdded(selector);

        return function unwatch() {
            callbacks.delete(callback);
            !callbacks.size && TARGETS.delete(selector);
        }
    }

    function getAdded(selector) {
        let not_watched = `${selector.trim()}:not([data-watch])`;
        let elements = document.querySelectorAll(not_watched);
        elements.forEach(elem => {
            elem.dataset.watch = true
        });

        return elements;
    }

    function triggerCallbacks(callbacks, selector) {
        let added = getAdded(selector);
        added.length && callbacks.forEach(cb => cb(added));
    }

    const bodyObserver = new MutationObserver( mutations => {
        if (!mutations.some(m => m.addedNodes.length)) return;
        TARGETS.forEach(triggerCallbacks);
    });

    bodyObserver.observe(document.body, { childList: true, subtree: true })
    return _watch;
})();

const POLICY = (function() {
    let policy = {
        createHTML: string => string
    }

    if (window.trustedTypes) {
        try {
            policy = window.trustedTypes.createPolicy('dangerouslySetInnerHTMLPolicy', policy);
        } catch (err) {
            console.warn("Could not create trustedTypes policy.");
        }
    }

    return policy;
})();

function main() {
    let refreshAll = debounce(() => {
        refreshDrafts();
        refreshMessages();
    }, 50)

    const themeObserver = new MutationObserver(refreshAll);

    watch(selectors.newMailButton, added => insertTexButton(added[0]));
    watch(selectors.draftBody, added => {
        themeObserver.observe(added[0], { attributeFilter: ["class"] });
        added.forEach(processDraft)
    });

    watch(selectors.messageBody, added => {
        added.forEach(msg => {
            themeObserver.observe(msg.firstChild, { childList: true, subtree: true, attributeFilter: ["style"] });
            processMessage(msg);
        })

        refreshMessages(added);
    })

    injectStyles();
    addShortcuts();
    GM_registerMenuCommand('Toggle LaTeX', toggleMessages);
    GM_registerMenuCommand('Toggle Default State', toggleDefaultMessagesState);
    GM_registerMenuCommand('Toggle $ & $$ Delimiters', toggleDollarDelimiters);
}

main();
