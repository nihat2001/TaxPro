/* TaxPro — autofill content script.
 * Passivdir: səhifəni skan etmir. Yalnız popup/background tərəfindən
 * {type: "AUTOFILL", fields} mesajı göndəriləndə işə düşür.
 *
 * Hər sahə üçün fallback zənciri:
 *   1) exact    — DB-dəki dəqiq CSS selector
 *   2) attr     — input[name|id|autocomplete*=keyword]
 *   3) label    — <label> mətninə görə
 *   4) placeholder — input[placeholder] mətninə görə
 */
(function () {
  "use strict";

  const FIELD_KEYWORDS = {
    nomre: ["nomre", "nömrə", "phone", "mobile", "mobil", "tel", "telefon", "telephone"],
    asan_id: ["asan", "asanid", "userid", "user id", "username", "istifadeci", "istifadəçi", "login"],
  };

  function normalize(s) {
    return String(s == null ? "" : s)
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function keywordsFor(field, selector) {
    const keys = new Set();
    for (const base of [field, selector]) {
      if (!base) continue;
      normalize(base)
        .split(/[^a-z0-9]+/)
        .filter((t) => t.length >= 3)
        .forEach((t) => keys.add(t));
    }
    (FIELD_KEYWORDS[field] || []).forEach((k) => keys.add(k));
    keys.delete("input");
    return [...keys];
  }

  function isEditable(el) {
    if (!el) return false;
    const tag = el.tagName.toLowerCase();
    return (
      tag === "input" ||
      tag === "textarea" ||
      tag === "select" ||
      el.isContentEditable
    );
  }

  /* DOM-a dəyər yazma.
   * React (və digər controlled framework-lər) element üzərində öz "value
   * tracker" setter-ini qurur. `el.value = x` yazsaq tracker də x olur və
   * dispatch etdiyimiz `input` hadisəsində React "dəyişiklik yoxdur" qənaətinə
   * gəlib onChange-i işə salmır; növbəti render-də köhnə dəyər geri qaytarılır.
   * Prototype descriptor-un `set`-i ilə yazmaq tracker-ı bypass edir və React
   * state düzgün yenilənir. Vanilla input-larda da adi `el.value` ilə eynidir.
   */
  function setNativeValue(el, value) {
    const tag = el.tagName ? el.tagName.toLowerCase() : "";
    let proto = window.HTMLInputElement && window.HTMLInputElement.prototype;
    if (tag === "textarea" && window.HTMLTextAreaElement) {
      proto = window.HTMLTextAreaElement.prototype;
    } else if (tag === "select" && window.HTMLSelectElement) {
      proto = window.HTMLSelectElement.prototype;
    }

    let setter = null;
    if (proto && Object.getOwnPropertyDescriptor) {
      const desc = Object.getOwnPropertyDescriptor(proto, "value");
      if (desc && desc.set) setter = desc.set;
    }

    if (setter) setter.call(el, value);
    else el.value = value;
  }

  function dispatchValueEvents(el) {
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function setElementValue(el, value) {
    const tag = el.tagName.toLowerCase();

    if (tag === "select") {
      const target = Array.from(el.options).find(
        (o) => o.value === value || o.text === value
      );
      if (!target) return false;
      setNativeValue(el, target.value);
    } else if (el.isContentEditable) {
      el.textContent = value;
    } else {
      setNativeValue(el, value);
    }

    dispatchValueEvents(el);
    return true;
  }

  /* ===== Telefon prefiks məntiqi =====
   * Nömrə kanonik 9 rəqəmdir (məs. "501234567"): operator kodu ilk 2 rəqəm
   * ("50"), yerli hissə son 7 rəqəm ("1234567").
   * Login formalarında prefiks ayrıca select/dropdown ola bilər ("050", "+994 50",
   * "50", "(050)") — hamısı normalize edilib müqayisə olunur.
   */
  function phoneParts(value) {
    let digits = String(value == null ? "" : value).replace(/\D/g, "");
    if (digits.length === 12 && digits.startsWith("994")) digits = digits.slice(3);
    else if (digits.length === 10 && digits.startsWith("0")) digits = digits.slice(1);
    if (digits.length !== 9) {
      return { prefix: null, local: null, full: digits || null };
    }
    return { prefix: digits.slice(0, 2), local: digits.slice(2), full: digits };
  }

  function normalizePrefixToken(raw) {
    let digits = String(raw == null ? "" : raw).replace(/\D/g, "");
    if (digits.startsWith("994")) digits = digits.slice(3);
    if (digits.startsWith("0")) digits = digits.slice(1);
    return digits;
  }

  function looksLikePrefixSelect(sel) {
    const options = Array.from(sel.options || []);
    if (options.length === 0 || options.length > 40) return false;
    let matches = 0;
    for (const o of options) {
      const tok = normalizePrefixToken(o.textContent);
      const val = normalizePrefixToken(o.value);
      if (
        (tok && tok.length >= 2 && tok.length <= 3) ||
        (val && val.length >= 2 && val.length <= 3)
      ) {
        matches += 1;
      }
    }
    return matches >= Math.max(1, Math.ceil(options.length / 2));
  }

  function findPrefixSelect(phoneInput, overrideSelector) {
    if (overrideSelector) {
      try {
        const el = document.querySelector(overrideSelector);
        if (el && el.tagName && el.tagName.toLowerCase() === "select") return el;
      } catch (err) {
        // keçərsiz selector — heuristic-ə keç
      }
    }
    let node = phoneInput;
    for (let i = 0; i < 4 && node; i++) {
      const scope = node.parentElement;
      if (!scope) break;
      for (const sel of scope.querySelectorAll("select")) {
        if (sel !== phoneInput && looksLikePrefixSelect(sel)) return sel;
      }
      node = scope;
    }
    return null;
  }

  function selectPrefix(sel, prefix) {
    for (const o of Array.from(sel.options || [])) {
      if (
        normalizePrefixToken(o.textContent) === prefix ||
        normalizePrefixToken(o.value) === prefix
      ) {
        setNativeValue(sel, o.value);
        dispatchValueEvents(sel);
        return true;
      }
    }
    return false;
  }

  function fillPhone(phoneInput, phoneValue, overrideSelector) {
    const parts = phoneParts(phoneValue);
    if (!parts.full) return false;
    if (parts.prefix) {
      const sel = findPrefixSelect(phoneInput, overrideSelector);
      if (sel && selectPrefix(sel, parts.prefix)) {
        return setElementValue(phoneInput, parts.local);
      }
    }
    return setElementValue(phoneInput, parts.full);
  }

  function escAttr(k) {
    return String(k).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  }

  function findByAttr(keys) {
    for (const k of keys) {
      const q =
        `input[name*="${escAttr(k)}"], input[id*="${escAttr(k)}"], ` +
        `input[autocomplete*="${escAttr(k)}"], ` +
        `textarea[name*="${escAttr(k)}"], textarea[id*="${escAttr(k)}"]`;
      try {
        const list = document.querySelectorAll(q);
        for (const el of list) if (isEditable(el)) return el;
      } catch (err) {
        // səhv selector sintaksisi — növbəti keyword-a keç
      }
    }
    return null;
  }

  function findByLabel(keys) {
    const labels = document.querySelectorAll("label");
    for (const lbl of labels) {
      const txt = normalize(lbl.textContent);
      if (!keys.some((k) => txt.includes(k))) continue;
      let target = null;
      if (lbl.htmlFor) target = document.getElementById(lbl.htmlFor);
      if (!target) target = lbl.querySelector("input, textarea, select");
      if (target && isEditable(target)) return target;
    }
    return null;
  }

  function findByPlaceholder(keys) {
    const els = document.querySelectorAll("input, textarea");
    for (const el of els) {
      if (!el.placeholder) continue;
      const txt = normalize(el.placeholder);
      if (keys.some((k) => txt.includes(k))) return el;
    }
    return null;
  }

  function findElement(field, selector) {
    if (selector) {
      try {
        const el = document.querySelector(selector);
        if (el && isEditable(el)) return { el, how: "exact" };
      } catch (err) {
        // keçərli olmayan selector — fallback davam edir
      }
    }

    const keys = keywordsFor(field, selector);

    const byAttr = findByAttr(keys);
    if (byAttr) return { el: byAttr, how: "attr" };

    const byLabel = findByLabel(keys);
    if (byLabel) return { el: byLabel, how: "label" };

    const byPh = findByPlaceholder(keys);
    if (byPh) return { el: byPh, how: "placeholder" };

    return null;
  }

  function fillFields(fields, phonePrefixSelector) {
    const entries = Array.isArray(fields)
      ? fields
      : Object.entries(fields || {}).map(([selector, value]) => ({
          field: selector,
          selector,
          value,
        }));

    const total = entries.length;
    const missing = [];
    const matchedBy = {};
    const results = [];
    let filled = 0;

    for (const entry of entries) {
      const field = entry.field || entry.selector;
      const value = entry.value;
      const found = findElement(field, entry.selector);

      if (!found) {
        missing.push(entry.selector || field);
        results.push({ field, selector: entry.selector || "", ok: false, how: null });
        continue;
      }

      const ok =
        field === "nomre"
          ? fillPhone(found.el, value, phonePrefixSelector)
          : setElementValue(found.el, value);

      if (ok) {
        filled += 1;
        matchedBy[field] = found.how;
        results.push({ field, selector: entry.selector || "", ok: true, how: found.how });
      } else {
        missing.push(entry.selector || field);
        results.push({ field, selector: entry.selector || "", ok: false, how: found.how });
      }
    }

    // `how`: "exact" = DB-dəki selector ilə; "attr"/"label"/"placeholder" = heuristik.
    return { filled, total, missing, matchedBy, results };
  }

  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg && msg.type === "AUTOFILL") {
      const result = fillFields(msg.fields, msg.phone_prefix_selector || null);
      sendResponse(result);
    }

    if (msg && msg.type === "START_RECORD") {
      startRecordMode();
      sendResponse({ ok: true });
    }

    if (msg && msg.type === "STOP_RECORD") {
      stopRecordMode();
      sendResponse({ ok: true });
    }
  });

  /* ===== Qeydiyyat rejimi (Record Mode) ===== */
  const RECORD_CLASS = "taxpro-record-highlight";
  let recordMode = false;

  function injectRecordStyle() {
    const style = document.createElement("style");
    style.textContent =
      "." + RECORD_CLASS +
      " { outline: 2px solid #ff4d4d !important; outline-offset: 1px !important; " +
      "box-shadow: 0 0 0 2px rgba(255, 77, 77, 0.35) !important; cursor: crosshair !important; }";
    document.head.appendChild(style);
  }

  function clearAllHighlights() {
    document
      .querySelectorAll("." + RECORD_CLASS)
      .forEach((el) => el.classList.remove(RECORD_CLASS));
  }

  function onRecordMouseOver(e) {
    const el = e.target && e.target.closest
      ? e.target.closest("input, textarea, select")
      : null;
    if (el && isEditable(el)) el.classList.add(RECORD_CLASS);
  }

  function onRecordMouseOut(e) {
    const el = e.target && e.target.closest
      ? e.target.closest("input, textarea, select")
      : null;
    if (el) el.classList.remove(RECORD_CLASS);
  }

  function onRecordClick(e) {
    try {
      const el = e.target && e.target.closest
        ? e.target.closest("input, textarea, select")
        : null;
      if (!el || !isEditable(el)) return;
      e.preventDefault();
      e.stopPropagation();
      const selector = computeUniqueSelector(el);
      el.classList.remove(RECORD_CLASS);
      console.log("[TaxPro] field captured:", selector);
      chrome.runtime.sendMessage({ type: "FIELD_CAPTURED", selector }).catch(() => {});
    } catch (err) {
      console.error("[TaxPro] record click failed:", err);
    }
  }

  function computeUniqueSelector(el) {
    const cssEscape = (v) => (typeof CSS !== "undefined" && CSS.escape
      ? CSS.escape(String(v))
      : String(v));

    if (el.id) {
      const s = "#" + cssEscape(el.id);
      try {
        if (document.querySelectorAll(s).length === 1) return s;
      } catch (err) {
        // davam et
      }
    }

    if (el.name) {
      const safe = String(el.name).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
      const s = el.tagName.toLowerCase() + '[name="' + safe + '"]';
      try {
        if (document.querySelectorAll(s).length === 1) return s;
      } catch (err) {
        // davam et
      }
    }

    const parts = [];
    let node = el;
    while (node && node.nodeType === 1 && node.tagName) {
      const tag = node.tagName.toLowerCase();
      if (tag === "body" || tag === "html") break;
      let sel = tag;
      const parent = node.parentNode;
      if (parent && parent.children) {
        const same = Array.from(parent.children).filter(
          (c) => c.tagName === node.tagName
        );
        if (same.length > 1) {
          sel += ":nth-of-type(" + (same.indexOf(node) + 1) + ")";
        }
      }
      parts.unshift(sel);
      node = parent;
    }
    return "body > " + parts.join(" > ");
  }

  function startRecordMode() {
    if (recordMode) return;
    recordMode = true;
    document.addEventListener("mouseover", onRecordMouseOver);
    document.addEventListener("mouseout", onRecordMouseOut);
    document.addEventListener("click", onRecordClick, true);
  }

  function stopRecordMode() {
    recordMode = false;
    document.removeEventListener("mouseover", onRecordMouseOver);
    document.removeEventListener("mouseout", onRecordMouseOut);
    document.removeEventListener("click", onRecordClick, true);
    clearAllHighlights();
  }

  if (document.head) injectRecordStyle();
})();
