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
    ad_soyad: ["ad soyad", "adsoyad", "fullname", "full name", "name", "tam ad"],
    voen: ["voen", "vatin", "vat", "tin", "tax"],
    nomre: ["nomre", "nömrə", "phone", "mobile", "mobil", "tel", "telefon", "telephone"],
    asan_id: ["asan", "asanid", "userid", "user id", "username", "istifadeci", "istifadəçi", "login"],
    pin1: ["pin", "parol", "shifre", "şifrə", "password"],
    pin2: ["pin2", "tekrar", "confirm"],
    puk: ["puk"],
    qeyd: ["qeyd", "note", "comment", "mulahize", "mülahizə"],
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

  function setElementValue(el, value) {
    const tag = el.tagName.toLowerCase();

    if (tag === "select") {
      const target = Array.from(el.options).find(
        (o) => o.value === value || o.text === value
      );
      if (!target) return false;
      el.value = target.value;
    } else if (el.isContentEditable) {
      el.textContent = value;
    } else {
      el.value = value;
    }

    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    return true;
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

  function fillFields(fields) {
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
    let filled = 0;

    for (const entry of entries) {
      const field = entry.field || entry.selector;
      const value = entry.value;
      const found = findElement(field, entry.selector);

      if (!found) {
        missing.push(entry.selector || field);
        continue;
      }

      if (setElementValue(found.el, value)) {
        filled += 1;
        matchedBy[field] = found.how;
      } else {
        missing.push(entry.selector || field);
      }
    }

    return { filled, total, missing, matchedBy };
  }

  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg && msg.type === "AUTOFILL") {
      const result = fillFields(msg.fields);
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
