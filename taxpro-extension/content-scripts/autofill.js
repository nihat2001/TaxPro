/* TaxPro — autofill content script.
 * Passivdir: səhifəni skan etmir. Yalnız popup/background tərəfindən
 * {type: "AUTOFILL", fields} mesajı göndəriləndə işə düşür.
 */
(function () {
  "use strict";

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

  function fillFields(fields) {
    const entries = Object.entries(fields || {});
    const total = entries.length;
    let filled = 0;
    const missing = [];

    for (const [selector, value] of entries) {
      let el = null;
      try {
        el = document.querySelector(selector);
      } catch (err) {
        missing.push(selector);
        continue;
      }
      if (!el || !isEditable(el)) {
        missing.push(selector);
        continue;
      }
      if (setElementValue(el, value)) filled += 1;
      else missing.push(selector);
    }

    return { filled, total, missing };
  }

  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg && msg.type === "AUTOFILL") {
      const result = fillFields(msg.fields);
      sendResponse(result);
    }
  });
})();
