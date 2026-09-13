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

  function digitsOnly(value) {
    return String(value == null ? "" : value).replace(/\D/g, "");
  }

  // Dəyərin input-da həqiqətən əks olunduğunu yoxlayır. Mask prefiks əlavə
  // edə bilər (məs. '+(994)'), ona görə rəqəm sufiksinə də baxırıq.
  function valueMatches(el, value) {
    const expected = String(value == null ? "" : value);
    const actual = String(el.value == null ? "" : el.value);
    if (expected === actual) return true;

    const expDigits = digitsOnly(expected);
    if (expDigits) {
      const actDigits = digitsOnly(actual);
      return actDigits === expDigits || actDigits.endsWith(expDigits);
    }
    return actual.trim().toLowerCase() === expected.trim().toLowerCase();
  }

  function keyEventInit(ch) {
    const isChar = typeof ch === "string" && ch.length === 1;
    const codePoint = isChar ? ch.charCodeAt(0) : 0;
    return {
      bubbles: true,
      cancelable: true,
      key: ch,
      code: /[0-9]/.test(ch) ? "Digit" + ch : ch,
      keyCode: codePoint,
      which: codePoint,
      charCode: codePoint,
    };
  }

  // `document.execCommand("insertText")` KÖHNƏLMİŞ (deprecated) API-dir, lakin
  // mask/controlled widget-ləri (DevExtreme, Cleave, imask, React) "real yazı"
  // kimi emal etmək üçün hazırda ən etibarlı yoldur: brauzer özü
  // `beforeinput(inputType:"insertText")` + `input` hadisələrini yaradır və
  // caret/selection-u düzgün idarə edir. Gələcəkdə silinsə, aşağıdakı fallback
  // (sintetik beforeinput + native setter) işə düşür.
  //
  // Əlavə: bəzi wrapper komponentlər (məs. ATB-nin `sb-text-box`-i) maskalı
  // dəyəri yalnız `keyup` hadisəsində oxuyub form state-inə ötürür; yalnız
  // `input` kifayət etmir. Ona görə real yazıdakı kimi keydown/keypress/keyup
  // da dispatch edilir.
  function insertChar(el, ch) {
    try {
      el.dispatchEvent(new KeyboardEvent("keydown", keyEventInit(ch)));
    } catch (err) {
      // KeyboardEvent dəstəklənmirsə, keç
    }

    let inserted = false;
    if (typeof document.execCommand === "function") {
      try {
        inserted = document.execCommand("insertText", false, ch);
      } catch (err) {
        inserted = false;
      }
    }

    if (!inserted) {
      try {
        el.dispatchEvent(
          new InputEvent("beforeinput", {
            bubbles: true,
            cancelable: true,
            inputType: "insertText",
            data: ch,
          })
        );
      } catch (err) {
        // InputEvent dəstəklənmirsə, keç
      }
      setNativeValue(el, el.value + ch);
      el.dispatchEvent(new Event("input", { bubbles: true }));
    }

    try {
      el.dispatchEvent(new KeyboardEvent("keypress", keyEventInit(ch)));
    } catch (err) {
      // keç
    }
    try {
      el.dispatchEvent(new KeyboardEvent("keyup", keyEventInit(ch)));
    } catch (err) {
      // keç
    }
    return true;
  }

  function clearElement(el) {
    try {
      el.focus();
    } catch (err) {
      // keç
    }

    if (typeof document.execCommand === "function") {
      try {
        document.execCommand("selectAll", false, null);
        document.execCommand("delete", false, null);
        return;
      } catch (err) {
        // fallback
      }
    }

    try {
      el.setSelectionRange(0, String(el.value || "").length);
    } catch (err) {
      // setSelectionRange yoxdursa keç
    }
    setNativeValue(el, "");
    dispatchValueEvents(el);
  }

  // Mask-lı/controlled komponentlər birbaşa value təyinatını qəbul etmir,
  // çünki prefiksi/ayrıcıları/state-i özləri idarə edir. Bu halda dəyəri real
  // istifadəçi yazısı kimi simvol-simvol daxil edirik.
  // Bəzi mask widget-ləri (məs. DevExtreme) real dəyəri gizli elementdə/öz
  // state-ində saxlayır, görünən input isə yalnız maskalı mətn göstərir.
  // Belə halda birbaşa value təyinatı "uğurlu" görünsə də, widget sonradan
  // dəyəri geri qaytarır. Bu əlamətlərlə aşkar edib həmişə real yazı simulyasiya
  // edirik (ümumi — konkret portala bağlı deyil).
  function looksLikeMaskedWidget(el) {
    try {
      const cls = el.classList;
      if (cls && typeof cls.contains === "function") {
        if (cls.contains("dx-texteditor-input") || cls.contains("dx-texteditor")) {
          return true;
        }
      }
      if (el.closest && el.closest(".dx-texteditor")) return true;

      const parent = el.parentElement;
      if (parent && typeof parent.querySelector === "function") {
        if (parent.querySelector('input[type="hidden"]')) return true;
      }
    } catch (err) {
      // yoxlama alınmadı — mask widget sayılmır
    }
    return false;
  }

  async function typeValue(el, value) {
    const target = String(value == null ? "" : value);

    clearElement(el);
    await nextTick();

    for (const ch of target) {
      insertChar(el, ch);
      await delay(0);
    }

    // Widget-lər (məs. DevExtreme) dəyəri `change`/`blur` zamanı commit edir.
    el.dispatchEvent(new Event("change", { bubbles: true }));
    try {
      el.dispatchEvent(new Event("blur", { bubbles: true }));
    } catch (err) {
      // keç
    }

    return valueMatches(el, target);
  }

  // Bəzi mask komponentləri (məs. ATB-nin DevExtreme `sb-text-box`-i) `input`/
  // `execCommand`/`keyup` hadisələrini qəbul etmir, yalnız sintetik `paste`
  // hadisəsini (ClipboardEvent + DataTransfer) real yapışdırma kimi emal edir.
  // Bu üsul debugger/icazə tələb etmir və real brauzer davranışını təqlid edir.
  async function pasteValue(el, value) {
    const target = String(value == null ? "" : value);

    clearElement(el);
    await nextTick();

    try {
      const dt = new DataTransfer();
      dt.setData("text/plain", target);
      el.dispatchEvent(
        new ClipboardEvent("paste", {
          bubbles: true,
          cancelable: true,
          clipboardData: dt,
        })
      );
    } catch (err) {
      return false;
    }

    el.dispatchEvent(new Event("change", { bubbles: true }));
    await nextTick();
    return valueMatches(el, target);
  }

  async function setElementValue(el, value) {
    const tag = el.tagName.toLowerCase();

    if (tag === "select") {
      const target = Array.from(el.options).find(
        (o) => o.value === value || o.text === value
      );
      if (!target) return false;
      setNativeValue(el, target.value);
      dispatchValueEvents(el);
      return true;
    }

    if (el.isContentEditable) {
      el.textContent = value;
      dispatchValueEvents(el);
      return true;
    }

    setNativeValue(el, value);
    dispatchValueEvents(el);
    await nextTick();

    // 1) Sürətli yol: React/vanilla (mask widget deyilsə).
    if (valueMatches(el, value) && !looksLikeMaskedWidget(el)) return true;

    // 2) Sintetik paste — paste-only mask-lar (məs. ATB / DevExtreme).
    if (await pasteValue(el, value)) return true;

    // 3) Son ehtiyat: simvol-simvol real yazı (məs. Access Bank / Cleave).
    return await typeValue(el, value);
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

  function isPrefixCapableInput(el) {
    if (!el || el.tagName !== "INPUT") return false;
    const type = (el.type || "text").toLowerCase();
    return type === "text" || type === "tel" || type === "number";
  }

  // Framework "Select/Combobox" trigger-i: portala açılan siyahıdan seçim
  // tələb edir. Yalnız aydın əlamətlərlə tanınır (yanlış elementə klikləməmək üçün).
  function looksLikeCombobox(el) {
    if (!el || el.tagName !== "INPUT") return false;
    if (el.getAttribute("role") === "combobox") return true;
    const popup = el.getAttribute("aria-haspopup");
    return popup === "listbox" || popup === "menu";
  }

  // Vue Select (v-select) prefiks komponentidir: toggle `role="combobox"` olan
  // `<div class="vs__dropdown-toggle">`, axtarış `input.vs__search`, option-lar
  // `.vs__dropdown-option`. Native <select> deyil.
  function isVueSelect(el) {
    if (!el) return false;
    try {
      if (
        el.closest &&
        (el.closest(".v-select") || el.closest(".vs__dropdown-toggle"))
      ) {
        return true;
      }
      const cls = el.classList;
      if (cls && typeof cls.contains === "function") {
        if (
          cls.contains("v-select") ||
          cls.contains("vs__dropdown-toggle") ||
          cls.contains("vs__search")
        ) {
          return true;
        }
      }
    } catch (err) {
      // yoxlama alınmadı
    }
    return false;
  }

  function isSimpleDropdown(el) {
    if (!el) return false;
    try {
      if (el.classList && typeof el.classList.contains === "function") {
        if (
          el.classList.contains("dropdown") ||
          el.classList.contains("selected-item")
        ) {
          return true;
        }
      }
      if (el.closest && (el.closest(".dropdown") || el.closest(".selected-item"))) {
        return true;
      }
      if (el.querySelector && el.querySelector(".selected-item")) return true;
    } catch (err) {
      // yoxlama alınmadı
    }
    return false;
  }

  function findPrefixControl(phoneInput, overrideSelector) {
    if (overrideSelector) {
      try {
        const el = document.querySelector(overrideSelector);
        if (el && el.tagName) {
          const tag = el.tagName.toLowerCase();
          if (tag === "select") return el;
          if (tag === "input" && (isPrefixCapableInput(el) || looksLikeCombobox(el))) {
            return el;
          }
          if (isVueSelect(el)) return el;
          if (isSimpleDropdown(el)) return el;
        }
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
      // Yalnız role=combobox / aria-haspopup=listbox olan aydın trigger.
      const combos = scope.querySelectorAll(
        'input[role="combobox"], input[aria-haspopup="listbox"]'
      );
      for (const c of combos) {
        if (c !== phoneInput) return c;
      }
      // Vue Select toggle (role=combobox olan div) və ya .vs__dropdown-toggle.
      const vueToggle = scope.querySelector(
        '.vs__dropdown-toggle, [role="combobox"]:not(input)'
      );
      if (vueToggle && vueToggle !== phoneInput) return vueToggle;
      const vueWrap = scope.querySelector(".v-select");
      if (vueWrap) {
        const t = vueWrap.querySelector(".vs__dropdown-toggle") || vueWrap;
        if (t !== phoneInput) return t;
      }
      // Sadə dropdown: <span class="selected-item"> + açılan .dropdown-menu.
      const selItem = scope.querySelector(".selected-item");
      if (selItem) {
        const container =
          (selItem.closest &&
            selItem.closest(".dropdown, [role='button'], .select")) ||
          selItem.parentElement ||
          selItem;
        if (container && container !== phoneInput) return container;
      }
      node = scope;
    }
    return null;
  }

  function prefixCandidates(prefix) {
    const set = new Set([prefix]);
    if (!prefix.startsWith("0")) set.add("0" + prefix);
    set.add("994" + prefix);
    return [...set];
  }

  function matchPrefixToken(text, candidates) {
    const norm = normalizePrefixToken(text);
    return candidates.some((c) => normalizePrefixToken(c) === norm);
  }

  function delay(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function nextTick() {
    return new Promise((resolve) => {
      if (typeof requestAnimationFrame === "function") {
        requestAnimationFrame(() => resolve());
      } else {
        setTimeout(resolve, 0);
      }
    });
  }

  function findComboboxOptions() {
    return Array.from(
      document.querySelectorAll(
        '[data-combobox-option], [role="listbox"] [role="option"], ' +
          '[role="option"], .vs__dropdown-option'
      )
    );
  }

  // Vue Select prefiksini açıb uyğun operator kodunu seçir.
  async function selectVueSelectOption(control, prefix) {
    let toggle = control;
    if (control && control.closest) {
      toggle = control.closest(".vs__dropdown-toggle") || control;
    }
    const search =
      (toggle.querySelector && toggle.querySelector(".vs__search")) || toggle;
    const candidates = prefixCandidates(prefix);

    let options = Array.from(document.querySelectorAll(".vs__dropdown-option"));
    if (options.length === 0) {
      try {
        search.focus();
      } catch (err) {
        // focus alınmadı
      }
      toggle.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
      toggle.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
      toggle.dispatchEvent(new MouseEvent("click", { bubbles: true }));

      const deadline = Date.now() + 900;
      while (Date.now() < deadline && options.length === 0) {
        await delay(40);
        options = Array.from(document.querySelectorAll(".vs__dropdown-option"));
      }
    }

    for (const opt of options) {
      const text = opt.textContent || "";
      if (matchPrefixToken(text, candidates)) {
        opt.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
        opt.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
        opt.click();
        await nextTick();
        return true;
      }
    }
    return false;
  }

  // Sadə dropdown prefiksi: bağlı `<span class="selected-item">` + açılanda
  // `.dropdown-menu .dropdown-item` (məs. Rabitəbank). Menyu yalnız açıq
  // olduqda DOM-da olur — ona görə option-ları gözləməklə axtarırıq.
  async function selectSimpleDropdownOption(control, prefix) {
    const candidates = prefixCandidates(prefix);

    const selectedSpan =
      (control.querySelector && control.querySelector(".selected-item")) ||
      (control.classList &&
      typeof control.classList.contains === "function" &&
      control.classList.contains("selected-item")
        ? control
        : null);

    const toggle =
      (control.closest &&
        (control.closest(".dropdown") ||
          control.closest('[role="button"]') ||
          control.closest(".select"))) ||
      (selectedSpan && selectedSpan.parentElement) ||
      control;

    const menuOptions = () =>
      Array.from(
        document.querySelectorAll(
          ".ps.dropdown-menu .dropdown-item, .dropdown-menu .dropdown-item"
        )
      );

    let options = menuOptions();
    if (options.length === 0) {
      toggle.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
      toggle.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
      toggle.dispatchEvent(new MouseEvent("click", { bubbles: true }));

      const deadline = Date.now() + 1000;
      while (Date.now() < deadline && options.length === 0) {
        await delay(40);
        options = menuOptions();
      }
    }

    for (const opt of options) {
      const text = opt.textContent || "";
      if (matchPrefixToken(text, candidates)) {
        opt.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
        opt.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
        opt.click();
        await nextTick();
        await delay(30);

        // Seçimin tətbiq olunduğunu yoxla (span.selected-item mətni dəyişdi?).
        const span =
          (selectedSpan &&
          typeof document.contains === "function" &&
          document.contains(selectedSpan)
            ? selectedSpan
            : null) || document.querySelector(".selected-item");
        if (!span) return true;
        return matchPrefixToken(span.textContent || "", candidates);
      }
    }
    return false;
  }

  // Mantine Select kimi combobox-ı açıb uyğun operator kodunu seçir.
  // Açmaq və seçim real istifadəçi hərəkətləri ilə simulyasiya olunur.
  async function selectComboboxOption(trigger, prefix) {
    const candidates = prefixCandidates(prefix);
    let options = findComboboxOptions();

    if (options.length === 0) {
      trigger.focus();
      trigger.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
      trigger.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
      trigger.dispatchEvent(new MouseEvent("click", { bubbles: true }));

      const deadline = Date.now() + 900;
      while (Date.now() < deadline && options.length === 0) {
        await delay(40);
        options = findComboboxOptions();
      }
    }

    for (const opt of options) {
      const text = opt.textContent || opt.getAttribute("value") || "";
      if (matchPrefixToken(text, candidates)) {
        opt.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
        opt.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
        opt.click();
        await nextTick(); // React state-in oturması
        return true;
      }
    }
    return false;
  }

  // Prefiks nəzarətinin növünə görə operator kodunu yazır.
  // select → uyğun option; combobox → siyahıdan seçim; input → format.
  async function applyPrefix(control, prefix) {
    if (!control) return false;
    const tag = control.tagName ? control.tagName.toLowerCase() : "";

    if (tag === "select") {
      for (const o of Array.from(control.options || [])) {
        if (
          normalizePrefixToken(o.textContent) === prefix ||
          normalizePrefixToken(o.value) === prefix
        ) {
          setNativeValue(control, o.value);
          dispatchValueEvents(control);
          return true;
        }
      }
      return false;
    }

    if (isVueSelect(control)) {
      return await selectVueSelectOption(control, prefix);
    }

    if (isSimpleDropdown(control)) {
      return await selectSimpleDropdownOption(control, prefix);
    }

    if (looksLikeCombobox(control)) {
      return await selectComboboxOption(control, prefix);
    }

    const maxlength = parseInt(control.getAttribute("maxlength") || "", 10);
    const placeholder = String(control.getAttribute("placeholder") || "").replace(
      /\D/g,
      ""
    );
    let value = prefix;
    if (placeholder.startsWith("994")) value = "994" + prefix;
    else if (
      placeholder.startsWith("0") ||
      (Number.isFinite(maxlength) && maxlength >= 3)
    ) {
      value = "0" + prefix;
    }

    return await setElementValue(control, value);
  }

  // Bəzi masklar yalnız milli 9 rəqəmi, bəziləri isə ölkə kodu daxil tam
  // 12 rəqəmi (994 + 9) gözləyir. Bu, element-səviyyəli siqnallarla aşkarlanır:
  // boş vəziyyətdə default "994", placeholder-də "994", yaxud maxLength >= 12.
  function needsCountryCode(el) {
    const raw = String(el.value || "");
    const digits = digitsOnly(raw);
    // Maska placeholder-i varsa (məs. "+994 (__) ___-__-__"), prefiks sabit
    // göstərilir və yalnız yerli 9 rəqəm gözlənilir -> ölkə kodu tələb olunmur.
    const hasMaskPlaceholder = /_/.test(raw);
    if (!hasMaskPlaceholder && /^994/.test(digits)) return true;
    if (/994/.test(String(el.placeholder || ""))) return true;
    const ml = parseInt(el.getAttribute("maxlength") || "", 10);
    if (Number.isFinite(ml) && ml >= 12) return true;
    return false;
  }

  async function fillPhone(phoneInput, phoneValue, overrideSelector) {
    const parts = phoneParts(phoneValue);
    if (!parts.full) return false;
    if (parts.prefix) {
      const control = findPrefixControl(phoneInput, overrideSelector);
      if (control && (await applyPrefix(control, parts.prefix))) {
        await nextTick();
        return await setElementValue(phoneInput, parts.local);
      }
      // Prefiks nəzarəti tapılmadı: sahə yerli hissə üçündürsə (maxLength tam
      // nömrədən kiçikdir), tam 9 rəqəmi yazmaq kəsilmiş səhv dəyər yaradar —
      // yazmırıq (düzgün "uğursuz" statusu).
      const maxlength = parseInt(phoneInput.getAttribute("maxlength") || "", 10);
      if (
        Number.isFinite(maxlength) &&
        maxlength > 0 &&
        maxlength < parts.full.length
      ) {
        return false;
      }
    }
    // Ölkə kodunu öz daxilində gözləyən mask (məs. Yelo: default "994 ").
    // 12 rəqəm göndərilir; CİDDİ yoxlama (===) sayəsində Access/Cleave kimi
    // masklarda ikiqat 994 "uğurlu" sayılmır və aşağıdakı 9 rəqəm yoluna düşür.
    if (needsCountryCode(phoneInput)) {
      const cc = "994" + parts.full;
      await setElementValue(phoneInput, cc);
      if (digitsOnly(phoneInput.value) === cc) return true;
    }
    return await setElementValue(phoneInput, parts.full);
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

  async function fillFields(fields, phonePrefixSelector) {
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

    for (const entry of entries) {
      const field = entry.field || entry.selector;
      const value = entry.value;
      const found = findElement(field, entry.selector);

      if (!found) {
        results.push({ field, selector: entry.selector || "", value, ok: false, how: null });
        continue;
      }

      let ok = false;
      try {
        ok =
          field === "nomre"
            ? await fillPhone(found.el, value, phonePrefixSelector)
            : await setElementValue(found.el, value);
      } catch (err) {
        // Bir sahənin xətası digər sahələri/statusu pozmasın.
        ok = false;
      }

      if (ok) matchedBy[field] = found.how;
      results.push({
        field,
        selector: entry.selector || "",
        value,
        ok,
        how: found.how,
      });
    }

    // İkinci yoxlama keçidi: yazı anında "uğursuz" görünsə də, DOM-da faktiki
    // dəyəri düzgün olan sahələri (mask gecikməsi / format false-negative)
    // uğurlu say. Element yenidən tapılır ki, köhnə referans istifadə olunmasın.
    if (results.some((item) => !item.ok)) {
      await delay(20); // widget/mask dəyəri gec commit edə bilər
    }
    for (const item of results) {
      if (item.ok) continue;
      try {
        const found = findElement(item.field, item.selector);
        if (found && valueMatches(found.el, item.value)) {
          item.ok = true;
          matchedBy[item.field] = found.how || item.how || "verified";
        }
      } catch (err) {
        // yoxlama alınmadı — uğursuz qalır
      }
    }

    let filled = 0;
    for (const item of results) {
      if (item.ok) filled += 1;
      else missing.push(item.selector || item.field);
    }

    // `how`: "exact" = DB-dəki selector ilə; "attr"/"label"/"placeholder" = heuristik.
    return { filled, total, missing, matchedBy, results };
  }

  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (msg && msg.type === "AUTOFILL") {
      // Prefiks combobox seçimi async olduğu üçün cavab gec verilir.
      fillFields(msg.fields, msg.phone_prefix_selector || null)
        .then(sendResponse)
        .catch(() => {
          const list = Array.isArray(msg.fields)
            ? msg.fields
            : Object.keys(msg.fields || {});
          sendResponse({
            filled: 0,
            total: list.length,
            missing: list.map((e) => (e && e.selector) || e || ""),
            matchedBy: {},
            results: [],
          });
        });
      return true; // async sendResponse kanalı
    }

    if (msg && msg.type === "START_RECORD") {
      startRecordMode();
      sendResponse({ ok: true });
      return false;
    }

    if (msg && msg.type === "STOP_RECORD") {
      stopRecordMode();
      sendResponse({ ok: true });
      return false;
    }

    return false;
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

  // Framework-lərin (Mantine/React/Radix və s.) generasiya etdiyi ID-lər hər
  // yükləmədə dəyişə bilər — onları selector kimi saxlamaq etibarsızdır və
  // sonradan tamam başqa elementə düşə bilər.
  function isUnstableId(id) {
    const v = String(id || "");
    if (!v) return true;
    if (/^mantine-/i.test(v)) return true;
    if (/^:r[0-9a-z]+:/i.test(v)) return true; // React useId: :r0:, :r3:
    if (/^(radix|headlessui|downshift|mui|chakra)-/i.test(v)) return true;
    if (/^[0-9a-f]{8,}$/i.test(v)) return true; // hash-a bənzər ID (tiresiz)
    // UUID / GUID (tireli) — Vue və digər framework-lərin generasiya etdiyi
    // hər yükləmədə dəyişən ID-lər (məs. Rabitəbank input-u).
    if (
      /^#?\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i.test(
        v
      )
    ) {
      return true;
    }
    return false;
  }

  function escapeAttrValue(v) {
    return String(v).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  }

  function uniqueSelector(selector) {
    if (!selector) return null;
    try {
      return document.querySelectorAll(selector).length === 1 ? selector : null;
    } catch (err) {
      return null;
    }
  }

  function computeUniqueSelector(el) {
    const cssEscape = (v) => (typeof CSS !== "undefined" && CSS.escape
      ? CSS.escape(String(v))
      : String(v));
    const tag = el.tagName ? el.tagName.toLowerCase() : "input";
    const getAttr = (name) =>
      el.getAttribute ? el.getAttribute(name) : null;

    // 1) Test atributları — ən sabit seçim.
    for (const attr of ["data-testid", "data-test", "data-cy", "data-qa"]) {
      const val = getAttr(attr);
      if (val) {
        const found = uniqueSelector(`[${attr}="${escapeAttrValue(val)}"]`);
        if (found) return found;
      }
    }

    // 2) name.
    if (el.name) {
      const found = uniqueSelector(
        `${tag}[name="${escapeAttrValue(el.name)}"]`
      );
      if (found) return found;
    }

    // 3) aria-label.
    const aria = getAttr("aria-label");
    if (aria) {
      const found = uniqueSelector(
        `${tag}[aria-label="${escapeAttrValue(aria)}"]`
      );
      if (found) return found;
    }

    // 4) autocomplete.
    const ac = getAttr("autocomplete");
    if (ac && ac !== "off") {
      const found = uniqueSelector(
        `${tag}[autocomplete="${escapeAttrValue(ac)}"]`
      );
      if (found) return found;
    }

    // 5) Stabil id (framework generasiyası istisna) — data-v-*-dan ÜSTÜN,
    // çünki sabit id build-lar arasında dəyişmir.
    if (el.id && !isUnstableId(el.id)) {
      const found = uniqueSelector("#" + cssEscape(el.id));
      if (found) return found;
    }

    // 5.5) Vue scoped atributu (data-v-*) — yalnız stabil id seçilməyəndə.
    // (data-v-* təkbaşına komponent səviyyəsindədir, unikal olmaya bilər.)
    if (typeof el.getAttributeNames === "function") {
      const names = el.getAttributeNames() || [];
      for (const attr of names) {
        if (!/^data-v-/.test(attr)) continue;
        const cls = String(el.className || "")
          .trim()
          .split(/\s+/)
          .filter(Boolean)[0];
        if (cls) {
          const found = uniqueSelector(
            `${tag}.${cssEscape(cls)}[${attr}]`
          );
          if (found) return found;
        }
        const found2 = uniqueSelector(`${tag}[${attr}]`);
        if (found2) return found2;
      }
    }

    // 6) placeholder.
    if (el.placeholder) {
      const found = uniqueSelector(
        `${tag}[placeholder="${escapeAttrValue(el.placeholder)}"]`
      );
      if (found) return found;
    }

    // 7) Struktur yol (son çarə).
    const parts = [];
    let node = el;
    while (node && node.nodeType === 1 && node.tagName) {
      const nodeTag = node.tagName.toLowerCase();
      if (nodeTag === "body" || nodeTag === "html") break;
      let sel = nodeTag;
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
