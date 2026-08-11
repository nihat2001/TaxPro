/* TaxPro — ortaq HITL (human-in-the-loop) komponenti.
 * LangGraph interrupt() çağıranda backend {status:"interrupt", interrupt:{...}} qaytarır.
 * Bu komponent preview-i göstərir və istifadəçiyə Təsdiqlə/Ləğv seçimi təqdim edir,
 * sonra /resume çağırıb nəticəni onDone callback-i ilə qaytarır.
 * Görünmə yalnız `hidden` atributu ilə idarə olunur (class="hidden" yox!).
 */
(function (global) {
  "use strict";

  function escapeHtml(str) {
    return String(str == null ? "" : str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function money(v) {
    const n = Number(v);
    if (!isFinite(n)) return "0.00";
    return n.toLocaleString("az-AZ", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  }

  /* Interrupt payload-ından GÜVƏNLİ preview sətirləri hazırlayır.
   * Yalnız məlum, həssas olmayan sahələr göstərilir — PIN/PUK kimi dəyərlər heç vaxt
   * preview mətninə düşmür. Naməlum açarlar avtomatik göstərilmir.
   */
  function previewLines(interrupt) {
    if (!interrupt || typeof interrupt !== "object") return [];
    const lines = [];
    if (interrupt.action === "create_profiles") {
      const n = Number(interrupt.profile_count);
      if (isFinite(n) && n > 0) {
        lines.push(`Yaradılacaq profil sayı: ${n}`);
      }
    } else if (interrupt.action === "update_profile_from_analysis") {
      const p = interrupt.analysis_preview || {};
      lines.push(`Ümumi: ${money(p.total_amount)} AZN`);
      lines.push(`ƏDV: ${money(p.vat_amount)} AZN`);
      lines.push(`Xərc: ${money(p.expense_amount)} AZN`);
      lines.push(`Gəlir: ${money(p.income_amount)} AZN`);
    }
    return lines;
  }

  /* container: daxilində blokun göstəriləcəyi element
   * opts.threadId  — /resume üçün thread id
   * opts.interrupt — backend interrupt obyekti
   * opts.onDone    — (response) => void  resume nəticəsindən sonra çağırılır
   */
  async function renderInterruptBlock(container, opts) {
    const api = global.TaxProAPI;
    const threadId = opts.threadId;
    const interrupt = opts.interrupt || {};
    const onDone = opts.onDone || function () {};

    container.innerHTML =
      `<div class="hitl-block">` +
      `<div class="hitl-title">Təsdiq tələb olunur</div>` +
      `<div class="hitl-message">${escapeHtml(interrupt.message || "Əməliyyatı təsdiqləyin.")}</div>` +
      `<div class="hitl-preview">` +
      previewLines(interrupt).map((l) => `<div class="hitl-line">${escapeHtml(l)}</div>`).join("") +
      `</div>` +
      `<div class="hitl-actions">` +
      `<button type="button" class="btn hitl-cancel">Ləğv et</button>` +
      `<button type="button" class="btn primary hitl-confirm">Təsdiqlə</button>` +
      `</div>` +
      `</div>`;

    const confirmBtn = container.querySelector(".hitl-confirm");
    const cancelBtn = container.querySelector(".hitl-cancel");
    const setBusy = (busy) => {
      confirmBtn.disabled = busy;
      cancelBtn.disabled = busy;
      if (busy) confirmBtn.textContent = "Gözləyin…";
      else confirmBtn.textContent = "Təsdiqlə";
    };

    const act = async (decision) => {
      setBusy(true);
      try {
        const response = await api.resume(threadId, decision);
        onDone(response);
      } catch (err) {
        container.innerHTML =
          `<div class="hitl-block">` +
          `<div class="hitl-title">Xəta</div>` +
          `<div class="hitl-message">${escapeHtml(err.message)}</div>` +
          `</div>`;
      }
    };

    cancelBtn.addEventListener("click", () => act("cancel"));
    confirmBtn.addEventListener("click", () => act("confirm"));
  }

  global.TaxProHITL = {
    renderInterruptBlock,
    previewLines,
    money,
    escapeHtml,
  };
})(typeof window !== "undefined" ? window : globalThis);
