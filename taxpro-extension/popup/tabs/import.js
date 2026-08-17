/* TaxPro — İdxal tab-ı (standart Excel idxalı).
 * Görünmə yalnız `hidden` atributu ilə idarə olunur (class="hidden" yox!).
 */
(function () {
  "use strict";

  const api = window.TaxProAPI;

  const els = {};
  const $ = (id) => (els[id] = els[id] || document.getElementById(id));

  const MSG = {
    needFile: "Fayl seçin.",
    uploading: "Yüklənir…",
  };

  function escapeHtml(str) {
    return String(str == null ? "" : str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function showToast(message, type) {
    const toast = document.getElementById("toast");
    if (!toast) return;
    toast.textContent = message;
    toast.className = "toast " + (type || "");
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => {
      toast.className = "toast hidden";
    }, 3200);
  }

  function renderErrors(errors) {
    if (!errors || errors.length === 0) return "";
    const shown = errors.slice(0, 20);
    const extra = errors.length - shown.length;
    return (
      `<div class="import-errors">` +
      `<div class="card-label">Atlanan sətirlər</div>` +
      shown.map((e) => `<div class="import-error">${escapeHtml(e)}</div>`).join("") +
      (extra > 0 ? `<div class="import-error muted">…və ${extra} daha</div>` : "") +
      `</div>`
    );
  }

  function renderDoneExcel(body) {
    const inserted = body.inserted || 0;
    const errors = body.sources || [];
    const text =
      inserted > 0
        ? `${inserted} profil əlavə olundu.`
        : "0 profil əlavə olundu. Xəta detalları:";
    return (
      `<div class="result-box ${inserted > 0 ? "ok" : "err"}">` +
      `<div>${escapeHtml(text)}</div>` +
      `</div>` +
      renderErrors(errors)
    );
  }

  async function runUpload() {
    const file = $("importExcelFile").files[0];
    const btn = $("importExcelUploadBtn");
    const resultBox = $("importExcelResult");
    resultBox.hidden = true;
    if (!file) {
      showToast(MSG.needFile, "err");
      return;
    }
    btn.disabled = true;
    btn.textContent = MSG.uploading;
    try {
      const body = await api.importExcel(file);
      resultBox.innerHTML = renderDoneExcel(body);
    } catch (err) {
      resultBox.innerHTML = `<div class="result-box err">${escapeHtml(err.message)}</div>`;
    } finally {
      btn.disabled = false;
      btn.textContent = "Yüklə";
      resultBox.hidden = false;
    }
  }

  function init() {
    if (!window.TaxProAPI || !document.getElementById("importExcelFile")) return;
    $("importExcelUploadBtn").addEventListener("click", runUpload);
  }

  document.addEventListener("DOMContentLoaded", init);
})();
