/* TaxPro — Kreditorlar/Debitorlar tab-ı (Gələn). */
(function () {
  "use strict";

  const api = window.TaxProAPI;

  const els = {};
  const $ = (id) => (els[id] = els[id] || document.getElementById(id));

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

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function runGenerate() {
    const gelen = $("credGelenFile").files[0];
    const depozit = $("credDepozitFile").files[0];
    const bank = $("credBankFile").files[0];
    const btn = $("credGenerateBtn");
    const resultBox = $("credResult");
    resultBox.hidden = true;

    if (!gelen || !depozit || !bank) {
      showToast("3 faylı da seçin (GƏLƏN, DEPOZİT, BANK ÇIXARIŞI).", "err");
      return;
    }

    btn.disabled = true;
    btn.textContent = "Yüklənir…";
    try {
      const blob = await api.generateCreditors(gelen, depozit, bank);
      downloadBlob(blob, "taxpro_kreditorlar.xlsx");
      resultBox.innerHTML =
        '<div class="result-box ok">KREDİTORLAR workbook-u endirildi.</div>';
    } catch (err) {
      resultBox.innerHTML =
        `<div class="result-box err">${escapeHtml(err.message)}</div>`;
    } finally {
      btn.disabled = false;
      btn.textContent = "Generasiya et";
      resultBox.hidden = false;
    }
  }

  function init() {
    if (!window.TaxProAPI || !document.getElementById("credGenerateBtn")) return;
    $("credGenerateBtn").addEventListener("click", runGenerate);
  }

  document.addEventListener("DOMContentLoaded", init);
})();
