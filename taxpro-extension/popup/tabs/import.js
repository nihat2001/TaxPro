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
      if ((body.inserted || 0) > 0) {
        window.dispatchEvent(new CustomEvent("taxpro:profiles-changed"));
      }
    } catch (err) {
      resultBox.innerHTML = `<div class="result-box err">${escapeHtml(err.message)}</div>`;
    } finally {
      btn.disabled = false;
      btn.textContent = "Yüklə";
      resultBox.hidden = false;
    }
  }

  async function downloadTemplate() {
    try {
      const blob = await api.downloadExcelTemplate();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "taxpro_profil_sablonu.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      showToast(err.message, "err");
    }
  }

  /* ===== Bank çıxarışı kateqoriyalaşdırması ===== */
  const CATEGORY_ORDER = [
    "Nağd çəkim komissiyası",
    "KOMİSİYA",
    "BANK FAİZ XƏRCİ",
    "ƏDV ÖDƏNİŞİ",
    "RÜSUM",
    "VERGİ MƏXARİC",
    "MAAŞ MƏXARİC",
    "TƏSİSÇİDƏN BORC",
    "DİGƏR ŞİRKƏT DAXİL OLMALAR",
    "DİGƏR ŞİRKƏT ÖDƏNİŞLƏRİ",
  ];

  function formatCem(value) {
    const n = Number(value);
    return (Number.isFinite(n) ? n : 0).toFixed(2);
  }

  function renderBankSummary(summary) {
    if (!summary) return "";
    const rows = CATEGORY_ORDER.map((name) => {
      const item = summary[name] || { cem: 0, say: 0 };
      return (
        `<div class="bank-summary-row">` +
        `<span class="bank-summary-name">${escapeHtml(name)}</span>` +
        `<span class="bank-summary-say">${item.say} əməliyyat</span>` +
        `<span class="bank-summary-cem">${escapeHtml(formatCem(item.cem))} AZN</span>` +
        `</div>`
      );
    }).join("");
    return (
      `<div class="result-box ok">` +
      `<div>Kateqoriyalaşdırılmış fayl endirildi. Xülasə:</div>` +
      `<div class="bank-summary">${rows}</div>` +
      `</div>`
    );
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  /* ===== Manual təsnifat (human-in-the-loop) ===== */
  let reviewFile = null;
  let reviewUnclassified = [];

  function reviewAmount(row) {
    const debit = Number(row.debit);
    const kredit = Number(row.kredit);
    if (Number.isFinite(debit)) return Math.abs(debit);
    if (Number.isFinite(kredit)) return Math.abs(kredit);
    return null;
  }

  function renderBankReview() {
    const list = $("bankReviewList");
    const rows = reviewUnclassified.map((u, i) => {
      const amount = reviewAmount(u);
      const amountText =
        amount === null ? "" : `${escapeHtml(formatCem(amount))} AZN`;
      const options = ["", ...CATEGORY_ORDER]
        .map(
          (c) =>
            `<option value="${escapeHtml(c)}">${c ? escapeHtml(c) : "Kateqoriyasız qalsın"}</option>`
        )
        .join("");
      return (
        `<div class="review-row" data-idx="${i}">` +
        `<div class="review-meta">` +
        `<div class="review-detail">${escapeHtml(u.detail || "—")}</div>` +
        `<div class="review-sender">${escapeHtml(u.sender || "")}</div>` +
        `</div>` +
        `<div class="review-amount">${amountText}</div>` +
        `<select class="review-select" data-field="review-${i}">${options}</select>` +
        `</div>`
      );
    }).join("");
    list.innerHTML = rows;
    $("bankReviewTitle").textContent =
      `${reviewUnclassified.length} əməliyyat tanınmadı, nəzərdən keçirin.`;
    $("bankReview").hidden = false;
  }

  function collectSelections() {
    const selections = {};
    reviewUnclassified.forEach((u, i) => {
      const sel = document.querySelector(`select[data-field="review-${i}"]`);
      const value = sel ? sel.value : "";
      selections[u.row_index] = value;
    });
    return selections;
  }

  async function confirmBankReview() {
    if (!reviewFile) return;
    const btn = $("bankReviewConfirmBtn");
    const resultBox = $("bankStatementResult");
    btn.disabled = true;
    btn.textContent = "Yüklənir…";
    try {
      const selections = collectSelections();
      const { blob, summary } = await api.finalizeBankStatement(reviewFile, selections);
      downloadBlob(blob, "taxpro_bank_cizarsisi_yekun.xlsx");
      $("bankReview").hidden = true;
      resultBox.innerHTML = renderBankSummary(summary);
    } catch (err) {
      resultBox.innerHTML = `<div class="result-box err">${escapeHtml(err.message)}</div>`;
    } finally {
      btn.disabled = false;
      btn.textContent = "Təsdiqlə və yenidən yüklə";
    }
  }

  async function runBankStatement() {
    const file = $("bankStatementFile").files[0];
    const btn = $("bankStatementBtn");
    const resultBox = $("bankStatementResult");
    resultBox.hidden = true;
    $("bankReview").hidden = true;
    if (!file) {
      showToast("Çıxarış faylı seçin.", "err");
      return;
    }
    btn.disabled = true;
    btn.textContent = "Yüklənir…";
    try {
      const { blob, summary, unclassified } = await api.categorizeBankStatement(file);
      reviewFile = file;
      reviewUnclassified = Array.isArray(unclassified) ? unclassified : [];

      if (reviewUnclassified.length === 0) {
        downloadBlob(blob, "taxpro_bank_cizarsisi_kateqoriyalasmis.xlsx");
        resultBox.innerHTML = renderBankSummary(summary);
      } else {
        renderBankReview();
        resultBox.innerHTML =
          `<div class="result-box ok">Kateqoriyalaşdırma hazırdır. ` +
          `${reviewUnclassified.length} əməliyyat nəzərdən keçirmə tələb edir.</div>`;
      }
    } catch (err) {
      resultBox.innerHTML = `<div class="result-box err">${escapeHtml(err.message)}</div>`;
    } finally {
      btn.disabled = false;
      btn.textContent = "Kateqoriyalaşdır və endir";
      resultBox.hidden = false;
    }
  }

  function init() {
    if (!window.TaxProAPI || !document.getElementById("importExcelFile")) return;
    $("importExcelUploadBtn").addEventListener("click", runUpload);
    $("importTemplateBtn").addEventListener("click", (e) => {
      e.preventDefault();
      downloadTemplate();
    });
    if (document.getElementById("bankStatementBtn")) {
      $("bankStatementBtn").addEventListener("click", runBankStatement);
    }
    if (document.getElementById("bankReviewConfirmBtn")) {
      $("bankReviewConfirmBtn").addEventListener("click", confirmBankReview);
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();
