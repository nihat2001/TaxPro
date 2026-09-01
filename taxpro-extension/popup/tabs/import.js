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
    "Terminal-Komissiya",
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
    // Blob URL-ni dərhal revoke etmək Chrome-da download-ı kəsə bilər.
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  /* ===== Manual təsnifat (human-in-the-loop) ===== */
  function reviewAmount(row) {
    const debit = Number(row.debit);
    const kredit = Number(row.kredit);
    if (Number.isFinite(debit)) return Math.abs(debit);
    if (Number.isFinite(kredit)) return Math.abs(kredit);
    return null;
  }

  function renderBankReview(state, prefix) {
    const list = $(prefix + "ReviewList");
    const rows = state.unclassified.map((u, i) => {
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
        `<select class="review-select" data-field="${prefix}-review-${i}">${options}</select>` +
        `</div>`
      );
    }).join("");
    list.innerHTML = rows;
    $(prefix + "ReviewTitle").textContent =
      `${state.unclassified.length} əməliyyat tanınmadı, nəzərdən keçirin.`;
    $(prefix + "Review").hidden = false;
  }

  function collectSelections(state, prefix) {
    const selections = {};
    const list = $(prefix + "ReviewList");
    state.unclassified.forEach((u, i) => {
      const sel = list.querySelector(`select[data-field="${prefix}-review-${i}"]`);
      const value = sel ? sel.value : "";
      selections[u.row_index] = value;
    });
    return selections;
  }

  /* Hər bank kartı öz prefix + state ilə bağlanır. */
  function bindBankCard(prefix, downloadName) {
    const state = { file: null, unclassified: [] };

    async function confirmReview() {
      if (!state.file) return;
      const btn = $(prefix + "ReviewConfirmBtn");
      const resultBox = $(prefix + "StatementResult");
      btn.disabled = true;
      btn.textContent = "Yüklənir…";
      try {
        const selections = collectSelections(state, prefix);
        const { blob, summary } = await api.finalizeBankStatement(state.file, selections);
        downloadBlob(blob, downloadName.replace(".xlsx", "_yekun.xlsx"));
        $(prefix + "Review").hidden = true;
        resultBox.innerHTML = renderBankSummary(summary);
      } catch (err) {
        resultBox.innerHTML = `<div class="result-box err">${escapeHtml(err.message)}</div>`;
      } finally {
        btn.disabled = false;
        btn.textContent = "Təsdiqlə və yenidən yüklə";
      }
    }

    async function run() {
      const file = $(prefix + "StatementFile").files[0];
      const btn = $(prefix + "StatementBtn");
      const resultBox = $(prefix + "StatementResult");
      resultBox.hidden = true;
      $(prefix + "Review").hidden = true;
      if (!file) {
        showToast("Çıxarış faylı seçin.", "err");
        return;
      }
      btn.disabled = true;
      btn.textContent = "Yüklənir…";
      try {
        const { blob, summary, unclassified } = await api.categorizeBankStatement(file);
        state.file = file;
        state.unclassified = Array.isArray(unclassified) ? unclassified : [];

        if (state.unclassified.length === 0) {
          downloadBlob(blob, downloadName);
          resultBox.innerHTML = renderBankSummary(summary);
        } else {
          renderBankReview(state, prefix);
          resultBox.innerHTML =
            `<div class="result-box ok">Kateqoriyalaşdırma hazırdır. ` +
            `${state.unclassified.length} əməliyyat nəzərdən keçirmə tələb edir.</div>`;
        }
      } catch (err) {
        resultBox.innerHTML = `<div class="result-box err">${escapeHtml(err.message)}</div>`;
      } finally {
        btn.disabled = false;
        btn.textContent = "Kateqoriyalaşdır və endir";
        resultBox.hidden = false;
      }
    }

    if (document.getElementById(prefix + "StatementBtn")) {
      $(prefix + "StatementBtn").addEventListener("click", run);
    }
    if (document.getElementById(prefix + "ReviewConfirmBtn")) {
      $(prefix + "ReviewConfirmBtn").addEventListener("click", confirmReview);
    }
  }

  function init() {
    if (!window.TaxProAPI || !document.getElementById("importExcelFile")) return;
    $("importExcelUploadBtn").addEventListener("click", runUpload);
    $("importTemplateBtn").addEventListener("click", (e) => {
      e.preventDefault();
      downloadTemplate();
    });
    bindBankCard("bank", "taxpro_bank_cizarsisi_kateqoriyalasmis.xlsx");
    bindBankCard("kapital", "taxpro_kapital_bank_cizarsisi.xlsx");
    bindBankCard("yelo", "taxpro_yelo_bank_cizarsisi.xlsx");
    bindBankCard("access", "taxpro_access_bank_cizarsisi.xlsx");
    bindBankCard("vtb", "taxpro_vtb_bank_cizarsisi.xlsx");
  }

  document.addEventListener("DOMContentLoaded", init);
})();
