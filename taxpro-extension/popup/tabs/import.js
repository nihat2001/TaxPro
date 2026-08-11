/* TaxPro — İdxal tab-ı (Excel idxal + Bank çıxarışı analizi).
 * Görünmə yalnız `hidden` atributu ilə idarə olunur (class="hidden" yox!).
 */
(function () {
  "use strict";

  const api = window.TaxProAPI;
  const hitl = window.TaxProHITL;

  const els = {};
  const $ = (id) => (els[id] = els[id] || document.getElementById(id));

  const MSG = {
    needFile: "Fayl seçin.",
    uploading: "Yüklənir…",
    noError: "Xəta yoxdur.",
  };

  function escapeHtml(str) {
    return String(str == null ? "" : str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function money(v) {
    return hitl ? hitl.money(v) : String(v);
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

  function setBusy(btn, busy, label) {
    btn.disabled = busy;
    btn.textContent = busy ? MSG.uploading : label;
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
    const errors = body.sources || [];
    const text = body.message || `${body.inserted || 0} profil əlavə olundu.`;
    return (
      `<div class="result-box ok">` +
      `<div>${escapeHtml(text)}</div>` +
      `</div>` +
      renderErrors(errors)
    );
  }

  function renderDoneBank(body) {
    const a = body.analysis || {};
    const cats = a.categories || {};
    const catNames = {
      edv: "ƏDV",
      xerc: "Xərc",
      medaxil: "Gəlir",
      diger: "Digər",
    };
    return (
      `<div class="result-box ok bank-result">` +
      `<div class="bank-grid">` +
      `<div class="bank-cell"><span class="card-label">Ümumi</span><span class="bank-val">${money(a.total_amount)}</span></div>` +
      `<div class="bank-cell"><span class="card-label">ƏDV</span><span class="bank-val">${money(a.vat_amount)}</span></div>` +
      `<div class="bank-cell"><span class="card-label">Xərc</span><span class="bank-val">${money(a.expense_amount)}</span></div>` +
      `<div class="bank-cell"><span class="card-label">Gəlir</span><span class="bank-val">${money(a.income_amount)}</span></div>` +
      `</div>` +
      (Object.keys(cats).length
        ? `<div class="bank-cats">` +
          Object.entries(cats)
            .map(
              ([k, v]) =>
                `<div class="bank-cat"><span>${escapeHtml(catNames[k] || k)}</span><span>${money(v)}</span></div>`
            )
            .join("") +
          `</div>`
        : "") +
      `</div>`
    );
  }

  async function runUpload(endpoint, file, resultBox, btn, renderDone) {
    resultBox.hidden = true;
    if (!file) {
      showToast(MSG.needFile, "err");
      return;
    }
    setBusy(btn, true);
    try {
      const body = await endpoint(file);
      if (body.status === "interrupt") {
        hitl.renderInterruptBlock(resultBox, {
          threadId: body.thread_id,
          interrupt: body.interrupt,
          onDone: (res) => {
            if (res.status === "interrupt") {
              hitl.renderInterruptBlock(resultBox, {
                threadId: res.thread_id,
                interrupt: res.interrupt,
                onDone: (res2) => runUploadFinal(resultBox, res2, renderDone),
              });
            } else {
              runUploadFinal(resultBox, res, renderDone);
            }
          },
        });
      } else if (body.status === "done") {
        resultBox.innerHTML = renderDone(body);
      } else {
        resultBox.innerHTML =
          `<div class="result-box err">${escapeHtml(body.message || "Uğursuz oldu.")}</div>`;
      }
    } catch (err) {
      resultBox.innerHTML = `<div class="result-box err">${escapeHtml(err.message)}</div>`;
    } finally {
      setBusy(btn, false);
      resultBox.hidden = false;
    }
  }

  function runUploadFinal(resultBox, body, renderDone) {
    if (body.status === "done") {
      resultBox.innerHTML = renderDone(body);
    } else {
      resultBox.innerHTML =
        `<div class="result-box err">${escapeHtml(body.message || "Uğursuz oldu.")}</div>`;
    }
  }

  function switchImportMode(name) {
    $("importExcelPanel").hidden = name !== "excel";
    $("importBankPanel").hidden = name !== "bank";
    $("importExcelBtn").classList.toggle("active", name === "excel");
    $("importBankBtn").classList.toggle("active", name === "bank");
  }

  function init() {
    if (!window.TaxProAPI || !document.getElementById("importExcelFile")) return;

    $("importExcelBtn").addEventListener("click", () => switchImportMode("excel"));
    $("importBankBtn").addEventListener("click", () => switchImportMode("bank"));

    $("importSmartToggle").addEventListener("change", (e) => {
      $("importSmartHint").hidden = !e.target.checked;
    });

    $("importExcelUploadBtn").addEventListener("click", async () => {
      const file = $("importExcelFile").files[0];
      const smart = $("importSmartToggle").checked;
      await runUpload(
        smart ? api.importSmart.bind(api) : api.importExcel.bind(api),
        file,
        $("importExcelResult"),
        $("importExcelUploadBtn"),
        renderDoneExcel
      );
    });

    $("importBankUploadBtn").addEventListener("click", async () => {
      const file = $("importBankFile").files[0];
      await runUpload(
        api.analyzeBankStatement.bind(api),
        file,
        $("importBankResult"),
        $("importBankUploadBtn"),
        renderDoneBank
      );
    });

    $("importBankFile").addEventListener("change", (e) => {
      const name = e.target.files[0] ? e.target.files[0].name : "";
      $("importBankFileName").textContent = name;
      $("importBankFileName").hidden = !name;
    });

    switchImportMode("excel");
  }

  document.addEventListener("DOMContentLoaded", init);
})();
