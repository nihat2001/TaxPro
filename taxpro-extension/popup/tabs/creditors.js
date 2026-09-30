/* TaxPro — Kreditorlar/Debitorlar tab-ı (Gələn/Gedən). */
(function () {
  "use strict";

  const api = window.TaxProAPI;
  const MAX_BANK_FILES = 20;

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

  function bytes(n) {
    if (n < 1024) return n + " B";
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
    return (n / (1024 * 1024)).toFixed(1) + " MB";
  }

  /* ===== Çoxlu bank faylı seçicisi ===== */
  function createBankPicker(inputId, counterId, listId, clearId) {
    const input = $(inputId);
    const counter = $(counterId);
    const list = $(listId);
    const clearBtn = $(clearId);
    let files = [];

    function key(f) {
      return f.name + "|" + f.size + "|" + f.lastModified;
    }

    function render() {
      counter.textContent = "Seçilib: " + files.length + " fayl";
      list.innerHTML = "";
      files.forEach((f, i) => {
        const li = document.createElement("li");
        li.className = "file-chip";

        const name = document.createElement("span");
        name.className = "file-name";
        name.textContent = f.name; // XSS-safe (textContent)

        const size = document.createElement("span");
        size.className = "file-size";
        size.textContent = bytes(f.size);

        const rm = document.createElement("button");
        rm.type = "button";
        rm.className = "file-remove";
        rm.textContent = "\u00d7";
        rm.setAttribute("aria-label", "Sil");
        rm.addEventListener("click", () => {
          files.splice(i, 1);
          render();
        });

        li.appendChild(name);
        li.appendChild(size);
        li.appendChild(rm);
        list.appendChild(li);
      });
    }

    function add(newFiles) {
      const existing = new Set(files.map(key));
      let overflow = false;
      for (const f of newFiles) {
        if (existing.has(key(f))) continue; // təkrar əlavə olunmur
        if (files.length >= MAX_BANK_FILES) {
          overflow = true;
          break;
        }
        files.push(f);
        existing.add(key(f));
      }
      if (overflow) {
        showToast(
          "Maksimum " + MAX_BANK_FILES + " bank faylı seçilə bilər.",
          "err"
        );
      }
      render();
    }

    input.addEventListener("change", () => {
      if (input.files && input.files.length) {
        add(Array.from(input.files));
      }
      // Silinmiş eyni faylın yenidən seçilə bilməsi üçün input sıfırlanır
      input.value = "";
    });

    clearBtn.addEventListener("click", () => {
      files = [];
      render();
    });

    render();
    return { getFiles: () => files };
  }

  const credPicker = createBankPicker(
    "credBankFile", "credBankCounter", "credBankList", "credBankClear"
  );
  const debPicker = createBankPicker(
    "debBankFile", "debBankCounter", "debBankList", "debBankClear"
  );

  async function runGenerate() {
    const gelen = $("credGelenFile").files[0];
    const depozit = $("credDepozitFile").files[0];
    const banks = credPicker.getFiles();
    const btn = $("credGenerateBtn");
    const resultBox = $("credResult");
    resultBox.hidden = true;

    if (!gelen || !depozit || !banks.length) {
      showToast(
        "GƏLƏN, DEPOZİT və ən azı 1 bank çıxarışı seçin.",
        "err"
      );
      return;
    }

    btn.disabled = true;
    btn.textContent = "Yüklənir…";
    try {
      const blob = await api.generateCreditors(gelen, depozit, banks);
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

  async function runGenerateDebitors() {
    const gelen = $("debGedenFile").files[0];
    const depozit = $("debDepozitFile").files[0];
    const banks = debPicker.getFiles();
    const btn = $("debGenerateBtn");
    const resultBox = $("debResult");
    resultBox.hidden = true;

    if (!gelen || !depozit || !banks.length) {
      showToast(
        "GEDƏN, DEPOZİT və ən azı 1 bank çıxarışı seçin.",
        "err"
      );
      return;
    }

    btn.disabled = true;
    btn.textContent = "Yüklənir…";
    try {
      const blob = await api.generateDebitors(gelen, depozit, banks);
      downloadBlob(blob, "taxpro_debitorlar.xlsx");
      resultBox.innerHTML =
        '<div class="result-box ok">DEBİTORLAR workbook-u endirildi.</div>';
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
    if (document.getElementById("debGenerateBtn")) {
      $("debGenerateBtn").addEventListener("click", runGenerateDebitors);
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();
