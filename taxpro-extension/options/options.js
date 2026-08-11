/* TaxPro — options tənzimləmələri. */
"use strict";

(function () {
  const api = window.TaxProAPI;

  function setStatus(text, kind) {
    const el = document.getElementById("status");
    el.textContent = text;
    el.className = "status " + (kind || "muted");
  }

  async function init() {
    const base = await api.getBaseUrl();
    document.getElementById("baseUrl").value = base;

    document.getElementById("saveBtn").addEventListener("click", async () => {
      const value = document.getElementById("baseUrl").value.trim();
      if (!value) {
        setStatus("Ünvan boş ola bilməz.", "err");
        return;
      }
      await api.setBaseUrl(value);
      setStatus("Saxlandı.", "ok");
    });

    document.getElementById("testBtn").addEventListener("click", async () => {
      setStatus("Yoxlanılır…", "muted");
      const value = document.getElementById("baseUrl").value.trim();
      try {
        await api.setBaseUrl(value);
        const data = await api.health();
        if (data && data.status === "ok") {
          setStatus("Backend işləyir. ✓", "ok");
        } else {
          setStatus("Cavab gözlənilməzdir.", "err");
        }
      } catch (err) {
        setStatus(err.message || "Əlaqə yoxdur.", "err");
      }
    });
  }

  document.addEventListener("DOMContentLoaded", init);
})();
