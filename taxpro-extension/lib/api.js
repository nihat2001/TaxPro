/* TaxPro — backend fetch wrapper.
 * Bütün endpoint çağırışları bu faylda cəmlənir ki, backend ünvanı/endpoint
 * dəyişəndə yalnız bu fayl düzəlsin.
 */
(function (global) {
  "use strict";

  const DEFAULT_BASE_URL = "http://localhost:8000";
  const BASE_URL_KEY = "taxpro_base_url";

  let cachedBaseUrl = null;

  async function getBaseUrl() {
    if (cachedBaseUrl) return cachedBaseUrl;
    try {
      const stored = await chrome.storage.local.get(BASE_URL_KEY);
      cachedBaseUrl = stored[BASE_URL_KEY] || DEFAULT_BASE_URL;
    } catch (err) {
      cachedBaseUrl = DEFAULT_BASE_URL;
    }
    return cachedBaseUrl;
  }

  async function setBaseUrl(url) {
    cachedBaseUrl = (url || "").trim().replace(/\/+$/, "") || DEFAULT_BASE_URL;
    await chrome.storage.local.set({ [BASE_URL_KEY]: cachedBaseUrl });
    return cachedBaseUrl;
  }

  function base64ToJson(str) {
    const bin = atob(str);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return JSON.parse(new TextDecoder("utf-8").decode(bytes));
  }

  function base64ToBlob(b64, mime) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new Blob([bytes], { type: mime || "application/octet-stream" });
  }

  function apiError(res, data) {
    let detail = `HTTP ${res.status}`;
    if (data) {
      if (typeof data.detail === "string") detail = data.detail;
      else if (data.detail && typeof data.detail === "object") {
        const first = Array.isArray(data.detail) ? data.detail[0] : data.detail;
        if (first && first.msg) detail = first.msg;
      } else if (data.message) {
        detail = data.message;
      }
    }
    const err = new Error(detail);
    err.status = res.status;
    return err;
  }

  async function api(path, options) {
    options = options || {};
    const base = await getBaseUrl();
    const url = base + path;
    const headers = { ...(options.headers || {}) };
    let body = options.body;

    if (body && !(body instanceof FormData) && typeof body === "object") {
      headers["Content-Type"] = "application/json";
      body = JSON.stringify(body);
    }

    let res;
    try {
      res = await fetch(url, {
        method: options.method || "GET",
        headers,
        body,
      });
    } catch (err) {
      const wrapped = new Error(
        `Backend ünvanına çatmaq mümkün olmadı: ${base} (server işləyir?)`
      );
      wrapped.cause = err;
      throw wrapped;
    }

    const text = await res.text();
    let data = null;
    if (text) {
      try {
        data = JSON.parse(text);
      } catch (err) {
        data = text;
      }
    }

    if (!res.ok) throw apiError(res, data);
    return data;
  }

  const TaxProAPI = {
    DEFAULT_BASE_URL,
    getBaseUrl,
    setBaseUrl,

    /* ===== Sağlamlıq ===== */
    health: () => api("/health"),

    /* ===== Portallar ===== */
    getPortals: () => api("/api/v1/portals"),
    updateSelectorMapping: (portalId, data) =>
      api(`/api/v1/portals/${portalId}/selector-mapping`, {
        method: "PATCH",
        body: data,
      }),

    /* ===== Profillər ===== */
    listProfiles: () => api("/api/v1/profiles"),
    createProfile: (data) => api("/api/v1/profiles", { method: "POST", body: data }),
    updateProfile: (id, data) =>
      api(`/api/v1/profiles/${id}`, { method: "PATCH", body: data }),
    deleteProfile: (id) => api(`/api/v1/profiles/${id}`, { method: "DELETE" }),
    getProfileAutoFill: (id) => api(`/api/v1/profiles/${id}/autofill`),

    /* ===== Doldur (auto-fill) ===== */
    getPortalAutoFill: (portalId, profileId) =>
      api(`/api/v1/portals/${portalId}/autofill/${profileId}`),

    /* ===== İdxal (standart Excel) ===== */
    importExcel: (file) =>
      api("/api/v1/import/excel", { method: "POST", body: fileForm(file) }),
    downloadExcelTemplate: async () => {
      const base = await getBaseUrl();
      const res = await fetch(base + "/api/v1/import/excel-template");
      if (!res.ok) throw apiError(res, await res.text().catch(() => null));
      return res.blob();
    },

    /* ===== Bank çıxarışı kateqoriyalaşdırması (JSON-based axın) ===== */
    categorizeBankStatement: async (file) => {
      const base = await getBaseUrl();
      const res = await fetch(
        base + "/api/v1/import/bank-statement/categorize",
        { method: "POST", body: fileForm(file) }
      );
      if (!res.ok) {
        const text = await res.text().catch(() => null);
        let data = null;
        if (text) {
          try {
            data = JSON.parse(text);
          } catch (err) {
            data = text;
          }
        }
        throw apiError(res, data);
      }
      const data = await res.json();
      const blob = base64ToBlob(
        data.file_base64,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      return {
        blob,
        summary: data.summary || null,
        unclassified: Array.isArray(data.unclassified) ? data.unclassified : [],
      };
    },

    /* ===== Bank çıxarışı finalize (manual təsnifat override) ===== */
    finalizeBankStatement: async (file, selections) => {
      const base = await getBaseUrl();
      const fd = new FormData();
      fd.append("file", file);
      fd.append("selections", JSON.stringify(selections || {}));
      const res = await fetch(
        base + "/api/v1/import/bank-statement/finalize",
        { method: "POST", body: fd }
      );
      if (!res.ok) {
        const text = await res.text().catch(() => null);
        let data = null;
        if (text) {
          try {
            data = JSON.parse(text);
          } catch (err) {
            data = text;
          }
        }
        throw apiError(res, data);
      }
      const blob = await res.blob();
      let summary = null;
      const hdr = res.headers.get("X-Summary");
      if (hdr) {
        try {
          summary = base64ToJson(hdr);
        } catch (err) {
          summary = null;
        }
      }
      return { blob, summary };
    },
  };

  function fileForm(file) {
    const fd = new FormData();
    fd.append("file", file);
    return fd;
  }

  global.TaxProAPI = TaxProAPI;
})(typeof window !== "undefined" ? window : globalThis);
