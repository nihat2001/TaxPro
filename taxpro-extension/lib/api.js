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

    /* ===== Faza 2 (AI) — yer saxlanılır ===== */
    importExcel: (file) =>
      api("/api/v1/import/excel", { method: "POST", body: file }),
    importSmart: (file) =>
      api("/api/v1/import/smart", { method: "POST", body: file }),
    analyzeBankStatement: (file) =>
      api("/api/v1/bank-statements/analyze", { method: "POST", body: file }),
    chat: (message, profileId) =>
      api("/api/v1/chat", { method: "POST", body: { message, profile_id: profileId || null } }),
    resume: (threadId, decision) =>
      api("/api/v1/resume", { method: "POST", body: { thread_id: threadId, decision } }),
  };

  global.TaxProAPI = TaxProAPI;
})(typeof window !== "undefined" ? window : globalThis);
