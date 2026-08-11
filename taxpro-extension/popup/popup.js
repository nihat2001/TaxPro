/* TaxPro — popup UI məntiqi. */
"use strict";

(function () {
  const api = window.TaxProAPI;

  let portals = [];
  let profiles = [];
  let editingProfileId = null;
  let currentView = "tabs"; // "tabs" | "profileForm"

  let recordActive = false;
  let recordPortalId = null;
  let recordMap = {};
  let recordPendingSelector = null;

  let activePortal = null;

  const els = {};
  const $ = (id) => (els[id] = els[id] || document.getElementById(id));

  /* ===== Köməkçilər ===== */
  function initials(name) {
    return String(name || "?")
      .trim()
      .split(/\s+/)
      .slice(0, 2)
      .map((w) => w[0] || "")
      .join("")
      .toUpperCase();
  }

  function hostOf(url) {
    try {
      return new URL(url).hostname.replace(/^www\./, "");
    } catch (err) {
      return "";
    }
  }

  function escapeHtml(str) {
    return String(str == null ? "" : str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function showToast(message, type) {
    const toast = $("toast");
    toast.textContent = message;
    toast.className = "toast " + (type || "");
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => {
      toast.className = "toast hidden";
    }, 3200);
  }

  function setStatus(ok, text) {
    $("apiStatus").className = "status-dot " + (ok ? "ok" : "err");
    $("apiStatusText").textContent = text;
  }

  function switchTab(name) {
    document.querySelectorAll(".tab-btn").forEach((b) => {
      b.classList.toggle("active", b.dataset.tab === name);
    });
    document.querySelectorAll(".tab-panel").forEach((p) => {
      p.classList.toggle("active", p.id === "tab-" + name);
    });
  }

  /* ===== Saytlar tabı ===== */
  function renderSites(filter) {
    const list = $("sitesList");
    filter = (filter || "").trim().toLowerCase();
    const shown = portals.filter((p) => {
      if (!filter) return true;
      return (
        p.name.toLowerCase().includes(filter) ||
        hostOf(p.url).toLowerCase().includes(filter)
      );
    });

    if (shown.length === 0) {
      list.innerHTML = '<div class="empty-state">Portal tapılmadı.</div>';
      return;
    }

    const groups = [
      { key: "dövlət", label: "Dövlət portalları" },
      { key: "bank", label: "Banklar" },
      { key: "__other", label: "Digər" },
    ];

    let html = "";
    for (const g of groups) {
      const items =
        g.key === "__other"
          ? shown.filter((p) => !["dövlət", "bank"].includes(String(p.category).toLowerCase()))
          : shown.filter((p) => String(p.category).toLowerCase() === g.key);
      if (items.length === 0) continue;
      html += `<div class="sites-group-title">${g.label}</div>`;
      for (const p of items) {
        html +=
          `<a class="site-row" href="${escapeHtml(p.url)}" target="_blank" rel="noopener noreferrer">` +
          `<span class="site-avatar">${escapeHtml(initials(p.name))}</span>` +
          `<span class="site-meta">` +
          `<span class="site-name">${escapeHtml(p.name)}</span>` +
          `</span></a>`;
      }
    }
    list.innerHTML = html;
  }

  function initSites() {
    $("siteSearch").addEventListener("input", (e) => renderSites(e.target.value));
  }

  /* ===== Profillər tabı ===== */
  function renderProfiles(filter) {
    const list = $("profilesList");
    filter = (filter || "").trim().toLowerCase();
    const shown = profiles.filter((p) => {
      if (!filter) return true;
      return (
        p.ad_soyad.toLowerCase().includes(filter) ||
        String(p.voen || "").includes(filter) ||
        String(p.asan_id || "").toLowerCase().includes(filter)
      );
    });

    if (shown.length === 0) {
      list.innerHTML = '<div class="empty-state">Profil tapılmadı.</div>';
      return;
    }

    let html = "";
    for (const p of shown) {
      const color = p.reng || "#3b82f6";
      html +=
        `<div class="profile-card" data-id="${p.id}">` +
        `<span class="profile-avatar" style="background:${escapeHtml(color)}">${escapeHtml(initials(p.ad_soyad))}</span>` +
        `<div class="profile-meta">` +
        `<div class="profile-name">${escapeHtml(p.ad_soyad)}</div>` +
        `<div class="profile-sub">VÖEN: ${escapeHtml(p.voen)} · Asan ID: ${escapeHtml(p.asan_id)}</div>` +
        `</div>` +
        `<div class="profile-actions">` +
        `<button type="button" class="icon-btn" data-act="edit" title="Redaktə">&#9998;</button>` +
        `<button type="button" class="icon-btn" data-act="delete" title="Sil">&#128465;</button>` +
        `</div></div>`;
    }
    list.innerHTML = html;
  }

  function initProfiles() {
    $("profileSearch").addEventListener("input", (e) => renderProfiles(e.target.value));
    $("newProfileBtn").addEventListener("click", () => openProfileModal(null));

    $("profilesList").addEventListener("click", async (e) => {
      const btn = e.target.closest("[data-act]");
      if (!btn) return;
      const card = btn.closest(".profile-card");
      const id = card.dataset.id;
      const profile = profiles.find((p) => p.id === id);
      if (!profile) return;

      if (btn.dataset.act === "edit") {
        openProfileModal(profile);
      } else if (btn.dataset.act === "delete") {
        if (!confirm(`"${profile.ad_soyad}" profilini silmək istəyirsiniz?`)) return;
        try {
          await api.deleteProfile(id);
          profiles = profiles.filter((p) => p.id !== id);
          renderProfiles($("profileSearch").value);
          showToast("Profil silindi.", "ok");
        } catch (err) {
          showToast(err.message, "err");
        }
      }
    });
  }

  function openProfileModal(profile) {
    editingProfileId = profile ? profile.id : null;
    $("profileModalTitle").textContent = profile ? "Profili Redaktə Et" : "Yeni Profil";

    $("pfId").value = profile ? profile.id : "";
    $("pfAdSoyad").value = profile ? profile.ad_soyad : "";
    $("pfVoen").value = profile ? profile.voen : "";
    $("pfNomre").value = profile ? profile.nomre : "";
    $("pfAsanId").value = profile ? profile.asan_id : "";
    $("pfPin1").value = "";
    $("pfPin2").value = "";
    $("pfPuk").value = "";
    $("pfBitmeTarixi").value = profile && profile.bitme_tarixi ? profile.bitme_tarixi : "";
    $("pfQeyd").value = profile && profile.qeyd ? profile.qeyd : "";
    $("pfReng").value = profile && profile.reng ? profile.reng : "#3b82f6";
    $("pfIkon").value = profile && profile.ikon ? profile.ikon : "user";

    const pinHints = document.querySelectorAll("#profileForm .hint");
    pinHints.forEach((h) => {
      h.style.display = profile ? "" : "none";
    });

    currentView = "profileForm";
    $("profileModal").hidden = false;
    $("pfAdSoyad").focus();
  }

  function closeProfileModal() {
    currentView = "tabs";
    $("profileModal").hidden = true;
    editingProfileId = null;
  }

  function collectForm() {
    const v = (id) => $(id).value;
    const data = {
      ad_soyad: v("pfAdSoyad").trim(),
      voen: v("pfVoen").trim(),
      nomre: v("pfNomre").trim(),
      asan_id: v("pfAsanId").trim(),
      reng: v("pfReng") || "#3b82f6",
      ikon: v("pfIkon") || "user",
    };

    const pin1 = v("pfPin1");
    const pin2 = v("pfPin2");
    const puk = v("pfPuk");
    if (pin1) data.pin1 = pin1;
    if (pin2) data.pin2 = pin2;
    if (puk) data.puk = puk;

    if (v("pfBitmeTarixi")) data.bitme_tarixi = v("pfBitmeTarixi");
    if (v("pfQeyd").trim()) data.qeyd = v("pfQeyd").trim();

    return data;
  }

  function initProfileModal() {
    $("profileCancelBtn").addEventListener("click", closeProfileModal);
    $("profileModal").addEventListener("click", (e) => {
      if (e.target === $("profileModal")) closeProfileModal();
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && currentView === "profileForm") {
        closeProfileModal();
      }
    });

    $("profileForm").addEventListener("submit", async (e) => {
      e.preventDefault();
      const data = collectForm();

      if (!editingProfileId && (!data.pin1 || !data.pin2 || !data.puk)) {
        showToast("PIN1, PIN2 və PUK tələb olunur.", "err");
        return;
      }

      try {
        if (editingProfileId) {
          await api.updateProfile(editingProfileId, data);
          showToast("Profil yeniləndi.", "ok");
        } else {
          await api.createProfile(data);
          showToast("Profil yaradıldı.", "ok");
        }
        closeProfileModal();
        await loadProfiles();
        await populateAutofillSelects();
      } catch (err) {
        showToast(err.message, "err");
      }
    });
  }

  /* ===== Doldur tabı ===== */
  function populateProfilesSelect() {
    const sel = $("autofillProfile");
    const current = sel.value;
    sel.innerHTML =
      '<option value="">— Profil seçin —</option>' +
      profiles
        .map((p) => `<option value="${p.id}">${escapeHtml(p.ad_soyad)}</option>`)
        .join("");
    if (current) sel.value = current;
    updateAutofillBtn();
  }

  const STOPWORDS = new Set([
    "bank", "banks", "az", "gov", "com", "net", "org", "ltd", "llc",
    "official", "portal", "portalı",
  ]);

  function normalizeHost(host) {
    return String(host || "")
      .toLowerCase()
      .replace(/^www\./, "")
      .replace(/social/g, "sosial");
  }

  function portalKeywords(p) {
    const keys = new Set();
    let host = "";
    try {
      host = new URL(p.url).hostname.toLowerCase().replace(/^www\./, "");
    } catch (err) {
      host = "";
    }

    const hostNorm = normalizeHost(host);
    keys.add(hostNorm);
    const hostNoTld = hostNorm.split(".").slice(0, -1).join(".");
    if (hostNoTld) keys.add(hostNoTld);
    for (const seg of hostNorm.split(/[.-]/)) {
      if (seg.length >= 4 && !STOPWORDS.has(seg)) keys.add(seg);
    }

    const nameTokens = String(p.name || "")
      .toLowerCase()
      .split(/[^a-z0-9əöüçşğı]+/);
    for (const t of nameTokens) {
      const tok = t.replace(/social/g, "sosial");
      if (tok.length >= 4 && !STOPWORDS.has(tok)) keys.add(tok);
    }
    return keys;
  }

  function matchPortalByUrl(tabUrl) {
    if (!tabUrl) return null;
    let tabHost = "";
    let tabPath = "";
    try {
      const u = new URL(tabUrl);
      tabHost = normalizeHost(u.hostname);
      tabPath = u.pathname.toLowerCase();
    } catch (err) {
      return null;
    }
    if (!tabHost) return null;

    let best = null;
    for (const p of portals) {
      let ph = "";
      let pp = "";
      try {
        const pu = new URL(p.url);
        ph = normalizeHost(pu.hostname);
        pp = (pu.pathname || "/").toLowerCase();
      } catch (err) {
        continue;
      }
      if (!ph) continue;

      let score = 0;
      if (tabHost === ph) {
        score = 4;
      } else if (tabHost.endsWith("." + ph)) {
        score = 3;
      } else if (p._keywords && [...p._keywords].some((k) => tabHost.includes(k))) {
        score = 2;
      }

      if (score === 0) continue;

      const hasDeepPath = pp !== "/" && pp !== "";
      if (hasDeepPath && tabPath.startsWith(pp)) score += 1;

      if (!best || score > best.score) {
        best = { portal: p, score };
      }
    }
    return best ? best.portal : null;
  }

  function updateAutofillBtn() {
    const profileOk = !!$("autofillProfile").value;
    const nonePortal = activePortal && activePortal.auth_type === "none";
    $("autofillBtn").disabled = !profileOk || !!nonePortal;
    if (nonePortal) {
      const resultBox = $("autofillResult");
      resultBox.className = "result-box err";
      resultBox.textContent = "Bu sayt üçün auto-fill tələb olunmur/dəstəklənmir.";
      resultBox.classList.remove("hidden");
    }
  }

  async function loadActiveTab() {
    const info = $("activeTabInfo");
    try {
      const resp = await chrome.runtime.sendMessage({ type: "GET_ACTIVE_TAB" });
      const tab = resp && resp.tab;
      if (!tab) {
        info.textContent = "Aktiv tab tapılmadı.";
        info.className = "active-tab-info muted";
        return;
      }
      const matched = matchPortalByUrl(tab.url);
      activePortal = matched || null;
      if (matched) {
        info.textContent = `${matched.name} (${hostOf(tab.url)})`;
        info.className = "active-tab-info";
      } else {
        info.textContent = `${tab.title || hostOf(tab.url) || "Səhifə"} — tanınmış portal deyil`;
        info.className = "active-tab-info muted";
      }
      updateAutofillBtn();
    } catch (err) {
      info.textContent = "Aktiv tab oxunmadı.";
      info.className = "active-tab-info muted";
    }
  }

  function initAutofill() {
    $("autofillProfile").addEventListener("change", updateAutofillBtn);

    $("autofillBtn").addEventListener("click", async () => {
      const profileId = $("autofillProfile").value;
      const resultBox = $("autofillResult");
      resultBox.className = "result-box hidden";

      if (!profileId) return;

      if (activePortal && activePortal.auth_type === "none") {
        resultBox.className = "result-box err";
        resultBox.textContent = "Bu sayt üçün auto-fill tələb olunmur/dəstəklənmir.";
        resultBox.classList.remove("hidden");
        return;
      }

      let activeTab = null;
      try {
        const resp = await chrome.runtime.sendMessage({ type: "GET_ACTIVE_TAB" });
        activeTab = resp && resp.tab;
      } catch (err) {
        activeTab = null;
      }

      const portal = activeTab ? matchPortalByUrl(activeTab.url) : null;
      if (!portal) {
        resultBox.className = "result-box err";
        resultBox.textContent =
          "Zəhmət olmasa, əvvəlcə Saytlar tabından portalın daxilolma səhifəsinə keçin.";
        resultBox.classList.remove("hidden");
        return;
      }

      try {
        const payload = await api.getPortalAutoFill(portal.id, profileId);
        const fields = Object.entries(payload.fields || {}).map(
          ([selector, value]) => ({
            field: (payload.field_map && payload.field_map[selector]) || selector,
            selector,
            value,
          })
        );
        if (fields.length === 0) {
          resultBox.className = "result-box err";
          resultBox.textContent = "Bu portal üçün doldurulacaq sahə yoxdur.";
          resultBox.classList.remove("hidden");
          return;
        }

        const outcome = await chrome.runtime.sendMessage({
          type: "AUTOFILL",
          fields,
        });

        if (!outcome.ok) {
          resultBox.className = "result-box err";
          resultBox.textContent = outcome.error || "Doldurma uğursuz oldu.";
          resultBox.classList.remove("hidden");
          return;
        }

        const r = outcome.result;
        const missingCount = r.missing ? r.missing.length : 0;
        let text = `${r.filled}/${r.total} sahə dolduruldu.`;
        if (missingCount > 0) {
          text += ` ${missingCount} sahə tapılmadı (sayt DOM-u fərqli ola bilər).`;
        }
        resultBox.className = "result-box " + (missingCount === 0 ? "ok" : "err");
        resultBox.textContent = text;
        resultBox.classList.remove("hidden");
      } catch (err) {
        resultBox.className = "result-box err";
        resultBox.textContent = err.message;
        resultBox.classList.remove("hidden");
      }
    });
  }

  async function populateAutofillSelects() {
    populateProfilesSelect();
  }

  /* ===== Qeydiyyat rejimi (Record Mode) ===== */
  const RECORD_FIELD_LABELS = {
    asan_id: "Asan ID",
    nomre: "Telefon",
    voen: "VÖEN",
    ad_soyad: "Ad Soyad",
    pin1: "PIN1",
    pin2: "PIN2",
    puk: "PUK",
    qeyd: "Qeyd",
  };

  function renderRecordList() {
    const list = $("recordList");
    const fields = Object.keys(recordMap);
    if (fields.length === 0) {
      list.innerHTML = '<div class="empty-state">Hələ sahə tutulmayıb.</div>';
      return;
    }
    list.innerHTML = fields
      .map(
        (f) =>
          `<div class="record-row" data-field="${escapeHtml(f)}">` +
          `<span class="record-row-field">${escapeHtml(RECORD_FIELD_LABELS[f] || f)}</span>` +
          `<code class="record-row-selector">${escapeHtml(recordMap[f])}</code>` +
          `<button type="button" class="icon-btn" data-remove="${escapeHtml(f)}" title="Sil">&#10005;</button>` +
          `</div>`
      )
      .join("");
  }

  async function startRecordMode() {
    if (recordActive) return;
    let activeTab = null;
    try {
      const resp = await chrome.runtime.sendMessage({ type: "GET_ACTIVE_TAB" });
      activeTab = resp && resp.tab;
    } catch (err) {
      activeTab = null;
    }
    const portal = activeTab ? matchPortalByUrl(activeTab.url) : null;
    if (!portal) {
      showToast("Əvvəlcə portalın daxilolma səhifəsində olun.", "err");
      return;
    }

    const resp = await chrome.runtime.sendMessage({ type: "START_RECORD" });
    if (!resp || !resp.ok) {
      showToast((resp && resp.error) || "Qeydiyyat rejimi başlamadı.", "err");
      return;
    }

    recordActive = true;
    recordPortalId = portal.id;
    recordMap = {};
    recordPendingSelector = null;
    $("recordPortalName").textContent = portal.name;
    $("recordPanel").hidden = false;
    $("recordPending").hidden = true;
    $("recordModeBtn").textContent = "🎯 Qeydiyyat rejimi: Aktiv";
    renderRecordList();
  }

  async function stopRecordMode() {
    if (recordActive) {
      try {
        await chrome.runtime.sendMessage({ type: "STOP_RECORD" });
      } catch (err) {
        // background cavabı vacib deyil
      }
    }
    recordActive = false;
    recordPortalId = null;
    recordMap = {};
    recordPendingSelector = null;
    $("recordPanel").hidden = true;
    $("recordModeBtn").textContent = "🎯 Qeydiyyat rejimi";
  }

  function initRecordMode() {
    $("recordModeBtn").addEventListener("click", () => {
      if (recordActive) {
        stopRecordMode();
      } else {
        startRecordMode();
      }
    });

    $("recordAddBtn").addEventListener("click", () => {
      if (!recordPendingSelector) return;
      const field = $("recordField").value;
      recordMap[field] = recordPendingSelector;
      recordPendingSelector = null;
      $("recordPending").hidden = true;
      renderRecordList();
      showToast("Sahə əlavə olundu.", "ok");
    });

    $("recordList").addEventListener("click", (e) => {
      const btn = e.target.closest("[data-remove]");
      if (!btn) return;
      delete recordMap[btn.dataset.remove];
      renderRecordList();
    });

    $("recordCancelBtn").addEventListener("click", stopRecordMode);

    $("recordSaveBtn").addEventListener("click", async () => {
      if (!recordPortalId) return;
      const body = {
        selector_mapping: recordMap,
        auth_type: $("recordAuthType").value,
      };
      try {
        await api.updateSelectorMapping(recordPortalId, body);
        const p = portals.find((x) => x.id === recordPortalId);
        if (p) {
          p.selector_mapping = recordMap;
          p.auth_type = body.auth_type;
        }
        showToast("Yadda saxlandı.", "ok");
        await stopRecordMode();
      } catch (err) {
        showToast(err.message, "err");
      }
    });

    chrome.runtime.onMessage.addListener((msg) => {
      if (!msg || msg.type !== "FIELD_CAPTURED") return;
      if (!recordActive) {
        console.warn("[TaxPro panel] FIELD_CAPTURED while record inactive");
        return;
      }
      console.log("[TaxPro panel] received FIELD_CAPTURED", msg.selector);
      recordPendingSelector = msg.selector;
      $("recordPendingSelector").textContent = msg.selector;
      $("recordField").value = "asan_id";
      $("recordPending").hidden = false;
    });
  }

  /* ===== Yükləmə ===== */
  async function loadPortals() {
    portals = await api.getPortals();
    portals.forEach((p) => {
      p._keywords = portalKeywords(p);
    });
    renderSites($("siteSearch").value);
  }

  async function loadProfiles() {
    profiles = await api.listProfiles();
    renderProfiles($("profileSearch").value);
    populateProfilesSelect();
  }

  async function checkHealth() {
    try {
      await api.health();
      const base = await api.getBaseUrl();
      setStatus(true, base);
    } catch (err) {
      setStatus(false, "Backend əlçatmaz");
    }
  }

  async function init() {
    currentView = "tabs";
    $("profileModal").hidden = true;

    initSites();
    initProfiles();
    initProfileModal();
    initAutofill();
    initRecordMode();

    document.querySelectorAll(".tab-btn").forEach((b) => {
      b.addEventListener("click", () => switchTab(b.dataset.tab));
    });
    $("optionsBtn").addEventListener("click", () => {
      chrome.runtime.openOptionsPage();
    });

    await checkHealth();
    try {
      await Promise.all([loadPortals(), loadProfiles()]);
    } catch (err) {
      showToast(err.message, "err");
    }

    chrome.tabs.onActivated.addListener(() => loadActiveTab());
    chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
      if (changeInfo.url) loadActiveTab();
    });
    await loadActiveTab();
  }

  document.addEventListener("DOMContentLoaded", init);
})();
