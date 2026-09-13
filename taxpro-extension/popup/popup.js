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

  // Doldurulan məntiqi sahələr (Record Mode etiketləri üçün).
  const FIELD_LABELS = {
    nomre: "Telefon",
    asan_id: "Asan ID",
    _phone_prefix_selector: "Prefiks",
  };

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

  function faviconUrl(portalUrl, size = 64) {
    try {
      const domain = new URL(portalUrl).hostname;
      return `https://www.google.com/s2/favicons?domain=${domain}&sz=${size}`;
    } catch (err) {
      return null;
    }
  }

  function attachLogoFallback(img) {
    img.onerror = () => {
      const span = document.createElement("span");
      span.className = "site-avatar-initials";
      span.textContent = img.dataset.initials || "?";
      img.replaceWith(span);
    };
  }

  function renderSiteAvatar(portal) {
    const url = portal.logo_url || faviconUrl(portal.url);
    const inner = url
      ? `<img class="site-logo" src="${escapeHtml(url)}" alt="" loading="lazy" data-initials="${escapeHtml(initials(portal.name))}">`
      : `<span class="site-avatar-initials">${escapeHtml(initials(portal.name))}</span>`;
    return `<span class="site-avatar">${inner}</span>`;
  }

  function portalLogoImg(portal) {
    const url = portal.logo_url || faviconUrl(portal.url);
    if (!url) return null;
    const img = document.createElement("img");
    img.className = "site-logo";
    img.src = url;
    img.alt = "";
    img.dataset.initials = initials(portal.name);
    attachLogoFallback(img);
    return img;
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

  function setStatus(ok, text, host) {
    $("apiStatus").className = "status-dot " + (ok ? "ok" : "err");
    $("apiStatusText").textContent = text;
    $("connStatus").title = host || "";
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
          renderSiteAvatar(p) +
          `<span class="site-meta">` +
          `<span class="site-name">${escapeHtml(p.name)}</span>` +
          `</span></a>`;
      }
    }
    list.innerHTML = html;
    list.querySelectorAll("img.site-logo").forEach(attachLogoFallback);
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
        String(p.asan_id || "").toLowerCase().includes(filter) ||
        String(p.nomre || "").includes(filter)
      );
    });

    if (shown.length === 0) {
      list.innerHTML = '<div class="empty-state">Profil tapılmadı.</div>';
      return;
    }

    let html = "";
    for (const p of shown) {
      html +=
        `<div class="profile-card" data-id="${p.id}">` +
        `<span class="profile-avatar" style="background:#3b82f6">${escapeHtml(initials(p.ad_soyad))}</span>` +
        `<div class="profile-meta">` +
        `<div class="profile-name">${escapeHtml(p.ad_soyad)}</div>` +
        `<div class="profile-sub">Asan ID: ${escapeHtml(p.asan_id)} · Tel: ${escapeHtml(p.nomre)}</div>` +
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
    $("pfNomre").value = profile ? profile.nomre : "";
    $("pfAsanId").value = profile ? profile.asan_id : "";

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
    return {
      ad_soyad: v("pfAdSoyad").trim(),
      nomre: v("pfNomre").trim(),
      asan_id: v("pfAsanId").trim(),
    };
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
    $("autofillBtn").disabled = !profileOk;
  }

  // Backend payload-ını content script-in gözlədiyi entry siyahısına çevirir.
  // Selector-u olmayan sahələr selector:"" ilə göndərilir ki, heuristik
  // fallback (attr → label → placeholder) işə düşsün.
  function buildAutofillEntries(payload) {
    const entries = [];
    const fieldMap = payload.field_map || {};
    for (const [selector, value] of Object.entries(payload.fields || {})) {
      entries.push({ selector, field: fieldMap[selector] || selector, value });
    }
    for (const [field, value] of Object.entries(payload.fallback_fields || {})) {
      entries.push({ selector: "", field, value });
    }
    return entries;
  }

  // Nəticə neytral göstərilir: yalnız "Dolduruldu" və ya "Uğursuz oldu".
  function showAutofillResult(resultBox, r) {
    const total = (r && r.total) || 0;
    const done = total > 0 && r.filled === total;
    resultBox.className = "result-box";
    resultBox.textContent = done ? "Dolduruldu" : "Uğursuz oldu";
    resultBox.classList.remove("hidden");
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
        info.textContent = "";
        const logo = portalLogoImg(matched);
        if (logo) info.appendChild(logo);
        info.append(document.createTextNode(`${matched.name} (${hostOf(tab.url)})`));
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
      resultBox.innerHTML = "";

      if (!profileId) return;

      let activeTab = null;
      try {
        const resp = await chrome.runtime.sendMessage({ type: "GET_ACTIVE_TAB" });
        activeTab = resp && resp.tab;
      } catch (err) {
        activeTab = null;
      }

      const portal = activeTab ? matchPortalByUrl(activeTab.url) : null;
      if (!portal) {
        resultBox.className = "result-box";
        resultBox.textContent =
          "Zəhmət olmasa, əvvəlcə Saytlar tabından portalın daxilolma səhifəsinə keçin.";
        resultBox.classList.remove("hidden");
        return;
      }

      try {
        const payload = await api.getPortalAutoFill(portal.id, profileId);
        const entries = buildAutofillEntries(payload);

        if (entries.length === 0) {
          resultBox.className = "result-box";
          resultBox.textContent = "Uğursuz oldu";
          resultBox.classList.remove("hidden");
          return;
        }

        const outcome = await chrome.runtime.sendMessage({
          type: "AUTOFILL",
          fields: entries,
          phone_prefix_selector: payload.phone_prefix_selector || null,
        });

        if (!outcome.ok) {
          resultBox.className = "result-box";
          resultBox.textContent = outcome.error || "Uğursuz oldu";
          resultBox.classList.remove("hidden");
          return;
        }

        showAutofillResult(resultBox, outcome.result);
      } catch (err) {
        resultBox.className = "result-box";
        resultBox.textContent = err.message || "Uğursuz oldu";
        resultBox.classList.remove("hidden");
      }
    });
  }

  async function populateAutofillSelects() {
    populateProfilesSelect();
  }

  /* ===== Qeydiyyat rejimi (Record Mode) ===== */
  const RECORD_FIELD_LABELS = FIELD_LABELS;

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
    const recName = $("recordPortalName");
    recName.textContent = "";
    const recLogo = portalLogoImg(portal);
    if (recLogo) recName.appendChild(recLogo);
    recName.append(document.createTextNode(portal.name));
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
      setStatus(true, "Qoşulu", base);
    } catch (err) {
      setStatus(false, "Əlaqə yoxdur");
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
    window.addEventListener("taxpro:profiles-changed", () => loadProfiles());
    await loadActiveTab();
  }

  document.addEventListener("DOMContentLoaded", init);
})();
