/* TaxPro — background service worker.
 * Mesaj marşrutlaşdırma: aktiv tab məlumatı, autofill icrası (lazım olduqda
 * content script inject edərək). Backend ünvanı popup tərəfindən oxunur.
 */
"use strict";

const DEFAULT_BASE_URL = "http://localhost:8000";

chrome.runtime.onInstalled.addListener(() => {
  chrome.storage.local.get("taxpro_base_url").then((stored) => {
    if (!stored.taxpro_base_url) {
      chrome.storage.local.set({ taxpro_base_url: DEFAULT_BASE_URL });
    }
  });
});

chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(() => {});

async function getActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab || null;
}

async function getFrameIds(tabId) {
  try {
    if (chrome.webNavigation && chrome.webNavigation.getAllFrames) {
      const frames = await chrome.webNavigation.getAllFrames({ tabId });
      if (Array.isArray(frames) && frames.length) {
        return frames.map((f) => f.frameId);
      }
    }
  } catch (err) {
    // webNavigation əlçatan deyil — yalnız top frame ilə davam et.
  }
  return [0];
}

async function ensureContentScript(tabId) {
  try {
    await chrome.scripting.executeScript({
      target: { tabId, allFrames: true },
      files: ["content-scripts/autofill.js"],
    });
  } catch (err) {
    // Qorunan səhifə (chrome://, web store və s.) və ya artıq inject olunub.
  }
}

// Bütün frame-lərə (o cümlədən iframe-lərə) mesaj göndərir.
async function sendToContent(tabId, msg) {
  const frameIds = await getFrameIds(tabId);
  let any = false;
  for (const frameId of frameIds) {
    try {
      await chrome.tabs.sendMessage(tabId, msg, { frameId });
      any = true;
    } catch (err) {
      // bu frame-də content script yoxdur — keç
    }
  }
  if (!any) {
    try {
      await chrome.tabs.sendMessage(tabId, msg);
      any = true;
    } catch (err) {
      // heç bir frame cavab vermir
    }
  }
  return any;
}

async function sendAutofillToFrames(tabId, payload) {
  const frameIds = await getFrameIds(tabId);
  let best = null;
  for (const frameId of frameIds) {
    try {
      const res = await chrome.tabs.sendMessage(tabId, payload, { frameId });
      if (res && typeof res.filled === "number") {
        if (!best || res.filled > best.filled) best = res;
      }
    } catch (err) {
      // frame-də content script yoxdur — keç
    }
  }
  return best;
}

async function runAutofill(tab, fields, phonePrefixSelector) {
  const payload = {
    type: "AUTOFILL",
    fields,
    phone_prefix_selector: phonePrefixSelector,
  };

  let best = await sendAutofillToFrames(tab.id, payload);
  if (!best) {
    await ensureContentScript(tab.id);
    best = await sendAutofillToFrames(tab.id, payload);
  }

  if (!best) {
    return {
      ok: false,
      error:
        "Bu səhifəyə giriş mümkün olmadı (qorunan səhifə və ya hələ yüklənməyib).",
    };
  }

  return { ok: true, result: best };
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  (async () => {
    if (!msg || typeof msg !== "object") return;

    if (msg.type === "GET_ACTIVE_TAB") {
      const tab = await getActiveTab();
      sendResponse(
        tab ? { tab: { id: tab.id, url: tab.url, title: tab.title } } : { tab: null }
      );
      return;
    }

    if (msg.type === "AUTOFILL") {
      const tab = await getActiveTab();
      if (!tab) {
        sendResponse({ ok: false, error: "Aktiv tab tapılmadı." });
        return;
      }
      const outcome = await runAutofill(
        tab,
        msg.fields || {},
        msg.phone_prefix_selector || null
      );
      sendResponse({ ...outcome, tab: { id: tab.id, url: tab.url } });
    }

    if (msg.type === "START_RECORD" || msg.type === "STOP_RECORD") {
      const tab = await getActiveTab();
      if (!tab) {
        sendResponse({ ok: false, error: "Aktiv tab tapılmadı." });
        return;
      }
      try {
        const ok = await sendToContent(tab.id, { type: msg.type });
        sendResponse(
          ok ? { ok: true } : { ok: false, error: "Səhifəyə giriş mümkün olmadı." }
        );
      } catch (err) {
        sendResponse({ ok: false, error: "Səhifəyə giriş mümkün olmadı." });
      }
    }

    if (msg.type === "FIELD_CAPTURED") {
      console.log("[TaxPro bg] forwarding FIELD_CAPTURED", msg.selector);
      const payload = {
        type: "FIELD_CAPTURED",
        selector: msg.selector,
        tabId: sender.tab ? sender.tab.id : null,
        tabUrl: sender.tab ? sender.tab.url : null,
      };
      chrome.runtime.sendMessage(payload).catch(() => {});
      sendResponse({ ok: true });
    }
  })();
  return true; // async sendResponse üçün kanalı açıq saxlayırıq
});
