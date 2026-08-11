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

async function ensureContentScript(tabId) {
  try {
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["content-scripts/autofill.js"],
    });
  } catch (err) {
    // Qorunan səhifə (chrome://, web store və s.) və ya artıq inject olunub.
  }
}

async function sendToContent(tabId, msg) {
  try {
    return await chrome.tabs.sendMessage(tabId, msg);
  } catch (err) {
    await ensureContentScript(tabId);
    return await chrome.tabs.sendMessage(tabId, msg);
  }
}

async function runAutofill(tab, fields) {
  let result = null;
  try {
    result = await chrome.tabs.sendMessage(tab.id, {
      type: "AUTOFILL",
      fields,
    });
  } catch (err) {
    result = null;
  }

  if (!result || typeof result.filled !== "number") {
    await ensureContentScript(tab.id);
    try {
      result = await chrome.tabs.sendMessage(tab.id, {
        type: "AUTOFILL",
        fields,
      });
    } catch (err) {
      return {
        ok: false,
        error:
          "Bu səhifəyə giriş mümkün olmadı (qorunan səhifə və ya hələ yüklənməyib).",
      };
    }
  }

  return { ok: true, result };
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
      const outcome = await runAutofill(tab, msg.fields || {});
      sendResponse({ ...outcome, tab: { id: tab.id, url: tab.url } });
    }

    if (msg.type === "START_RECORD" || msg.type === "STOP_RECORD") {
      const tab = await getActiveTab();
      if (!tab) {
        sendResponse({ ok: false, error: "Aktiv tab tapılmadı." });
        return;
      }
      try {
        await sendToContent(tab.id, { type: msg.type });
        sendResponse({ ok: true });
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
