/* TaxPro — Çat tab-ı (AI sual-cavab).
 * thread_id sessiya daxilində saxlanılır ki, eyni AI söhbəti davam etsin.
 * Görünmə yalnız `hidden` atributu ilə idarə olunur (class="hidden" yox!).
 */
(function () {
  "use strict";

  const api = window.TaxProAPI;
  const hitl = window.TaxProHITL;

  const els = {};
  const $ = (id) => (els[id] = els[id] || document.getElementById(id));

  let chatThreadId = null;
  let chatProfiles = [];
  let chatProfileId = null;

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

  function scrollBottom() {
    const list = $("chatList");
    if (list) list.scrollTop = list.scrollHeight;
  }

  function addBubble(role, text) {
    const list = $("chatList");
    if (!list) return;
    const div = document.createElement("div");
    div.className = "chat-bubble " + (role === "user" ? "user" : "ai");
    div.textContent = text;
    list.appendChild(div);
    scrollBottom();
  }

  function addBubbleHtml(role, html) {
    const list = $("chatList");
    if (!list) return;
    const div = document.createElement("div");
    div.className = "chat-bubble " + (role === "user" ? "user" : "ai");
    div.innerHTML = html;
    list.appendChild(div);
    scrollBottom();
  }

  function setInputEnabled(enabled) {
    $("chatInput").disabled = !enabled;
    $("chatSendBtn").disabled = !enabled;
  }

  async function loadProfiles() {
    try {
      chatProfiles = await api.listProfiles();
      const sel = $("chatProfile");
      const current = sel.value;
      sel.innerHTML =
        '<option value="">— Kontekst yoxdur —</option>' +
        chatProfiles
          .map((p) => `<option value="${p.id}">${escapeHtml(p.ad_soyad)}</option>`)
          .join("");
      if (current) sel.value = current;
      chatProfileId = sel.value || null;
    } catch (err) {
      chatProfiles = [];
    }
  }

  async function sendMessage() {
    const text = $("chatInput").value.trim();
    if (!text) return;
    $("chatInput").value = "";
    addBubble("user", text);
    setInputEnabled(false);

    try {
      const body = await api.chat(text, chatProfileId, chatThreadId);
      if (body.thread_id) chatThreadId = body.thread_id;

      if (body.status === "interrupt") {
        const holder = document.createElement("div");
        holder.className = "chat-hitl";
        $("chatList").appendChild(holder);
        scrollBottom();
        hitl.renderInterruptBlock(holder, {
          threadId: body.thread_id,
          interrupt: body.interrupt,
          onDone: (res) => {
            if (res.thread_id) chatThreadId = res.thread_id;
            if (res.status === "interrupt") {
              holder.innerHTML = "";
              hitl.renderInterruptBlock(holder, {
                threadId: res.thread_id,
                interrupt: res.interrupt,
                onDone: (res2) => {
                  addBubble("ai", res2.message || "Təsdiq edildi.");
                  holder.remove();
                },
              });
            } else {
              addBubble("ai", res.message || "Təsdiq edildi.");
              holder.remove();
            }
          },
        });
      } else {
        addBubble("ai", body.message || "");
      }
    } catch (err) {
      addBubble("ai", "Xəta: " + (err.message || "Bilinməyən xəta"));
    } finally {
      setInputEnabled(true);
      $("chatInput").focus();
    }
  }

  function init() {
    if (!window.TaxProAPI || !document.getElementById("chatInput")) return;

    $("chatSendBtn").addEventListener("click", sendMessage);
    $("chatInput").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        sendMessage();
      }
    });
    $("chatProfile").addEventListener("change", (e) => {
      chatProfileId = e.target.value || null;
    });

    loadProfiles();
  }

  document.addEventListener("DOMContentLoaded", init);
})();
