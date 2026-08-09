/* ============================================================
   DECK — Command Console
   Fill in the CONFIG block below, then open via a real http(s)
   origin (not file://). See README.md for full setup steps.
   ============================================================ */

const CONFIG = {
  // Google Cloud Console → APIs & Services → Credentials → OAuth Client ID (Web application)
  GOOGLE_CLIENT_ID: "240635416449-csnre23fkrn7sbh82f9b74l3ld9s888n.apps.googleusercontent.com",

  GOOGLE_SCOPES: [
    "openid",
    "email",
    "profile",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/chat.spaces.readonly",
    "https://www.googleapis.com/auth/chat.messages"
  ].join(" "),

  // Zoom Meeting SDK — https://marketplace.zoom.us (create a "Meeting SDK" app).
  // Use the app's "Client ID" here (older guides call this "SDK Key" — same thing).
  ZOOM_SDK_KEY: "aSJq531ERmKZUXvjFZZZMw",
  // Signature must be generated server-side (Client Secret can't live in browser JS).
  // Run the included zoom-signature-server (see zoom-signature-server/README or
  // the main README.md) and point this at it — defaults to the local dev port.
  ZOOM_SIGNATURE_ENDPOINT: "http://localhost:3001/zoom-signature",

  // ElevenLabs Conversational AI — https://elevenlabs.io/app/conversational-ai
  ELEVENLABS_AGENT_ID: "agent_6001kxdf8rnfedwttcj81hq8wez2",

  // How often to refresh mail/calendar/chat (ms)
  REFRESH_INTERVAL: 60000,

  // How often (ms) to check if this browser window has moved/resized, so a
  // docked Meet/Teams window can be kept aligned continuously rather than
  // only on explicit resize events
  WINDOW_TRACK_INTERVAL: 500
};

/* ============================================================ */

let accessToken = null;
let tokenClient = null;
let calendarEvents = [];
let externalMeetingWin = null;
let externalMeetingProvider = null;
let externalPollTimer = null;
let windowTrackTimer = null;
let lastKnownWinRect = null;
let externalMeetingRepositionBlocked = false;
let mailIndex = {};
let currentModalSubmit = null;
let currentUserEmail = "";
let currentAttachments = [];
let calendarWeekOffset = 0;
let mailPageTokens = [null];
let mailPageIndex = 0;
let chatPageTokens = [null];
let chatPageIndex = 0;

/* ---------- boot ---------- */
window.addEventListener("load", () => {
  startClock();
  mountAssistant();
  initGoogleSignIn();
  initBoardSwitch();
  document.getElementById("fullscreenBtn").onclick = toggleFullscreen;
  document.getElementById("joinBtn").onclick = joinMeeting;
  document.getElementById("providerSelect").onchange = updateHostToggleVisibility;
  updateHostToggleVisibility();
  document.getElementById("leaveBtn").onclick = leaveMeeting;
  document.getElementById("refreshMail").onclick = () => withSpin(document.getElementById("refreshMail"), loadMail);
  document.getElementById("refreshCal").onclick = () => withSpin(document.getElementById("refreshCal"), loadCalendar);
  document.getElementById("prevWeekBtn").onclick = () => { calendarWeekOffset--; loadCalendar(); };
  document.getElementById("nextWeekBtn").onclick = () => { calendarWeekOffset++; loadCalendar(); };
  document.getElementById("todayWeekBtn").onclick = () => { calendarWeekOffset = 0; loadCalendar(); };
  document.getElementById("refreshChat").onclick = () => withSpin(document.getElementById("refreshChat"), loadChat);
  document.getElementById("mailPrevBtn").onclick = mailPrevPage;
  document.getElementById("mailNextBtn").onclick = mailNextPage;
  document.getElementById("chatPrevBtn").onclick = chatPrevPage;
  document.getElementById("chatNextBtn").onclick = chatNextPage;
  document.getElementById("openChat").onclick = () => window.open(withAuthUser("https://chat.google.com"), "_blank");
  document.getElementById("recenterBtn").onclick = repositionExternalMeeting;
  document.getElementById("closeExternalBtn").onclick = leaveMeeting;
  document.getElementById("composeBtn").onclick = openComposeModal;
  document.getElementById("newEventBtn").onclick = openNewEventModal;
  document.getElementById("modalClose").onclick = closeModal;
  document.getElementById("modalCancel").onclick = closeModal;
  document.getElementById("modalOverlay").onclick = (e) => {
    if (e.target.id === "modalOverlay") closeModal();
  };
  document.getElementById("modalSubmit").onclick = handleModalSubmit;
  window.addEventListener("resize", () => {
    if (externalMeetingWin && !externalMeetingWin.closed) repositionExternalMeeting();
  });
});

/* ---------- dashboard switcher (Ops / Dev) ---------- */
function initBoardSwitch() {
  const buttons = document.querySelectorAll(".board-btn");
  const saved = localStorage.getItem("deck.board") || "ops";
  setBoard(saved);
  buttons.forEach(btn => {
    btn.onclick = () => setBoard(btn.dataset.board);
  });
}

function setBoard(board) {
  document.querySelectorAll(".board-btn").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.board === board);
  });
  document.querySelectorAll("[data-board]").forEach(el => {
    el.hidden = el.dataset.board !== board;
  });
  localStorage.setItem("deck.board", board);
}

/* ---------- clock ---------- */
function startClock() {
  const el = document.getElementById("clock");
  const tick = () => { el.textContent = new Date().toLocaleTimeString([], { hour12: false }); };
  tick();
  setInterval(tick, 1000);
}

/* ---------- Google sign-in ---------- */
function initGoogleSignIn() {
  const btn = document.getElementById("signInBtn");
  const errEl = document.getElementById("gateError");

  console.log("[DECK] Waiting for Google Identity Services library...");
  const waitForGis = setInterval(() => {
    if (!window.google || !google.accounts) return;
    clearInterval(waitForGis);
    console.log("[DECK] Google Identity Services library loaded");

    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CONFIG.GOOGLE_CLIENT_ID,
      scope: CONFIG.GOOGLE_SCOPES,
      callback: (resp) => {
        console.log("[DECK] token callback fired:", resp);
        if (resp.error) {
          console.error("[DECK] token error:", resp.error, resp);
          showGateError(resp.error);
          return;
        }
        accessToken = resp.access_token;
        console.log("[DECK] got access token, calling onSignedIn()");
        onSignedIn().catch(err => {
          console.error("[DECK] onSignedIn() threw:", err);
        });
      }
    });

    btn.disabled = false;
    btn.onclick = () => {
      console.log("[DECK] Sign in clicked. CLIENT_ID =", CONFIG.GOOGLE_CLIENT_ID);
      if (CONFIG.GOOGLE_CLIENT_ID.startsWith("YOUR_")) {
        showGateError("Set GOOGLE_CLIENT_ID in app.js first (see README.md).");
        return;
      }
      tokenClient.requestAccessToken({ prompt: "" });
    };
  }, 100);
}

function showGateError(msg) {
  const el = document.getElementById("gateError");
  el.textContent = typeof msg === "string" ? msg : "Sign-in failed. Check console for details.";
  el.hidden = false;
}

async function onSignedIn() {
  console.log("[DECK] onSignedIn() start");
  document.getElementById("gate").hidden = true;
  document.getElementById("console").hidden = false;
  console.log("[DECK] gate hidden, console shown. gate.hidden =", document.getElementById("gate").hidden, "console.hidden =", document.getElementById("console").hidden);

  await loadProfile();
  await loadMail();
  await loadCalendar();
  await loadChat();

  setInterval(loadMail, CONFIG.REFRESH_INTERVAL);
  setInterval(loadCalendar, CONFIG.REFRESH_INTERVAL);
  setInterval(loadChat, CONFIG.REFRESH_INTERVAL);
  setInterval(updateNextEvent, 15000);
  setInterval(updateMeetingPanelUpcoming, 15000);
  console.log("[DECK] onSignedIn() complete");
}

async function gfetch(url, retryCount = 0) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
  if (res.status === 401) {
    // token expired — re-prompt silently
    tokenClient.requestAccessToken({ prompt: "" });
    throw new Error("Token expired, re-authenticating…");
  }
  if (res.status === 429 && retryCount < 2) {
    // Rate-limited — back off briefly and retry, per Google's own guidance
    await new Promise(r => setTimeout(r, 500 * (retryCount + 1)));
    return gfetch(url, retryCount + 1);
  }
  if (!res.ok) throw new Error(`Request failed: ${res.status}`);
  return res.json();
}

// Runs `fn` over `items` with at most `limit` calls in flight at once,
// instead of firing everything simultaneously (which is what triggers
// Google's rate limiting on larger batches like a 50-message mail page).
async function mapWithConcurrencyLimit(items, limit, fn) {
  const results = new Array(items.length);
  let nextIndex = 0;
  async function worker() {
    while (nextIndex < items.length) {
      const i = nextIndex++;
      results[i] = await fn(items[i]);
    }
  }
  const workers = Array.from({ length: Math.min(limit, items.length) }, worker);
  await Promise.all(workers);
  return results;
}

/* ---------- profile ---------- */
async function loadProfile() {
  try {
    const me = await gfetch("https://www.googleapis.com/oauth2/v2/userinfo");
    document.getElementById("userName").textContent = me.name || me.email;
    document.getElementById("userAvatar").src = me.picture || "";
    currentUserEmail = me.email || "";
  } catch (e) { console.warn(e); }
}

/* ---------- mail ---------- */
function formatDateTime(date) {
  const datePart = date.toLocaleDateString([], { day: "2-digit", month: "short" });
  const timePart = date.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  return `${datePart} ${timePart}`;
}

async function loadMail() {
  const body = document.getElementById("emailBody");
  try {
    const token = mailPageTokens[mailPageIndex];
    const url = `https://gmail.googleapis.com/gmail/v1/users/me/messages?maxResults=50&labelIds=INBOX${token ? `&pageToken=${token}` : ""}`;
    const list = await gfetch(url);

    // cache the token for the next page the first time we see it, so "next"
    // works without the API providing a true "previous" token — going back
    // just steps down this locally-built stack instead
    if (list.nextPageToken && mailPageTokens.length === mailPageIndex + 1) {
      mailPageTokens.push(list.nextPageToken);
    }
    updateMailPaginationUI(list.resultSizeEstimate, list.messages?.length || 0, !!list.nextPageToken);

    if (!list.messages) {
      body.innerHTML = mailPageIndex === 0
        ? `<div class="empty-state">Inbox is empty.</div>`
        : `<div class="empty-state">No more messages.</div>`;
      return;
    }

    // Fetching 50 messages' details all at once (one request per message)
    // bursts past Gmail API's per-user rate limit and triggers 429 errors —
    // this caps it to a handful in flight at a time instead.
    const items = await mapWithConcurrencyLimit(list.messages, 8, m =>
      gfetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages/${m.id}?format=metadata&metadataHeaders=From&metadataHeaders=Subject&metadataHeaders=Message-ID&metadataHeaders=To&metadataHeaders=Cc`)
    );

    mailIndex = {};
    let unread = 0;
    body.innerHTML = items.map(msg => {
      const headers = msg.payload?.headers || [];
      const fromHeader = headers.find(h => h.name === "From")?.value || "Unknown";
      const { name: fromName, email: fromEmail } = parseFromHeader(fromHeader);
      const subject = headers.find(h => h.name === "Subject")?.value || "(no subject)";
      const messageId = headers.find(h => h.name === "Message-ID")?.value || "";
      const toHeader = headers.find(h => h.name === "To")?.value || "";
      const ccHeader = headers.find(h => h.name === "Cc")?.value || "";
      const isUnread = msg.labelIds?.includes("UNREAD");
      if (isUnread) unread++;
      const time = formatDateTime(new Date(Number(msg.internalDate)));

      mailIndex[msg.id] = { fromEmail, subject, messageId, threadId: msg.threadId, toHeader, ccHeader };

      return `
        <div class="mail-item" onclick="openMailViewModal('${msg.id}')">
          <div class="mail-top">
            <span class="mail-from ${isUnread ? "unread" : ""}">${escapeHtml(fromName)}</span>
            <span class="mail-time">${time}</span>
          </div>
          <div class="mail-subject">${escapeHtml(subject)}</div>
          <div class="mail-snippet">${escapeHtml(msg.snippet || "")}</div>
          <div class="mail-actions">
            <button class="mail-reply-btn" onclick="event.stopPropagation(); openReplyModal('${msg.id}')">Reply</button>
            <button class="mail-delete-btn" title="Delete" onclick="event.stopPropagation(); trashEmail('${msg.id}')">🗑</button>
          </div>
        </div>`;
    }).join("");

    document.getElementById("unreadChip").textContent = `${unread} unread`;
  } catch (e) {
    console.warn(e);
    body.innerHTML = `<div class="empty-state">Couldn't load mail — ${escapeHtml(e.message)}</div>`;
  }
}

function updateMailPaginationUI(estimateTotal, pageCount, hasNext) {
  const start = mailPageIndex * 50 + (pageCount ? 1 : 0);
  const end = mailPageIndex * 50 + pageCount;
  const rangeEl = document.getElementById("mailRange");
  rangeEl.textContent = pageCount
    ? `${start}–${end}${estimateTotal ? ` of ~${estimateTotal.toLocaleString()}` : ""}`
    : "—";
  document.getElementById("mailPrevBtn").disabled = mailPageIndex === 0;
  document.getElementById("mailNextBtn").disabled = !hasNext;
}

function mailPrevPage() {
  if (mailPageIndex > 0) { mailPageIndex--; loadMail(); }
}

function mailNextPage() {
  if (mailPageIndex < mailPageTokens.length - 1) { mailPageIndex++; loadMail(); }
}

async function trashEmail(id) {
  try {
    const res = await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages/${id}/trash`, {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    if (!res.ok) throw new Error(`Delete failed (${res.status}): ${await res.text()}`);
    loadMail();
  } catch (e) {
    console.error(e);
    alert("Couldn't delete that email: " + e.message);
  }
}

/* ---------- calendar ---------- */
async function loadCalendar() {
  const body = document.getElementById("calendarBody");
  try {
    const today = new Date();
    const monday = getMonday(today);
    monday.setDate(monday.getDate() + calendarWeekOffset * 7);
    const friday = new Date(monday);
    friday.setDate(friday.getDate() + 4);
    friday.setHours(23, 59, 59, 999);

    document.getElementById("weekRange").textContent =
      `${monday.toLocaleDateString([], { month: "short", day: "numeric" })} – ${friday.toLocaleDateString([], { month: "short", day: "numeric" })}`;
    document.getElementById("todayWeekBtn").hidden = calendarWeekOffset === 0;

    const url = `https://www.googleapis.com/calendar/v3/calendars/primary/events?timeMin=${monday.toISOString()}&timeMax=${friday.toISOString()}&singleEvents=true&orderBy=startTime&maxResults=100`;
    const data = await gfetch(url);
    calendarEvents = data.items || [];

    const days = [];
    for (let i = 0; i < 5; i++) {
      const d = new Date(monday);
      d.setDate(d.getDate() + i);
      days.push(d);
    }
    const todayKey = dateKey(today);

    body.innerHTML = `<div class="week-grid">${
      days.map(d => {
        const key = dateKey(d);
        const dayEvents = calendarEvents.filter(ev => {
          const start = ev.start?.dateTime ? new Date(ev.start.dateTime) : (ev.start?.date ? new Date(ev.start.date + "T00:00:00") : null);
          return start && dateKey(start) === key;
        });
        const isToday = key === todayKey;
        const dayLabel = d.toLocaleDateString([], { weekday: "short" }).toUpperCase();
        const dateLabel = d.getDate();

        const eventsHtml = dayEvents.length
          ? dayEvents.map(ev => {
              const start = ev.start?.dateTime ? new Date(ev.start.dateTime) : null;
              const timeStr = start ? start.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }) : "All day";
              const now = new Date();
              const isNow = start && ev.end?.dateTime && now >= start && now <= new Date(ev.end.dateTime);
              return `
                <div class="week-event ${isNow ? "now" : ""}" title="${escapeHtml(ev.summary || "(untitled)")}" onclick="openEventViewModal('${ev.id}')">
                  <div class="week-event-time">${timeStr}</div>
                  <div class="week-event-title">${escapeHtml(ev.summary || "(untitled)")}</div>
                </div>`;
            }).join("")
          : `<div class="week-day-empty">—</div>`;

        return `
          <div class="week-day ${isToday ? "today" : ""}">
            <div class="week-day-head">
              <span class="week-day-name">${dayLabel}</span>
              <span class="week-day-date">${dateLabel}</span>
            </div>
            <div class="week-day-events">${eventsHtml}</div>
          </div>`;
      }).join("")
    }</div>`;

    if (calendarWeekOffset === 0) { updateNextEvent(); updateMeetingPanelUpcoming(); }
  } catch (e) {
    console.warn(e);
    body.innerHTML = `<div class="empty-state">Couldn't load calendar — ${escapeHtml(e.message)}</div>`;
  }
}

function getMonday(d) {
  const date = new Date(d);
  const day = date.getDay();
  const diff = (day === 0 ? -6 : 1) - day;
  date.setDate(date.getDate() + diff);
  date.setHours(0, 0, 0, 0);
  return date;
}

function dateKey(d) {
  return `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
}

function updateNextEvent() {
  const now = new Date();
  const upcoming = calendarEvents.find(ev => {
    const start = ev.start?.dateTime && new Date(ev.start.dateTime);
    return start && start > now;
  });
  const textEl = document.getElementById("nextEventText");
  const cdEl = document.getElementById("nextEventCountdown");
  if (!upcoming) { textEl.textContent = "Nothing else this week"; cdEl.textContent = ""; return; }

  const start = new Date(upcoming.start.dateTime);
  const diffMs = start - now;
  const mins = Math.round(diffMs / 60000);
  textEl.textContent = upcoming.summary || "(untitled)";
  cdEl.textContent = mins <= 0 ? "now" : mins < 60 ? `in ${mins}m` : `in ${Math.round(mins / 60)}h`;
}

/* Surfaces the current-or-next calendar event inside the meeting panel
   itself while it's idle, instead of leaving that space purely decorative.
   Only touches the panel when it's actually idle (placeholder visible) —
   never interferes with a live Zoom/Meet/Teams session. */
function findRelevantEvent() {
  const now = new Date();
  const ongoing = calendarEvents.find(ev => {
    const s = ev.start?.dateTime && new Date(ev.start.dateTime);
    const e = ev.end?.dateTime && new Date(ev.end.dateTime);
    return s && e && now >= s && now <= e;
  });
  if (ongoing) return { event: ongoing, status: "now" };

  const upcoming = calendarEvents
    .filter(ev => ev.start?.dateTime && new Date(ev.start.dateTime) > now)
    .sort((a, b) => new Date(a.start.dateTime) - new Date(b.start.dateTime))[0];
  return upcoming ? { event: upcoming, status: "next" } : null;
}

function updateMeetingPanelUpcoming() {
  const placeholder = document.getElementById("meetingPlaceholder");
  const card = document.getElementById("upcomingCard");
  if (!placeholder || placeholder.hidden) return; // a call is live — leave it alone

  const found = findRelevantEvent();
  if (!found) {
    card.hidden = true;
    document.getElementById("phSub").textContent = "No active meeting — join a call to bring this panel online";
    return;
  }

  const { event, status } = found;
  const now = new Date();
  const start = new Date(event.start.dateTime);
  const mins = Math.round((start - now) / 60000);
  const timeLabel = status === "now"
    ? "Happening now"
    : mins < 60 ? `Starts in ${mins}m` : `Starts in ${Math.round(mins / 60)}h`;

  document.getElementById("phSub").textContent = "";
  card.hidden = false;
  document.getElementById("upcomingTitle").textContent = event.summary || "(untitled)";
  document.getElementById("upcomingTime").textContent =
    `${start.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })} · ${timeLabel}`;

  const guestNames = (event.attendees || []).map(a => a.displayName || a.email).filter(Boolean);
  document.getElementById("upcomingGuests").textContent = guestNames.length
    ? `With ${guestNames.slice(0, 4).join(", ")}${guestNames.length > 4 ? ` +${guestNames.length - 4}` : ""}`
    : "";

  const link = detectMeetingLink(event);
  const joinBtn = document.getElementById("upcomingJoinBtn");
  if (link) {
    const label = link.provider === "zoom" ? "Zoom" : link.provider === "teams" ? "Teams" : "Google Meet";
    joinBtn.hidden = false;
    joinBtn.textContent = `Join ${label} now`;
    joinBtn.onclick = () => joinFromCalendarLink(link);
  } else {
    joinBtn.hidden = true;
  }
}

/* ---------- calendar: view / edit an event, detect & join meeting links ---------- */
function detectMeetingLink(event) {
  if (event.hangoutLink) return { provider: "meet", url: event.hangoutLink };

  const entryPoints = event.conferenceData?.entryPoints;
  if (entryPoints) {
    const video = entryPoints.find(e => e.entryPointType === "video");
    if (video?.uri) {
      if (/zoom\.us/.test(video.uri)) return { provider: "zoom", url: video.uri };
      if (/teams\.microsoft\.com/.test(video.uri)) return { provider: "teams", url: video.uri };
      return { provider: "meet", url: video.uri };
    }
  }

  const text = `${event.location || ""} ${event.description || ""}`;
  const zoomMatch = text.match(/https?:\/\/[a-zA-Z0-9.-]*zoom\.us\/j\/[^\s"'<>]+/);
  if (zoomMatch) return { provider: "zoom", url: zoomMatch[0] };
  const teamsMatch = text.match(/https?:\/\/teams\.microsoft\.com\/[^\s"'<>]+/);
  if (teamsMatch) return { provider: "teams", url: teamsMatch[0] };
  const meetMatch = text.match(/https?:\/\/meet\.google\.com\/[^\s"'<>]+/);
  if (meetMatch) return { provider: "meet", url: meetMatch[0] };
  return null;
}

function joinFromCalendarLink(link) {
  document.getElementById("providerSelect").value = link.provider;
  updateHostToggleVisibility();
  if (link.provider === "zoom") {
    const idMatch = link.url.match(/\/j\/(\d+)/);
    document.getElementById("meetingUrlInput").value = idMatch ? idMatch[1] : link.url;
  } else {
    document.getElementById("meetingUrlInput").value = link.url;
  }
  joinMeeting();
}

function openEventViewModal(id) {
  const event = calendarEvents.find(e => e.id === id);
  if (!event) return;

  const start = event.start?.dateTime ? new Date(event.start.dateTime) : (event.start?.date ? new Date(event.start.date + "T00:00:00") : null);
  const end = event.end?.dateTime ? new Date(event.end.dateTime) : (event.end?.date ? new Date(event.end.date + "T00:00:00") : null);
  const dateStr = start ? `${start.getFullYear()}-${String(start.getMonth() + 1).padStart(2, "0")}-${String(start.getDate()).padStart(2, "0")}` : "";
  const startStr = start && event.start.dateTime ? start.toTimeString().slice(0, 5) : "09:00";
  const endStr = end && event.end.dateTime ? end.toTimeString().slice(0, 5) : "17:00";
  const guests = (event.attendees || []).map(a => a.email).filter(Boolean).join(", ");
  const link = detectMeetingLink(event);
  const linkLabel = link ? (link.provider === "zoom" ? "Zoom" : link.provider === "teams" ? "Microsoft Teams" : "Google Meet") : "";

  const html = `
    <div><div class="field-label">TITLE</div><input class="field-input" id="fldEvTitle" value="${escapeHtml(event.summary || "")}" /></div>
    <div class="field-row">
      <div><div class="field-label">DATE</div><input class="field-input" type="date" id="fldEvDate" value="${dateStr}" /></div>
      <div><div class="field-label">START</div><input class="field-input" type="time" id="fldEvStart" value="${startStr}" /></div>
      <div><div class="field-label">END</div><input class="field-input" type="time" id="fldEvEnd" value="${endStr}" /></div>
    </div>
    ${link ? `
      <div class="meeting-link-row">
        <button type="button" class="mini-btn" id="joinFromEventBtn">Join ${linkLabel} in meeting panel ↗</button>
        <button type="button" class="mini-btn ghost" id="copyLinkBtn">Copy link</button>
      </div>` : ""}
    <div><div class="field-label">LOCATION</div><input class="field-input" id="fldEvLocation" value="${escapeHtml(event.location || "")}" /></div>
    <div><div class="field-label">GUESTS (comma-separated emails)</div><input class="field-input" id="fldEvGuests" value="${escapeHtml(guests)}" /></div>
    <div class="grow-field"><div class="field-label">DESCRIPTION</div><textarea class="field-textarea field-textarea-grow" id="fldEvDesc">${escapeHtml(event.description || "")}</textarea></div>
  `;

  openModal(event.summary || "(untitled event)", html, "Save", async () => {
    const title = document.getElementById("fldEvTitle").value.trim();
    const date = document.getElementById("fldEvDate").value;
    const startT = document.getElementById("fldEvStart").value;
    const endT = document.getElementById("fldEvEnd").value;
    const location = document.getElementById("fldEvLocation").value.trim();
    const guestsVal = document.getElementById("fldEvGuests").value.trim();
    const desc = document.getElementById("fldEvDesc").value;
    if (!title || !date) throw new Error("Title and date are required");
    await updateEvent(event.id, { title, date, start: startT, end: endT, location, guests: guestsVal, desc });
    loadCalendar();
  }, { large: true });

  if (link) {
    document.getElementById("joinFromEventBtn").onclick = () => {
      closeModal();
      joinFromCalendarLink(link);
    };
    document.getElementById("copyLinkBtn").onclick = async () => {
      const btn = document.getElementById("copyLinkBtn");
      try {
        await navigator.clipboard.writeText(link.url);
        const original = btn.textContent;
        btn.textContent = "Copied";
        setTimeout(() => { btn.textContent = original; }, 1200);
      } catch (e) {
        console.warn("Clipboard write failed:", e);
      }
    };
  }
}

async function updateEvent(eventId, { title, date, start, end, location, guests, desc }) {
  const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const body = {
    summary: title,
    start: { dateTime: `${date}T${start}:00`, timeZone: tz },
    end: { dateTime: `${date}T${end}:00`, timeZone: tz },
    location: location || "",
    description: desc || "",
    attendees: guests ? guests.split(",").map(e => ({ email: e.trim() })).filter(a => a.email) : []
  };

  const res = await fetch(
    `https://www.googleapis.com/calendar/v3/calendars/primary/events/${eventId}?sendUpdates=all`,
    {
      method: "PATCH",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify(body)
    }
  );
  if (!res.ok) throw new Error(`Update failed (${res.status}): ${await res.text()}`);
  return res.json();
}

/* ---------- chat ---------- */
async function loadChat() {
  const body = document.getElementById("chatBody");
  try {
    const token = chatPageTokens[chatPageIndex];
    const spacesRes = await gfetch(`https://chat.googleapis.com/v1/spaces?pageSize=50${token ? `&pageToken=${token}` : ""}`);
    const spaces = spacesRes.spaces || [];

    if (spacesRes.nextPageToken && chatPageTokens.length === chatPageIndex + 1) {
      chatPageTokens.push(spacesRes.nextPageToken);
    }
    updateChatPaginationUI(spaces.length, !!spacesRes.nextPageToken);

    if (!spaces.length) {
      body.innerHTML = chatPageIndex === 0
        ? `<div class="empty-state">No Chat spaces found for this account.</div>`
        : `<div class="empty-state">No more conversations.</div>`;
      return;
    }

    const rows = await Promise.all(spaces.map(async (space) => {
      try {
        const msgs = await gfetch(
          `https://chat.googleapis.com/v1/${space.name}/messages?pageSize=1&orderBy=createTime desc`
        );
        const latest = msgs.messages?.[0];
        return { space, latest };
      } catch (e) {
        return { space, latest: null };
      }
    }));

    // spaces with no accessible messages fall to the bottom, most recent first
    rows.sort((a, b) => {
      const ta = a.latest ? new Date(a.latest.createTime).getTime() : 0;
      const tb = b.latest ? new Date(b.latest.createTime).getTime() : 0;
      return tb - ta;
    });

    body.innerHTML = rows.map(({ space, latest }) => {
      const name = space.displayName || (space.spaceType === "DIRECT_MESSAGE" ? "Direct message" : "Chat space");
      const sender = latest?.sender?.displayName || "";
      const snippet = latest?.text || latest?.formattedText || "No recent messages";
      const time = latest ? formatDateTime(new Date(latest.createTime)) : "";
      const safeName = escapeHtml(name).replace(/'/g, "\\'");
      return `
        <div class="chat-item" data-space="${space.name}" onclick="openChatViewModal('${space.name}', '${safeName}')">
          <div class="chat-top">
            <span class="chat-space">${escapeHtml(name)}</span>
            <span class="chat-time">${time}</span>
          </div>
          ${sender ? `<div class="chat-sender">${escapeHtml(sender)}</div>` : ""}
          <div class="chat-snippet">${escapeHtml(snippet)}</div>
        </div>`;
    }).join("");
  } catch (e) {
    console.warn(e);
    body.innerHTML = `<div class="empty-state">
      Couldn't load Chat — ${escapeHtml(e.message)}.<br/><br/>
      Make sure the Chat API is enabled in Google Cloud Console and the OAuth
      consent screen is set to Internal for protem.solutions.
    </div>`;
  }
}

function updateChatPaginationUI(pageCount, hasNext) {
  const rangeEl = document.getElementById("chatRange");
  rangeEl.textContent = pageCount ? `Page ${chatPageIndex + 1}` : "—";
  document.getElementById("chatPrevBtn").disabled = chatPageIndex === 0;
  document.getElementById("chatNextBtn").disabled = !hasNext;
}

function chatPrevPage() {
  if (chatPageIndex > 0) { chatPageIndex--; loadChat(); }
}

function chatNextPage() {
  if (chatPageIndex < chatPageTokens.length - 1) { chatPageIndex++; loadChat(); }
}

/* ---------- chat: view/reply in a large popout, like mail ---------- */
function openChatViewModal(spaceName, displayName) {
  openModal(displayName, `<div class="empty-state">Loading conversation…</div>`, "Send", async () => {
    const input = document.getElementById("chatModalInput");
    const text = input.value.trim();
    if (!text) throw new Error("Type a message first");
    await sendChatMessageTo(spaceName, text);
    input.value = "";
    await renderChatThread(spaceName, displayName);
    return "KEEP_OPEN";
  }, { large: true });

  renderChatThread(spaceName, displayName);
}

async function renderChatThread(spaceName, displayName) {
  try {
    const data = await gfetch(`https://chat.googleapis.com/v1/${spaceName}/messages?pageSize=25&orderBy=createTime%20desc`);
    const messages = (data.messages || []).slice().reverse();

    const threadHtml = messages.length
      ? messages.map(m => `
          <div class="chat-thread-msg">
            <div class="chat-thread-meta">
              <span class="chat-thread-sender">${escapeHtml(m.sender?.displayName || "Unknown")}</span>
              <span class="chat-thread-time">${new Date(m.createTime).toLocaleString([], { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" })}</span>
            </div>
            <div class="chat-thread-text">${escapeHtml(m.text || m.formattedText || "")}</div>
          </div>`).join("")
      : `<div class="empty-state">No messages yet — say hello.</div>`;

    document.getElementById("modalBody").innerHTML = `
      <div class="chat-thread" id="chatThreadList">${threadHtml}</div>
      <input class="field-input" id="chatModalInput" placeholder="Send a message…" />
    `;

    const list = document.getElementById("chatThreadList");
    if (list) list.scrollTop = list.scrollHeight;

    const input = document.getElementById("chatModalInput");
    if (input) {
      input.focus();
      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") { e.preventDefault(); handleModalSubmit(); }
      });
    }
  } catch (e) {
    console.warn(e);
    document.getElementById("modalBody").innerHTML =
      `<div class="empty-state">Couldn't load the conversation — ${escapeHtml(e.message)}</div>`;
  }
}

async function sendChatMessageTo(spaceName, text) {
  const res = await fetch(`https://chat.googleapis.com/v1/${spaceName}/messages`, {
    method: "POST",
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
    body: JSON.stringify({ text })
  });
  if (!res.ok) throw new Error(`Send failed (${res.status}): ${await res.text()}`);
  return res.json();
}

/* ---------- generic modal (Compose / Reply / New event) ---------- */
function openModal(title, bodyHtml, submitLabel, onSubmit, opts = {}) {
  document.getElementById("modalTitle").textContent = title;
  document.getElementById("modalBody").innerHTML = bodyHtml;
  document.getElementById("modalSubmit").textContent = submitLabel;
  document.getElementById("modal").classList.toggle("large", !!opts.large);
  const statusEl = document.getElementById("modalStatus");
  statusEl.textContent = "";
  statusEl.className = "modal-status";
  currentModalSubmit = onSubmit;
  document.getElementById("modalOverlay").hidden = false;
}

function closeModal() {
  document.getElementById("modalOverlay").hidden = true;
  currentModalSubmit = null;
}

async function handleModalSubmit() {
  if (!currentModalSubmit) return;
  const statusEl = document.getElementById("modalStatus");
  const btn = document.getElementById("modalSubmit");
  statusEl.textContent = "Sending…";
  statusEl.className = "modal-status";
  btn.disabled = true;
  try {
    const result = await currentModalSubmit();
    if (result === "KEEP_OPEN") {
      statusEl.textContent = "";
      btn.disabled = false;
      return;
    }
    statusEl.textContent = "Done";
    statusEl.className = "modal-status success";
    setTimeout(closeModal, 500);
  } catch (e) {
    console.error(e);
    statusEl.textContent = e.message || "Something went wrong";
    statusEl.className = "modal-status error";
  } finally {
    btn.disabled = false;
  }
}

/* ---------- mail: view full message ---------- */
function findMailBodyData(payload) {
  if (!payload) return null;
  if (payload.parts) {
    for (const part of payload.parts) {
      if (part.mimeType === "text/plain" && part.body?.data) return { data: part.body.data, html: false };
    }
    for (const part of payload.parts) {
      const nested = findMailBodyData(part);
      if (nested) return nested;
    }
    for (const part of payload.parts) {
      if (part.mimeType === "text/html" && part.body?.data) return { data: part.body.data, html: true };
    }
    return null;
  }
  if (payload.body?.data) return { data: payload.body.data, html: payload.mimeType === "text/html" };
  return null;
}

function decodeBase64Url(data) {
  const base64 = data.replace(/-/g, "+").replace(/_/g, "/");
  const binary = atob(base64);
  const percentEncoded = binary.split("").map(c => "%" + c.charCodeAt(0).toString(16).padStart(2, "0")).join("");
  return decodeURIComponent(percentEncoded);
}

function htmlToPlainText(html) {
  // Strip style/script/head blocks entirely so their raw CSS/JS text can't
  // leak into the visible output — textContent otherwise includes them.
  let cleaned = html
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<head[\s\S]*?<\/head>/gi, "");

  // Turn common block-level boundaries into line breaks before stripping tags,
  // so paragraphs/rows/list items don't all run together.
  cleaned = cleaned
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|li|h[1-6]|table)>/gi, "\n")
    .replace(/<li[^>]*>/gi, "• ");

  const tmp = document.createElement("div");
  tmp.innerHTML = cleaned;
  let text = tmp.textContent || "";

  return text
    .replace(/[ \t]+/g, " ")
    .split("\n").map(l => l.trim()).join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function openMailViewModal(id) {
  const meta = mailIndex[id];
  if (!meta) return;

  openModal(meta.subject || "(no subject)", `<div class="empty-state">Loading message…</div>`, "Reply", () => {
    openReplyModal(id, "reply");
    return "KEEP_OPEN";
  }, { large: true });

  gfetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages/${id}?format=full`)
    .then(full => {
      const found = findMailBodyData(full.payload);
      let text = found ? decodeBase64Url(found.data) : (full.snippet || "");
      if (found?.html) {
        text = htmlToPlainText(text);
      }
      document.getElementById("modalBody").innerHTML = `
        <div class="field-label">FROM</div>
        <div class="mail-view-from">${escapeHtml(meta.fromEmail)}</div>
        <div class="mail-view-actions">
          <button class="mini-btn ghost" onclick="openReplyModal('${id}', 'replyAll')">Reply All</button>
        </div>
        <div class="mail-view-body">${escapeHtml(text)}</div>
      `;
    })
    .catch(e => {
      console.warn(e);
      document.getElementById("modalBody").innerHTML =
        `<div class="empty-state">Couldn't load the full message — ${escapeHtml(e.message)}</div>`;
    });
}

/* ---------- mail: compose / reply (full editor: cc/bcc, attachments) ---------- */
function parseFromHeader(value) {
  const match = /<([^>]+)>/.exec(value || "");
  const email = match ? match[1] : (value || "").trim();
  const name = (value || "").split("<")[0].trim() || email;
  return { name, email };
}

function parseAddressList(headerValue) {
  if (!headerValue) return [];
  return headerValue.split(",").map(s => parseFromHeader(s).email).filter(Boolean);
}

function composeFieldsHtml({ to = "", subject = "" } = {}) {
  return `
    <div><div class="field-label">TO</div><input class="field-input" id="fldTo" value="${escapeHtml(to)}" placeholder="name@company.com" /></div>
    <div class="cc-bcc-toggle">
      <button type="button" class="link-btn" id="toggleCc">+Cc</button>
      <button type="button" class="link-btn" id="toggleBcc">+Bcc</button>
    </div>
    <div id="ccRow" hidden><div class="field-label">CC</div><input class="field-input" id="fldCc" /></div>
    <div id="bccRow" hidden><div class="field-label">BCC</div><input class="field-input" id="fldBcc" /></div>
    <div><div class="field-label">SUBJECT</div><input class="field-input" id="fldSubject" value="${escapeHtml(subject)}" /></div>
    <div class="grow-field"><div class="field-label">MESSAGE</div><textarea class="field-textarea field-textarea-grow" id="fldBody"></textarea></div>
    <div>
      <div class="field-label">ATTACHMENTS</div>
      <input type="file" id="fldAttachInput" multiple hidden />
      <button type="button" class="mini-btn ghost" id="addAttachBtn">+ Add files</button>
      <div id="attachList" class="attach-list"></div>
    </div>
  `;
}

function setupComposeFields() {
  currentAttachments = [];
  document.getElementById("toggleCc").onclick = () => {
    document.getElementById("ccRow").hidden = false;
    document.getElementById("toggleCc").style.display = "none";
  };
  document.getElementById("toggleBcc").onclick = () => {
    document.getElementById("bccRow").hidden = false;
    document.getElementById("toggleBcc").style.display = "none";
  };
  const fileInput = document.getElementById("fldAttachInput");
  document.getElementById("addAttachBtn").onclick = () => fileInput.click();
  fileInput.onchange = async () => {
    for (const file of Array.from(fileInput.files)) {
      try {
        const base64 = await fileToBase64(file);
        currentAttachments.push({ name: file.name, type: file.type, size: file.size, base64 });
      } catch (e) {
        console.warn("Couldn't read file:", file.name, e);
      }
    }
    fileInput.value = "";
    renderAttachList();
  };
  renderAttachList();
}

function renderAttachList() {
  const el = document.getElementById("attachList");
  if (!el) return;
  if (!currentAttachments.length) { el.innerHTML = ""; return; }
  el.innerHTML = currentAttachments.map((a, i) => `
    <div class="attach-chip">
      <span class="attach-name">${escapeHtml(a.name)}</span>
      <span class="attach-size">${formatBytes(a.size)}</span>
      <button type="button" onclick="removeAttachment(${i})">✕</button>
    </div>`).join("");
}

function removeAttachment(i) {
  currentAttachments.splice(i, 1);
  renderAttachList();
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1024 / 1024).toFixed(1) + " MB";
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function openComposeModal() {
  openModal("Compose", composeFieldsHtml(), "Send", async () => {
    const to = document.getElementById("fldTo").value.trim();
    const cc = document.getElementById("fldCc")?.value.trim();
    const bcc = document.getElementById("fldBcc")?.value.trim();
    const subject = document.getElementById("fldSubject").value.trim();
    const body = document.getElementById("fldBody").value;
    if (!to) throw new Error("Add a recipient");
    await sendEmail({ to, cc, bcc, subject, body, attachments: currentAttachments });
    loadMail();
  }, { large: true });
  setupComposeFields();
}

function openReplyModal(id, mode = "reply") {
  const msg = mailIndex[id];
  if (!msg) return;
  const cleanSubject = (msg.subject || "").replace(/^Re:\s*/i, "");

  let prefillCc = "";
  if (mode === "replyAll") {
    const all = [...parseAddressList(msg.toHeader), ...parseAddressList(msg.ccHeader)];
    const selfEmail = (currentUserEmail || "").toLowerCase();
    const fromEmail = (msg.fromEmail || "").toLowerCase();
    const filtered = all.filter(e => e.toLowerCase() !== selfEmail && e.toLowerCase() !== fromEmail);
    prefillCc = [...new Set(filtered)].join(", ");
  }

  openModal(
    mode === "replyAll" ? "Reply All" : "Reply",
    composeFieldsHtml({ to: msg.fromEmail, subject: "Re: " + cleanSubject }),
    "Send",
    async () => {
      const to = document.getElementById("fldTo").value.trim();
      const cc = document.getElementById("fldCc")?.value.trim();
      const bcc = document.getElementById("fldBcc")?.value.trim();
      const subject = document.getElementById("fldSubject").value.trim();
      const body = document.getElementById("fldBody").value;
      if (!to) throw new Error("Add a recipient");
      await sendEmail({
        to, cc, bcc, subject, body,
        threadId: msg.threadId, inReplyTo: msg.messageId, attachments: currentAttachments
      });
      loadMail();
    },
    { large: true }
  );
  setupComposeFields();

  if (prefillCc) {
    document.getElementById("ccRow").hidden = false;
    document.getElementById("fldCc").value = prefillCc;
    document.getElementById("toggleCc").style.display = "none";
  }
}

function wrapBase64(b64) {
  return (b64.match(/.{1,76}/g) || []).join("\r\n");
}

function buildRawMessage({ to, cc, bcc, subject, body, inReplyTo, attachments = [] }) {
  const lines = [`To: ${to}`];
  if (cc) lines.push(`Cc: ${cc}`);
  if (bcc) lines.push(`Bcc: ${bcc}`);
  lines.push(`Subject: ${subject}`);
  if (inReplyTo) {
    lines.push(`In-Reply-To: ${inReplyTo}`);
    lines.push(`References: ${inReplyTo}`);
  }
  lines.push("MIME-Version: 1.0");

  if (attachments.length) {
    const boundary = "deck_" + Math.random().toString(36).slice(2);
    lines.push(`Content-Type: multipart/mixed; boundary="${boundary}"`, "");
    lines.push(`--${boundary}`);
    lines.push('Content-Type: text/plain; charset="UTF-8"', "", body, "");
    for (const att of attachments) {
      lines.push(`--${boundary}`);
      lines.push(`Content-Type: ${att.type || "application/octet-stream"}; name="${att.name}"`);
      lines.push("Content-Transfer-Encoding: base64");
      lines.push(`Content-Disposition: attachment; filename="${att.name}"`, "");
      lines.push(wrapBase64(att.base64), "");
    }
    lines.push(`--${boundary}--`);
  } else {
    lines.push('Content-Type: text/plain; charset="UTF-8"', "", body);
  }

  const raw = lines.join("\r\n");
  return btoa(unescape(encodeURIComponent(raw)))
    .replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

async function sendEmail({ to, cc, bcc, subject, body, threadId, inReplyTo, attachments }) {
  const payload = { raw: buildRawMessage({ to, cc, bcc, subject, body, inReplyTo, attachments }) };
  if (threadId) payload.threadId = threadId;
  const res = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/messages/send", {
    method: "POST",
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  if (!res.ok) throw new Error(`Send failed (${res.status}): ${await res.text()}`);
  return res.json();
}

/* ---------- calendar: new event ---------- */
function openNewEventModal() {
  const today = new Date().toISOString().slice(0, 10);
  const html = `
    <div><div class="field-label">TITLE</div><input class="field-input" id="fldTitle" /></div>
    <div class="field-row">
      <div><div class="field-label">DATE</div><input class="field-input" type="date" id="fldDate" value="${today}" /></div>
      <div><div class="field-label">START</div><input class="field-input" type="time" id="fldStart" value="09:00" /></div>
      <div><div class="field-label">END</div><input class="field-input" type="time" id="fldEnd" value="09:30" /></div>
    </div>
    <div><div class="field-label">LOCATION (optional)</div><input class="field-input" id="fldLocation" /></div>
    <div><div class="field-label">GUESTS (optional, comma-separated emails)</div><input class="field-input" id="fldGuests" /></div>
    <div><div class="field-label">DESCRIPTION (optional)</div><textarea class="field-textarea" id="fldDesc"></textarea></div>
  `;
  openModal("New event", html, "Create", async () => {
    const title = document.getElementById("fldTitle").value.trim();
    const date = document.getElementById("fldDate").value;
    const start = document.getElementById("fldStart").value;
    const end = document.getElementById("fldEnd").value;
    const location = document.getElementById("fldLocation").value.trim();
    const guests = document.getElementById("fldGuests").value.trim();
    const desc = document.getElementById("fldDesc").value;
    if (!title || !date || !start || !end) throw new Error("Fill in title, date, start and end");
    await createEvent({ title, date, start, end, location, guests, desc });
    loadCalendar();
  });
}

async function createEvent({ title, date, start, end, location, guests, desc }) {
  const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const body = {
    summary: title,
    start: { dateTime: `${date}T${start}:00`, timeZone: tz },
    end: { dateTime: `${date}T${end}:00`, timeZone: tz }
  };
  if (location) body.location = location;
  if (desc) body.description = desc;
  if (guests) body.attendees = guests.split(",").map(e => ({ email: e.trim() })).filter(a => a.email);

  const res = await fetch(
    "https://www.googleapis.com/calendar/v3/calendars/primary/events?sendUpdates=all",
    {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify(body)
    }
  );
  if (!res.ok) throw new Error(`Create failed (${res.status}): ${await res.text()}`);
  return res.json();
}

/* ---------- meeting panel ---------- */
function withAuthUser(url) {
  if (!currentUserEmail) return url;
  const sep = url.includes("?") ? "&" : "?";
  return `${url}${sep}authuser=${encodeURIComponent(currentUserEmail)}`;
}

function updateHostToggleVisibility() {
  const provider = document.getElementById("providerSelect").value;
  document.getElementById("zoomHostToggleLabel").hidden = provider !== "zoom";
}

function joinMeeting() {
  const provider = document.getElementById("providerSelect").value;
  const input = document.getElementById("meetingUrlInput").value.trim();
  if (!input) return;

  document.getElementById("meetingPlaceholder").hidden = true;
  document.getElementById("joinBtn").hidden = true;
  document.getElementById("leaveBtn").hidden = false;
  setMeetingChip("LIVE", true);

  if (provider === "zoom") {
    const isHost = document.getElementById("zoomHostToggle").checked;
    joinZoom(input, isHost);
  } else {
    let url = provider === "meet"
      ? (input.startsWith("http") ? input : `https://meet.google.com/${input}`)
      : input;
    if (provider === "meet") url = withAuthUser(url);
    openExternalMeeting(provider, url);
  }
}

function leaveMeeting() {
  document.getElementById("meetingPlaceholder").hidden = false;
  document.getElementById("joinBtn").hidden = false;
  document.getElementById("leaveBtn").hidden = true;
  document.getElementById("zoomRoot").hidden = true;
  document.getElementById("externalStatus").hidden = true;
  setMeetingChip("IDLE", false);
  updateMeetingPanelUpcoming();

  if (window.zoomClient) { try { window.zoomClient.leaveMeeting(); } catch (e) {} }

  clearInterval(externalPollTimer);
  stopWindowTracking();
  if (externalMeetingWin && !externalMeetingWin.closed) {
    try { externalMeetingWin.close(); } catch (e) {}
  }
  externalMeetingWin = null;
  externalMeetingProvider = null;
}

/* ---------- positioning Meet/Teams windows over the meeting panel ----------
   Neither Google nor Microsoft allow their live call UI to be embedded in a
   third-party page, so this opens the call in its own OS window and sizes /
   places that window to sit directly over the meeting panel's on-screen
   rectangle. Browsers don't expose their own chrome (title bar, address bar)
   dimensions precisely, so this is a close best-effort alignment rather than
   a pixel-perfect embed — use "Recenter window" if it drifts. */

function getPanelScreenRect() {
  const panel = document.getElementById("meetingBody");
  const rect = panel.getBoundingClientRect();
  const chromeW = Math.max(0, window.outerWidth - window.innerWidth);
  const chromeH = Math.max(0, window.outerHeight - window.innerHeight);
  return {
    left: Math.round(window.screenX + chromeW + rect.left),
    top: Math.round(window.screenY + chromeH + rect.top),
    width: Math.round(rect.width),
    height: Math.round(rect.height)
  };
}

async function openExternalMeeting(provider, url) {
  // window.open() cannot create a genuinely separate window while the page
  // itself is in true fullscreen mode — the browser owns the entire physical
  // display exclusively in that state, so a "popup" has nowhere to appear
  // except as a tab within that same fullscreen window. Exiting fullscreen
  // first restores normal windowed behavior, which is what actually lets it
  // dock as a real window instead.
  if (document.fullscreenElement) {
    try { await document.exitFullscreen(); } catch (e) {}
    // give the browser a moment to finish restoring window chrome/dimensions
    // before we measure the panel's screen position
    await new Promise(r => setTimeout(r, 250));
  }

  const r = getPanelScreenRect();
  // "popup=1" is the modern, explicit signal Chrome uses to decide
  // window-vs-tab — relying only on width/height/toolbar=no (the old way)
  // has become unreliable on current Chrome versions, which is very likely
  // why this was opening as a plain tab despite those legacy flags.
  const features = `popup=1,left=${r.left},top=${r.top},width=${r.width},height=${r.height},` +
    `toolbar=no,location=no,menubar=no,status=no,scrollbars=yes,resizable=yes`;

  // A unique window name each time forces a genuine new popup window. Reusing
  // a fixed name risks Chrome quietly treating a prior instance as a regular
  // tab (e.g. if it ever got merged into the main window), after which
  // moveTo/resizeTo become no-ops — which looks exactly like "it just opens
  // as a normal tab and won't dock."
  const win = window.open(url, `deck_meeting_${Date.now()}`, features);
  if (!win) {
    alert(`Pop-up blocked — allow pop-ups for this site so ${provider === "meet" ? "Google Meet" : "MS Teams"} can dock over the meeting panel.`);
    leaveMeeting();
    return;
  }

  externalMeetingWin = win;
  externalMeetingProvider = provider;
  externalMeetingRepositionBlocked = false;
  document.getElementById("recenterBtn").hidden = false;
  document.getElementById("extSub").textContent = "Running in its own window, docked over this panel";
  try { win.focus(); } catch (e) {}

  // give the popup a moment to finish opening, then nudge it precisely into
  // place — repeated several times, since Meet/Teams often resize their own
  // window once as the page transitions from the "ready to join" lobby to
  // the live call view, which would otherwise undo a single early attempt
  [300, 800, 1500, 2500, 4000].forEach(delay => setTimeout(repositionExternalMeeting, delay));

  document.getElementById("externalStatus").hidden = false;
  document.getElementById("externalStatusLabel").textContent =
    provider === "meet" ? "GOOGLE MEET" : "MICROSOFT TEAMS";

  clearInterval(externalPollTimer);
  externalPollTimer = setInterval(() => {
    if (!externalMeetingWin || externalMeetingWin.closed) {
      clearInterval(externalPollTimer);
      leaveMeeting();
    }
  }, 1000);

  startWindowTracking();
}

/* Browsers fire a 'resize' event for the dashboard's own window but not a
   'move' event, so dragging the browser window (or the whole app to a
   different monitor) wouldn't otherwise be noticed. This polls the
   dashboard window's screen position/size at a light interval and nudges
   the docked meeting window back into place whenever it changes. */
function startWindowTracking() {
  stopWindowTracking();
  lastKnownWinRect = { x: window.screenX, y: window.screenY, w: window.outerWidth, h: window.outerHeight };
  windowTrackTimer = setInterval(() => {
    if (!externalMeetingWin || externalMeetingWin.closed) { stopWindowTracking(); return; }
    const current = { x: window.screenX, y: window.screenY, w: window.outerWidth, h: window.outerHeight };
    const moved = !lastKnownWinRect ||
      current.x !== lastKnownWinRect.x || current.y !== lastKnownWinRect.y ||
      current.w !== lastKnownWinRect.w || current.h !== lastKnownWinRect.h;
    if (moved) {
      lastKnownWinRect = current;
      repositionExternalMeeting();
    }
  }, CONFIG.WINDOW_TRACK_INTERVAL);
}

function stopWindowTracking() {
  clearInterval(windowTrackTimer);
  windowTrackTimer = null;
  lastKnownWinRect = null;
}

function repositionExternalMeeting() {
  if (!externalMeetingWin || externalMeetingWin.closed) return;
  if (externalMeetingRepositionBlocked) return;
  const r = getPanelScreenRect();
  try {
    externalMeetingWin.moveTo(r.left, r.top);
    externalMeetingWin.resizeTo(r.width, r.height);
    externalMeetingWin.focus();
  } catch (e) {
    // Once the popup navigates to Meet/Teams' own domain, browsers block the
    // opener from moving or resizing it — a deliberate cross-origin security
    // boundary, not something fixable from here. Stop retrying (it will
    // never succeed) and be upfront about it in the UI instead of silently
    // failing on every poll.
    console.warn("Browser blocked repositioning the meeting window — this is expected once it navigates to a different domain:", e);
    externalMeetingRepositionBlocked = true;
    stopWindowTracking();
    const sub = document.getElementById("extSub");
    if (sub) sub.textContent = "Running as its own independent window — browsers block auto-repositioning once it loads a different site. Move or resize it manually if needed.";
    const recenterBtn = document.getElementById("recenterBtn");
    if (recenterBtn) recenterBtn.hidden = true;
  }
}

function setMeetingChip(text, live) {
  const chip = document.getElementById("meetingChip");
  chip.textContent = text;
  chip.classList.toggle("live", live);
}

async function joinZoom(meetingInput, isHost = false) {
  if (CONFIG.ZOOM_SDK_KEY.startsWith("YOUR_")) {
    alert("Add your Zoom Meeting SDK key + signature endpoint in app.js (see README.md) to embed live Zoom calls.");
    leaveMeeting();
    return;
  }
  try {
    const sigRes = await fetch(CONFIG.ZOOM_SIGNATURE_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ meetingNumber: meetingInput, role: isHost ? 1 : 0 })
    });
    const { signature } = await sigRes.json();

    const root = document.getElementById("zoomRoot");
    root.hidden = false;
    const client = window.ZoomMtgEmbedded.createClient();

    // Detects the meeting ending from *inside* Zoom's own UI (host ends it,
    // or the user clicks Zoom's native "Leave"/"End" button) — without this,
    // only our own "Leave" button in the panel header would ever reset the
    // panel, leaving it blank if the call ends any other way.
    client.on("connection-change", (payload) => {
      if (payload.state === "Closed") {
        leaveMeeting();
      }
    });

    await client.init({
      zoomAppRoot: root,
      language: "en-US",
      patchJsMedia: true,
      leaveOnPageUnload: true
    });
    await client.join({
      signature,
      meetingNumber: meetingInput,
      userName: document.getElementById("userName").textContent || "Guest"
    });
    window.zoomClient = client;
  } catch (e) {
    console.error("[DECK] Zoom join failed — full error object:", e);
    if (e?.reason === "Meeting has not started" || e?.errorCode === 3008) {
      alert("That meeting hasn't been started by the host yet — try again once it's actually live, or enable \"Join before host\" in the meeting's Zoom settings.");
    } else {
      const detail = e?.message || e?.reason || e?.errorMessage || JSON.stringify(e) || "Unknown error";
      alert("Zoom join failed: " + detail);
    }
    leaveMeeting();
  }
}

/* ---------- ElevenLabs assistant ---------- */
function mountAssistant() {
  const mount = document.getElementById("assistantMount");
  if (CONFIG.ELEVENLABS_AGENT_ID.startsWith("YOUR_")) {
    mount.innerHTML = `<div class="empty-state">Add your ElevenLabs agent ID in app.js to activate the assistant.</div>`;
    return;
  }
  const widget = document.createElement("elevenlabs-convai");
  widget.setAttribute("agent-id", CONFIG.ELEVENLABS_AGENT_ID);
  mount.appendChild(widget);
}

/* ---------- utils ---------- */
async function withSpin(btn, fn) {
  btn.classList.add("spinning");
  btn.disabled = true;
  try {
    await fn();
  } finally {
    btn.classList.remove("spinning");
    btn.disabled = false;
  }
}

function toggleFullscreen() {
  if (!document.fullscreenElement) document.documentElement.requestFullscreen();
  else document.exitFullscreen();
}

function escapeHtml(str) {
  const d = document.createElement("div");
  d.textContent = str;
  return d.innerHTML;
}
