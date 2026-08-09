# DECK — Command Console

A dark, "flight-deck" style personal dashboard: live meeting panel, Gmail, Google Calendar,
Google Chat launcher, and an ElevenLabs conversational assistant docked on the side.

## What this can and can't do (read this first)

- **Gmail & Calendar** are pulled live via the official Google APIs and rendered in the
  dashboard's own UI. This works fully once you've set up OAuth (below).
- **Zoom** can run *truly embedded* inside the meeting panel via the Zoom Meeting SDK.
- **Google Meet and Microsoft Teams cannot be embedded live** inside a third-party page —
  neither Google nor Microsoft expose that capability to outside developers. For those, the
  dashboard opens the call in its own window and shows a status card in the meeting panel.
  This is a platform limitation, not something extra code can fix.
- **Google Chat** requires Workspace-admin-approved OAuth scopes to read messages via API.
  Until your admin (or you, if you're the admin at protem.solutions) approves the app, the
  Chat panel just gives you a one-click launch into chat.google.com.

## 1. Create a Google OAuth Client

1. Go to [Google Cloud Console](https://console.cloud.google.com/) → create (or pick) a project.
2. **APIs & Services → Library** → enable:
   - Gmail API
   - Google Calendar API
   - Google Chat API
3. **APIs & Services → OAuth consent screen** → set **User type** to **Internal**. Since
   you're the Workspace admin for protem.solutions, Internal apps skip Google's public
   verification process entirely — including for the Chat scopes, which are otherwise
   "restricted" and normally require a verification review. This only works because the app
   is limited to your own domain's users.
4. **APIs & Services → Credentials → Create Credentials → OAuth client ID**
   - Application type: **Web application**
   - Authorized JavaScript origins: add the URL you'll host this on, e.g.
     `https://deck.yourdomain.com` or `http://localhost:5500` for local testing.
5. Copy the Client ID into `app.js` → `CONFIG.GOOGLE_CLIENT_ID`.

Google's sign-in library will not run from a `file://` path — you must serve this over
`http://localhost` or a real domain (see hosting section below).

## 2. Google Chat feed

Once the Chat API is enabled and the consent screen is Internal (step 1 above), the Chat
panel works with no extra setup — it lists your spaces, pulls the latest message from each,
and links out to `chat.google.com` when you click one. If it shows an error instead:
- Double check the Chat API is enabled for the project (not just Gmail/Calendar).
- Confirm the consent screen is Internal, not External/Testing.
- Re-authenticate (sign out and back in) so the browser picks up the new Chat scopes — token
  requests only include scopes granted at sign-in time.

## 3. Enable live Zoom embedding (recommended — this is the one platform that truly embeds)

Unlike Meet/Teams, Zoom has a genuine embeddable SDK, so this is worth setting up if you use
Zoom at all — it's the only provider where the call actually runs *inside* the dashboard with
real, working controls, not a docked separate window.

1. Go to the [Zoom App Marketplace](https://marketplace.zoom.us/) → **Develop → Build App**
   → choose **Meeting SDK**.
2. On the app's **App Credentials** page, copy the **Client ID** (older guides call this
   "SDK Key" — Zoom renamed it in mid-2026, same thing) into `app.js` → `CONFIG.ZOOM_SDK_KEY`.
3. The **Client Secret** must never sit in browser code — it signs every join request and has
   to stay server-side. A ready-to-run signer is included in `zoom-signature-server/`:
   ```bash
   cd zoom-signature-server
   npm install
   cp .env.example .env
   # edit .env and paste in your Client ID and Client Secret
   node server.js
   ```
   Leave that running in its own Terminal window/tab alongside your `python3 -m http.server`
   for the main dashboard. By default it listens on `http://localhost:3001`, which is already
   what `CONFIG.ZOOM_SIGNATURE_ENDPOINT` in `app.js` points to — no change needed unless you
   run it on a different port or deploy it somewhere else later.

If you skip this, the Zoom option will prompt you to add it and the panel will fall back to
opening the call in a new window instead, same as Meet/Teams.

## 4. Enable the ElevenLabs assistant

1. In [ElevenLabs](https://elevenlabs.io/app/conversational-ai), create a Conversational AI

   agent and copy its **Agent ID**.
2. Paste it into `app.js` → `CONFIG.ELEVENLABS_AGENT_ID`.

## 4. Run it

**Locally (quick test):**
```bash
cd dashboard
python3 -m http.server 5500
# open http://localhost:5500
```
Add `http://localhost:5500` as an authorized JavaScript origin in step 1.4 above.

**Properly hosted (recommended for daily use):**
Deploy the folder as a static site (Vercel, Netlify, Cloudflare Pages, or your own server),
add that domain as an authorized origin in Google Cloud, and bookmark it / pin it as a PWA
so it opens full-screen.

## About the Meet/Teams window docking

Google Meet and Microsoft Teams don't offer any way to run their live call UI *inside*
another site — there's no embed SDK for either, unlike Zoom. So instead, the dashboard opens
the call in its own OS window and sizes/positions it over the meeting panel's on-screen
rectangle using `window.open` with explicit `left/top/width/height`.

**Important limitation, discovered in practice:** browsers only allow the opener page to move
or resize that window while it's still on the *same origin* it was opened from. The instant it
navigates to `meet.google.com` or `teams.microsoft.com`, Chrome (and other browsers) blocks
any further `moveTo`/`resizeTo` calls with a `SecurityError` — this is deliberate cross-origin
protection against sites hijacking windows they don't control, not a bug in this code, and
there's no JavaScript-side workaround for it. In practice this means:
- The window opens at roughly the right size/position for an instant.
- The moment it finishes loading Meet/Teams (which often resizes itself, e.g. moving from the
  "ready to join" lobby to the live call view), the dashboard can no longer correct that —
  Recenter window will silently stop working too, and the panel will say so honestly rather
  than keep failing quietly.
- From that point on it behaves as a fully independent window: move or resize it yourself as
  needed, same as any other window.

Zoom doesn't have this problem, since it's genuinely embedded via the Zoom Meeting SDK rather
than opened as a separate window — the video area you see for Zoom calls really is inside the
dashboard, with real, working Zoom controls (mute, camera, etc.), not a visual approximation.

Other caveats:
- Your browser must allow pop-ups for this site (add an exception if it's blocked).
- The window is still a real, separate OS window — the person on the call can drag it away,
  and it will show above other apps if you alt-tab.
- **Fullscreen and docking don't mix.** If the dashboard's own tab is in true browser
  fullscreen (via the ⛶ button, or the OS-level fullscreen control), a fullscreen window
  occupies the *entire physical display* — there's no screen space left for a second window
  to appear, so browsers fall back to opening the "popup" as a plain tab instead. This is
  correct, deliberate browser behavior, not a bug. The dashboard now automatically exits
  fullscreen right before opening a Meet/Teams window so docking still works — you'll notice
  the dashboard briefly drop out of fullscreen when you join. Re-enter fullscreen afterward
  with the ⛶ button if you want it back; browsers require a fresh click to re-enter
  fullscreen, so this can't be done automatically once the call ends.
- Closing that window (or ending the call) automatically returns the panel to idle.

The dashboard also tracks its own position and size continuously (not just on resize events),
so dragging the browser to a different monitor or resizing it repositions the docked window
to match, without you needing to hit Recenter manually — that's just there for edge cases the
polling doesn't catch quickly enough, or if the browser blocks the automatic move.

## Ops / Dev board switcher

The **OPS / DEV** toggle in the top bar swaps the calendar/mail/chat panels for a second set
of placeholder panels (Repos, Pipelines, Environments) in the same grid slots — the meeting
panel and assistant sidebar stay put either way, since a call or the assistant might be
relevant regardless of which board you're on. The Dev panels are empty stubs for now; when
you're ready to wire them up, follow the same pattern as `loadMail()` / `loadCalendar()` /
`loadChat()` in `app.js` — fetch from whatever API you're integrating, then render into the
matching `panel-body` element. Your last-selected board is remembered locally between visits.

## Using it

- **Sign in** with your Google account on first load.
- **Meeting panel**: pick Zoom, Meet, or Teams from the dropdown, paste the meeting link/ID,
  hit Join. The placeholder HUD animation shows whenever nothing is live.
- **Calendar & Mail** refresh automatically every 60s (configurable via
  `CONFIG.REFRESH_INTERVAL`), plus manual refresh buttons.
- **Assistant** sits docked on the right at all times, independent of the main grid.
- Press the ⛶ icon top-right for true browser fullscreen.

## Full interactivity (reply, compose, create, send)

Beyond just viewing feeds, the dashboard can now reply to and compose email, create
calendar events, and send Chat messages — directly from the panels. This needed broader
Google permissions than the original read-only setup:

- Gmail: `gmail.readonly`, `gmail.send`, **and** `gmail.modify` (needed for the delete/trash
  button — moving a message to Trash counts as a modify action, not just sending)
- Calendar: full `calendar` access (not just `calendar.readonly`)
- Chat: `chat.spaces.readonly` **and** `chat.messages` (read + write)

If you set this dashboard up before this feature existed, two things need doing:

1. **Google Cloud Console → APIs & Services → OAuth consent screen → Data Access** (or
   under Google Auth Platform → Data Access in the newer console) — add the four scopes
   above if they're not already listed. Internal apps still need scopes declared here even
   though they skip the public verification review.
2. **Sign out and back in** on the dashboard itself. Google only grants the scopes you
   consented to at sign-in time, so a previously-granted session won't pick up the new
   permissions automatically — the next "Sign in with Google" click will show an updated
   consent screen listing the new permissions.

### What each panel can do now

- **Mail**: click "Reply" on any message for a quote-free reply (threaded correctly via
  `In-Reply-To`/`References` headers), or "+ Compose" for a new email from scratch.
- **Calendar**: "+ New event" opens a form for title, date, start/end time, location, and
  optional guests — created via the Calendar API and (if guests are added) sends real
  invites.
- **Chat**: click any space to select it (highlighted with a cyan left border) — a message
  box appears at the bottom of the panel to send into that space.

### Still read/launch-only for now

- Editing or deleting existing calendar events, archiving/deleting email, and RSVPing to
  invites aren't wired up yet — they'd follow the same pattern as the code above
  (`createEvent()`, `sendEmail()`) if you want to extend it later.
- Rich formatting (HTML email, mentions in Chat) isn't included — everything sends as
  plain text for simplicity.

## Files

- `index.html` — structure
- `styles.css` — dark high-tech theme (all colors as CSS variables at the top)
- `app.js` — all logic; your keys go in the `CONFIG` block at the top
- `README.md` — this file

## Security notes

- Access tokens are held in memory only (never written to localStorage), so refreshing the
  page requires re-authenticating — this is intentional for a shared/always-on screen.
- Never put a Zoom SDK *Secret*, or any Google *Client Secret*, in this frontend code. Only
  the Client ID (public) and SDK Key (public) belong here.
- If you later add the Google Chat API, treat its scope request separately — it's the one
  piece here that touches company-wide data and should go through your Workspace admin.
