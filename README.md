# TN-Reservoir

**Live demo:** https://tioluwanimofe.github.io/Tn-reservoir-pwa/ (demo mode, admin PIN `1234`)

A small installable web app (PWA) for reserving locations. Teams send a request for a location, date and time, and an admin approves, declines or cancels it. Everything is stored in a Google Sheet through Google Apps Script, so there is no server to run.

## What it does

- **Book:** enter your name, pick your team, a location and a time. It shows what's already booked that day and warns about clashes.
- **Locations:** every location, grouped by campus, with whether it's free right now and what's coming up.
- **Calendar:** a month view of all bookings, filterable by location.
- **Day view:** a timeline of every location for one day.
- **My requests:** requests sent from this device, with their status. People can cancel their own requests without an account.
- **Admin:** behind a PIN. Approve, decline or cancel requests (with a note), and add, hide or delete locations. It won't let two approved bookings overlap.

Teams: NF1–NF4, TN1–TN4, Center Workers 15/17/18/19/20, MWG.
Starting locations: Vandy Buttrick Hall, Vandy Calhoun, Vandy Student Center, Vandy Wilson Hall (Vanderbilt) and Belmont Barbara Massey (Belmont).

With `API_URL` empty in `app/config.js`, the app runs in **demo mode**: example data kept in the browser, admin PIN `1234`.

## Repo layout

```
app/            The PWA. Deploy this folder.
  index.html    The whole app (HTML, CSS, JS in one file)
  config.js     Where the Apps Script URL goes
  sw.js         Service worker (caches the app shell, never the API)
  manifest.webmanifest, icons
  demo-body.html  Same app without the <head> wrapper (used for the hosted demo)
apps-script/
  Code.gs       Google Apps Script backend
```

## Setup

1. **Sheet + backend**
   - Create a Google Sheet, then open Extensions → Apps Script and paste in `apps-script/Code.gs`.
   - Set `ADMIN_PIN` at the top to your own PIN. Don't commit your real PIN back to this repo.
   - Run `setup` once and approve permissions. This creates the `Buildings` and `Bookings` tabs and adds the starting locations.
   - Deploy → New deployment → Web app. Execute as **Me**, access **Anyone**. Copy the `/exec` URL.
2. **Connect the app:** paste that URL into `API_URL` in `app/config.js`.
3. **Host it** (either one works):
   - **Netlify:** in Netlify, choose **Add new site → Import an existing project → GitHub** and pick this repo. `netlify.toml` already tells Netlify to publish `app/` with no build step, so just click **Deploy**. Every push to `main` redeploys.
   - **GitHub Pages:** already set up. `.github/workflows/pages.yml` deploys `app/` on every push to `main`.
4. Open it on your phone and choose **Add to Home Screen**.

After changing `Code.gs`: Deploy → Manage deployments → edit → Version: New version → Deploy. The URL stays the same.

To change the teams, edit `TEAM_GROUPS` in `app/index.html` and `TEAMS` in `apps-script/Code.gs`.

## How the backend works

The app sends every call as a `POST` with a `text/plain` JSON body, which avoids a CORS preflight against Apps Script. Writes run under `LockService` so two people can't grab the same slot at once. Each request gets a private token stored on the device that sent it; that's what lets people cancel their own request without signing in.
