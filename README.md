# MB & VB Group Capacity Planner

A garment manufacturing capacity planning, booking, and reporting tool for MB & VB Group —
built as a React app.

## What's in here

- `src/App.jsx` — the entire application (data model, capacity engine, all screens).
- `src/main.jsx` — entry point. Includes a `localStorage`-backed shim for `window.storage`,
  since that API only exists inside Claude.ai's artifact sandbox where this was originally
  built. With the shim, the app persists its data in the browser's local storage instead —
  same behavior, works in any normal deployment.
- `index.html`, `vite.config.js`, `package.json` — a minimal Vite project wrapper so the JSX
  can actually be built and served.

## Run it locally

```bash
npm install
npm run dev
```

Opens at `http://localhost:5173`.

## Push to GitHub

From inside this folder:

```bash
git init
git add .
git commit -m "Initial commit — MB & VB Group Capacity Planner"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo-name>.git
git push -u origin main
```

(Create the empty repo on GitHub first, then use the URL it gives you in place of the one above.)

## Deploy it somewhere people can actually open it

A GitHub repo alone doesn't serve a live website — pick one of these:

**Vercel or Netlify (easiest)**
Connect the GitHub repo in either dashboard, build command `npm run build`, output directory
`dist`. Every push to `main` auto-deploys. No config needed beyond that.

**GitHub Pages**
1. `npm install` (pulls in the `gh-pages` package already listed in `package.json`)
2. In `vite.config.js`, set `base: "/your-repo-name/"` (matching your actual repo name)
3. `npm run build && npm run deploy`
4. In the repo's Settings → Pages, set the source to the `gh-pages` branch

## Important: data storage is per-browser, not shared

Since this build uses `localStorage` (via the shim), everyone who opens the deployed site gets
their **own independent copy** of the data, stored only in their own browser. Bookings one
person enters won't be visible to anyone else, and clearing browser data wipes it. That's fine
for trying the tool out or single-user use, but if you need the whole PPC/Merchandising team
working off one shared, always-up-to-date dataset, this needs a real backend and database next
— the front-end is ready for that, it just currently has no server to talk to.

## Known limitations carried over from the original build

- The PPC password gate is a UI-level convenience, not real security — no backend enforces it.
- "Booking confirmation email" opens the browser's own mail client with the message pre-filled;
  it doesn't send email automatically from a server.
