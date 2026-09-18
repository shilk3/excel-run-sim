# Cell Grind — Excel Esports Manager

A mobile-first simulation game: manage a rising Excel esports competitor. Every
day has 24 hours — split them across Excel Training, Running, Cross Training,
Sleep, and Relaxation. Everything is connected:

- **Excel Training** raises your Excel Skill, the core competitive stat.
- **Running** and **Cross Training** raise Physical Health — but low Physical
  Health caps how much your Excel Training actually helps.
- **Sleep** restores Energy and pays down Sleep Debt. Skimp on sleep and
  Sleep Debt quietly wears down your Physical Health, even if you train hard.
- **Relaxation** relieves Stress. Train hard with no downtime and Stress
  builds until you burn out, tanking your effectiveness until you rest.
- Overtrain physically (too much Running + Cross Training) and you risk an
  injury that locks out physical activities for several days.

Every 7 days you compete in a ranked match. Win to climb the rank ladder and
earn prize money, then spend it in the Coaching Shop on permanent upgrades.

## Playing it on your iPhone

This is a installable web app (PWA) — no App Store or Xcode needed.

1. Host the folder (see below) or open `index.html` directly in Safari.
2. In Safari, tap the **Share** icon → **Add to Home Screen** → **Add**.
3. Launch it from your home screen — it runs full-screen and saves your
   progress on-device (`localStorage`), with basic offline support via a
   service worker.

### Running it locally

It's a static site with no build step or dependencies:

```sh
python3 -m http.server 8000
# then open http://localhost:8000 on your phone (same network) or in a browser
```

For access from an iPhone on the same Wi-Fi, serve it from your computer and
open `http://<your-computer-ip>:8000` in Safari, or deploy the folder as-is
to any static host (GitHub Pages, Netlify, Vercel, etc.) and open that HTTPS
URL in Safari.

## Files

- `index.html` — markup/layout
- `style.css` — mobile-first dark UI
- `game.js` — game state, daily simulation, match/shop logic, rendering
- `manifest.webmanifest`, `sw.js`, `icons/` — PWA install + offline support
