# Cell Grind — Excel Esports Manager

A mobile-first simulation game: manage a rising Excel esports competitor. Every
day has up to 24 hours — split them across Skill Training, Exercise, Sleep,
Relaxation, and Nutrition. Everything is connected:

- **Skill Training** raises one of 7 case specialties — Data Analysis,
  Mapping, Text Processing, Game Logic, Math & Formulas, Time & Dates, and
  Cards & Random. Only 1-3 skills are "active" (trainable) each round,
  revealed at the start of that round's week — every match that week tests
  exactly those skills. The other 4-6 sit locked and slowly rust.
- **Exercise** raises Physical Health — but low Physical Health caps how much
  your skill training actually helps.
- **Sleep** builds Rest. Skimp on sleep and Rest drains, which quietly wears
  down your Physical Health even if you train hard.
- **Relaxation** builds Composure. Train hard with no downtime and Composure
  drains until you burn out, tanking your effectiveness until you recover.
- **Nutrition** keeps tomorrow's day at full length — see below.
- **Decay is universal**: every stat needs upkeep or it slips — an
  inactive skill, Exercise, Sleep, Relaxation, or Nutrition below its
  threshold hours (shown as a marker on each slider) causes that stat to
  fall instead of rise.
- **Rest** swings training itself: well-rested days train up to 200%
  as effectively. **Composure** swings match day specifically: low
  Composure can cut your active skills' effect on a match in half.
  **Nutrition** swings how many hours you get at all: below 40 your day
  shrinks to 16h, sliding up to a full 24h at 90+.
- Overtrain physically and you risk an injury that locks out Exercise for
  several days.
- Each skill caps at 50 until you invest match winnings in that skill's
  dedicated Coach (5 levels, Coaching Shop) — and the effective ceiling is
  also capped by the highest league you've ever reached, from 60 in League 5
  up to 100 in League 1. Both gates must be cleared to hit 100. Physical
  Health caps at 70 until you invest in Sports Physio.

**The season**: a 14-day preseason to train, then a 39-round regular season —
one match a week against a named rival, the full schedule known in advance.
Finish in the top 16 of your 40-competitor league to reach the knockout
playoffs (Round of 16 → Quarterfinal → Semifinal → Final, single elimination).
Lose a playoff match and you're out; win the Final and you're champion. Miss
the playoffs and you get a 28-day training camp instead of a short offseason,
so missing the cut is never a worse deal than qualifying and getting knocked
out early.

**Leagues**: 5 tiers (League 1 at the top, League 5 at the bottom — new
careers start at the bottom). Every league has a persistent roster of 199
named rivals total across all 5 tiers, whose ratings evolve from real
simulated results every season, exactly like yours — it's a living world,
not scenery. Finish top 4 of your league's table and you're promoted a tier;
finish bottom 4 and you're relegated. This applies to every competitor in
every league, not just you, so rivals you've never played can climb or fall
in the background too. Promotion/relegation is based purely on table
position — the playoffs are a separate prize, unrelated to which league you
play in next year. Check the Leagues screen any time to see all 5 tables.

Rank, Cash, stats, and Coaching Shop upgrades all carry over between years
and leagues.

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
