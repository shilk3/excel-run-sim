# Cell Grind — Excel Esports Manager

A mobile-first simulation game: manage a rising Excel esports competitor. Every
day has up to 24 hours — split them across Skill Training, Gym, Sleep,
Relax, and Food. Everything is connected:

- **Skill Training** raises one of 7 case specialties — Data Analysis,
  Mapping, Text Processing, Game Logic, Math & Formulas, Time & Dates, and
  Cards & Random. During the season only 1-3 skills are "active"
  (trainable) each round, revealed at the start of that round's week —
  every match that week tests exactly those skills. The other 4-6 sit
  locked and slowly rust. Preseason and the off-season (including training
  camp) have no matches, so all 7 are open.
- **Gym** raises Physical Health — but low Physical Health caps how much
  your skill training actually helps.
- **Sleep** builds Rest. Skimp on sleep and Rest drains, which quietly wears
  down your Physical Health even if you train hard.
- **Relax** builds Composure. Train hard with no downtime and Composure
  drains until you burn out, tanking your effectiveness until you recover.
- **Food** builds Nutrition, which keeps tomorrow's day at full length — see below.
- **Decay is universal**: every stat needs upkeep or it slips — an
  inactive skill, Gym, Sleep, Relax, or Food below its
  threshold hours (shown as a marker on each slider) causes that stat to
  fall instead of rise.
- **Rest** swings training itself: well-rested days train up to 200%
  as effectively. **Composure** swings match day specifically: low
  Composure can cut your active skills' effect on a match in half.
  **Nutrition** swings how many hours you get at all: below 40 your day
  shrinks to 16h, sliding up to a full 24h at 90+.
- Every Gym hour carries a 2.5% injury chance that day (8h, the max, is
  20%; Sports Physio cuts it). An injury costs 5–10 Health and locks the
  Gym for 3–5 days, but never stops you playing matches.
- Each skill caps at 50 until you invest match winnings in that skill's
  dedicated Coach (5 levels, Coaching Shop) — and the effective ceiling is
  also capped by the highest league you've ever reached, from 60 in League 5
  up to 100 in League 1. Both gates must be cleared to hit 100. Physical
  Health caps at 70 until you invest in Sports Physio.
- **Work** funds everything else. 9h/day, every phase, no exceptions — you
  start with 3.0 chances, and the bar shows exactly what today's shortfall
  will cost before you end the day. Pay is tied to still *having* the job,
  not to hitting the exact hour target every day: come up short and you
  lose a chunk of a chance scaled to the shortfall (a near-miss costs
  barely anything, skipping the day entirely costs a full chance,
  regained after 3 weeks), but you're still paid in full. Only running out
  of chances actually costs you income — you're fired, and the same slider
  becomes a Job Search until you log 30 cumulative hours. Reach League 2
  with $5,000 banked while employed and you go pro automatically: Work
  drops to 5h/day of Pro Duties, same chances rule. Pros also have to stay
  current — a new Excel technique appears roughly every 30 days needing
  25-40h to master (any Pro Duties hours beyond the 5h minimum go toward
  it); falling behind never costs you progress, but every not-yet-mastered
  technique costs 5% match performance.
- **Cost of living**: $50/day, charged no matter what — employed or not,
  every phase. Pay isn't flat: Work starts at $70/day and Pro Duties at
  $100/day, each rising $10/year for your first 5 years in that role before
  plateauing at $120 and $150. Lose the job or get dropped from Pro and
  that role's pay resets to its minimum for next time — seniority isn't
  carried over. Every new year opens with a cash-flow summary of the year
  just finished — pay earned, match winnings, expenses paid, the net, and
  any raise you just earned.

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
simulated results every week, exactly like yours — it's a living world, not
scenery. Every tier runs a genuine round-robin schedule in lockstep with
your own matches, so all 5 tables are live and visible from round 1 of the
season, not just once it ends. Four go up from every league below League 1:
the playoff champion, plus the top 3 of the table other than the champion
— so a top-3 finish is always promoted, and anyone who makes the playoffs
can still win their way up. Finish bottom 4 and you're relegated. This
applies to every competitor in every league, not just you — every league
plays out its own knockout too — so rivals you've never played can climb or
fall in the background. Promotion/relegation is applied once the playoffs
are over. Check the Leagues screen any time to see all 5 tables, or jump
there directly from the Career modal.

Rating, Cash, stats, and Coaching Shop upgrades all carry over between years
and leagues.

## Playing it on your iPhone

This is a installable web app (PWA) — no App Store or Xcode needed.

1. Host the folder (see below) or open `index.html` directly in Safari.
2. In Safari, tap the **Share** icon → **Add to Home Screen** → **Add**.
3. Launch it from your home screen — it runs full-screen and saves your
   progress on-device (`localStorage`), with basic offline support via a
   service worker.
4. Safari and the Home Screen app keep **separate saves**. To move a career
   between them (or to another device), open ☰ → **💾 Export / Import Save**,
   tap **Copy save code** (or Share / Save to Files), then paste it into
   **Import** on the other side. The code is your whole career, gzipped and
   base64'd (~20 KB).

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
