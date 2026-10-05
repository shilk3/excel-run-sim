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
- **Sleep** builds Rest: 7h holds it, each hour short costs 1.3 a day,
  each hour over adds 1.4. Rest below 80 wears down Physical Health and
  Calm, more the lower it goes. Sleep also adds +1.5 Calm a day
  at 7h, +3 at 8h or more.
- **Relax** builds Calm on its own (training doesn't drain it): 3h
  holds it, each hour short costs 2.4 a day, each hour over adds 2.4.
  Relax past 100 isn't wasted: Calm past 100 becomes 😎 Chill, a reserve (half rate,
  up to 130, fading 1 a day) that drains first on short-Relax days — match
  day still counts Calm as 100 at most. Let it hit 0 and you burn out,
  tanking your training until you recover.
- **Food** builds Nutrition, which keeps tomorrow's day at full length — see below.
  **Carb loading:** once Nutrition is already at 100, each Food hour past 1h banks one extra hour for the next
  day only (up to +3h at 4h of Food) — it moves hours between days, 1:1.
  End Week stops on a carb-loaded day so the extra hours can be planned.
- **Decay is universal**: every stat needs upkeep or it slips — an
  inactive skill, Gym, Sleep, Relax, or Food below its
  threshold hours (shown as a marker on each slider) causes that stat to
  fall instead of rise.
- **Rest** swings training itself (skills and Gym), smoothly: ×1 at 70
  Rest or below, rising to ×2 at 100 (85 Rest trains at ×1.5).
  **Overcharge:** with Rest already at 100, sleeping past 7h pushes that
  day's training higher still — ×3 at 10h, ×4 at 12h, linear between. **Calm** swings match day specifically: low
  Calm can cut your active skills' effect on a match in half.
  **Nutrition** swings how many hours you get at all: below 40 your day
  shrinks to 16h, sliding up to a full 24h at 90+.
- Every Gym hour carries a 2.5% injury chance that day (8h, the max, is
  20%; Sports Physio cuts it). An injury costs 5–10 Health and locks the
  Gym for 3–5 days, but never stops you playing matches.
- **Energy Items** (in Staff & Items) are instant top-ups, each usable
  once a week: ☕ Coffee $40 (+10 Rest, −4 Rest the next day), 🧃 Energy
  Drink $90 (+15 Rest, +10 Calm, then −5 Rest, −3 Calm), 💆 Spa
  Day $250 (+10 Rest, +30 Calm, no crash). Nothing goes past 100.
- **Equipment** (also in Staff & Items) is bought outright and works until
  the current year ends, then wears out — never pro-rated, so buy early:
  🖥️ Second Monitor $700 (+5% skill training), 💻 Gaming PC $5,000
  (+10%), 🪑 Ergonomic Chair $2,400 (−20% Gym injury risk), 🎧
  Noise-cancelling Headphones $900 (+0.5 Calm a day), 🛏️ Memory-foam
  Mattress $2,800 (−25% Rest lost from short sleep).
- **Prize money**: a win pays $60 + your new rating ÷ 20 (about $85 at
  rating 500); a loss pays nothing. Winning the playoffs adds a $1,000
  champion bonus (+40 rating).
- Skills train up to 50 on your own. **Staff** are hired a week at a time
  (a week ends after each match), paid up front, pro-rated if hired
  mid-week. Every coach and support-team member has an optional
  🔁 Auto-rehire switch that keeps them on at the same level each new week
  (a coach even when their skill isn't in that week's focus)
  while you can afford it (if you can't, they're off that week, the switch
  turns itself off and you're told). A hired Coach lifts that
  skill's ceiling (Lv1 60 … Lv5 100), speeds training and stops that skill
  rusting while hired (even in a week you don't train it), but never past
  your league cap (60 in League 5 up to 100 in League 1). Above the ceiling
  a skill holds with an hour a day and slips with less. Physical Health caps
  at 70 without a hired Sports Physio. Higher staff levels need a one-off
  fee and a high enough league, and cost more per week; wages are tuned so
  you can afford roughly 90% of what you'd want in a 1-skill week and 40% in
  a 3-skill week.
- **Work** funds everything else. 9h/day, every phase, no exceptions (up
  to 4h of overtime on top pays half your hourly rate — $4/h at the
  starting $70/day, rising with raises) — you
  start with 3.0 chances, and the bar shows exactly what today's shortfall
  will cost before you end the day. Pay is tied to still *having* the job,
  not to hitting the exact hour target every day: come up short and you
  lose a chunk of a chance scaled to the shortfall (a near-miss costs
  barely anything, skipping the day entirely costs a full chance,
  regained after 3 weeks), but you're still paid in full. Only running out
  of chances actually costs you income — you're fired, and the same slider
  becomes a Job Search until you log a random 10–40 cumulative hours
  (rolled each time you lose the job, and shown on the slider) — at least 1h
  of searching a day, or the day can't end, so you can't idle into debt. Reach League 2
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
  carried over. The new-year screen includes a cash-flow summary of the
  year just finished — pay earned, match winnings, expenses paid, the net,
  and any raise you just earned.

**End Day / To Match**: End Day plays one day. To Match (End Week outside the
season, up to 7 days) repeats today's plan up to and including the next match,
then shows the result plus a week summary of how every stat moved. It stops
early so you can re-plan on an injury, losing/finding a job or going pro,
burnout, a new technique to master, or Nutrition shrinking the day below what
your plan needs. Each of those events also gets its own screen whenever it happens —
after End Day too — saying what changed and what to do about it.

**Match day**: both sides get a match-day rating = rating + performance +
luck, and the higher one wins. Your performance comes from your stats (each
point above 70 adds 3, below 70 costs 3); rivals get a performance on the same
scale, tracking their rating. Luck is random for both sides every match,
drawn so the odds of winning match the classic Elo win chance for the rating
gap — upsets always stay possible — and the scores shown are scaled so most
games read as close. The result
screen shows what changed (rating, table position, cash, record), both sides'
numbers side by side, and collapsible explanations. League tables rank by points, then
head-to-head between tied players, then rating; they show W and L too.

**The season**: a 14-day preseason to train, then a 39-round regular season —
one match a week against a named rival, the full schedule known in advance.
Finish in the top 16 of your 40-competitor league to reach the knockout
playoffs (Round of 16 → Quarterfinal → Semifinal → Final, single elimination).
Lose a playoff match and you're out; win the Final and you're champion. Miss
the playoffs and you get a 28-day training camp instead of a short offseason,
so missing the cut is never a worse deal than qualifying and getting knocked
out early. Get knocked out and you go to training camp too, for the rest of
the playoff window (Round of 16 exit: 21 days, Quarterfinal: 14, Semifinal or
Final: 7), so an early exit gets its time back as training, just like missing
the cut. Only the champion gets a plain 7-day offseason.

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

**Phase screens**: every change of phase (preseason → season, season →
playoffs or training camp, playoffs → camp/offseason, offseason → the new
year) opens a screen that sums up the phase just finished — table finish,
rating, cash and biggest skill gains — and says what's coming next. A new
career starts with a five-page quick tutorial right after you name your
player; replay it any time from the menu (📖 Quick Tutorial).

**Match History** (menu, the Log header, Leagues and Career) has an
**Upcoming** tab — the rest of the regular season with each opponent's rating,
table position and your win chance, then your path through the playoff
bracket once it starts, plus the whole playoff bracket (seeds, results, the
champion) during the playoffs and offseason — and keeps every match
you play — tap one to reopen its full result screen — plus every result in
all 5 leagues, round by round, for this season and last. Tap any rival's
name (there or in a league table) to see their season so far.

Rating, Cash, stats, and unlocked staff levels all carry over between years
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
