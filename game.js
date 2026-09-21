/* Cell Grind — Excel Esports Manager
 * Single-file game engine: state, daily simulation, season/league structure,
 * matches, shop, UI rendering.
 */

const APP_VERSION = "4.1.0";
const SAVE_KEY = "cellgrind_save_v1";

/* ---------------------------------------------------------------------- */
/* Balance constants — tune game feel here                                */
/* ---------------------------------------------------------------------- */
const BAL = {
  idealSleep: 8, // also the Rest decay threshold: sleep below this drains Rest
  relaxComposureThreshold: 3, // Composure's decay threshold, in relaxation hours
  skillDecayThresholdHours: 1, // an active skill/Exercise below this hour count rusts
  skillGainBase: 0.24,
  exerciseGainBase: 0.6,
  softFatigueCap: 8, // hours per activity before in-day fatigue kicks in
  hardFatigueCap: 12,
  fatigueMultSoft: 0.6, // effectiveness for hours between soft and hard cap
  fatigueMultHard: 0.3, // effectiveness for hours beyond hard cap
  energyDrainPerHour: { skill: 1.0, exercise: 1.3 },
  energyRestorePerSleepHour: 10.5,
  relaxEnergyRestore: 1.2,
  relaxComposureRelief: 2.4,
  composureLoadPerHour: 0.85,
  restGoodSleepBonus: 3, // extra composure relief when sleepHours >= ideal
  overtrainThreshold: 10, // exercise hours before injury risk starts
  injuryChancePerExcessHour: 0.045,
  injuryPhysLoss: [15, 25],
  injuryDaysRange: [3, 5],
  burnoutComposureThreshold: 0,
  burnoutRecoverThreshold: 35,
  burnoutEffectivenessMult: 0.2,
  detrainThresholdHours: 1, // exercise hours below this triggers slow detraining
  detrainDecay: 0.35,
  // Rest's training-effectiveness upside: well-rested days train harder, on
  // top of (not instead of) Energy's own effect on focus.
  restTrainingBoostThreshold: 80, // <=80 Rest: 100% effective
  restTrainingBoostHigh: 90, // 81-90 Rest: 150%; >90 Rest: 200%
  // Composure's match-day effect: low Composure halves your skill on the
  // day it matters most, recovering in the same step pattern in reverse.
  composureMatchLow: 10, // <10 Composure: 50% skill
  composureMatchMid: 20, // 10-19: 75% skill; >=20: 100% skill
  // Season structure
  preseasonDays: 14,
  seasonRounds: 39,
  roundIntervalDays: 7,
  playoffSize: 16,
  offseasonDays: 7,
  // Missing the playoffs ends your season 4 rounds early compared to a
  // Final run — give that time back as an explicit training camp instead
  // of just a short generic break, so missing the cut isn't strictly worse
  // for preparing next season than qualifying and getting knocked out fast.
  trainingCampDays: 28,
  statCapBase: 70, // Physical Health ceiling with zero relevant upgrades
  statCapPerLevel: 10, // + this much per Physio level (max 3 levels -> +30 -> 100)
  // Skill ceilings stack two independent gates: the Coaching Shop (per
  // skill, 5 levels) and the league you've ever reached (peak, not
  // current — getting relegated doesn't lower it). Effective cap is the
  // lower of the two.
  skillShopCapBase: 50,
  skillShopCapPerLevel: 10, // levels 1-5 -> 60/70/80/90/100
  // League structure: 5 tiers, 40 competitors each (199 persistent rivals +
  // the player, who occupies one slot in whichever tier they're currently in).
  leagueCount: 5,
  leagueSize: 40,
  promotionCount: 4,
  relegationCount: 4,
};

// Skill cap granted by the highest league tier ever reached (1 = top).
const LEAGUE_SKILL_CAP = { 1: 100, 2: 90, 3: 80, 4: 70, 5: 60 };

/* ---------------------------------------------------------------------- */
/* The 7 case specialties                                                 */
/* ---------------------------------------------------------------------- */
const SKILLS = [
  { key: "data", name: "Data Analysis", icon: "📈" },
  { key: "map", name: "Mapping", icon: "🗺️" },
  { key: "text", name: "Text Processing", icon: "📝" },
  { key: "games", name: "Game Logic", icon: "🎲" },
  { key: "math", name: "Math & Formulas", icon: "🔢" },
  { key: "time", name: "Time & Dates", icon: "⏱️" },
  { key: "cards", name: "Cards & Random", icon: "🃏" },
];
const SKILL_KEYS = SKILLS.map((s) => s.key);
function skillMeta(key) {
  return SKILLS.find((s) => s.key === key);
}
function coachKey(key) {
  return `coach_${key}`;
}
// Each round focuses on 1-3 randomly chosen skills, revealed only once the
// previous round's week is over — the other 4-6 skills can't be trained
// (and quietly rust) until they come up again.
function rollActiveSkills() {
  const n = randInt(1, 3);
  const pool = SKILL_KEYS.slice();
  const picked = [];
  for (let i = 0; i < n; i++) {
    const idx = randInt(0, pool.length - 1);
    picked.push(pool.splice(idx, 1)[0]);
  }
  return picked;
}
function isPreseason() {
  return state.seasonPhase === "preseason";
}
// Preseason has no matches and no case-type rotation yet — every skill is
// open for training. Outside preseason, only this week's revealed 1-3
// active skills are trainable.
function trainableSkills() {
  return isPreseason() ? SKILL_KEYS : state.activeSkills;
}

// Initial rating bands per league tier (1 = top, 5 = bottom) — only used to
// seed a fresh roster. Ratings drift from there via real simulated results.
const LEAGUE_RATING_BANDS = {
  1: [1350, 1900],
  2: [1000, 1400],
  3: [750, 1050],
  4: [550, 800],
  5: [350, 600],
};

// Each upgrade has up to 3 purchasable levels (5 for the 7 skill coaches).
// state.upgrades[key] stores the current level (0 = not purchased). Buying
// goes 0->1->2->... in order; each level's fields describe that level's
// total (not additive) effect.
const SKILL_COACH_LEVELS = [
  { cost: 150, bonus: 0.05, capAdd: 10 },
  { cost: 300, bonus: 0.1, capAdd: 20 },
  { cost: 500, bonus: 0.15, capAdd: 30 },
  { cost: 800, bonus: 0.22, capAdd: 40 },
  { cost: 1300, bonus: 0.3, capAdd: 50 },
];
const UPGRADES = {};
SKILLS.forEach((sk) => {
  UPGRADES[coachKey(sk.key)] = {
    name: `${sk.name} Coach`,
    icon: sk.icon,
    levels: SKILL_COACH_LEVELS.map((lvl) => ({
      cost: lvl.cost,
      bonus: lvl.bonus,
      desc: `+${Math.round(lvl.bonus * 100)}% ${sk.name} training gains, skill ceiling ${BAL.skillShopCapBase + lvl.capAdd}`,
    })),
  };
});
Object.assign(UPGRADES, {
  physio: {
    name: "Sports Physio",
    icon: "🩺",
    levels: [
      { cost: 300, injuryReduceMult: 0.8, exerciseBonus: 0.05, desc: "-20% injury risk, +5% Exercise gains, health ceiling 80" },
      { cost: 650, injuryReduceMult: 0.65, exerciseBonus: 0.1, desc: "-35% injury risk, +10% Exercise gains, health ceiling 90" },
      { cost: 1200, injuryReduceMult: 0.5, exerciseBonus: 0.18, desc: "-50% injury risk, +18% Exercise gains, health ceiling 100" },
    ],
  },
  sleepApp: {
    name: "Sleep Coach App",
    icon: "📱",
    levels: [
      { cost: 250, sleepDebtMult: 0.85, desc: "-15% Rest lost from poor sleep" },
      { cost: 550, sleepDebtMult: 0.75, desc: "-25% Rest lost from poor sleep" },
      { cost: 1000, sleepDebtMult: 0.6, desc: "-40% Rest lost from poor sleep" },
    ],
  },
  nutritionist: {
    name: "Nutritionist",
    icon: "🥗",
    levels: [
      { cost: 300, decayMult: 0.85, desc: "-15% physical decay from low Rest" },
      { cost: 650, decayMult: 0.75, desc: "-25% physical decay from low Rest" },
      { cost: 1150, decayMult: 0.6, desc: "-40% physical decay from low Rest" },
    ],
  },
  meditation: {
    name: "Meditation Coach",
    icon: "🧘",
    levels: [
      { cost: 200, reliefMult: 1.2, desc: "+20% Composure relief from Relaxation" },
      { cost: 450, reliefMult: 1.4, desc: "+40% Composure relief from Relaxation" },
      { cost: 850, reliefMult: 1.65, desc: "+65% Composure relief from Relaxation" },
    ],
  },
  recovery: {
    name: "Recovery Program",
    icon: "🧊",
    levels: [
      { cost: 350, injuryDaysReduce: 1, detrainMult: 0.7, desc: "-1 day injury duration, -30% detraining" },
      { cost: 750, injuryDaysReduce: 2, detrainMult: 0.45, desc: "-2 days injury duration, -55% detraining" },
      { cost: 1400, injuryDaysReduce: 3, detrainMult: 0.2, desc: "-3 days injury duration, -80% detraining" },
    ],
  },
  manager: {
    name: "Team Manager",
    icon: "💼",
    levels: [
      { cost: 400, rankLossMult: 0.9, cashBonusMult: 1.0, desc: "-10% rank lost on defeat" },
      { cost: 800, rankLossMult: 0.8, cashBonusMult: 1.05, desc: "-20% rank lost on defeat, +5% prize money" },
      { cost: 1500, rankLossMult: 0.65, cashBonusMult: 1.1, desc: "-35% rank lost on defeat, +10% prize money" },
    ],
  },
});

function upgradeLevel(key) {
  return state.upgrades[key] || 0;
}
function upgradeEffect(key) {
  const lvl = upgradeLevel(key);
  return lvl > 0 ? UPGRADES[key].levels[lvl - 1] : null;
}
function physCap() {
  return BAL.statCapBase + BAL.statCapPerLevel * upgradeLevel("physio");
}
// Effective skill ceiling stacks two independent gates: the Coaching Shop
// (per skill, 5 levels) and the highest league ever reached (peak tier).
function skillShopCap(key) {
  return BAL.skillShopCapBase + BAL.skillShopCapPerLevel * upgradeLevel(coachKey(key));
}
function leagueSkillCap() {
  return LEAGUE_SKILL_CAP[state.peakLeagueTier] || LEAGUE_SKILL_CAP[5];
}
function skillCap(key) {
  return Math.min(skillShopCap(key), leagueSkillCap());
}
function restTrainingMultiplier(rest) {
  if (rest > BAL.restTrainingBoostHigh) return 2.0;
  if (rest > BAL.restTrainingBoostThreshold) return 1.5;
  return 1.0;
}
function composureMatchMultiplier(composure) {
  if (composure >= BAL.composureMatchMid) return 1.0;
  if (composure >= BAL.composureMatchLow) return 0.75;
  return 0.5;
}

/* ---------------------------------------------------------------------- */
/* Opponent name generation                                               */
/* ---------------------------------------------------------------------- */
const NAME_PARTS_A = [
  "Pivot", "Macro", "VLOOKUP", "CellBlock", "Formula", "Ctrl", "AutoFill", "RangeLord",
  "ByteSheet", "TabKing", "GridIron", "SUMIFS", "IndexMatch", "ChartWiz", "ClipboardX",
  "FreezePane", "HotKey", "DataDaemon", "SheetStorm", "RowRunner", "QuickSort", "QueryQueen",
  "QuietPivot", "QuantumCell", "FlashFill", "ArrayForm", "PowerQuery", "SolverSage",
];
const NAME_PARTS_B = [
  "Vex", "Prime", "Zero", "Byte", "Nova", "Reaper", "Ace", "Ghost", "Blaze", "Cross",
  "Wolf", "Fox", "Ninja", "Prodigy", "Legend", "Master", "Kid", "Pro", "Storm", "Flux",
  "Edge", "Rush", "Spark", "Wraith", "King", "Queen", "Titan", "Phantom",
];
function generateOpponentName() {
  const a = NAME_PARTS_A[randInt(0, NAME_PARTS_A.length - 1)];
  const b = NAME_PARTS_B[randInt(0, NAME_PARTS_B.length - 1)];
  return a + b;
}

/* ---------------------------------------------------------------------- */
/* Persistent league roster                                               */
/* ---------------------------------------------------------------------- */
// 199 persistent rivals distributed across the 5 leagues (40 per league),
// leaving exactly one slot in League 5 for the player to occupy. Every
// rival keeps its identity and rating for the life of the career, drifting
// via real simulated results and moving between tiers via promotion and
// relegation — same as the player.
function generateRivalRoster(playerTier) {
  const used = new Set();
  const rivals = [];
  let id = 0;
  // Whichever tier the player occupies gets 39 rivals instead of 40, leaving
  // the player's slot open; every other tier is a full 40.
  const countsByTier = { 1: 40, 2: 40, 3: 40, 4: 40, 5: 40 };
  countsByTier[playerTier] = 39;
  for (let tier = 1; tier <= 5; tier++) {
    const [lo, hi] = LEAGUE_RATING_BANDS[tier];
    for (let i = 0; i < countsByTier[tier]; i++) {
      let name;
      do {
        name = generateOpponentName();
      } while (used.has(name));
      used.add(name);
      rivals.push({ id: id++, name, rating: randInt(lo, hi), league: tier, wins: 0, losses: 0, promotions: 0, relegations: 0 });
    }
  }
  return rivals;
}

// One-time placement for saves migrating in from before leagues existed —
// buckets an already-earned rank into the tier whose band it best fits.
function rankToLeagueTier(rank) {
  if (rank >= LEAGUE_RATING_BANDS[1][0]) return 1;
  if (rank >= LEAGUE_RATING_BANDS[2][0]) return 2;
  if (rank >= LEAGUE_RATING_BANDS[3][0]) return 3;
  if (rank >= LEAGUE_RATING_BANDS[4][0]) return 4;
  return 5;
}

// Pulls the current 39 other members of the given league tier from the
// persistent roster and shuffles them into this season's fixture order.
function buildLeagueSchedule(rivals, leagueTier) {
  const others = rivals.filter((r) => r.league === leagueTier);
  const shuffled = others.slice();
  for (let i = shuffled.length - 1; i > 0; i--) {
    const j = randInt(0, i);
    const tmp = shuffled[i];
    shuffled[i] = shuffled[j];
    shuffled[j] = tmp;
  }
  return shuffled.map((r) => ({ rivalId: r.id, name: r.name, rating: r.rating, played: false, win: null }));
}

function findRival(id) {
  return state.rivals.find((r) => r.id === id);
}

/* ---------------------------------------------------------------------- */
/* State                                                                  */
/* ---------------------------------------------------------------------- */
function freshState() {
  const leagueTier = 5;
  const rivals = generateRivalRoster(leagueTier);
  return {
    day: 1, // total career days played — flavor/log only
    playerName: null,
    year: 1,
    seasonPhase: "preseason", // preseason | regular | playoffs | offseason
    phaseDay: 0, // days elapsed in the current phase
    roundIndex: 0, // next regular-season fixture index (0-38)
    leagueTier, // 1 (top) - 5 (bottom); new careers start at the bottom
    peakLeagueTier: leagueTier, // numerically lowest (best) tier ever reached
    rivals, // 199 persistent named rivals, spanning all 5 leagues
    leagueStandings: { 1: null, 2: null, 3: null, 4: null, 5: null }, // last completed season per tier, for the Leagues viewer
    schedule: buildLeagueSchedule(rivals, leagueTier),
    seasonResults: [], // { opponent, win } per completed regular-season round
    lastStandings: null,
    lastPlayerPosition: null,
    lastStandingsTier: null, // which league lastStandings was for (may differ from leagueTier after a promotion/relegation)
    offseasonDays: BAL.offseasonDays,
    offseasonReason: null, // "missed" | "playoffs" — set when entering offseason
    playoff: null, // { stage, currentRound, eliminated, champion, playerSeed, tier }
    cash: 100,
    rank: 480,
    peakRank: 480,
    wins: 0,
    losses: 0,
    stats: {
      skills: Object.fromEntries(SKILL_KEYS.map((k) => [k, 8])),
      phys: 65,
      energy: 80,
      rest: 100, // higher = better; drains when sleep < idealSleep
      composure: 92, // higher = better; drains when relax < relaxComposureThreshold
    },
    allocation: {
      skills: Object.fromEntries(SKILL_KEYS.map((k) => [k, 0])),
      exercise: 1,
      sleep: 8,
      relax: 2,
    },
    activeSkills: rollActiveSkills(), // 1-3 skills trainable this round; rerolled weekly
    skillCycleDay: 0, // days elapsed in the current 7-day skill-focus cycle
    injury: { active: false, daysLeft: 0 },
    burnout: { active: false, daysLeft: 0 },
    upgrades: Object.fromEntries(Object.keys(UPGRADES).map((k) => [k, 0])),
    logEntries: [],
  };
}

function storageAvailable() {
  try {
    const testKey = "__cellgrind_test__";
    localStorage.setItem(testKey, "1");
    localStorage.removeItem(testKey);
    return true;
  } catch (e) {
    return false;
  }
}

const STORAGE_OK = storageAvailable();
let lastLoadedFromSave = false;
let state = loadState();

function migrateSave(parsed) {
  // v1.2.0 merged the separate Running/Cross Training sliders into one
  // Exercise slider. Old saves still have { running, cross } instead.
  const alloc = parsed.allocation;
  if (alloc && alloc.exercise === undefined && (alloc.running !== undefined || alloc.cross !== undefined)) {
    alloc.exercise = clamp((alloc.running || 0) + (alloc.cross || 0), 0, 12);
    delete alloc.running;
    delete alloc.cross;
  }

  // v1.3.0 turned upgrades from owned:boolean into owned:level (0-3), and
  // added two new upgrade types. Rebuild upgrades from scratch so old
  // booleans convert to level 1 and any new/missing keys default to 0.
  const oldUpgrades = parsed.upgrades || {};
  const normalizedUpgrades = {};
  Object.keys(UPGRADES).forEach((key) => {
    const v = oldUpgrades[key];
    if (typeof v === "boolean") normalizedUpgrades[key] = v ? 1 : 0;
    else if (typeof v === "number") normalizedUpgrades[key] = v;
    else normalizedUpgrades[key] = 0;
  });
  parsed.upgrades = normalizedUpgrades;

  // v2.0.0 introduced the season/league structure (preseason -> 39-round
  // regular season -> top-16 playoffs -> offseason -> new year). Any save
  // that predates it needs a Year 1 reset; handled together with the
  // v3.0.0 migration below since every such save also predates leagues.
  if (!parsed.seasonPhase) {
    parsed.year = 1;
  }

  // v3.0.0 introduced 5 persistent leagues (199 rivals + the player, 40 per
  // tier) with promotion/relegation, replacing the old per-season randomly
  // generated 39-rival schedule. Any save without a roster gets one now,
  // placed into whichever tier its existing rank best fits, and starts a
  // fresh preseason there — reconciling an in-flight old-style season
  // against the new league structure isn't worth the complexity. This also
  // covers saves from before v2.0.0, which equally lack a roster.
  if (!parsed.rivals) {
    const tier = rankToLeagueTier(parsed.rank || 480);
    const rivals = generateRivalRoster(tier);
    parsed.rivals = rivals;
    parsed.leagueTier = tier;
    parsed.peakLeagueTier = tier;
    parsed.leagueStandings = { 1: null, 2: null, 3: null, 4: null, 5: null };
    parsed.seasonPhase = "preseason";
    parsed.phaseDay = 0;
    parsed.roundIndex = 0;
    parsed.seasonResults = [];
    parsed.lastStandings = null;
    parsed.lastPlayerPosition = null;
    parsed.offseasonDays = BAL.offseasonDays;
    parsed.offseasonReason = null;
    parsed.playoff = null;
    parsed.schedule = buildLeagueSchedule(rivals, tier);
  }

  // v4.0.0 replaced the single Excel Skill stat with 7 case specialties
  // (only 1-3 "active" and trainable each round), and renamed/inverted
  // Sleep Debt -> Rest and Stress -> Composure (both higher-is-better now,
  // matching every other stat). Any save still on the old single-skill
  // shape gets migrated: every new skill starts at the old Excel Skill
  // level (simplest continuity, no worse than a fresh grind), Rest/Composure
  // are seeded from their old inverse, and cash is refunded for whatever
  // was invested in the old single Personal Coach upgrade (it no longer
  // exists — 5 flat levels: 300/650/1200 cumulative) since that progress
  // can't carry over 1:1 into 7 separate per-skill coaches.
  if (!parsed.stats || parsed.stats.skills === undefined) {
    const oldStats = parsed.stats || { excel: 8, phys: 65, energy: 80, sleepDebt: 0, stress: 8 };
    const oldExcel = oldStats.excel !== undefined ? oldStats.excel : 8;
    // Every skill starts fresh (level 0) in the new per-skill Coaching Shop,
    // so the migrated value can't exceed the shop-base/league-cap gate any
    // veteran's real progress would still have to clear.
    const migratedSkillCap = Math.min(BAL.skillShopCapBase, LEAGUE_SKILL_CAP[parsed.peakLeagueTier] || LEAGUE_SKILL_CAP[5]);
    const seededSkill = clamp(oldExcel, 0, migratedSkillCap);
    parsed.stats = {
      skills: Object.fromEntries(SKILL_KEYS.map((k) => [k, seededSkill])),
      phys: oldStats.phys !== undefined ? oldStats.phys : 65,
      energy: oldStats.energy !== undefined ? oldStats.energy : 80,
      rest: oldStats.sleepDebt !== undefined ? clamp(100 - oldStats.sleepDebt, 0, 100) : 100,
      composure: oldStats.stress !== undefined ? clamp(100 - oldStats.stress, 0, 100) : 92,
    };

    const oldAlloc = parsed.allocation || {};
    parsed.allocation = {
      skills: Object.fromEntries(SKILL_KEYS.map((k) => [k, 0])),
      exercise: oldAlloc.exercise !== undefined ? oldAlloc.exercise : 1,
      sleep: oldAlloc.sleep !== undefined ? oldAlloc.sleep : 8,
      relax: oldAlloc.relax !== undefined ? oldAlloc.relax : 2,
    };

    const OLD_COACH_CUMULATIVE = [0, 300, 950, 2150];
    const oldCoachLvl = clamp(oldUpgrades.coach || 0, 0, 3);
    parsed.cash = (parsed.cash || 0) + OLD_COACH_CUMULATIVE[oldCoachLvl];

    parsed.activeSkills = rollActiveSkills();
    parsed.skillCycleDay = 0;
  }

  return parsed;
}

function loadState() {
  if (!STORAGE_OK) return freshState();
  try {
    const raw = localStorage.getItem(SAVE_KEY);
    if (!raw) return freshState();
    const parsed = migrateSave(JSON.parse(raw));
    lastLoadedFromSave = true;
    return Object.assign(freshState(), parsed);
  } catch (e) {
    return freshState();
  }
}

function saveState() {
  if (!STORAGE_OK) return false;
  try {
    localStorage.setItem(SAVE_KEY, JSON.stringify(state));
    return true;
  } catch (e) {
    return false;
  }
}

/* ---------------------------------------------------------------------- */
/* Helpers                                                                */
/* ---------------------------------------------------------------------- */
function clamp(v, lo, hi) {
  return Math.max(lo, Math.min(hi, v));
}

function effectiveHours(h, cfg = BAL) {
  const soft = cfg.softFatigueCap;
  const hard = cfg.hardFatigueCap;
  const base = Math.min(h, soft);
  const mid = Math.max(0, Math.min(h, hard) - soft) * cfg.fatigueMultSoft;
  const over = Math.max(0, h - hard) * cfg.fatigueMultHard;
  return base + mid + over;
}

// Composure no longer factors into training focus — it's purely a match-day
// effect now (see composureMatchMultiplier). Rest supplies training's own
// separate upside multiplier (see restTrainingMultiplier).
function focusMultiplier(energy) {
  const energyFactor = 0.4 + 0.6 * (energy / 100);
  return clamp(energyFactor, 0.15, 1.0);
}

function physSynergy(phys) {
  return 0.5 + 0.5 * (phys / 100);
}

function skillDiminish(excel) {
  // A steep, accelerating falloff (not linear) so the last stretch to 100
  // takes meaningfully longer than the climb to ~80 — mastery should be a
  // long tail, not a wall you hit in a couple of months.
  return clamp(1 - 0.85 * Math.pow(excel / 100, 1.5), 0.12, 1.0);
}

function physDiminish(phys) {
  return clamp(1 - 0.75 * Math.pow(phys / 100, 1.3), 0.18, 1.0);
}

function randInt(lo, hi) {
  return Math.floor(lo + Math.random() * (hi - lo + 1));
}

function fmt(n) {
  return Math.round(n).toString();
}

function fmtSigned(n, decimals = 1) {
  const r = Number(n.toFixed(decimals));
  return (r >= 0 ? "+" : "") + r;
}

/* ---------------------------------------------------------------------- */
/* Day resolution — stats only                                            */
/* ---------------------------------------------------------------------- */
function resolveDay() {
  const a = state.allocation;
  const s = state.stats;
  const physioEff = upgradeEffect("physio");
  const sleepAppEff = upgradeEffect("sleepApp");
  const nutritionistEff = upgradeEffect("nutritionist");
  const meditationEff = upgradeEffect("meditation");
  const recoveryEff = upgradeEffect("recovery");
  const events = [];

  // ---- Injury / burnout lockouts: enforce before computing effects ----
  let exerciseH = state.injury.active ? 0 : a.exercise;
  const sleepH = a.sleep;
  const relaxH = a.relax;

  const burnoutActive = state.burnout.active;
  const focusMult = focusMultiplier(s.energy) * (burnoutActive ? BAL.burnoutEffectivenessMult : 1);
  const physMult = physSynergy(s.phys);
  const restMult = restTrainingMultiplier(s.rest);

  // ---- The 7 case specialties: outside preseason, only this round's 1-3
  // active skills can be trained; the rest sit locked (0h, forced) and
  // quietly rust. Preseason has no rotation yet — everything is open. ----
  let totalSkillH = 0;
  const trainable = trainableSkills();
  SKILL_KEYS.forEach((key) => {
    const meta = skillMeta(key);
    const isActive = trainable.includes(key);
    const hours = isActive ? a.skills[key] || 0 : 0;
    const cap = skillCap(key);
    const wasAtCap = s.skills[key] >= cap - 0.05;
    const coachEff = upgradeEffect(coachKey(key));
    const coachMult = coachEff ? 1 + coachEff.bonus : 1.0;

    if (hours >= BAL.skillDecayThresholdHours) {
      totalSkillH += hours;
      const eff = effectiveHours(hours);
      const gain = eff * BAL.skillGainBase * focusMult * physMult * restMult * skillDiminish(s.skills[key]) * coachMult;
      s.skills[key] = clamp(s.skills[key] + gain, 0, cap);
      let note = "";
      if (wasAtCap && cap < 100) note = ` (capped at ${cap})`;
      else if (focusMult < 0.5) note = " (low energy hurt your training)";
      else if (physMult < 0.75) note = " (low physical health capped your gains)";
      events.push({ type: gain > 0.5 ? "good" : "neutral", text: `${meta.icon} ${meta.name} (${hours}h): ${fmtSigned(gain)} skill${note}` });
    } else {
      totalSkillH += hours;
      const rust = Math.min(0.15, s.skills[key] * 0.003);
      if (rust > 0) {
        s.skills[key] = clamp(s.skills[key] - rust, 0, cap);
        if (isActive) events.push({ type: "bad", text: `${meta.icon} No ${meta.name} training: skill rusted slightly (${fmtSigned(-rust)})` });
      }
    }
  });

  // ---- Physical health ----
  const pCap = physCap();
  const physWasAtCap = s.phys >= pCap - 0.05;
  const physioMult = physioEff ? 1 + physioEff.exerciseBonus : 1.0;
  const exerciseGain = effectiveHours(exerciseH) * BAL.exerciseGainBase * physioMult * restMult * physDiminish(s.phys);
  const restDecayFactor = nutritionistEff ? nutritionistEff.decayMult : 1.0;
  const physDecayFromRest = (100 - s.rest) * 0.045 * restDecayFactor;
  const detrainMult = recoveryEff ? recoveryEff.detrainMult : 1.0;
  const detrain = exerciseH < BAL.detrainThresholdHours ? BAL.detrainDecay * detrainMult : 0;
  const physDelta = exerciseGain - physDecayFromRest - detrain;
  s.phys = clamp(s.phys + physDelta, 0, pCap);
  if (exerciseH > 0) {
    let note = physWasAtCap && pCap < 100 ? ` (capped at ${pCap} — upgrade Sports Physio for a higher ceiling)` : "";
    events.push({ type: physDelta > 0 ? "good" : "neutral", text: `🏃 Exercise (${exerciseH}h): ${fmtSigned(physDelta)} physical health${note}` });
  }
  if (physDecayFromRest > 1) {
    events.push({ type: "bad", text: `🌙 Poor Rest is wearing down your body (${fmtSigned(-physDecayFromRest)} health)` });
  }

  // ---- Injury roll (only if not already injured) ----
  if (!state.injury.active && exerciseH > BAL.overtrainThreshold) {
    const injuryReduceMult = physioEff ? physioEff.injuryReduceMult : 1.0;
    const excess = exerciseH - BAL.overtrainThreshold;
    const chance = clamp(excess * BAL.injuryChancePerExcessHour * injuryReduceMult, 0, 0.6);
    if (Math.random() < chance) {
      const loss = randInt(BAL.injuryPhysLoss[0], BAL.injuryPhysLoss[1]);
      const daysReduce = recoveryEff ? recoveryEff.injuryDaysReduce : 0;
      const days = Math.max(1, randInt(BAL.injuryDaysRange[0], BAL.injuryDaysRange[1]) - daysReduce);
      s.phys = clamp(s.phys - loss, 0, pCap);
      state.injury = { active: true, daysLeft: days };
      events.push({ type: "bad", text: `🤕 Overtraining injury! -${loss} physical health. Exercise disabled for ${days} days.` });
    }
  }

  // ---- Energy ----
  // How rested you feel today is driven mainly by *last night's* sleep, not
  // a slowly-accumulating bank — otherwise a well-rested surplus can quietly
  // absorb a bad night and Energy never visibly drops. Low Rest degrades how
  // restorative sleep is; today's activity then spends down whatever that
  // sleep gave you.
  const sleepQualityFactor = clamp(0.5 + 0.5 * (s.rest / 100), 0.5, 1);
  const energyFromSleep = clamp((sleepH / BAL.idealSleep) * 100, 0, 115) * sleepQualityFactor;
  const energySpend = totalSkillH * BAL.energyDrainPerHour.skill + exerciseH * BAL.energyDrainPerHour.exercise;
  const relaxEnergyBonus = relaxH * BAL.relaxEnergyRestore;
  s.energy = clamp(energyFromSleep - energySpend + relaxEnergyBonus, 0, 100);

  // ---- Rest (renamed/inverted Sleep Debt: higher = better) ----
  const sleepAppMult = sleepAppEff ? sleepAppEff.sleepDebtMult : 1.0;
  let restDelta;
  if (sleepH < BAL.idealSleep) {
    restDelta = -((BAL.idealSleep - sleepH) * 1.3 * sleepAppMult);
  } else {
    restDelta = Math.min(sleepH - BAL.idealSleep, 3) * 1.4;
  }
  s.rest = clamp(s.rest + restDelta, 0, 100);
  if (sleepH < 6) {
    events.push({ type: "bad", text: `🌙 Only slept ${sleepH}h: Rest dropping fast (${fmtSigned(restDelta)})` });
  } else if (sleepH >= BAL.idealSleep) {
    events.push({ type: "good", text: `🌙 Slept ${sleepH}h: well rested (Rest ${fmtSigned(restDelta)})` });
  }

  // ---- Composure (renamed/inverted Stress: higher = better) ----
  const meditationMult = meditationEff ? meditationEff.reliefMult : 1.0;
  const composureLoad = (totalSkillH + exerciseH) * BAL.composureLoadPerHour;
  const composureRelief = relaxH * BAL.relaxComposureRelief * meditationMult + (sleepH >= BAL.idealSleep ? BAL.restGoodSleepBonus : 0);
  const composurePenaltyFromRest = (100 - s.rest) * 0.1;
  const composureDelta = composureRelief - composureLoad - composurePenaltyFromRest;
  s.composure = clamp(s.composure + composureDelta, 0, 100);

  if (relaxH < BAL.relaxComposureThreshold && composureLoad > 5) {
    events.push({ type: "bad", text: `🔥 Under ${BAL.relaxComposureThreshold}h relaxation: Composure fell to ${fmt(s.composure)}` });
  } else if (relaxH >= BAL.relaxComposureThreshold) {
    events.push({ type: "good", text: `🎮 Relaxed ${relaxH}h: Composure ${fmtSigned(composureDelta)}` });
  }

  // ---- Burnout state transitions ----
  if (!state.burnout.active && s.composure <= BAL.burnoutComposureThreshold) {
    state.burnout = { active: true, daysLeft: 1 };
    events.push({ type: "bad", text: `⚠️ BURNOUT! You've pushed too hard with too little rest. Training and exercise are far less effective until your Composure recovers — relax more.` });
  } else if (state.burnout.active && s.composure >= BAL.burnoutRecoverThreshold) {
    state.burnout = { active: false, daysLeft: 0 };
    events.push({ type: "good", text: `✅ Recovered from burnout. You're focused again.` });
  }

  // ---- Injury countdown ----
  if (state.injury.active) {
    state.injury.daysLeft -= 1;
    if (state.injury.daysLeft <= 0) {
      state.injury = { active: false, daysLeft: 0 };
      events.push({ type: "good", text: `✅ Injury healed. Exercise is available again.` });
    }
  }

  return events;
}

/* ---------------------------------------------------------------------- */
/* Match simulation                                                       */
/* ---------------------------------------------------------------------- */
// Every match tests only this round's active 1-3 skills — a live case
// competition on those specific specialties, not the full roster of 7.
function activeSkillAverage() {
  const active = state.activeSkills;
  if (!active.length) return 0;
  const sum = active.reduce((acc, k) => acc + state.stats.skills[k], 0);
  return sum / active.length;
}

function performanceScore() {
  const s = state.stats;
  // Low Composure hits you hardest where it matters most: match day. This
  // stacks on top of (doesn't replace) Composure's own weighted contribution
  // below, same as before.
  const skillComponent = activeSkillAverage() * composureMatchMultiplier(s.composure);
  return clamp(skillComponent * 0.55 + s.phys * 0.25 + s.composure * 0.15 + s.rest * 0.05, 0, 100);
}

// Resolves one real match for the player against a specific opponent rating.
// Used for both regular-season fixtures and playoff matches.
function resolveMatch(opponentRating) {
  const managerEff = upgradeEffect("manager");
  const rankLossMult = managerEff ? managerEff.rankLossMult : 1.0;
  const cashBonusMult = managerEff ? managerEff.cashBonusMult : 1.0;

  if (state.injury.active) {
    const loss = Math.round(10 * rankLossMult);
    state.rank = Math.max(400, state.rank - loss);
    return { forfeit: true, win: false, text: `You're injured and had to forfeit. Rank -${loss}.`, opponentRating: Math.round(opponentRating) };
  }

  const perf = performanceScore();
  // Perfect performance (100) gives a strong but no longer overwhelming form
  // bonus — enough to matter, not enough to guarantee a win against a tough
  // opponent.
  const matchRating = state.rank + (perf - 70) * 3;
  const winProb = 1 / (1 + Math.pow(10, (opponentRating - matchRating) / 400));
  const win = Math.random() < winProb;
  const K = 24;
  const actual = win ? 1 : 0;
  let ratingChange = Math.round(K * (actual - winProb));
  if (ratingChange < 0) ratingChange = Math.round(ratingChange * rankLossMult);
  state.rank = clamp(state.rank + ratingChange, 400, 5000);
  state.peakRank = Math.max(state.peakRank, state.rank);

  const cashReward = Math.round((win ? 150 + state.rank / 10 : 40) * cashBonusMult);
  state.cash += cashReward;
  if (win) state.wins += 1;
  else state.losses += 1;

  return {
    forfeit: false,
    win,
    perf,
    opponentRating: Math.round(opponentRating),
    winProb: Math.round(winProb * 100),
    ratingChange,
    cashReward,
  };
}

// Cheap win/lose roll for matches that don't involve the player (other
// league members' simulated season records, and NPC-vs-NPC bracket games).
function simulateNpcMatch(ratingA, ratingB) {
  const winProbA = 1 / (1 + Math.pow(10, (ratingB - ratingA) / 400));
  return Math.random() < winProbA;
}

function eloChange(ratingA, ratingB, aWon, K = 24) {
  const winProbA = 1 / (1 + Math.pow(10, (ratingB - ratingA) / 400));
  return Math.round(K * ((aWon ? 1 : 0) - winProbA));
}

// Resolves one NPC-vs-NPC match and updates both rivals' persistent ratings
// and lifetime records — this is what makes the other 199 rivals a real,
// evolving world rather than static names.
function resolveNpcMatch(rivalA, rivalB) {
  const aWon = simulateNpcMatch(rivalA.rating, rivalB.rating);
  const change = eloChange(rivalA.rating, rivalB.rating, aWon);
  rivalA.rating = clamp(rivalA.rating + change, 300, 3000);
  rivalB.rating = clamp(rivalB.rating - change, 300, 3000);
  if (aWon) {
    rivalA.wins += 1;
    rivalB.losses += 1;
  } else {
    rivalB.wins += 1;
    rivalA.losses += 1;
  }
  return aWon;
}

// Simulates a full round-robin (every pair plays once) for a league tier
// that doesn't contain the player, returning that tier's final standings.
function simulateLeagueRoundRobin(tier) {
  const members = state.rivals.filter((r) => r.league === tier);
  const points = new Map(members.map((r) => [r.id, 0]));
  for (let i = 0; i < members.length; i++) {
    for (let j = i + 1; j < members.length; j++) {
      const aWon = resolveNpcMatch(members[i], members[j]);
      const winner = aWon ? members[i] : members[j];
      points.set(winner.id, points.get(winner.id) + 3);
    }
  }
  return members
    .map((r) => ({ name: r.name, rating: r.rating, points: points.get(r.id), isPlayer: false, rivalId: r.id }))
    .sort((a, b) => b.points - a.points || b.rating - a.rating);
}

// Same idea for the player's own league: the player's 39 real match results
// are already known (state.seasonResults), so only the 39x38/2 NPC-vs-NPC
// pairs among their rivals need simulating; the player's record merges in
// directly rather than being re-simulated.
function simulatePlayerLeagueStandings() {
  const members = state.rivals.filter((r) => r.league === state.leagueTier);
  const points = new Map(members.map((r) => [r.id, 0]));
  for (let i = 0; i < members.length; i++) {
    for (let j = i + 1; j < members.length; j++) {
      const aWon = resolveNpcMatch(members[i], members[j]);
      const winner = aWon ? members[i] : members[j];
      points.set(winner.id, points.get(winner.id) + 3);
    }
  }
  const npcStandings = members.map((r) => ({ name: r.name, rating: r.rating, points: points.get(r.id), isPlayer: false, rivalId: r.id }));
  const playerWins = state.seasonResults.filter((r) => r.win).length;
  const playerEntry = { name: state.playerName || "You", rating: state.rank, points: playerWins * 3, isPlayer: true, rivalId: null };
  return [...npcStandings, playerEntry].sort((a, b) => b.points - a.points || b.rating - a.rating);
}

// Promotes the top BAL.promotionCount and relegates the bottom
// BAL.relegationCount of every league simultaneously (coordinated across
// all 5 tiers so each stays at exactly 40 members), based on the standings
// already computed for every tier this season. Returns the player's own
// resulting tier.
function applyPromotionRelegation(allStandings) {
  const promoted = {};
  const relegated = {};
  const stayed = {};

  for (let tier = 1; tier <= BAL.leagueCount; tier++) {
    const standings = allStandings[tier];
    const promoCount = tier > 1 ? BAL.promotionCount : 0;
    const relCount = tier < BAL.leagueCount ? BAL.relegationCount : 0;
    promoted[tier] = promoCount > 0 ? standings.slice(0, promoCount) : [];
    relegated[tier] = relCount > 0 ? standings.slice(standings.length - relCount) : [];
    const movedIds = new Set([...promoted[tier], ...relegated[tier]].map((e) => e.rivalId));
    stayed[tier] = standings.filter((e) => !movedIds.has(e.rivalId));
  }

  let playerNewTier = state.leagueTier;
  if (promoted[state.leagueTier].some((e) => e.isPlayer)) playerNewTier = state.leagueTier - 1;
  else if (relegated[state.leagueTier].some((e) => e.isPlayer)) playerNewTier = state.leagueTier + 1;

  for (let tier = 1; tier <= BAL.leagueCount; tier++) {
    stayed[tier].forEach((e) => {
      if (e.rivalId != null) findRival(e.rivalId).league = tier;
    });
    if (tier > 1) {
      promoted[tier].forEach((e) => {
        if (e.rivalId != null) {
          const r = findRival(e.rivalId);
          r.league = tier - 1;
          r.promotions += 1;
        }
      });
    }
    if (tier < BAL.leagueCount) {
      relegated[tier].forEach((e) => {
        if (e.rivalId != null) {
          const r = findRival(e.rivalId);
          r.league = tier + 1;
          r.relegations += 1;
        }
      });
    }
  }

  return playerNewTier;
}

// Called once the player's 39th regular-season match resolves. Simulates a
// full round-robin for every league (not just the player's), so the whole
// 200-competitor world — all 5 tables — advances together every season.
function processLeagueSeasonEnd() {
  const allStandings = {};
  for (let tier = 1; tier <= BAL.leagueCount; tier++) {
    allStandings[tier] = tier === state.leagueTier ? simulatePlayerLeagueStandings() : simulateLeagueRoundRobin(tier);
  }

  const playerStandings = allStandings[state.leagueTier];
  const playerPosition = playerStandings.findIndex((s) => s.isPlayer) + 1;
  const qualified = playerPosition <= BAL.playoffSize;

  for (let tier = 1; tier <= BAL.leagueCount; tier++) {
    state.leagueStandings[tier] = allStandings[tier];
  }

  const playerNewTier = applyPromotionRelegation(allStandings);

  return { standings: playerStandings, playerPosition, qualified, playerNewTier };
}

// Standard 16-bracket seeding pairs so top seeds are spread across the draw
// (1 and 2 can only meet in the Final, etc).
const SEED_PAIRS_16 = [[1, 16], [8, 9], [5, 12], [4, 13], [6, 11], [3, 14], [7, 10], [2, 15]];

function buildPlayoffBracket(standings, tier) {
  const top = standings.slice(0, BAL.playoffSize);
  const bySeed = {};
  top.forEach((team, idx) => {
    bySeed[idx + 1] = team;
  });
  const currentRound = [];
  SEED_PAIRS_16.forEach(([a, b]) => {
    currentRound.push(bySeed[a], bySeed[b]);
  });
  const playerSeed = top.findIndex((t) => t.isPlayer) + 1;
  return { stage: "r16", currentRound, eliminated: false, champion: false, playerSeed, tier };
}

const PLAYOFF_ROUND_NAMES = { r16: "Round of 16", qf: "Quarterfinal", sf: "Semifinal", f: "Final" };
const PLAYOFF_NEXT_STAGE = { r16: "qf", qf: "sf", sf: "f" };

function resolvePlayoffRound() {
  const p = state.playoff;
  const roundName = PLAYOFF_ROUND_NAMES[p.stage];
  const currentTeams = p.currentRound;
  const winners = [];
  let matchResult = null;
  let summary = "";

  for (let i = 0; i < currentTeams.length; i += 2) {
    const teamA = currentTeams[i];
    const teamB = currentTeams[i + 1];
    if (teamA.isPlayer || teamB.isPlayer) {
      const player = teamA.isPlayer ? teamA : teamB;
      const opp = teamA.isPlayer ? teamB : teamA;
      matchResult = resolveMatch(opp.rating);
      matchResult.opponentName = opp.name;
      matchResult.roundLabel = roundName;
      const win = !matchResult.forfeit && matchResult.win;
      winners.push(win ? player : opp);
      summary = win ? `🏆 ${roundName} WIN vs ${opp.name}!` : `💔 Eliminated in the ${roundName} by ${opp.name}.`;
      if (!win) p.eliminated = true;
    } else {
      // Uses the persistent-rating-updating resolver so playoff results
      // also feed back into the rivals' ongoing world, same as the regular
      // season's round-robin does.
      const rivalA = findRival(teamA.rivalId);
      const rivalB = findRival(teamB.rivalId);
      const aWins = rivalA && rivalB ? resolveNpcMatch(rivalA, rivalB) : simulateNpcMatch(teamA.rating, teamB.rating);
      winners.push(aWins ? teamA : teamB);
    }
  }

  let seasonOver = false;
  if (p.eliminated) {
    seasonOver = true;
  } else if (p.stage === "f") {
    p.champion = true;
    seasonOver = true;
    summary = `👑 CHAMPION! You won the Year ${state.year} League ${p.tier} Final!`;
    state.cash += 2000;
    state.rank += 40;
    state.peakRank = Math.max(state.peakRank, state.rank);
  } else {
    p.stage = PLAYOFF_NEXT_STAGE[p.stage];
    p.currentRound = winners;
  }

  return { matchResult, summary, seasonOver };
}

/* ---------------------------------------------------------------------- */
/* Season/day phase engine                                                */
/* ---------------------------------------------------------------------- */
// Advances the season state machine by one day. Call after resolveDay().
// Returns { matchResult, phaseEvent } — either may be null.
function processDayEnd() {
  state.phaseDay += 1;
  let matchResult = null;
  let phaseEvent = null;

  if (state.seasonPhase === "preseason") {
    if (state.phaseDay >= BAL.preseasonDays) {
      state.seasonPhase = "regular";
      state.phaseDay = 0;
      state.roundIndex = 0;
      phaseEvent = `🏁 Preseason over — the ${BAL.seasonRounds}-round regular season begins.`;
    }
    return { matchResult, phaseEvent };
  }

  if (state.seasonPhase === "regular") {
    if (state.phaseDay >= BAL.roundIntervalDays) {
      state.phaseDay = 0;
      const fixture = state.schedule[state.roundIndex];
      matchResult = resolveMatch(fixture.rating);
      matchResult.opponentName = fixture.name;
      matchResult.roundLabel = `Round ${state.roundIndex + 1}/${BAL.seasonRounds}`;
      const win = !matchResult.forfeit && matchResult.win;
      fixture.played = true;
      fixture.win = win;
      state.seasonResults.push({ opponent: fixture.name, win });
      state.roundIndex += 1;

      if (state.roundIndex >= BAL.seasonRounds) {
        const playedTier = state.leagueTier;
        const result = processLeagueSeasonEnd();
        state.lastStandings = result.standings;
        state.lastPlayerPosition = result.playerPosition;
        state.lastStandingsTier = playedTier;
        state.leagueTier = result.playerNewTier;
        state.peakLeagueTier = Math.min(state.peakLeagueTier, state.leagueTier);
        const moveText =
          state.leagueTier < playedTier
            ? ` Promoted to League ${state.leagueTier}!`
            : state.leagueTier > playedTier
            ? ` Relegated to League ${state.leagueTier}.`
            : "";
        if (result.qualified) {
          state.seasonPhase = "playoffs";
          state.playoff = buildPlayoffBracket(result.standings, playedTier);
          phaseEvent = `🏆 Regular season complete! Finished #${result.playerPosition} of ${result.standings.length} in League ${playedTier} — through to the playoffs as seed ${state.playoff.playerSeed}.${moveText}`;
        } else {
          state.seasonPhase = "offseason";
          state.offseasonDays = BAL.trainingCampDays;
          state.offseasonReason = "missed";
          phaseEvent = `📋 Regular season complete. Finished #${result.playerPosition} of ${result.standings.length} in League ${playedTier} — missed the top ${BAL.playoffSize} playoff cutoff.${moveText} ${BAL.trainingCampDays}-day training camp starts now to get ready for next season.`;
        }
      }
    }
    return { matchResult, phaseEvent };
  }

  if (state.seasonPhase === "playoffs") {
    if (state.phaseDay >= BAL.roundIntervalDays) {
      state.phaseDay = 0;
      const result = resolvePlayoffRound();
      matchResult = result.matchResult;
      phaseEvent = result.summary;
      if (result.seasonOver) {
        state.seasonPhase = "offseason";
        state.offseasonDays = BAL.offseasonDays;
        state.offseasonReason = "playoffs";
      }
    }
    return { matchResult, phaseEvent };
  }

  if (state.seasonPhase === "offseason") {
    if (state.phaseDay >= (state.offseasonDays || BAL.offseasonDays)) {
      state.phaseDay = 0;
      state.year += 1;
      state.seasonPhase = "preseason";
      state.roundIndex = 0;
      state.seasonResults = [];
      state.schedule = buildLeagueSchedule(state.rivals, state.leagueTier);
      state.playoff = null;
      // Offseason recovery — a clean slate for the new year.
      state.stats.composure = 100;
      state.stats.rest = 100;
      state.stats.energy = 100;
      phaseEvent = `🎉 Year ${state.year} begins! A fresh ${BAL.seasonRounds}-round season has been scheduled — good luck.`;
    }
    return { matchResult, phaseEvent };
  }

  return { matchResult, phaseEvent };
}

// Qualitative read on a specific known opponent vs the player's current
// rank — concrete now that the schedule/bracket tells us exactly who's
// next, rather than a vague matchmaking estimate.
function difficultyLabel(opponentRating, rank) {
  const diff = opponentRating - rank;
  if (diff <= -80) return "Favorable";
  if (diff < 60) return "Even";
  if (diff < 180) return "Tough";
  return "Elite";
}

function getNextMatchInfo() {
  const s = state;
  if (s.seasonPhase === "preseason") {
    const daysUntil = BAL.preseasonDays - s.phaseDay;
    return { kind: "preseason", label: `Season starts in ${daysUntil}d` };
  }
  if (s.seasonPhase === "regular") {
    const daysUntil = BAL.roundIntervalDays - s.phaseDay;
    const fixture = s.schedule[s.roundIndex];
    if (!fixture) return { kind: "regular-end", label: "Season wrapping up…" };
    const roundNum = s.roundIndex + 1;
    const diff = difficultyLabel(fixture.rating, s.rank);
    return {
      kind: "fixture",
      daysUntil,
      opponentName: fixture.name,
      opponentRating: Math.round(fixture.rating),
      difficulty: diff,
      label: daysUntil <= 0 ? `Round ${roundNum} today vs ${fixture.name}! (${diff})` : `Round ${roundNum}/${BAL.seasonRounds} in ${daysUntil}d vs ${fixture.name} (${diff})`,
    };
  }
  if (s.seasonPhase === "playoffs") {
    const daysUntil = BAL.roundIntervalDays - s.phaseDay;
    const p = s.playoff;
    const idx = p.currentRound.findIndex((t) => t.isPlayer);
    const opp = idx >= 0 ? p.currentRound[idx % 2 === 0 ? idx + 1 : idx - 1] : null;
    const roundName = PLAYOFF_ROUND_NAMES[p.stage];
    const diff = opp ? difficultyLabel(opp.rating, s.rank) : null;
    return {
      kind: "playoff",
      daysUntil,
      opponentName: opp ? opp.name : "?",
      opponentRating: opp ? Math.round(opp.rating) : null,
      difficulty: diff,
      label: daysUntil <= 0 ? `${roundName} today vs ${opp ? opp.name : "?"}! (${diff})` : `${roundName} in ${daysUntil}d vs ${opp ? opp.name : "?"} (${diff})`,
    };
  }
  if (s.seasonPhase === "offseason") {
    const daysUntil = (s.offseasonDays || BAL.offseasonDays) - s.phaseDay;
    let outcome = "Season over";
    if (s.offseasonReason === "missed") outcome = "Training camp";
    else if (s.playoff && s.playoff.champion) outcome = "🏆 Champion!";
    else if (s.playoff && s.playoff.eliminated) outcome = "Eliminated";
    return { kind: "offseason", daysUntil, label: `${outcome} — Year ${s.year + 1} in ${daysUntil}d` };
  }
  return { kind: "unknown", label: "" };
}

// Every 7 days, reroll which 1-3 skills are trainable for the coming week —
// revealed only now, at the end of the previous cycle, never in advance.
// Hours pending on the old active skills are cleared since they're about to
// become untrainable; the player reassigns under the new focus.
function advanceSkillCycle() {
  // Preseason trains everything (see trainableSkills()), so the weekly
  // reveal is frozen until the regular season actually begins.
  if (isPreseason()) return null;
  state.skillCycleDay += 1;
  if (state.skillCycleDay < 7) return null;
  state.skillCycleDay = 0;
  state.activeSkills = rollActiveSkills();
  SKILL_KEYS.forEach((k) => {
    state.allocation.skills[k] = 0;
  });
  const names = state.activeSkills.map((k) => `${skillMeta(k).icon} ${skillMeta(k).name}`).join(", ");
  return `📋 New focus for the week ahead: ${names}.`;
}

function phaseLabelText() {
  const s = state;
  const league = `League ${s.leagueTier} · `;
  if (s.seasonPhase === "preseason") return league + "Preseason";
  if (s.seasonPhase === "regular") return league + "Regular Season";
  if (s.seasonPhase === "playoffs") return league + "Playoffs";
  if (s.seasonPhase === "offseason") return league + (s.offseasonReason === "missed" ? "Training Camp" : "Offseason");
  return "";
}

/* ---------------------------------------------------------------------- */
/* UI rendering                                                           */
/* ---------------------------------------------------------------------- */
const $ = (id) => document.getElementById(id);

const ACT_KEYS = ["exercise", "sleep", "relax"];
const ACT_INPUT_IDS = { exercise: "exerciseHours", sleep: "sleepHours", relax: "relaxHours" };
const ACT_MAX = { exercise: 12, sleep: 12, relax: 12 };
const ACT_DECAY_MARKER = { exercise: BAL.skillDecayThresholdHours, sleep: BAL.idealSleep, relax: BAL.relaxComposureThreshold };
const SKILL_MAX_HOURS = 16;

function markerPct(threshold, max) {
  return clamp((threshold / max) * 100, 0, 100);
}

function renderTopbar() {
  $("yearNum").textContent = state.year;
  $("cashVal").textContent = fmt(state.cash);
  $("rankVal").textContent = fmt(state.rank);
  $("phaseLabel").textContent = phaseLabelText();
  $("matchCounter").textContent = getNextMatchInfo().label;
}

function renderStats() {
  const s = state.stats;
  const trainable = trainableSkills();
  const skillRows = trainable
    .map((key) => {
      const meta = skillMeta(key);
      const cap = skillCap(key);
      return `
      <div class="stat" data-stat="skill-${key}">
        <div class="stat-label"><span class="stat-icon">${meta.icon}</span>${meta.name}<span class="stat-val" id="skill_${key}Val">0</span><span class="stat-cap">/${cap}</span></div>
        <div class="bar"><div class="bar-fill skill" id="skill_${key}Bar" style="width:0%"></div></div>
      </div>`;
    })
    .join("");
  $("skillStatRows").innerHTML = skillRows;
  trainable.forEach((key) => setBar(`skill_${key}`, s.skills[key], skillCap(key)));

  setBar("phys", s.phys, 100);
  setBar("energy", s.energy, 100);
  setBar("rest", s.rest, 100);
  setBar("composure", s.composure, 100);

  const banner = $("warningBanner");
  const msgs = [];
  if (state.burnout.active) msgs.push("🔥 Burnout active — training & exercise are far less effective. Relax to recover.");
  else if (s.composure <= 10) msgs.push("🔥 Composure critical — burnout imminent. Schedule relaxation soon.");
  if (state.injury.active) msgs.push(`🤕 Injured — Exercise disabled for ${state.injury.daysLeft} more day(s).`);
  if (s.rest <= 30) msgs.push("🌙 Severely low Rest — your body is breaking down. Sleep more.");
  if (s.phys <= 20) msgs.push("💪 Physical health critically low — it's capping your training performance.");

  if (msgs.length) {
    banner.innerHTML = msgs.join("<br>");
    banner.classList.remove("hidden");
  } else {
    banner.classList.add("hidden");
  }
}

function setBar(key, val, max) {
  $(`${key}Val`).textContent = fmt(val);
  $(`${key}Bar`).style.width = `${clamp((val / max) * 100, 0, 100)}%`;
}

function totalAssigned() {
  const a = state.allocation;
  const skillSum = trainableSkills().reduce((sum, k) => sum + (a.skills[k] || 0), 0);
  return skillSum + a.exercise + a.sleep + a.relax;
}

function renderSkillActivityRows() {
  const a = state.allocation;
  const rows = trainableSkills()
    .map((key) => {
      const meta = skillMeta(key);
      const lvl = upgradeLevel(coachKey(key));
      const shopTag = lvl > 0 ? `<span class="skill-shop-tag">${meta.icon}Lv${lvl}</span>` : `<span class="skill-shop-tag skill-shop-tag-none">no coach</span>`;
      const hours = a.skills[key] || 0;
      const pct = markerPct(BAL.skillDecayThresholdHours, SKILL_MAX_HOURS);
      return `
      <div class="activity skill-activity" data-skill="${key}">
        <div class="activity-row">
          <span class="activity-icon">${meta.icon}</span>
          <span class="activity-name">${meta.name}</span>
          ${shopTag}
          <span class="activity-hours"><span id="skillHoursVal_${key}">${hours}</span>h</span>
        </div>
        <div class="stepper">
          <button class="step-btn" data-skill="${key}" data-dir="-1">−</button>
          <div class="slider-wrap">
            <input type="range" min="0" max="${SKILL_MAX_HOURS}" step="1" value="${hours}" id="skillHours_${key}" data-skill="${key}" />
            <div class="slider-marker" style="left:${pct}%"></div>
          </div>
          <button class="step-btn" data-skill="${key}" data-dir="1">+</button>
        </div>
      </div>`;
    })
    .join("");
  $("skillActivityRows").innerHTML = rows;
}

function renderPlanner() {
  const a = state.allocation;
  renderSkillActivityRows();

  ACT_KEYS.forEach((k) => {
    $(ACT_INPUT_IDS[k]).value = a[k];
    $(`${k}HoursVal`).textContent = a[k];
  });

  const exercise = $("exerciseHours");
  const exerciseRow = document.querySelector('.activity[data-act="exercise"]');
  exercise.disabled = state.injury.active;
  exerciseRow.style.opacity = state.injury.active ? 0.45 : 1;

  const left = 24 - totalAssigned();
  const hoursLeftEl = $("hoursLeft");
  hoursLeftEl.textContent = left;
  hoursLeftEl.classList.toggle("over", left < 0);
}

function appendLog(html, cls) {
  const entry = document.createElement("div");
  entry.className = `log-entry ${cls || ""}`.trim();
  entry.innerHTML = html;
  $("log").appendChild(entry);
}

function renderFullLog() {
  $("log").innerHTML = "";
  state.logEntries.slice(-200).forEach((e) => appendLog(e.html, e.cls));
}

let saveToastTimer = null;
function showSaveToast(saved) {
  const toast = $("saveToast");
  if (!toast) return;
  toast.textContent = saved ? "✓ Progress saved to this device" : "⚠️ Could not save — progress may be lost";
  toast.classList.toggle("toast-warn", !saved);
  toast.classList.remove("hidden");
  toast.classList.add("show");
  clearTimeout(saveToastTimer);
  saveToastTimer = setTimeout(() => {
    toast.classList.remove("show");
  }, saved ? 1800 : 4000);
}

/* ---------------------------------------------------------------------- */
/* Modal helpers                                                          */
/* ---------------------------------------------------------------------- */
function openModal(html) {
  $("modalBody").innerHTML = html;
  $("modalOverlay").classList.remove("hidden");
}
function closeModal() {
  $("modalOverlay").classList.add("hidden");
}
$("modalClose").addEventListener("click", closeModal);
$("modalOverlay").addEventListener("click", (e) => {
  if (e.target.id === "modalOverlay") closeModal();
});

/* ---------------------------------------------------------------------- */
/* Match modal                                                            */
/* ---------------------------------------------------------------------- */
function showMatchModal(result) {
  let html;
  const context = result.roundLabel ? `${result.roundLabel}${result.opponentName ? " vs " + result.opponentName : ""}` : "";
  if (result.forfeit) {
    html = `
      <div class="match-card">
        <div class="match-result loss">FORFEIT</div>
        <div class="match-sub">${context ? context + " — " : ""}${result.text}</div>
        <button class="primary-btn" id="matchOk">Continue</button>
      </div>`;
  } else {
    html = `
      <div class="match-card">
        <div class="match-result ${result.win ? "win" : "loss"}">${result.win ? "VICTORY" : "DEFEAT"}</div>
        <div class="match-sub">${context ? context + "<br>" : ""}Opponent rating: ${result.opponentRating} · You had a ${result.winProb}% win chance</div>
        <div class="match-stats">
          <div><b>${fmt(result.perf)}</b>Performance</div>
          <div><b>${fmtSigned(result.ratingChange, 0)}</b>Rank</div>
          <div><b>$${result.cashReward}</b>Prize</div>
        </div>
        <div class="match-sub">New rank: ${fmt(state.rank)}</div>
        <button class="primary-btn" id="matchOk">Continue</button>
      </div>`;
  }
  openModal(html);
  $("matchOk").addEventListener("click", closeModal);
}

/* ---------------------------------------------------------------------- */
/* Shop / Menu                                                            */
/* ---------------------------------------------------------------------- */
function shopItemHtml(key, u) {
  const lvl = upgradeLevel(key);
  const maxLvl = u.levels.length;
  const isMax = lvl >= maxLvl;
  const next = isMax ? null : u.levels[lvl];
  const current = lvl > 0 ? u.levels[lvl - 1] : null;
  const canAfford = next && state.cash >= next.cost;
  const desc = isMax ? `${current.desc} — MAX` : next.desc + (current ? ` <span class="shop-item-current">(now: ${current.desc})</span>` : "");
  return `
  <div class="shop-item">
    <div class="shop-item-icon">${u.icon}</div>
    <div class="shop-item-info">
      <div class="shop-item-name">${u.name}${lvl > 0 ? ` <span class="shop-item-level">Lv.${lvl}</span>` : ""}</div>
      <div class="shop-item-desc">${desc}</div>
    </div>
    <button class="shop-item-btn ${isMax ? "owned" : ""}" data-upgrade="${key}" ${isMax || !canAfford ? "disabled" : ""}>
      ${isMax ? "MAX" : "$" + next.cost}
    </button>
  </div>`;
}

function shopHtml() {
  const skillItems = SKILLS.map((sk) => shopItemHtml(coachKey(sk.key), UPGRADES[coachKey(sk.key)])).join("");
  const supportKeys = ["physio", "sleepApp", "nutritionist", "meditation", "recovery", "manager"];
  const supportItems = supportKeys.map((key) => shopItemHtml(key, UPGRADES[key])).join("");
  return `
    <h2>Coaching Shop</h2>
    <div class="modal-section">
      <h3>Cash: $${fmt(state.cash)}</h3>
    </div>
    <div class="modal-section">
      <h3>Skill Coaches</h3>
      <p class="modal-sub">Each skill's ceiling is also capped by your highest league reached — a maxed-out coach alone won't get you past that.</p>
      ${skillItems}
    </div>
    <div class="modal-section">
      <h3>Support Team</h3>
      ${supportItems}
    </div>`;
}

function openShop() {
  openModal(shopHtml());
  document.querySelectorAll("[data-upgrade]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-upgrade");
      const u = UPGRADES[key];
      const lvl = upgradeLevel(key);
      if (lvl >= u.levels.length) return;
      const next = u.levels[lvl];
      if (state.cash < next.cost) return;
      state.cash -= next.cost;
      state.upgrades[key] = lvl + 1;
      saveState();
      openShop();
      renderTopbar();
    });
  });
}

function openMenu() {
  const html = `
    <h2>Menu</h2>
    <div class="menu-row" id="menuShop"><span>🛒 Coaching Shop</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuCareer"><span>📈 Career &amp; Season</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuLeagues"><span>🏅 Leagues</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuRename"><span>✏️ Rename Player</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuHow"><span>❓ How to Play</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuInstall"><span>📲 Add to Home Screen</span><span class="arrow">›</span></div>
    <button class="ghost-btn" id="menuReset">Reset Career</button>
    <div class="version-tag">Cell Grind v${APP_VERSION}</div>
  `;
  openModal(html);
  $("menuShop").addEventListener("click", openShop);
  $("menuCareer").addEventListener("click", openCareer);
  $("menuLeagues").addEventListener("click", () => openLeagues(state.leagueTier));
  $("menuRename").addEventListener("click", () => openNameModal(false));
  $("menuHow").addEventListener("click", openHowTo);
  $("menuInstall").addEventListener("click", openInstall);
  $("menuReset").addEventListener("click", () => {
    if (confirm("Start a new career? This wipes all progress.")) {
      state = freshState();
      saveState();
      renderAll();
      closeModal();
    }
  });
}

function openNameModal(isFirstTime) {
  const html = `
    <h2>${isFirstTime ? "Name Your Player" : "Rename Player"}</h2>
    <div class="modal-section">
      <p>${isFirstTime ? "What should we call your Excel esports pro?" : "Enter a new name:"}</p>
      <input type="text" id="playerNameInput" maxlength="20" placeholder="Your name" value="${state.playerName ? state.playerName.replace(/"/g, "&quot;") : ""}" class="name-input" />
      <button class="primary-btn" id="nameSaveBtn">Save</button>
    </div>`;
  openModal(html);
  const input = $("playerNameInput");
  input.focus();
  const save = () => {
    const val = input.value.trim().slice(0, 20);
    if (!val) return;
    state.playerName = val;
    saveState();
    closeModal();
  };
  $("nameSaveBtn").addEventListener("click", save);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") save();
  });
}

function openCareer() {
  const info = getNextMatchInfo();
  let nextSection;
  if (info.kind === "fixture" || info.kind === "playoff") {
    nextSection = `<p>${info.label}<br>Opponent rating: ${info.opponentRating}<br>Your performance score right now: ${fmt(performanceScore())} / 100</p>`;
  } else {
    nextSection = `<p>${info.label}</p>`;
  }

  const seasonWins = state.seasonResults.filter((r) => r.win).length;
  const seasonPlayed = state.seasonResults.length;

  let standingsSection = "";
  if (state.lastStandings) {
    const top5 = state.lastStandings.slice(0, 5);
    standingsSection = `
    <div class="modal-section">
      <h3>Last Season — League ${state.lastStandingsTier} Final Standings</h3>
      <p>You finished <b>#${state.lastPlayerPosition}</b> of ${state.lastStandings.length}.</p>
      <ol class="standings-list">
        ${top5.map((t) => `<li class="${t.isPlayer ? "standings-you" : ""}">${t.name}${t.isPlayer ? " (You)" : ""} — ${t.points} pts</li>`).join("")}
      </ol>
    </div>`;
  }

  let playoffSection = "";
  if (state.playoff) {
    const statusText = state.playoff.champion ? "🏆 Champion" : state.playoff.eliminated ? "Eliminated" : PLAYOFF_ROUND_NAMES[state.playoff.stage];
    playoffSection = `
    <div class="modal-section">
      <h3>Playoffs</h3>
      <p>League ${state.playoff.tier} · Seed #${state.playoff.playerSeed} of ${BAL.playoffSize} · ${statusText}</p>
    </div>`;
  }

  const html = `
    <h2>Career &amp; Season</h2>
    <div class="modal-section">
      <h3>Next Up</h3>
      ${nextSection}
    </div>
    <div class="modal-section">
      <h3>This Season</h3>
      <p>Year ${state.year} · League ${state.leagueTier} · Round ${seasonPlayed}/${BAL.seasonRounds} · Record ${seasonWins}-${seasonPlayed - seasonWins}<br>
      Highest league reached: League ${state.peakLeagueTier}</p>
    </div>
    ${playoffSection}
    ${standingsSection}
    <div class="modal-section">
      <h3>Career Record</h3>
      <p>${state.playerName || "Player"} · Rank ${fmt(state.rank)} (peak ${fmt(state.peakRank)})<br>
      All-time: ${state.wins}W – ${state.losses}L<br>
      Cash: $${fmt(state.cash)}</p>
    </div>
    <div class="modal-section">
      <h3>Skills — ${isPreseason() ? "preseason: train anything" : `this week's focus: ${state.activeSkills.map((k) => skillMeta(k).name).join(", ")}`}</h3>
      <p>
      ${SKILLS.map((sk) => {
        const cap = skillCap(sk.key);
        const active = trainableSkills().includes(sk.key) ? " 🟢" : "";
        return `${sk.icon} ${sk.name}: ${fmt(state.stats.skills[sk.key])} / ${cap}${active}`;
      }).join("<br>")}
      </p>
    </div>
    <div class="modal-section">
      <h3>Current Stats</h3>
      <p>Physical Health: ${fmt(state.stats.phys)} / ${physCap()}${physCap() < 100 ? " (upgrade Physio for more)" : ""}<br>
      Energy: ${fmt(state.stats.energy)} / 100<br>
      Rest: ${fmt(state.stats.rest)} / 100<br>
      Composure: ${fmt(state.stats.composure)} / 100</p>
    </div>`;
  openModal(html);
}

function leagueTableHtml(tier) {
  const standings = state.leagueStandings[tier];
  if (!standings) {
    return `<p class="modal-sub">Not yet available — this table fills in once a season completes while you're in (or have been in) League ${tier}.</p>`;
  }
  const rows = standings
    .map((t, i) => {
      const pos = i + 1;
      const zone = pos <= BAL.promotionCount ? "zone-promo" : pos > standings.length - BAL.relegationCount ? "zone-releg" : "";
      return `
      <div class="league-row ${zone} ${t.isPlayer ? "league-row-you" : ""}">
        <span class="league-pos">${pos}</span>
        <span class="league-name">${t.name}${t.isPlayer ? " (You)" : ""}</span>
        <span class="league-rating">${fmt(t.rating)}</span>
        <span class="league-points">${t.points}</span>
      </div>`;
    })
    .join("");
  return `
    <div class="league-table-header">
      <span class="league-pos">#</span>
      <span class="league-name">Name</span>
      <span class="league-rating">Rating</span>
      <span class="league-points">Pts</span>
    </div>
    <div class="league-table">${rows}</div>
    <p class="modal-sub league-legend"><span class="legend-dot legend-promo"></span> Promotion zone · <span class="legend-dot legend-releg"></span> Relegation zone</p>`;
}

function openLeagues(startTier) {
  const html = `
    <h2>Leagues</h2>
    <div class="league-tabs" id="leagueTabs">
      ${[1, 2, 3, 4, 5]
        .map(
          (t) =>
            `<button class="league-tab ${t === startTier ? "active" : ""}" data-tier="${t}">L${t}${t === state.leagueTier ? " (You)" : ""}</button>`
        )
        .join("")}
    </div>
    <div id="leagueTableContainer">${leagueTableHtml(startTier)}</div>`;
  openModal(html);
  document.querySelectorAll(".league-tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".league-tab").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      const tier = Number(btn.getAttribute("data-tier"));
      $("leagueTableContainer").innerHTML = leagueTableHtml(tier);
    });
  });
}

function openHowTo() {
  const html = `
    <h2>How to Play</h2>
    <div class="modal-section">
      <p>You manage a rising Excel esports competitor. Every day has 24 hours — split them across:</p>
      <p>
      📈🗺️📝🎲🔢⏱️🃏 <b>Skill Training</b> — 7 case specialties (Data, Mapping, Text, Game Logic, Math, Time, Cards). During preseason, all 7 are open for training. Once the regular season starts, only 1-3 are "active" each round, revealed at the start of that round's week — the rest can't be trained until they come up again.<br>
      🏃 <b>Exercise</b> — raises Physical Health.<br>
      🌙 <b>Sleep</b> — restores Energy and builds Rest.<br>
      🎮 <b>Relaxation</b> — builds Composure and prevents burnout.
      </p>
      <p><b>It's all connected:</b> low Rest wears down Physical Health even if you train well, and low Physical Health caps how much your skill training actually helps. Training hard without Relaxation drains Composure — hit 0 and you burn out, tanking your effectiveness until it recovers.</p>
      <p><b>Decay:</b> every stat needs upkeep or it slips. Any skill that isn't active this round rusts; an active skill still rusts below ${BAL.skillDecayThresholdHours}h of training. Exercise below ${BAL.skillDecayThresholdHours}h detrains Physical Health. Sleep below ${BAL.idealSleep}h drains Rest. Relaxation below ${BAL.relaxComposureThreshold}h drains Composure. Each slider shows a marker at its threshold.</p>
      <p><b>Rest</b> also swings training itself: above ${BAL.restTrainingBoostThreshold} it's 150% effective, above ${BAL.restTrainingBoostHigh} it's 200% effective. <b>Composure</b> hits match day specifically — below ${BAL.composureMatchMid} your active skills count for only 75%, below ${BAL.composureMatchLow} just 50%.</p>
      <p>Overtraining physically (too much Exercise) risks injury, which locks out Exercise for several days.</p>
      <p><b>Stat ceilings:</b> each skill caps at ${BAL.skillShopCapBase} until you invest in that skill's dedicated Coach (5 levels, Coaching Shop) — but the effective ceiling is also capped by the highest league you've ever reached (peak, not current), from 60 in League 5 up to 100 in League 1. Both gates must be cleared to hit 100. Physical Health caps at ${BAL.statCapBase} until you invest in Sports Physio.</p>
      <p><b>The season:</b> a ${BAL.preseasonDays}-day preseason to train, then a ${BAL.seasonRounds}-round regular season — one match a week against a named rival, all scheduled in advance, each testing that week's active skills. Finish in the top ${BAL.playoffSize} of your ${BAL.seasonRounds + 1}-competitor league to reach the knockout playoffs. Lose a playoff match and you're out; win the Final and you're champion.</p>
      <p>Miss the playoffs and your season ends early — but training never stops. You get a ${BAL.trainingCampDays}-day training camp to prepare for next year, the same amount of time a full playoff run would have taken, so missing the cut isn't a worse deal than making it and getting knocked out early.</p>
      <p><b>Leagues:</b> there are ${BAL.leagueCount} leagues, League 1 at the top and League 5 at the bottom — you start in League 5. Every league has a persistent roster of named rivals whose ratings evolve from real simulated results year after year, same as yours. Finish top ${BAL.promotionCount} of your league's table at season's end and you're promoted a tier; finish bottom ${BAL.relegationCount} and you're relegated — this applies to every competitor in every league, not just you, so the standings you see are a living world, not scenery. Check the Leagues screen any time to see all ${BAL.leagueCount} tables. Promotion and relegation are based purely on table position — the playoffs are a separate prize, unrelated to which league you're in next year.</p>
      <p>Cash and Rank carry across seasons and leagues — spend cash in the Coaching Shop any time.</p>
    </div>`;
  openModal(html);
}

function openInstall() {
  const html = `
    <h2>Add to Home Screen</h2>
    <div class="modal-section">
      <p>To install Cell Grind as an app icon on your iPhone:</p>
      <p>1. Open this page in <b>Safari</b>.<br>
      2. Tap the <b>Share</b> icon (square with an arrow).<br>
      3. Scroll down and tap <b>Add to Home Screen</b>.<br>
      4. Tap <b>Add</b>.</p>
      <p>It'll launch full-screen from your home screen, and your progress is saved on this device.</p>
    </div>`;
  openModal(html);
}

/* ---------------------------------------------------------------------- */
/* End day flow                                                           */
/* ---------------------------------------------------------------------- */
function endDay() {
  const dayHeaderHtml = `Day ${state.day} — Results`;
  const entry = { html: dayHeaderHtml, cls: "day-header" };
  state.logEntries.push(entry);
  appendLog(entry.html, entry.cls);

  const events = resolveDay();
  events.forEach((ev) => {
    const cls = ev.type === "good" ? "event-good" : ev.type === "bad" ? "event-bad" : "";
    const e = { html: ev.text, cls };
    state.logEntries.push(e);
    appendLog(e.html, e.cls);
  });

  const skillCycleEvent = advanceSkillCycle();
  if (skillCycleEvent) {
    const e = { html: skillCycleEvent, cls: "event-season" };
    state.logEntries.push(e);
    appendLog(e.html, e.cls);
  }

  const { matchResult, phaseEvent } = processDayEnd();

  if (matchResult) {
    const oppText = matchResult.opponentName ? ` vs ${matchResult.opponentName}` : "";
    const label = matchResult.roundLabel || "Match";
    const summary = matchResult.forfeit
      ? `🏆 ${label} forfeited (injured)${oppText}. Rank now ${fmt(state.rank)}.`
      : `🏆 ${label}${oppText} — ${matchResult.win ? "WON" : "LOST"} (rating ${matchResult.opponentRating}). Rank ${fmtSigned(matchResult.ratingChange, 0)} → ${fmt(state.rank)}. +$${matchResult.cashReward}.`;
    const e = { html: summary, cls: "event-match" };
    state.logEntries.push(e);
    appendLog(e.html, e.cls);
  }

  if (phaseEvent) {
    const e = { html: phaseEvent, cls: "event-season" };
    state.logEntries.push(e);
    appendLog(e.html, e.cls);
  }

  state.day += 1;
  const saved = saveState();
  showSaveToast(saved);

  renderTopbar();
  renderStats();
  renderPlanner();

  $("log").scrollTop = $("log").scrollHeight;

  if (matchResult) {
    showMatchModal(matchResult);
  }
}

/* ---------------------------------------------------------------------- */
/* Input wiring                                                           */
/* ---------------------------------------------------------------------- */
function isSkillKey(key) {
  return SKILL_KEYS.includes(key);
}
function getAllocHours(key) {
  return isSkillKey(key) ? state.allocation.skills[key] || 0 : state.allocation[key];
}
function setAllocHoursRaw(key, val) {
  if (isSkillKey(key)) state.allocation.skills[key] = val;
  else state.allocation[key] = val;
}
function setAllocation(key, val) {
  const maxForKey = isSkillKey(key) ? SKILL_MAX_HOURS : ACT_MAX[key];
  const others = totalAssigned() - getAllocHours(key);
  const maxAllowed = Math.min(maxForKey, 24 - others);
  setAllocHoursRaw(key, clamp(val, 0, Math.max(0, maxAllowed)));
  renderPlanner();
}

function wireInputs() {
  ACT_KEYS.forEach((k) => {
    const input = $(ACT_INPUT_IDS[k]);
    input.addEventListener("input", () => setAllocation(k, Number(input.value)));
  });

  document.querySelectorAll(".activity[data-act] .step-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const act = btn.getAttribute("data-act");
      const dir = Number(btn.getAttribute("data-dir"));
      if (act === "exercise" && state.injury.active) return;
      setAllocation(act, state.allocation[act] + dir);
    });
  });

  // Skill rows are rebuilt on every renderPlanner() call (their set changes
  // weekly), so their inputs are wired via delegation on the stable
  // container rather than direct listeners that would go stale.
  const skillContainer = $("skillActivityRows");
  skillContainer.addEventListener("input", (e) => {
    if (!e.target.matches("input[type=range]")) return;
    const key = e.target.getAttribute("data-skill");
    if (key) setAllocation(key, Number(e.target.value));
  });
  skillContainer.addEventListener("click", (e) => {
    const btn = e.target.closest(".step-btn");
    if (!btn) return;
    const key = btn.getAttribute("data-skill");
    const dir = Number(btn.getAttribute("data-dir"));
    setAllocation(key, getAllocHours(key) + dir);
  });

  $("endDayBtn").addEventListener("click", endDay);
  $("menuBtn").addEventListener("click", openMenu);
  $("cashChip").addEventListener("click", openShop);
  $("rankChip").addEventListener("click", openCareer);
}

/* ---------------------------------------------------------------------- */
/* Init                                                                   */
/* ---------------------------------------------------------------------- */
function renderAll() {
  renderTopbar();
  renderStats();
  renderPlanner();
  renderFullLog();
  if (state.logEntries.length === 0) {
    const welcome = {
      html: "Welcome to Cell Grind. Plan your first day, then tap <b>End Day</b>.",
      cls: "",
    };
    state.logEntries.push(welcome);
    appendLog(welcome.html, welcome.cls);
  } else if (lastLoadedFromSave) {
    appendLog(`Welcome back — resumed from Day ${state.day}.`, "event-good");
  }

  if (!STORAGE_OK) {
    appendLog(
      "⚠️ This browser isn't allowing saves (private/incognito mode, or storage is blocked). You can still play, but progress won't persist after you close this tab.",
      "event-bad"
    );
  }
}

function init() {
  wireInputs();
  renderAll();

  if (!state.playerName) {
    openNameModal(true);
  }

  if ("serviceWorker" in navigator) {
    // updateViaCache: "none" stops the browser from serving a cached copy
    // of sw.js itself when checking for updates; reg.update() forces that
    // check immediately on every load instead of waiting on the browser's
    // own (much lazier) update heuristic.
    navigator.serviceWorker
      .register("sw.js", { updateViaCache: "none" })
      .then((reg) => reg.update())
      .catch(() => {});
    // If a newer service worker takes over (a fresh deploy was installed),
    // reload once so the page's own HTML/JS is the new version too, instead
    // of new cached assets running against this tab's already-loaded code.
    let refreshedForUpdate = false;
    navigator.serviceWorker.addEventListener("controllerchange", () => {
      if (refreshedForUpdate) return;
      refreshedForUpdate = true;
      window.location.reload();
    });
  }
}

document.addEventListener("DOMContentLoaded", init);
