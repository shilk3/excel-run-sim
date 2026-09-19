/* Cell Grind — Excel Esports Manager
 * Single-file game engine: state, daily simulation, season/league structure,
 * matches, shop, UI rendering.
 */

const APP_VERSION = "2.1.0";
const SAVE_KEY = "cellgrind_save_v1";

/* ---------------------------------------------------------------------- */
/* Balance constants — tune game feel here                                */
/* ---------------------------------------------------------------------- */
const BAL = {
  idealSleep: 8,
  excelGainBase: 0.24,
  exerciseGainBase: 0.6,
  softFatigueCap: 8, // hours per activity before in-day fatigue kicks in
  hardFatigueCap: 12,
  fatigueMultSoft: 0.6, // effectiveness for hours between soft and hard cap
  fatigueMultHard: 0.3, // effectiveness for hours beyond hard cap
  energyDrainPerHour: { excel: 1.0, exercise: 1.3 },
  energyRestorePerSleepHour: 10.5,
  relaxEnergyRestore: 1.2,
  relaxStressRelief: 2.4,
  stressLoadPerHour: 0.85,
  sleepDebtGoodBonus: 3, // stress relief when sleepHours >= ideal
  overtrainThreshold: 10, // exercise hours before injury risk starts
  injuryChancePerExcessHour: 0.045,
  injuryPhysLoss: [15, 25],
  injuryDaysRange: [3, 5],
  burnoutStressThreshold: 100,
  burnoutRecoverThreshold: 65,
  burnoutEffectivenessMult: 0.2,
  detrainThresholdHours: 1, // exercise hours below this triggers slow detraining
  detrainDecay: 0.35,
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
  statCapBase: 70, // Excel Skill / Physical Health ceiling with zero relevant upgrades
  statCapPerLevel: 10, // + this much per Coach / Physio level (max 3 levels -> +30 -> 100)
};

// Each upgrade has up to 3 purchasable levels. state.upgrades[key] stores
// the current level (0 = not purchased). Buying goes 0->1->2->3 in order;
// each level's fields describe that level's total (not additive) effect.
const UPGRADES = {
  coach: {
    name: "Personal Coach",
    icon: "🧑‍🏫",
    levels: [
      { cost: 300, bonus: 0.1, desc: "+10% Excel training gains, skill ceiling 80" },
      { cost: 650, bonus: 0.2, desc: "+20% Excel training gains, skill ceiling 90" },
      { cost: 1200, bonus: 0.35, desc: "+35% Excel training gains, skill ceiling 100" },
    ],
  },
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
      { cost: 250, sleepDebtMult: 0.85, desc: "-15% Sleep Debt build-up" },
      { cost: 550, sleepDebtMult: 0.75, desc: "-25% Sleep Debt build-up" },
      { cost: 1000, sleepDebtMult: 0.6, desc: "-40% Sleep Debt build-up" },
    ],
  },
  nutritionist: {
    name: "Nutritionist",
    icon: "🥗",
    levels: [
      { cost: 300, decayMult: 0.85, desc: "-15% physical decay from Sleep Debt" },
      { cost: 650, decayMult: 0.75, desc: "-25% physical decay from Sleep Debt" },
      { cost: 1150, decayMult: 0.6, desc: "-40% physical decay from Sleep Debt" },
    ],
  },
  meditation: {
    name: "Meditation Coach",
    icon: "🧘",
    levels: [
      { cost: 200, reliefMult: 1.2, desc: "+20% stress relief from Relaxation" },
      { cost: 450, reliefMult: 1.4, desc: "+40% stress relief from Relaxation" },
      { cost: 850, reliefMult: 1.65, desc: "+65% stress relief from Relaxation" },
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
};

function upgradeLevel(key) {
  return state.upgrades[key] || 0;
}
function upgradeEffect(key) {
  const lvl = upgradeLevel(key);
  return lvl > 0 ? UPGRADES[key].levels[lvl - 1] : null;
}
function excelCap() {
  return BAL.statCapBase + BAL.statCapPerLevel * upgradeLevel("coach");
}
function physCap() {
  return BAL.statCapBase + BAL.statCapPerLevel * upgradeLevel("physio");
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
/* State                                                                  */
/* ---------------------------------------------------------------------- */
function freshState() {
  return {
    day: 1, // total career days played — flavor/log only
    playerName: null,
    year: 1,
    seasonPhase: "preseason", // preseason | regular | playoffs | offseason
    phaseDay: 0, // days elapsed in the current phase
    roundIndex: 0, // next regular-season fixture index (0-38)
    schedule: generateSeasonSchedule(800),
    seasonResults: [], // { opponent, win } per completed regular-season round
    lastStandings: null,
    lastPlayerPosition: null,
    offseasonDays: BAL.offseasonDays,
    offseasonReason: null, // "missed" | "playoffs" — set when entering offseason
    playoff: null, // { stage, currentRound, eliminated, champion, playerSeed }
    cash: 100,
    rank: 800,
    peakRank: 800,
    wins: 0,
    losses: 0,
    stats: { excel: 8, phys: 65, energy: 80, sleepDebt: 0, stress: 8 },
    allocation: { excel: 4, exercise: 1, sleep: 8, relax: 2 },
    injury: { active: false, daysLeft: 0 },
    burnout: { active: false, daysLeft: 0 },
    upgrades: { coach: 0, physio: 0, sleepApp: 0, nutritionist: 0, meditation: 0, recovery: 0, manager: 0 },
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
  // regular season vs named rivals -> top-16 playoffs -> offseason -> new
  // year). Any save that predates it starts fresh at Year 1 preseason with
  // a new schedule anchored to whatever rank they'd already earned.
  if (!parsed.seasonPhase) {
    parsed.year = 1;
    parsed.seasonPhase = "preseason";
    parsed.phaseDay = 0;
    parsed.roundIndex = 0;
    parsed.seasonResults = [];
    parsed.lastStandings = null;
    parsed.lastPlayerPosition = null;
    parsed.offseasonDays = BAL.offseasonDays;
    parsed.offseasonReason = null;
    parsed.playoff = null;
    parsed.schedule = generateSeasonSchedule(parsed.rank || 800);
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

function focusMultiplier(energy, stress) {
  const energyFactor = 0.4 + 0.6 * (energy / 100);
  const stressFactor = 1 - 0.5 * (stress / 100);
  return clamp(energyFactor * stressFactor, 0.15, 1.0);
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
  const coachEff = upgradeEffect("coach");
  const physioEff = upgradeEffect("physio");
  const sleepAppEff = upgradeEffect("sleepApp");
  const nutritionistEff = upgradeEffect("nutritionist");
  const meditationEff = upgradeEffect("meditation");
  const recoveryEff = upgradeEffect("recovery");
  const events = [];

  // ---- Injury / burnout lockouts: enforce before computing effects ----
  let excelH = a.excel;
  let exerciseH = state.injury.active ? 0 : a.exercise;
  const sleepH = a.sleep;
  const relaxH = a.relax;

  const burnoutActive = state.burnout.active;
  const focusMult = focusMultiplier(s.energy, s.stress) * (burnoutActive ? BAL.burnoutEffectivenessMult : 1);
  const physMult = physSynergy(s.phys);

  // ---- Excel skill ----
  const cap = excelCap();
  const wasAtCap = s.excel >= cap - 0.05;
  const coachMult = coachEff ? 1 + coachEff.bonus : 1.0;
  const excelEff = effectiveHours(excelH);
  const excelGain = excelEff * BAL.excelGainBase * focusMult * physMult * skillDiminish(s.excel) * coachMult;
  const excelRust = excelH === 0 ? Math.min(0.15, s.excel * 0.003) : 0;
  const excelDelta = excelGain - excelRust;
  s.excel = clamp(s.excel + excelDelta, 0, cap);
  if (excelH > 0) {
    let note = "";
    if (wasAtCap && cap < 100) note = ` (capped at ${cap} — upgrade Personal Coach for a higher ceiling)`;
    else if (focusMult < 0.5) note = " (focus was poor — low energy or high stress hurt your training)";
    else if (physMult < 0.75) note = " (low physical health capped your gains)";
    events.push({ type: excelDelta > 0.5 ? "good" : "neutral", text: `📊 Excel Training (${excelH}h): ${fmtSigned(excelDelta)} skill${note}` });
  } else if (excelRust > 0) {
    events.push({ type: "bad", text: `📊 No training today: skill rusted slightly (${fmtSigned(-excelRust)})` });
  }

  // ---- Physical health ----
  const pCap = physCap();
  const physWasAtCap = s.phys >= pCap - 0.05;
  const physioMult = physioEff ? 1 + physioEff.exerciseBonus : 1.0;
  const exerciseGain = effectiveHours(exerciseH) * BAL.exerciseGainBase * physioMult * physDiminish(s.phys);
  const sleepDebtDecayFactor = nutritionistEff ? nutritionistEff.decayMult : 1.0;
  const physDecayFromSleep = s.sleepDebt * 0.045 * sleepDebtDecayFactor;
  const detrainMult = recoveryEff ? recoveryEff.detrainMult : 1.0;
  const detrain = exerciseH < BAL.detrainThresholdHours ? BAL.detrainDecay * detrainMult : 0;
  const physDelta = exerciseGain - physDecayFromSleep - detrain;
  s.phys = clamp(s.phys + physDelta, 0, pCap);
  if (exerciseH > 0) {
    let note = physWasAtCap && pCap < 100 ? ` (capped at ${pCap} — upgrade Sports Physio for a higher ceiling)` : "";
    events.push({ type: physDelta > 0 ? "good" : "neutral", text: `🏃 Exercise (${exerciseH}h): ${fmtSigned(physDelta)} physical health${note}` });
  }
  if (physDecayFromSleep > 1) {
    events.push({ type: "bad", text: `🌙 Poor sleep is wearing down your body (${fmtSigned(-physDecayFromSleep)} health from sleep debt)` });
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
  // absorb a bad night and Energy never visibly drops. Chronic sleep debt
  // degrades how restorative sleep is; today's activity then spends down
  // whatever that sleep gave you.
  const sleepQualityFactor = clamp(1 - Math.min(0.5, s.sleepDebt / 100), 0.5, 1);
  const energyFromSleep = clamp((sleepH / BAL.idealSleep) * 100, 0, 115) * sleepQualityFactor;
  const energySpend = excelH * BAL.energyDrainPerHour.excel + exerciseH * BAL.energyDrainPerHour.exercise;
  const relaxEnergyBonus = relaxH * BAL.relaxEnergyRestore;
  s.energy = clamp(energyFromSleep - energySpend + relaxEnergyBonus, 0, 100);

  // ---- Sleep debt ----
  const sleepAppMult = sleepAppEff ? sleepAppEff.sleepDebtMult : 1.0;
  let sleepDebtDelta;
  if (sleepH < BAL.idealSleep) {
    sleepDebtDelta = (BAL.idealSleep - sleepH) * 1.3 * sleepAppMult;
  } else {
    sleepDebtDelta = -Math.min(sleepH - BAL.idealSleep, 3) * 1.4;
  }
  s.sleepDebt = clamp(s.sleepDebt + sleepDebtDelta, 0, 100);
  if (sleepH < 6) {
    events.push({ type: "bad", text: `🌙 Only slept ${sleepH}h: sleep debt rising fast (${fmtSigned(sleepDebtDelta)})` });
  } else if (sleepH >= BAL.idealSleep) {
    events.push({ type: "good", text: `🌙 Slept ${sleepH}h: well rested (sleep debt ${fmtSigned(sleepDebtDelta)})` });
  }

  // ---- Stress ----
  const meditationMult = meditationEff ? meditationEff.reliefMult : 1.0;
  const stressLoad = (excelH + exerciseH) * BAL.stressLoadPerHour;
  const stressRelief = relaxH * BAL.relaxStressRelief * meditationMult + (sleepH >= BAL.idealSleep ? BAL.sleepDebtGoodBonus : 0);
  const stressFromDebt = s.sleepDebt * 0.1;
  const stressDelta = stressLoad - stressRelief + stressFromDebt;
  s.stress = clamp(s.stress + stressDelta, 0, 100);

  if (relaxH === 0 && stressLoad > 5) {
    events.push({ type: "bad", text: `🔥 No downtime today: stress climbed to ${fmt(s.stress)}` });
  } else if (relaxH > 0) {
    events.push({ type: "good", text: `🎮 Relaxed ${relaxH}h: stress relief ${fmtSigned(-stressRelief, 1)}` });
  }

  // ---- Burnout state transitions ----
  if (!state.burnout.active && s.stress >= BAL.burnoutStressThreshold) {
    state.burnout = { active: true, daysLeft: 1 };
    events.push({ type: "bad", text: `⚠️ BURNOUT! You've pushed too hard with too little rest. Training and exercise are far less effective until your stress drops — relax more.` });
  } else if (state.burnout.active && s.stress <= BAL.burnoutRecoverThreshold) {
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
function performanceScore() {
  const s = state.stats;
  return clamp(s.excel * 0.55 + s.phys * 0.25 + (100 - s.stress) * 0.15 + (100 - s.sleepDebt) * 0.05, 0, 100);
}

// A rating-tier toughness bump used only to *anchor* how strong a season's
// 39 named rivals are when the schedule is generated (higher career rank ->
// tougher league). It no longer drives individual match matchmaking, since
// matches are now against a fixed, known schedule.
function tierToughness(rank) {
  return Math.floor(Math.max(0, rank - 800) / 150) * 12;
}

function generateSeasonSchedule(baseRank) {
  const used = new Set();
  const schedule = [];
  for (let i = 0; i < BAL.seasonRounds; i++) {
    let name;
    do {
      name = generateOpponentName();
    } while (used.has(name));
    used.add(name);
    const spread = randInt(-260, 300);
    const rating = clamp(baseRank + tierToughness(baseRank) + spread, 400, 5000);
    schedule.push({ name, rating, played: false, win: null });
  }
  // Shuffle fixture order so difficulty isn't predictably sorted.
  for (let i = schedule.length - 1; i > 0; i--) {
    const j = randInt(0, i);
    const tmp = schedule[i];
    schedule[i] = schedule[j];
    schedule[j] = tmp;
  }
  return schedule;
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

// Called once the player's 39th regular-season match resolves. The other 39
// rivals' season records are approximated (not a full pairwise round robin)
// by giving each a win probability from their rating vs the league average,
// then rolling 39 Bernoulli trials — cheap, and still rating-correlated.
function computeSeasonStandings() {
  const avgRating = state.schedule.reduce((sum, o) => sum + o.rating, 0) / state.schedule.length;
  const npcResults = state.schedule.map((o) => {
    const p = 1 / (1 + Math.pow(10, (avgRating - o.rating) / 400));
    let wins = 0;
    for (let i = 0; i < BAL.seasonRounds; i++) {
      if (Math.random() < p) wins++;
    }
    return { name: o.name, rating: o.rating, points: wins * 3, isPlayer: false };
  });
  const playerWins = state.seasonResults.filter((r) => r.win).length;
  const playerEntry = { name: state.playerName || "You", rating: state.rank, points: playerWins * 3, isPlayer: true };
  const standings = [...npcResults, playerEntry].sort((a, b) => b.points - a.points || b.rating - a.rating);
  const playerPosition = standings.findIndex((s) => s.isPlayer) + 1;
  const qualified = playerPosition <= BAL.playoffSize;
  return { standings, playerPosition, qualified };
}

// Standard 16-bracket seeding pairs so top seeds are spread across the draw
// (1 and 2 can only meet in the Final, etc).
const SEED_PAIRS_16 = [[1, 16], [8, 9], [5, 12], [4, 13], [6, 11], [3, 14], [7, 10], [2, 15]];

function buildPlayoffBracket(standings) {
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
  return { stage: "r16", currentRound, eliminated: false, champion: false, playerSeed };
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
      const aWins = simulateNpcMatch(teamA.rating, teamB.rating);
      winners.push(aWins ? teamA : teamB);
    }
  }

  let seasonOver = false;
  if (p.eliminated) {
    seasonOver = true;
  } else if (p.stage === "f") {
    p.champion = true;
    seasonOver = true;
    summary = `👑 CHAMPION! You won the Year ${state.year} Final!`;
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
        const { standings, playerPosition, qualified } = computeSeasonStandings();
        state.lastStandings = standings;
        state.lastPlayerPosition = playerPosition;
        if (qualified) {
          state.seasonPhase = "playoffs";
          state.playoff = buildPlayoffBracket(standings);
          phaseEvent = `🏆 Regular season complete! Finished #${playerPosition} of ${standings.length} — through to the playoffs as seed ${state.playoff.playerSeed}.`;
        } else {
          state.seasonPhase = "offseason";
          state.offseasonDays = BAL.trainingCampDays;
          state.offseasonReason = "missed";
          phaseEvent = `📋 Regular season complete. Finished #${playerPosition} of ${standings.length} — missed the top ${BAL.playoffSize} playoff cutoff. ${BAL.trainingCampDays}-day training camp starts now to get ready for next season.`;
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
      state.schedule = generateSeasonSchedule(state.rank);
      state.playoff = null;
      // Offseason recovery — a clean slate for the new year.
      state.stats.stress = 0;
      state.stats.sleepDebt = 0;
      state.stats.energy = 100;
      phaseEvent = `🎉 Year ${state.year} begins! A fresh ${BAL.seasonRounds}-round season has been scheduled — good luck.`;
    }
    return { matchResult, phaseEvent };
  }

  return { matchResult, phaseEvent };
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
    return {
      kind: "fixture",
      daysUntil,
      opponentName: fixture.name,
      opponentRating: Math.round(fixture.rating),
      label: daysUntil <= 0 ? `Round ${roundNum} today vs ${fixture.name}!` : `Round ${roundNum}/${BAL.seasonRounds} in ${daysUntil}d vs ${fixture.name}`,
    };
  }
  if (s.seasonPhase === "playoffs") {
    const daysUntil = BAL.roundIntervalDays - s.phaseDay;
    const p = s.playoff;
    const idx = p.currentRound.findIndex((t) => t.isPlayer);
    const opp = idx >= 0 ? p.currentRound[idx % 2 === 0 ? idx + 1 : idx - 1] : null;
    const roundName = PLAYOFF_ROUND_NAMES[p.stage];
    return {
      kind: "playoff",
      daysUntil,
      opponentName: opp ? opp.name : "?",
      opponentRating: opp ? Math.round(opp.rating) : null,
      label: daysUntil <= 0 ? `${roundName} today vs ${opp ? opp.name : "?"}!` : `${roundName} in ${daysUntil}d vs ${opp ? opp.name : "?"}`,
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

function phaseLabelText() {
  const s = state;
  if (s.seasonPhase === "preseason") return `Preseason (Day ${s.phaseDay}/${BAL.preseasonDays})`;
  if (s.seasonPhase === "regular") return `Round ${Math.min(s.roundIndex + 1, BAL.seasonRounds)}/${BAL.seasonRounds}`;
  if (s.seasonPhase === "playoffs") return `Playoffs: ${PLAYOFF_ROUND_NAMES[s.playoff.stage]}`;
  if (s.seasonPhase === "offseason") {
    return s.offseasonReason === "missed" ? `Training Camp (Day ${s.phaseDay}/${s.offseasonDays})` : "Offseason";
  }
  return "";
}

/* ---------------------------------------------------------------------- */
/* UI rendering                                                           */
/* ---------------------------------------------------------------------- */
const $ = (id) => document.getElementById(id);

const ACT_KEYS = ["excel", "exercise", "sleep", "relax"];
const ACT_INPUT_IDS = { excel: "excelHours", exercise: "exerciseHours", sleep: "sleepHours", relax: "relaxHours" };
const ACT_MAX = { excel: 16, exercise: 12, sleep: 12, relax: 12 };

function renderTopbar() {
  $("yearNum").textContent = state.year;
  $("cashVal").textContent = fmt(state.cash);
  $("rankVal").textContent = fmt(state.rank);
  $("phaseLabel").textContent = phaseLabelText();
  $("matchCounter").textContent = getNextMatchInfo().label;
}

function renderStats() {
  const s = state.stats;
  setBar("excel", s.excel, 100);
  setBar("phys", s.phys, 100);
  setBar("energy", s.energy, 100);
  setBar("sleepDebt", s.sleepDebt, 100);
  setBar("stress", s.stress, 100);

  const banner = $("warningBanner");
  const msgs = [];
  if (state.burnout.active) msgs.push("🔥 Burnout active — training & exercise are far less effective. Relax to recover.");
  else if (s.stress >= 80) msgs.push("🔥 Stress critical — burnout imminent. Schedule relaxation soon.");
  if (state.injury.active) msgs.push(`🤕 Injured — Exercise disabled for ${state.injury.daysLeft} more day(s).`);
  if (s.sleepDebt >= 70) msgs.push("🌙 Severe sleep debt — your body is breaking down. Sleep more.");
  if (s.phys <= 20) msgs.push("💪 Physical health critically low — it's capping your Excel performance.");

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
  return a.excel + a.exercise + a.sleep + a.relax;
}

function renderPlanner() {
  const a = state.allocation;
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
function shopHtml() {
  const items = Object.entries(UPGRADES)
    .map(([key, u]) => {
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
    })
    .join("");
  return `
    <h2>Coaching Shop</h2>
    <div class="modal-section">
      <h3>Cash: $${fmt(state.cash)}</h3>
      ${items}
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
    <div class="menu-row" id="menuRename"><span>✏️ Rename Player</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuHow"><span>❓ How to Play</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuInstall"><span>📲 Add to Home Screen</span><span class="arrow">›</span></div>
    <button class="ghost-btn" id="menuReset">Reset Career</button>
    <div class="version-tag">Cell Grind v${APP_VERSION}</div>
  `;
  openModal(html);
  $("menuShop").addEventListener("click", openShop);
  $("menuCareer").addEventListener("click", openCareer);
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
      <h3>Last Season — Final Standings</h3>
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
      <p>Seed #${state.playoff.playerSeed} of ${BAL.playoffSize} · ${statusText}</p>
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
      <p>Year ${state.year} · Round ${seasonPlayed}/${BAL.seasonRounds} · Record ${seasonWins}-${seasonPlayed - seasonWins}</p>
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
      <h3>Current Stats</h3>
      <p>Excel Skill: ${fmt(state.stats.excel)} / ${excelCap()}${excelCap() < 100 ? " (upgrade Coach for more)" : ""}<br>
      Physical Health: ${fmt(state.stats.phys)} / ${physCap()}${physCap() < 100 ? " (upgrade Physio for more)" : ""}<br>
      Energy: ${fmt(state.stats.energy)} / 100<br>
      Sleep Debt: ${fmt(state.stats.sleepDebt)} / 100<br>
      Stress: ${fmt(state.stats.stress)} / 100</p>
    </div>`;
  openModal(html);
}

function openHowTo() {
  const html = `
    <h2>How to Play</h2>
    <div class="modal-section">
      <p>You manage a rising Excel esports competitor. Every day has 24 hours — split them across:</p>
      <p>
      📊 <b>Excel Training</b> — raises Excel Skill, your core competitive stat.<br>
      🏃 <b>Exercise</b> (running, cross training) — raises Physical Health.<br>
      🌙 <b>Sleep</b> — restores Energy and pays down Sleep Debt.<br>
      🎮 <b>Relaxation</b> — relieves Stress and prevents burnout.
      </p>
      <p><b>It's all connected:</b> poor sleep builds Sleep Debt, which wears down Physical Health even if you train well. Low Physical Health caps how much your Excel Training actually helps. Training hard without Relaxation builds Stress — hit 100 and you burn out, tanking your effectiveness until you rest.</p>
      <p>Overtraining physically (too much Exercise) risks injury, which locks out Exercise for several days.</p>
      <p><b>Stat ceilings:</b> Excel Skill and Physical Health cap at ${BAL.statCapBase} until you invest in the Coaching Shop — each level of Personal Coach raises your skill ceiling by ${BAL.statCapPerLevel}, each level of Sports Physio raises your health ceiling the same way. Maxing out at 100 in either stat requires buying every level.</p>
      <p><b>The season:</b> a ${BAL.preseasonDays}-day preseason to train, then a ${BAL.seasonRounds}-round regular season — one match a week against a named rival, all scheduled in advance. Finish in the top ${BAL.playoffSize} of the ${BAL.seasonRounds + 1}-competitor league (you + your rivals) to reach the knockout playoffs. Lose a playoff match and you're out; win the Final and you're champion.</p>
      <p>Miss the playoffs and your season ends early — but training never stops. You get a ${BAL.trainingCampDays}-day training camp to prepare for next year, the same amount of time a full playoff run would have taken, so missing the cut isn't a worse deal than making it and getting knocked out early. Either way, a new season with a fresh set of rivals begins once the year turns over.</p>
      <p>Cash and Rank carry across seasons — spend cash in the Coaching Shop any time.</p>
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
function setAllocation(act, val) {
  const a = state.allocation;
  const others = ACT_KEYS.filter((k) => k !== act).reduce((sum, k) => sum + a[k], 0);
  const maxAllowed = Math.min(ACT_MAX[act], 24 - others);
  a[act] = clamp(val, 0, Math.max(0, maxAllowed));
  renderPlanner();
}

function wireInputs() {
  ACT_KEYS.forEach((k) => {
    const input = $(ACT_INPUT_IDS[k]);
    input.addEventListener("input", () => setAllocation(k, Number(input.value)));
  });

  document.querySelectorAll(".step-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const act = btn.getAttribute("data-act");
      const dir = Number(btn.getAttribute("data-dir"));
      if (act === "exercise" && state.injury.active) return;
      setAllocation(act, state.allocation[act] + dir);
    });
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
    navigator.serviceWorker.register("sw.js").catch(() => {});
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
