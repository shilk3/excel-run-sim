/* Cell Grind — Excel Esports Manager
 * Single-file game engine: state, daily simulation, matches, shop, UI rendering.
 */

const APP_VERSION = "1.3.0";
const SAVE_KEY = "cellgrind_save_v1";

/* ---------------------------------------------------------------------- */
/* Balance constants — tune game feel here                                */
/* ---------------------------------------------------------------------- */
const BAL = {
  idealSleep: 8,
  excelGainBase: 0.65,
  exerciseGainBase: 1.15,
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
  matchIntervalDays: 7,
};

// Each upgrade has up to 3 purchasable levels. state.upgrades[key] stores
// the current level (0 = not purchased). Buying goes 0->1->2->3 in order;
// each level's fields describe that level's total (not additive) effect.
const UPGRADES = {
  coach: {
    name: "Personal Coach",
    icon: "🧑‍🏫",
    levels: [
      { cost: 300, bonus: 0.1, desc: "+10% Excel training gains" },
      { cost: 650, bonus: 0.2, desc: "+20% Excel training gains" },
      { cost: 1200, bonus: 0.35, desc: "+35% Excel training gains" },
    ],
  },
  physio: {
    name: "Sports Physio",
    icon: "🩺",
    levels: [
      { cost: 300, injuryReduceMult: 0.8, exerciseBonus: 0.05, desc: "-20% injury risk, +5% Exercise gains" },
      { cost: 650, injuryReduceMult: 0.65, exerciseBonus: 0.1, desc: "-35% injury risk, +10% Exercise gains" },
      { cost: 1200, injuryReduceMult: 0.5, exerciseBonus: 0.18, desc: "-50% injury risk, +18% Exercise gains" },
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

/* ---------------------------------------------------------------------- */
/* State                                                                  */
/* ---------------------------------------------------------------------- */
function freshState() {
  return {
    day: 1,
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
  return clamp(1 - 0.6 * (excel / 100), 0.2, 1.0);
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
/* Day resolution                                                         */
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
  const coachMult = coachEff ? 1 + coachEff.bonus : 1.0;
  const excelEff = effectiveHours(excelH);
  const excelGain = excelEff * BAL.excelGainBase * focusMult * physMult * skillDiminish(s.excel) * coachMult;
  const excelRust = excelH === 0 ? Math.min(0.15, s.excel * 0.003) : 0;
  const excelDelta = excelGain - excelRust;
  s.excel = clamp(s.excel + excelDelta, 0, 100);
  if (excelH > 0) {
    let note = "";
    if (focusMult < 0.5) note = " (focus was poor — low energy or high stress hurt your training)";
    else if (physMult < 0.75) note = " (low physical health capped your gains)";
    events.push({ type: excelDelta > 0.5 ? "good" : "neutral", text: `📊 Excel Training (${excelH}h): ${fmtSigned(excelDelta)} skill${note}` });
  } else if (excelRust > 0) {
    events.push({ type: "bad", text: `📊 No training today: skill rusted slightly (${fmtSigned(-excelRust)})` });
  }

  // ---- Physical health ----
  const physioMult = physioEff ? 1 + physioEff.exerciseBonus : 1.0;
  const exerciseGain = effectiveHours(exerciseH) * BAL.exerciseGainBase * physioMult;
  const sleepDebtDecayFactor = nutritionistEff ? nutritionistEff.decayMult : 1.0;
  const physDecayFromSleep = s.sleepDebt * 0.045 * sleepDebtDecayFactor;
  const detrainMult = recoveryEff ? recoveryEff.detrainMult : 1.0;
  const detrain = exerciseH < BAL.detrainThresholdHours ? BAL.detrainDecay * detrainMult : 0;
  const physDelta = exerciseGain - physDecayFromSleep - detrain;
  s.phys = clamp(s.phys + physDelta, 0, 100);
  if (exerciseH > 0) {
    events.push({ type: physDelta > 0 ? "good" : "neutral", text: `🏃 Exercise (${exerciseH}h): ${fmtSigned(physDelta)} physical health` });
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
      s.phys = clamp(s.phys - loss, 0, 100);
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

// The ladder gets tougher the higher you climb, independent of your own
// rating swings — every 150 rating above the starting 800, opponents skew
// systematically stronger. This is what stops "cap every stat once, then
// win forever": a maxed-out performance score gives a strong but bounded
// edge (see matchRating below), while the opposition keeps getting harder,
// so sustained results — not a one-time stat cap — are what keeps you
// winning.
function tierToughness(rank) {
  return Math.floor(Math.max(0, rank - 800) / 150) * 12;
}

function estimateOpponent() {
  const toughness = tierToughness(state.rank);
  const mid = state.rank + toughness;
  const low = clamp(mid - 130, 400, 5000);
  const high = clamp(mid + 150, 400, 5000);
  let label;
  if (toughness === 0) label = "Even matchup";
  else if (toughness < 40) label = "Competitive";
  else if (toughness < 90) label = "Tough matchup";
  else label = "Elite competition";
  return { low, high, label };
}

function daysUntilNextMatch() {
  const next = Math.ceil(state.day / BAL.matchIntervalDays) * BAL.matchIntervalDays;
  return next - state.day;
}

function simulateMatch() {
  const managerEff = upgradeEffect("manager");
  const rankLossMult = managerEff ? managerEff.rankLossMult : 1.0;
  const cashBonusMult = managerEff ? managerEff.cashBonusMult : 1.0;

  if (state.injury.active) {
    const loss = Math.round(10 * rankLossMult);
    state.rank = Math.max(400, state.rank - loss);
    return {
      forfeit: true,
      text: `You're injured and had to forfeit this week's match. Rank -${loss}.`,
    };
  }

  const perf = performanceScore();
  const opponentRating = clamp(state.rank + tierToughness(state.rank) + randInt(-130, 150), 400, 5000);
  // Perfect performance (100) gives a strong but no longer overwhelming form
  // bonus — enough to matter, not enough to guarantee a win against a tough
  // roll once the ladder has gotten hard.
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
    ratingChange,
    cashReward,
  };
}

/* ---------------------------------------------------------------------- */
/* UI rendering                                                           */
/* ---------------------------------------------------------------------- */
const $ = (id) => document.getElementById(id);

const ACT_KEYS = ["excel", "exercise", "sleep", "relax"];
const ACT_INPUT_IDS = { excel: "excelHours", exercise: "exerciseHours", sleep: "sleepHours", relax: "relaxHours" };
const ACT_MAX = { excel: 16, exercise: 12, sleep: 12, relax: 12 };

function renderTopbar() {
  $("dayNum").textContent = state.day;
  $("cashVal").textContent = fmt(state.cash);
  $("rankVal").textContent = fmt(state.rank);
  const week = Math.floor((state.day - 1) / BAL.matchIntervalDays) + 1;
  const phase = state.day <= BAL.matchIntervalDays ? "Preseason" : `Season Wk ${week}`;
  const daysToMatch = daysUntilNextMatch();
  const matchText = daysToMatch === 0 ? "Match today!" : `Match in ${daysToMatch}d`;
  $("phaseLabel").textContent = `${phase} · ${matchText}`;
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
  if (result.forfeit) {
    html = `
      <div class="match-card">
        <div class="match-result loss">FORFEIT</div>
        <div class="match-sub">${result.text}</div>
        <button class="primary-btn" id="matchOk">Continue</button>
      </div>`;
  } else {
    html = `
      <div class="match-card">
        <div class="match-result ${result.win ? "win" : "loss"}">${result.win ? "VICTORY" : "DEFEAT"}</div>
        <div class="match-sub">Opponent rating: ${result.opponentRating}</div>
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
    <div class="menu-row" id="menuCareer"><span>📈 Career Stats</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuHow"><span>❓ How to Play</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuInstall"><span>📲 Add to Home Screen</span><span class="arrow">›</span></div>
    <button class="ghost-btn" id="menuReset">Reset Career</button>
    <div class="version-tag">Cell Grind v${APP_VERSION}</div>
  `;
  openModal(html);
  $("menuShop").addEventListener("click", openShop);
  $("menuCareer").addEventListener("click", openCareer);
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

function openCareer() {
  const days = daysUntilNextMatch();
  const opp = estimateOpponent();
  const nextMatchDay = state.day + days;
  const html = `
    <h2>Career Stats</h2>
    <div class="modal-section">
      <h3>Next Match</h3>
      <p>${days === 0 ? "Today, after you end this day" : `In ${days} day${days === 1 ? "" : "s"}`} (Day ${nextMatchDay})<br>
      Expected opponent rating: ~${fmt(opp.low)}–${fmt(opp.high)} — <b>${opp.label}</b><br>
      Your performance score right now: ${fmt(performanceScore())} / 100</p>
    </div>
    <div class="modal-section">
      <h3>Record</h3>
      <p>Day ${state.day} · Rank ${fmt(state.rank)} (peak ${fmt(state.peakRank)})<br>
      Wins: ${state.wins} · Losses: ${state.losses}<br>
      Cash earned to date reflects current balance: $${fmt(state.cash)}</p>
    </div>
    <div class="modal-section">
      <h3>Current Stats</h3>
      <p>Excel Skill: ${fmt(state.stats.excel)} / 100<br>
      Physical Health: ${fmt(state.stats.phys)} / 100<br>
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
      <p>Every ${BAL.matchIntervalDays} days you compete in a ranked match. Win to climb the rank ladder and earn prize money — spend it in the Coaching Shop on upgrades, each with up to 3 levels. Tap your rank in the top bar any time to preview your next opponent.</p>
      <p>The ladder gets tougher the higher you climb — opponents get systematically stronger as your rank rises, so capping your stats once isn't the finish line. Staying on top takes sustained good management and shop investment.</p>
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

  let matchResult = null;
  if (state.day % BAL.matchIntervalDays === 0) {
    matchResult = simulateMatch();
    const summary = matchResult.forfeit
      ? `🏆 Match forfeited (injured). Rank now ${fmt(state.rank)}.`
      : `🏆 Match ${matchResult.win ? "won" : "lost"} vs rating ${matchResult.opponentRating}. Rank ${fmtSigned(matchResult.ratingChange, 0)} → ${fmt(state.rank)}. +$${matchResult.cashReward}.`;
    const e = { html: summary, cls: "event-match" };
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
