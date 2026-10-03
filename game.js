/* Cell Grind — Excel Esports Manager
 * Single-file game engine: state, daily simulation, season/league structure,
 * matches, shop, UI rendering.
 */

const APP_VERSION = "4.36.0";
const SAVE_KEY = "cellgrind_save_v1";

/* ---------------------------------------------------------------------- */
/* Balance constants — tune game feel here                                */
/* ---------------------------------------------------------------------- */
const BAL = {
  idealSleep: 7, // also the Rest decay threshold: sleep below this drains Rest
  relaxComposureThreshold: 3, // Composure's decay threshold, in relaxation hours
  skillDecayThresholdHours: 1, // an active skill/Exercise below this hour count rusts
  skillGainBase: 0.24,
  exerciseGainBase: 0.6,
  nutritionGainBase: 5, // per hour allocated, before diminishing returns near the cap
  nutritionDecayFlat: 9, // lost per day when under the 1h threshold
  softFatigueCap: 8, // hours per activity before in-day fatigue kicks in
  hardFatigueCap: 12,
  fatigueMultSoft: 0.6, // effectiveness for hours between soft and hard cap
  fatigueMultHard: 0.3, // effectiveness for hours beyond hard cap
  relaxComposureRelief: 2.4,
  composureLoadPerHour: 0.85,
  restGoodSleepBonus: 3, // extra composure relief when sleepHours >= ideal
  injuryChancePerHour: 0.025, // every Gym hour adds this much injury chance that day — only 0h is risk-free
  gymMaxHours: 8,
  injuryPhysLoss: [5, 10],
  injuryDaysRange: [3, 5],
  burnoutComposureThreshold: 0,
  burnoutRecoverThreshold: 35,
  burnoutEffectivenessMult: 0.2,
  detrainThresholdHours: 1, // exercise hours below this triggers slow detraining
  detrainDecay: 0.35,
  // Rest is the only stat with a training-effectiveness effect.
  restTrainingBoostThreshold: 80, // <=80 Rest: 100% effective
  restTrainingBoostHigh: 90, // 81-90 Rest: 150%; >90 Rest: 200%
  // Composure's match-day effect: low Composure halves your skill on the
  // day it matters most, recovering in the same step pattern in reverse.
  composureMatchLow: 10, // <10 Composure: 50% skill
  composureMatchMid: 20, // 10-19: 75% skill; >=20: 100% skill
  // Nutrition governs the total hours available to allocate each day —
  // neglect it and the day itself gets shorter, not just less effective.
  nutritionHoursCapLow: 40, // <=40 Nutrition: only the floor below
  nutritionHoursCapHigh: 90, // >=90 Nutrition: the full day
  dailyHoursFloor: 16,
  dailyHoursCeiling: 24,
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
  // Promotion: the playoff champion plus the top 3 of the table *other than*
  // the champion — always 4, matching relegationCount so every tier stays at
  // exactly leagueSize. The top 3 are therefore always promoted.
  promotionTablePlaces: 3,
  oppPerfRange: [55, 85], // rivals' cosmetic match-day performance (see resolveMatch)
  relegationCount: 4,
  // Rating bounds — the same for the player and every rival. The floor only
  // stops a long losing run going negative; the ceiling is well above any
  // rating the game produces.
  ratingMin: 0,
  ratingMax: 5000,
  // Employment: a mandatory day job funds the whole career until you can
  // go pro. Work/Pro Duties share one slider and one strike system —
  // only what's "required" and what it's called changes with status.
  workHoursRequired: 9, // Employed: Work hours/day required to avoid a strike
  proDutyHoursRequired: 5, // Pro: Pro Duties hours/day required; hours above this master techniques
  // Pay isn't flat: it starts at the minimum for whichever role you're in and
  // rises $10/year for 5 years (then plateaus), tracked independently per
  // role (state.employment.workPay / .proPay) so Pro pays more at every
  // tenure than a day job does. Getting fired or dropped resets that one
  // role's rate back to its minimum for next time — seniority is lost, not
  // carried over.
  workPayMin: 70, // Employed: starting (and reset) daily pay
  proPayMin: 100, // Pro: starting (and reset) daily pay
  payRaisePerYear: 10,
  payRaiseMaxYears: 5, // raises stop after this many years of unbroken tenure in the same role
  dailyExpenses: 50, // cost of living, charged every single day regardless of employment status —
  // this is what makes losing your job actually cost you money, not just stall your income
  strikeWindowDays: 21, // 3 weeks — each strike expires this many days after it's earned
  strikesToFire: 3,
  jobSearchHoursRange: [10, 40], // Unemployed: cumulative hours to get re-hired — rolled fresh each time you lose a job
  jobSearchMinHours: 1, // Unemployed: at least this much Job Search a day, or the day can't end (no idling into endless debt)
  goProLeagueTier: 2, // Employed + leagueTier <= this + cash >= goProCash -> Pro, automatically
  goProCash: 5000,
  // Pros must stay current: a new technique queues up periodically: only
  // the front of the queue is "in progress" at once, taking hours above
  // the Pro Duties minimum. Falling behind never loses progress, but every
  // not-yet-mastered technique costs match performance.
  techniqueIntervalDays: 30,
  techniqueMinHours: 25,
  techniqueMaxHours: 40,
  techniquePenaltyPerUnmastered: 0.05,
};

// Flavor pool for technique-queue entries — real Excel features, cycled
// without repeating one already in the queue.
const EXCEL_TECHNIQUES = [
  "XLOOKUP", "Dynamic Arrays", "LAMBDA Functions", "Power Query", "Power Pivot",
  "FILTER & SORT Functions", "LET Functions", "Spill Ranges", "Data Tables",
  "Conditional Formatting Rules", "Array Formulas", "Pivot Table Slicers",
  "What-If Analysis", "Macros & VBA Basics", "Named Ranges", "INDEX-MATCH Mastery",
  "TEXTSPLIT & TEXTJOIN", "XMATCH", "Structured Table References", "Power Automate Flows",
];
function pickTechniqueName() {
  const inQueue = new Set(state.employment.techniqueQueue.map((t) => t.name));
  const pool = EXCEL_TECHNIQUES.filter((n) => !inQueue.has(n));
  const choices = pool.length ? pool : EXCEL_TECHNIQUES;
  return choices[randInt(0, choices.length - 1)];
}

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
// Regular season and playoffs: the stretch with a match every week.
function inSeason() {
  return state.seasonPhase === "regular" || state.seasonPhase === "playoffs";
}
// Only in-season weeks have a match to prepare for, so only they restrict
// training to that week's revealed 1-3 active skills. Preseason and the
// off-season (including training camp) open all 7.
function trainableSkills() {
  return inSeason() ? state.activeSkills : SKILL_KEYS;
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
// Staff are hired a week at a time (see hireStaff). Each level is a one-off
// fee to unlock — only once your peak league allows it — plus a weekly wage
// to have that level on hand. Level 1 is free to unlock. Wages are tuned so
// your weekly spare money covers ~90% of what you'd want in a 1-skill week
// and ~40% in a 3-skill week, in every league at that league's best level.
const SKILL_COACH_LEVELS = [
  { cost: 0, wage: 185, unlock: 5, bonus: 0.05, capAdd: 10 },
  { cost: 600, wage: 245, unlock: 4, bonus: 0.1, capAdd: 20 },
  { cost: 1000, wage: 300, unlock: 3, bonus: 0.15, capAdd: 30 },
  { cost: 1600, wage: 360, unlock: 2, bonus: 0.22, capAdd: 40 },
  { cost: 2500, wage: 470, unlock: 1, bonus: 0.3, capAdd: 50 },
];
const SUPPORT_LEVEL_FEES = [0, 500, 1200];
const SUPPORT_LEVEL_UNLOCK = [5, 3, 1];
const SUPPORT_WAGES = {
  physio: [30, 48, 75],
  nutritionist: [20, 32, 50],
  manager: [20, 32, 50],
  sleepApp: [15, 24, 38],
  meditation: [15, 24, 38],
  recovery: [15, 24, 38],
};
const UPGRADES = {};
SKILLS.forEach((sk) => {
  UPGRADES[coachKey(sk.key)] = {
    name: `${sk.name} Coach`,
    icon: sk.icon,
    levels: SKILL_COACH_LEVELS.map((lvl) => ({
      cost: lvl.cost,
      wage: lvl.wage,
      unlock: lvl.unlock,
      bonus: lvl.bonus,
      desc: `+${Math.round(lvl.bonus * 100)}% training, ceiling ${BAL.skillShopCapBase + lvl.capAdd}`,
    })),
  };
});
Object.assign(UPGRADES, {
  physio: {
    name: "Sports Physio",
    icon: "🩺",
    levels: [
      { cost: 300, injuryReduceMult: 0.8, exerciseBonus: 0.05, desc: "-20% injury risk, +5% Gym gains, Health ceiling 80" },
      { cost: 650, injuryReduceMult: 0.65, exerciseBonus: 0.1, desc: "-35% injury risk, +10% Gym gains, Health ceiling 90" },
      { cost: 1200, injuryReduceMult: 0.5, exerciseBonus: 0.18, desc: "-50% injury risk, +18% Gym gains, Health ceiling 100" },
    ],
  },
  sleepApp: {
    name: "Sleep Coach App",
    icon: "💤",
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
      { cost: 300, nutritionGainMult: 1.15, nutritionDecayMult: 0.85, desc: "+15% Nutrition gain, -15% Nutrition lost when neglected" },
      { cost: 650, nutritionGainMult: 1.3, nutritionDecayMult: 0.7, desc: "+30% Nutrition gain, -30% Nutrition lost when neglected" },
      { cost: 1150, nutritionGainMult: 1.5, nutritionDecayMult: 0.5, desc: "+50% Nutrition gain, -50% Nutrition lost when neglected" },
    ],
  },
  meditation: {
    name: "Meditation Coach",
    icon: "🧘",
    levels: [
      { cost: 200, reliefMult: 1.2, desc: "+20% Composure relief from Relax" },
      { cost: 450, reliefMult: 1.4, desc: "+40% Composure relief from Relax" },
      { cost: 850, reliefMult: 1.65, desc: "+65% Composure relief from Relax" },
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
    icon: "🧑‍💼",
    levels: [
      { cost: 400, rankLossMult: 0.9, cashBonusMult: 1.0, desc: "-10% rating lost on defeat" },
      { cost: 800, rankLossMult: 0.8, cashBonusMult: 1.05, desc: "-20% rating lost on defeat, +5% prize money" },
      { cost: 1500, rankLossMult: 0.65, cashBonusMult: 1.1, desc: "-35% rating lost on defeat, +10% prize money" },
    ],
  },
});
Object.keys(SUPPORT_WAGES).forEach((key) => {
  UPGRADES[key].levels.forEach((lvl, i) => {
    lvl.cost = SUPPORT_LEVEL_FEES[i];
    lvl.wage = SUPPORT_WAGES[key][i];
    lvl.unlock = SUPPORT_LEVEL_UNLOCK[i];
  });
});

// Highest level unlocked (fee paid; level 1 is free, so always at least 1).
function upgradeLevel(key) {
  return Math.max(1, state.upgrades[key] || 0);
}
// Level hired for this week (0 = not hired). Staff only do anything while
// hired, so every effect below reads this, not the unlocked level.
function hiredLevel(key) {
  return (state.staff && state.staff.hired && state.staff.hired[key]) || 0;
}
function upgradeEffect(key) {
  const lvl = hiredLevel(key);
  return lvl > 0 ? UPGRADES[key].levels[lvl - 1] : null;
}
// This week's wage for a level, pro-rated by the days left in the week
// (today included) and rounded up to the dollar.
function staffCost(key, level) {
  const wage = UPGRADES[key].levels[level - 1].wage;
  return Math.ceil((wage * daysLeftInWeek()) / 7);
}
function canUnlockLevel(key, level) {
  const lvl = UPGRADES[key].levels[level - 1];
  return !!lvl && state.peakLeagueTier <= lvl.unlock;
}
function hireStaff(key, level) {
  if (hiredLevel(key) || level < 1 || level > upgradeLevel(key)) return false;
  const cost = staffCost(key, level);
  if (state.cash < cost) return false;
  state.cash -= cost;
  state.yearCashFlow.expenses += cost;
  state.staff.hired[key] = level;
  return true;
}
function unlockStaffLevel(key) {
  const next = upgradeLevel(key) + 1;
  const lvl = UPGRADES[key].levels[next - 1];
  if (!lvl || !canUnlockLevel(key, next) || state.cash < lvl.cost) return false;
  state.cash -= lvl.cost;
  state.yearCashFlow.expenses += lvl.cost;
  state.upgrades[key] = next;
  return true;
}
// Today's injury chance from Gym hours — shared by the roll in resolveDay()
// and the "#% 🤕" readout on the Gym row so the two always agree.
function injuryChance(gymHours) {
  const physioEff = upgradeEffect("physio");
  const mult = physioEff ? physioEff.injuryReduceMult : 1.0;
  return clamp(gymHours * BAL.injuryChancePerHour * mult, 0, 1);
}
function physCap() {
  return BAL.statCapBase + BAL.statCapPerLevel * hiredLevel("physio");
}
// Effective skill ceiling stacks two independent gates: the Coaching Shop
// (per skill, 5 levels) and the highest league ever reached (peak tier).
function skillShopCap(key) {
  return BAL.skillShopCapBase + BAL.skillShopCapPerLevel * hiredLevel(coachKey(key));
}
function leagueSkillCap() {
  return LEAGUE_SKILL_CAP[state.peakLeagueTier] || LEAGUE_SKILL_CAP[5];
}
function skillCap(key) {
  return Math.min(skillShopCap(key), leagueSkillCap());
}
// Which gate is holding a skill's ceiling down — tells the player whether
// the fix is hiring a coach (Staff) or a promotion (peak league).
function skillCapCause(key) {
  const shop = skillShopCap(key);
  const league = leagueSkillCap();
  if (Math.min(shop, league) >= 100) return null;
  const lvl = hiredLevel(coachKey(key));
  const coachText = lvl > 0 ? `Coach Lv${lvl}` : "no coach";
  const leagueText = `League ${state.peakLeagueTier}`;
  if (shop < league) return coachText;
  if (league < shop) return leagueText;
  return `${coachText} + ${leagueText}`;
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
// How many hours the day actually has to allocate, driven by Nutrition —
// linear between the floor (at/below the low anchor) and a full 24h
// (at/above the high anchor).
function dailyHoursCap(nutrition) {
  const { nutritionHoursCapLow: lo, nutritionHoursCapHigh: hi, dailyHoursFloor: floor, dailyHoursCeiling: ceil } = BAL;
  if (nutrition <= lo) return floor;
  if (nutrition >= hi) return ceil;
  const t = (nutrition - lo) / (hi - lo);
  return Math.round(floor + t * (ceil - floor));
}

/* ---------------------------------------------------------------------- */
/* Employment: Work -> Pro, with Unemployed as the failure/recovery state  */
/* ---------------------------------------------------------------------- */
function checkGoProEligible() {
  return state.leagueTier <= BAL.goProLeagueTier && state.cash >= BAL.goProCash;
}
// Fractional strikes for falling short of the daily requirement — indexed by
// hours actually worked. Hand-tuned, not a formula: a near-miss costs far
// less than blowing the day off entirely, which always costs a full strike.
const WORK_STRIKE_TABLE = [1, 0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.1, 0]; // Work: 9h/day required
const PRO_STRIKE_TABLE = [1, 0.7, 0.4, 0.2, 0.1, 0]; // Pro Duties: 5h/day required
function strikeAmount(hours, isPro) {
  const table = isPro ? PRO_STRIKE_TABLE : WORK_STRIKE_TABLE;
  return table[clamp(Math.floor(hours), 0, table.length - 1)];
}
function employmentStrikeCount() {
  const raw = state.employment.strikes.reduce((sum, s) => sum + (s.amount != null ? s.amount : 1), 0);
  return Math.round(raw * 10) / 10;
}
// Lives remaining before firing — the same number as the strike count, just
// framed as what's left (full at the start) instead of what's been lost.
function livesRemaining() {
  return Math.max(0, Math.round((BAL.strikesToFire - employmentStrikeCount()) * 10) / 10);
}
// A synthetic 0-100 "job security" readout for the Work/Pro Duties bar —
// full with no strikes, shrinking toward 0 as they stack up.
function jobSecurityPct() {
  return clamp((livesRemaining() / BAL.strikesToFire) * 100, 0, 100);
}
// Today's projected hit to job security if the day resolved right now at the
// current Work/Pro Duties slider position — feeds the bar's preview overlay
// so a shortfall visibly shows how much of today's lives it would cost.
function jobSecurityPreviewPct(hours, isPro) {
  const required = isPro ? BAL.proDutyHoursRequired : BAL.workHoursRequired;
  const loss = hours >= required ? 0 : strikeAmount(hours, isPro);
  const previewLives = clamp(livesRemaining() - loss, 0, BAL.strikesToFire);
  return clamp((previewLives / BAL.strikesToFire) * 100, 0, 100);
}
function workMaxHours(st = state.employment.status) {
  if (st === "pro") return 12; // 5 required + headroom to push a technique
  if (st === "unemployed") return 16; // no requirement, just a generous daily ceiling
  return BAL.workHoursRequired; // Employed: no benefit to going beyond the requirement
}
function techniqueStatusText() {
  const q = state.employment.techniqueQueue;
  if (!q.length) return "caught up";
  const cur = q[0];
  const behind = q.length > 1 ? ` (+${q.length - 1} queued)` : "";
  return `${cur.name} ${fmt(cur.hoursDone)}/${cur.hoursNeeded}h${behind}`;
}

/* ---------------------------------------------------------------------- */
/* Opponent name generation                                               */
/* ---------------------------------------------------------------------- */
// Hand-picked rival tags: Excel functions, features and puns paired with
// esports-style words. No word appears in more than one name, so a league
// table never reads as twenty variations on "Pivot". Each career draws 199
// of these at random (see generateRivalRoster).
const RIVAL_NAMES = [
  "TrimReaper", "CtrlFreak", "ShiftHappens", "RoundRobin", "CellShock", "RangeRover", "AbsoluteUnit",
  "PasteSpecialist", "FormatPainter", "EscapeArtist", "DivByZero", "InfiniteLoop", "NameBoxNebula",
  "TurnTheTables", "FormulaOne", "SumOfAllFears", "DAYDreamer", "HourGlass", "MinuteMan", "SecondSight",
  "TRUEGrit", "FALSEProphet", "DROPKick", "SmallFry", "BlankCanvas", "Excelsior", "TooltipSorcerer",
  "PVPredator", "NPERNoScope", "DDBDoubleTap", "REPTRepeater", "GridlockGorilla", "TracePrecedents",
  "VLOOKUPVandal", "XLOOKUPXenon", "HLOOKUPHawk", "SUMIFSensei", "COUNTIFCountess", "IFERRORImp",
  "LAMBDALynx", "FILTERFalcon", "SORTBYSamurai", "UNIQUEUnicorn", "SEQUENCESerpent", "TEXTJOINTitan",
  "CONCATCobra", "PROPERPaladin", "LEFTLeopard", "MIDMonk", "LENLlama", "RANDBandit", "OFFSETOutlaw",
  "IndirectIbis", "CHOOSEChimera", "SwitchSwift", "MAXIFSMaverick", "MINIFSMinotaur", "IFSIfrit",
  "XORXerxes", "NOTNecromancer", "ANDAndroid", "ISNUMBERNarwhal", "ISEVENIbex", "ADDRESSAdmiral",
  "INDEXAvalanche", "FINDFirefly", "SEARCHSentinel", "EXACTEnigma", "VALUEVanguard", "CLEANTempest",
  "SUBSTITUTESquid", "CODECobalt", "UNICHARUmbra", "NPVNomad", "IRRInferno", "PMTPanther", "XIRRXenomorph",
  "MIRRMirage", "FVFireball", "RATERampage", "SLNSlingshot", "MedianMantis", "ModeMongoose", "STDEVStallion",
  "PRODUCTSultan", "RANKRaven", "LARGELegend", "FrequencyFox", "QuantumQuartile", "PercentilePirate",
  "CORRELCougar", "LINESTRaccoon", "ForecastFerret", "TrendlineTiger", "OutlierOcelot", "VarianceVulcan",
  "TodayTyrant", "EOMonthEagle", "WorkdayWarlock", "DATEDIFDruid", "WeekdayWarrior", "YEARFRACYak",
  "WEEKNUMMoth", "EDATEEcho", "TRANSPOSETroll", "MMULTMammoth", "BYCOLBerserker", "MAPMagpie",
  "REDUCERaider", "SCANScavenger", "LETMaestro", "TAKETrickshot", "VSTACKViking", "HSTACKHurricane",
  "TOCOLToucan", "EXPANDElk", "PivotPoltergeist", "MacroMauler", "SlicerSpartan", "PowerQueryPilot",
  "FlashFillFinch", "AutoFitAce", "SparklineSprite", "GridlineGhost", "RibbonRogue", "SheetStorm",
  "WorkbookWizard", "TabTamer", "ColumnCrusher", "RowRider", "HeaderHunter", "FooterFiend", "FreezePaneYeti",
  "ChartChampion", "GanttGoblin", "WaterfallWarden", "HistogramHydra", "ScatterShark", "HeatmapHornet",
  "DashboardDuke", "TemplateTemplar", "HyperlinkHero", "GoalSeekGoose", "SolverSage", "ScenarioScout",
  "SpillSpecter", "ArrayArchitect", "VolatileViper", "CircularOrbit", "RelativeRonin", "ConditionalCrane",
  "GradientGryphon", "MergeMercenary", "WrapWraith", "BorderBison", "BoldBadger", "ItalicIguana",
  "CalibriCaptain", "ArialArcher", "UndoUndertaker", "RedoRaptor", "AltAvenger", "F4Fury", "F9Phoenix",
  "EnterEnforcer", "ClipboardKraken", "CursorCorsair", "CrosshairHeron", "AmpersandAsp", "NANinja",
  "DebugDemon", "ImmediateBear", "ModuleMarauder", "VariantVoyager", "WatchWindowWasp", "EvaluateEnvoy",
  "AuditAssassin", "ProtectedPython", "PadlockPuma", "PasswordPuffin", "HiddenHyena", "GroupGuru",
  "OutlineOtter", "SubtotalSphinx", "ConsolidateCrow", "ValidationVole", "PicklistDingo", "CheckboxCheetah",
  "ConnectionQueen", "RefreshRhino", "DAXDragon", "DelimiterDynamo", "CSVCyclone", "XLSXExile",
  "StatusBarSherpa", "ZoomZealot", "PageBreakPanda", "PrintAreaPelican", "LandscapeLion", "DataDaemon",
  "A1Alpha", "R1C1Rebel", "XFDFrontier", "BandedTrooper", "TimelineTerrapin", "IconSetImpala",
  "ColorScaleCoyote", "TopTenTurtle", "StructuredStork", "CalcChainCaracal", "BackstageMoose",
  "IterationIris", "CubeCrusader",
];
function shuffledRivalNames() {
  const names = RIVAL_NAMES.slice();
  for (let i = names.length - 1; i > 0; i--) {
    const j = randInt(0, i);
    [names[i], names[j]] = [names[j], names[i]];
  }
  return names;
}

// Saves from before v4.33.0 used names built from 28 + 28 repeating parts.
// Give every rival a name from the new list, and carry it through every
// stored copy of the old one (fixtures, standings snapshots, brackets,
// champions, season results) and the log, so nothing is left pointing at a
// name that no longer exists. Ratings, records and league places are kept.
function renameRivalsToCurrentList(s) {
  const pool = shuffledRivalNames();
  const map = new Map();
  s.rivals.forEach((r, i) => {
    const fresh = pool[i % pool.length];
    map.set(r.name, fresh);
    r.name = fresh;
  });
  const NAME_KEYS = ["name", "opponent", "opponentName"];
  const walk = (node) => {
    if (!node || typeof node !== "object") return;
    if (Array.isArray(node)) return node.forEach(walk);
    if (!node.isPlayer) {
      NAME_KEYS.forEach((k) => {
        if (typeof node[k] === "string" && map.has(node[k])) node[k] = map.get(node[k]);
      });
    }
    Object.keys(node).forEach((k) => walk(node[k]));
  };
  Object.keys(s).forEach((k) => {
    if (k !== "rivals" && k !== "logEntries") walk(s[k]);
  });
  if (Array.isArray(s.logEntries) && map.size) {
    const olds = [...map.keys()].sort((x, y) => y.length - x.length).map((n) => n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
    const re = new RegExp(`(?<![A-Za-z0-9])(${olds.join("|")})(?![A-Za-z0-9])`, "g");
    s.logEntries.forEach((e) => {
      if (typeof e.html === "string") e.html = e.html.replace(re, (m) => map.get(m) || m);
    });
  }
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
  const names = shuffledRivalNames();
  const rivals = [];
  let id = 0;
  // Whichever tier the player occupies gets 39 rivals instead of 40, leaving
  // the player's slot open; every other tier is a full 40.
  const countsByTier = { 1: 40, 2: 40, 3: 40, 4: 40, 5: 40 };
  countsByTier[playerTier] = 39;
  for (let tier = 1; tier <= 5; tier++) {
    const [lo, hi] = LEAGUE_RATING_BANDS[tier];
    for (let i = 0; i < countsByTier[tier]; i++) {
      const name = names[id % names.length];
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

// Standard circle-method round-robin: N members (even), N-1 rounds, every
// member paired with every other exactly once, once per round. Members mix
// numeric rival ids and the "player" sentinel for whichever tier they're in.
function generateRoundRobinSchedule(memberIds) {
  const n = memberIds.length;
  const rounds = [];
  const fixed = memberIds[0];
  const rotating = memberIds.slice(1);
  for (let r = 0; r < n - 1; r++) {
    const ring = [fixed, ...rotating];
    const pairs = [];
    for (let i = 0; i < n / 2; i++) {
      pairs.push([ring[i], ring[n - 1 - i]]);
    }
    rounds.push(pairs);
    rotating.unshift(rotating.pop());
  }
  return rounds;
}

// Builds a fresh round-robin schedule and a zeroed points table for every
// league tier at once (not just the player's) — this is what makes every
// league's standings live and visible from round 1, rather than only
// computable in one bulk pass once a season ends. Also derives the
// player's own 39-fixture schedule from their tier's round-robin so the
// rest of the game (getNextMatchInfo, processDayEnd, etc.) is unaffected.
function buildSeasonLeagueData(rivals, leagueTier) {
  const roundRobins = {};
  const points = {};
  for (let tier = 1; tier <= BAL.leagueCount; tier++) {
    const memberIds = rivals.filter((r) => r.league === tier).map((r) => r.id);
    if (tier === leagueTier) memberIds.push("player");
    roundRobins[tier] = generateRoundRobinSchedule(memberIds);
    points[tier] = {};
    memberIds.forEach((id) => {
      points[tier][id] = 0;
    });
  }
  const schedule = roundRobins[leagueTier].map((round) => {
    const pair = round.find((p) => p.includes("player"));
    const rivalId = pair[0] === "player" ? pair[1] : pair[0];
    const rival = rivals.find((r) => r.id === rivalId);
    return { rivalId, name: rival.name, rating: rival.rating, played: false, win: null };
  });
  return { roundRobins, points, schedule };
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
  const leagueData = buildSeasonLeagueData(rivals, leagueTier);
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
    rivalNamesVersion: 2, // 2 = names from RIVAL_NAMES (see renameRivalsToCurrentList)
    matchHistory: { mine: [], league: {} }, // see recordLeagueResult / recordMyMatch
    staff: { hired: {} }, // { [upgradeKey]: level } hired for this week only
    leagueStandings: { 1: null, 2: null, 3: null, 4: null, 5: null }, // last fully completed season per tier
    leagueRoundRobins: leagueData.roundRobins, // this season's full fixture list per tier, all 5 at once
    leaguePoints: leagueData.points, // this season's live running points per tier, updated every round
    schedule: leagueData.schedule,
    seasonResults: [], // { opponent, win } per completed regular-season round
    lastStandings: null,
    lastPlayerPosition: null,
    lastStandingsTier: null, // which league lastStandings was for (may differ from leagueTier after a promotion/relegation)
    offseasonDays: BAL.offseasonDays,
    offseasonReason: null, // "missed" | "eliminated" (both training camp) | "playoffs" (champion's break) — set when entering offseason
    playoff: null, // { stage, currentRound, eliminated, champion, playerSeed, tier }
    playoffChampions: null, // { [tier]: { name, rivalId, isPlayer } } — last completed knockout per tier
    seasonMovementApplied: false, // promotion/relegation already run for the season just finished
    cash: 100,
    rank: 480,
    peakRank: 480,
    wins: 0,
    losses: 0,
    yearCashFlow: { workPay: 0, matchCash: 0, expenses: 0 }, // resets every year; summarized at year-end
    stats: {
      skills: Object.fromEntries(SKILL_KEYS.map((k) => [k, 8])),
      phys: 65,
      rest: 100, // higher = better; drains when sleep < idealSleep
      composure: 92, // higher = better; drains when relax < relaxComposureThreshold
      nutrition: 75, // higher = better; drains below 1h/day, and governs the daily hours budget
    },
    allocation: {
      skills: Object.fromEntries(SKILL_KEYS.map((k) => [k, 0])),
      exercise: 1,
      sleep: 7,
      relax: 2,
      nutrition: 1,
      work: BAL.workHoursRequired,
    },
    employment: {
      status: "employed", // "employed" | "unemployed" | "pro"
      strikes: [], // [{day}] — each expires strikeWindowDays after it's earned
      jobSearchHours: 0, // Unemployed: cumulative progress toward jobSearchHoursNeeded
      jobSearchHoursNeeded: null, // Unemployed: this search's target, rolled from jobSearchHoursRange when the job is lost
      techniqueQueue: [], // Pro: [{name, hoursNeeded, hoursDone}], FIFO — only the front is "in progress"
      techniqueDayCounter: 0, // days elapsed toward the next technique; only ticks while Pro
      workPay: BAL.workPayMin, // current Employed daily pay — rises with tenure, resets on firing
      proPay: BAL.proPayMin, // current Pro daily pay — rises with tenure, resets on being dropped
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
    const leagueData = buildSeasonLeagueData(rivals, tier);
    parsed.schedule = leagueData.schedule;
    parsed.leagueRoundRobins = leagueData.roundRobins;
    parsed.leaguePoints = leagueData.points;
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
    const oldStats = parsed.stats || { excel: 8, phys: 65, sleepDebt: 0, stress: 8 };
    const oldExcel = oldStats.excel !== undefined ? oldStats.excel : 8;
    // Every skill starts fresh (level 0) in the new per-skill Coaching Shop,
    // so the migrated value can't exceed the shop-base/league-cap gate any
    // veteran's real progress would still have to clear.
    const migratedSkillCap = Math.min(BAL.skillShopCapBase, LEAGUE_SKILL_CAP[parsed.peakLeagueTier] || LEAGUE_SKILL_CAP[5]);
    const seededSkill = clamp(oldExcel, 0, migratedSkillCap);
    parsed.stats = {
      skills: Object.fromEntries(SKILL_KEYS.map((k) => [k, seededSkill])),
      phys: oldStats.phys !== undefined ? oldStats.phys : 65,
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

  // v4.2.0 removed Energy — Rest already fully covered training
  // effectiveness — and replaced it with Nutrition, which instead governs
  // the day's total hours budget. There's no meaningful way to derive
  // Nutrition from the old Energy number since they measure different
  // things, so any save missing it just gets the same starting value a
  // fresh career gets.
  if (parsed.stats && parsed.stats.nutrition === undefined) {
    delete parsed.stats.energy;
    parsed.stats.nutrition = 75;
  }
  if (parsed.allocation && parsed.allocation.nutrition === undefined) {
    parsed.allocation.nutrition = 1;
  }

  // v4.4.0 added mandatory employment (Work -> Pro, with Unemployed as the
  // recovery state). Every existing career just starts Employed — there's
  // no prior data to derive a status from, and starting employed means no
  // save is retroactively punished by a surprise strike. Any save already
  // meeting the go-pro thresholds is promoted automatically the next time
  // a day resolves, same as a freshly-qualifying career.
  if (!parsed.employment) {
    parsed.employment = {
      status: "employed",
      strikes: [],
      jobSearchHours: 0,
      techniqueQueue: [],
      techniqueDayCounter: 0,
    };
  }
  if (parsed.allocation && parsed.allocation.work === undefined) {
    parsed.allocation.work = BAL.workHoursRequired;
  }

  // v4.9.0 replaced flat pay with per-role rates that rise with tenure and
  // reset on job loss. Any save missing them starts both at their minimum —
  // under-crediting tenure already built up is the safe direction (never
  // retroactively grants raises that weren't earned under the old flat rate).
  if (parsed.employment && parsed.employment.workPay === undefined) {
    parsed.employment.workPay = BAL.workPayMin;
  }
  if (parsed.employment && parsed.employment.proPay === undefined) {
    parsed.employment.proPay = BAL.proPayMin;
  }

  // v4.5.0 replaced "simulate every league in one bulk pass at season end"
  // with a real week-by-week round-robin for all 5 tiers, so standings are
  // live and visible from round 1 instead of hidden for an entire 39-round
  // season. Any save missing that structure gets a fresh one for its
  // current league: there's no way to recover what the old random-shuffle
  // schedule would have played next, so only the unplayed tail of the
  // fixture list is replaced — already-played rounds and their results
  // stay exactly as recorded. The player's own points so far are credited
  // from their real record; other members start this season's live table
  // at 0, since the old model never tracked interim standings for anyone
  // but the player.
  if (!parsed.leagueRoundRobins || !parsed.leaguePoints) {
    const tier = parsed.leagueTier;
    const leagueData = buildSeasonLeagueData(parsed.rivals, tier);
    parsed.leagueRoundRobins = leagueData.roundRobins;
    parsed.leaguePoints = leagueData.points;

    const playedCount = parsed.roundIndex || 0;
    const oldSchedule = parsed.schedule || [];
    parsed.schedule = oldSchedule.slice(0, playedCount).concat(leagueData.schedule.slice(playedCount));

    const playerWinsSoFar = (parsed.seasonResults || []).filter((r) => r.win).length;
    parsed.leaguePoints[tier].player = playerWinsSoFar * 3;
  }

  // Promotion now waits for the playoffs (champion + top 3). A save already
  // in the playoffs or offseason when that shipped had its movement applied
  // at season end under the old table-only rule — mark it done so the end
  // of its playoffs doesn't move everyone a second time.
  if (parsed.seasonMovementApplied === undefined) {
    parsed.seasonMovementApplied = parsed.seasonPhase === "playoffs" || parsed.seasonPhase === "offseason";
  }
  if (parsed.playoffChampions === undefined) parsed.playoffChampions = null;

  // Injured saves from before the Gym slider was cleared on injury still
  // have hours locked in it — free them.
  if (parsed.injury && parsed.injury.active && parsed.allocation) parsed.allocation.exercise = 0;
  // The Gym slider's max dropped from 12h to 8h.
  if (parsed.allocation) parsed.allocation.exercise = Math.min(parsed.allocation.exercise || 0, BAL.gymMaxHours);
  // Job search targets became a random 10-40h rolled on job loss; a save
  // that's mid-search keeps the fixed 30h it started with.
  if (parsed.employment && parsed.employment.jobSearchHoursNeeded === undefined) {
    parsed.employment.jobSearchHoursNeeded = parsed.employment.status === "unemployed" ? 30 : null;
  }
  // A rehire used to keep Job Search's up-to-16h on a Work slider that caps
  // at 9, counting phantom hours against the day (and blocking End Day).
  if (parsed.allocation && parsed.employment) {
    parsed.allocation.work = Math.min(parsed.allocation.work || 0, workMaxHours(parsed.employment.status));
  }

  // In-season, skillCycleDay and phaseDay always advance together — both
  // start at 0 with the regular season, same every-day increment, same
  // reset-at-7. A save written while a prior bug let them drift (the weekly
  // skill reveal landing a day before that week's match) gets them
  // realigned here instead of carrying the offset forever.
  if ((parsed.seasonPhase === "regular" || parsed.seasonPhase === "playoffs") && parsed.skillCycleDay !== parsed.phaseDay) {
    parsed.skillCycleDay = parsed.phaseDay;
  }

  // v4.36.0: staff are hired weekly (owned levels stay unlocked).
  if (!parsed.staff) parsed.staff = { hired: {} };

  // v4.35.0: match history. Earlier matches this season only kept opponent
  // and result, so they come in as simple rows; league-wide results start now.
  if (!parsed.matchHistory) {
    const tier = parsed.lastStandingsTier || parsed.leagueTier;
    parsed.matchHistory = {
      mine: (parsed.seasonResults || []).map((r, i) => ({
        legacy: true,
        year: parsed.year,
        tier,
        roundLabel: `Round ${i + 1}/${BAL.seasonRounds}`,
        opponentName: r.opponent,
        win: r.win,
      })),
      league: {},
    };
  }

  // v4.33.0: rivals get names from the hand-picked list (no repeated words).
  if (parsed.rivalNamesVersion !== 2 && Array.isArray(parsed.rivals) && parsed.rivals.length) {
    renameRivalsToCurrentList(parsed);
    parsed.rivalNamesVersion = 2;
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

// Only the last 200 entries are ever shown; without a cap the saved log
// grows ~180KB a year and eventually fills localStorage.
const LOG_KEEP = 300;

function saveState() {
  if (!STORAGE_OK) return false;
  if (state.logEntries.length > LOG_KEEP) state.logEntries = state.logEntries.slice(-LOG_KEEP);
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

// Short percentage for tight rows: one decimal under 10% ("2.5%"), whole
// numbers from 10% up ("13%"), so it never grows past 4 characters.
function fmtPct(fraction) {
  const pct = fraction * 100;
  return pct < 10 ? `${+pct.toFixed(1)}%` : `${Math.round(pct)}%`;
}

function fmtMoney(n) {
  return `${n < 0 ? "−" : ""}$${fmt(Math.abs(n))}`;
}

// The player's name is the only user-typed text that reaches innerHTML.
function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
}

function fmt1(n) {
  return (Math.round(n * 10) / 10).toFixed(1);
}

function fmtSigned(n, decimals = 1) {
  const r = Number(n.toFixed(decimals));
  return (r >= 0 ? "+" : "") + r;
}

/* ---------------------------------------------------------------------- */
/* Day resolution — stats only                                            */
/* ---------------------------------------------------------------------- */
// Pure, non-mutating: computes what a day WOULD do to every stat given a
// (allocation, stats) pair, with no randomness (injury is a separate roll,
// kept out of this function — see resolveDay()). Shared by the real
// end-of-day resolution and the live "tomorrow" preview on each slider, so
// the two can never drift apart.
function computeDayResult(allocation, stats, injuryActive, burnoutActive) {
  const a = allocation;
  const s = stats;
  const physioEff = upgradeEffect("physio");
  const sleepAppEff = upgradeEffect("sleepApp");
  const nutritionistEff = upgradeEffect("nutritionist");
  const meditationEff = upgradeEffect("meditation");
  const recoveryEff = upgradeEffect("recovery");

  const exerciseH = injuryActive ? 0 : a.exercise;
  const sleepH = a.sleep;
  const relaxH = a.relax;

  const focusMult = burnoutActive ? BAL.burnoutEffectivenessMult : 1;
  const physMult = physSynergy(s.phys);
  const restMult = restTrainingMultiplier(s.rest);

  // ---- The 7 case specialties ----
  let totalSkillH = 0;
  const trainable = trainableSkills();
  const skills = {};
  const skillWasAtCap = {};
  SKILL_KEYS.forEach((key) => {
    const isActive = trainable.includes(key);
    const hours = isActive ? a.skills[key] || 0 : 0;
    const cap = skillCap(key);
    skillWasAtCap[key] = s.skills[key] >= cap - 0.05;
    const coachEff = upgradeEffect(coachKey(key));
    const coachMult = coachEff ? 1 + coachEff.bonus : 1.0;
    totalSkillH += hours;

    // Above this week's ceiling (e.g. 57 with no coach, ceiling 50): an
    // hour or more holds it steady, less lets it rust — it's never cut down
    // to the ceiling just because the coach isn't on hand this week.
    if (hours >= BAL.skillDecayThresholdHours) {
      const eff = effectiveHours(hours);
      const gain = eff * BAL.skillGainBase * focusMult * physMult * restMult * skillDiminish(s.skills[key]) * coachMult;
      skills[key] = s.skills[key] >= cap ? s.skills[key] : Math.min(s.skills[key] + gain, cap);
    } else {
      const rust = Math.min(0.15, s.skills[key] * 0.003);
      skills[key] = Math.max(0, s.skills[key] - rust);
    }
  });

  // ---- Physical health ----
  const pCap = physCap();
  const physWasAtCap = s.phys >= pCap - 0.05;
  const physioMult = physioEff ? 1 + physioEff.exerciseBonus : 1.0;
  const exerciseGain = effectiveHours(exerciseH) * BAL.exerciseGainBase * physioMult * restMult * physDiminish(s.phys);
  const physDecayFromRest = (100 - s.rest) * 0.045;
  const detrainMult = recoveryEff ? recoveryEff.detrainMult : 1.0;
  const detrain = exerciseH < BAL.detrainThresholdHours ? BAL.detrainDecay * detrainMult : 0;
  const physDelta = exerciseGain - physDecayFromRest - detrain;
  // Same rule as skills: gains stop at the ceiling, but Health already above
  // it (physio not hired this week) only goes down if the day's net is down.
  const phys = physDelta > 0 ? Math.min(s.phys + physDelta, Math.max(pCap, s.phys)) : Math.max(0, s.phys + physDelta);

  // ---- Nutrition: governs tomorrow's total hours budget (dailyHoursCap),
  // not training effectiveness — decays below the 1h threshold same as
  // every other trainable stat, grows with diminishing returns otherwise.
  const nutritionH = a.nutrition || 0;
  const nutritionGainMult = nutritionistEff ? nutritionistEff.nutritionGainMult : 1.0;
  const nutritionDecayMult = nutritionistEff ? nutritionistEff.nutritionDecayMult : 1.0;
  let nutrition;
  if (nutritionH >= BAL.skillDecayThresholdHours) {
    const gain = effectiveHours(nutritionH) * BAL.nutritionGainBase * skillDiminish(s.nutrition) * nutritionGainMult;
    nutrition = clamp(s.nutrition + gain, 0, 100);
  } else {
    nutrition = clamp(s.nutrition - BAL.nutritionDecayFlat * nutritionDecayMult, 0, 100);
  }

  // ---- Rest (renamed/inverted Sleep Debt: higher = better) ----
  const sleepAppMult = sleepAppEff ? sleepAppEff.sleepDebtMult : 1.0;
  const restDelta =
    sleepH < BAL.idealSleep ? -((BAL.idealSleep - sleepH) * 1.3 * sleepAppMult) : Math.min(sleepH - BAL.idealSleep, 3) * 1.4;
  const rest = clamp(s.rest + restDelta, 0, 100);

  // ---- Composure (renamed/inverted Stress: higher = better) ----
  const meditationMult = meditationEff ? meditationEff.reliefMult : 1.0;
  const composureLoad = (totalSkillH + exerciseH) * BAL.composureLoadPerHour;
  const composureRelief = relaxH * BAL.relaxComposureRelief * meditationMult + (sleepH >= BAL.idealSleep ? BAL.restGoodSleepBonus : 0);
  const composurePenaltyFromRest = (100 - s.rest) * 0.1;
  const composureDelta = composureRelief - composureLoad - composurePenaltyFromRest;
  const composure = clamp(s.composure + composureDelta, 0, 100);

  return {
    skills,
    phys,
    rest,
    composure,
    nutrition,
    exerciseH,
    sleepH,
    relaxH,
    totalSkillH,
    focusMult,
    physMult,
    restDelta,
    composureDelta,
    composureLoad,
    composurePenaltyFromRest,
    physDecayFromRest,
    skillWasAtCap,
    physWasAtCap,
    pCap,
  };
}

// Non-mutating projection of what tomorrow's stats would be if today
// resolved exactly as currently planned — drives the "tomorrow" preview
// overlay on each stat bar. Injury is stochastic and deliberately not
// previewed (see computeDayResult).
function previewTomorrow() {
  return computeDayResult(state.allocation, state.stats, state.injury.active, state.burnout.active);
}

function resolveDay() {
  const a = state.allocation;
  const s = state.stats;
  const recoveryEff = upgradeEffect("recovery");
  const events = [];
  // An injury picked up today starts counting down tomorrow — otherwise a
  // "3 day" injury only locks the Gym for 2, and a 1-day one for none.
  const wasInjured = state.injury.active;

  const result = computeDayResult(a, s, state.injury.active, state.burnout.active);
  const trainable = trainableSkills();

  // ---- The 7 case specialties ----
  SKILL_KEYS.forEach((key) => {
    const meta = skillMeta(key);
    const isActive = trainable.includes(key);
    const hours = isActive ? a.skills[key] || 0 : 0;
    const cap = skillCap(key);
    const before = s.skills[key];
    s.skills[key] = result.skills[key];
    const delta = result.skills[key] - before;

    if (hours >= BAL.skillDecayThresholdHours) {
      let note = "";
      if (result.skillWasAtCap[key] && cap < 100) note = ` (capped at ${cap})`;
      else if (result.focusMult < 1) note = " (burnout hurt your training)";
      else if (result.physMult < 0.75) note = " (low Health capped your gains)";
      events.push({ type: delta > 0.5 ? "good" : "neutral", text: `${meta.icon} ${meta.name} (${hours}h): ${fmtSigned(delta)} skill${note}` });
    } else if (delta < 0 && isActive) {
      events.push({ type: "bad", text: `${meta.icon} No ${meta.name} training: skill rusted slightly (${fmtSigned(delta)})` });
    }
  });

  // ---- Physical health ----
  const physBefore = s.phys;
  s.phys = result.phys;
  const physDelta = result.phys - physBefore;
  if (result.exerciseH > 0) {
    let note = result.physWasAtCap && result.pCap < 100 ? ` (capped at ${result.pCap} — upgrade Sports Physio for a higher ceiling)` : "";
    events.push({ type: physDelta > 0 ? "good" : "neutral", text: `🏃 Gym (${result.exerciseH}h): Health ${fmtSigned(physDelta)}${note}` });
  }
  if (result.physDecayFromRest > 1) {
    events.push({ type: "bad", text: `🌙 Poor Rest is wearing down your body (Health ${fmtSigned(-result.physDecayFromRest)})` });
  }

  // ---- Injury roll (only if not already injured) ----
  if (!state.injury.active && result.exerciseH > 0) {
    if (Math.random() < injuryChance(result.exerciseH)) {
      const loss = randInt(BAL.injuryPhysLoss[0], BAL.injuryPhysLoss[1]);
      const daysReduce = recoveryEff ? recoveryEff.injuryDaysReduce : 0;
      const days = Math.max(1, randInt(BAL.injuryDaysRange[0], BAL.injuryDaysRange[1]) - daysReduce);
      s.phys = clamp(s.phys - loss, 0, result.pCap);
      state.injury = { active: true, daysLeft: days };
      // The Gym slider locks while injured — free its hours rather than
      // leaving them stuck in the day's budget doing nothing.
      state.allocation.exercise = 0;
      events.push({ type: "bad", text: `🤕 Gym injury! -${loss} Health. Gym locked for ${days} ${days === 1 ? "day" : "days"}.` });
    }
  }

  // ---- Nutrition ----
  const nutritionBefore = s.nutrition;
  s.nutrition = result.nutrition;
  const nutritionH = a.nutrition || 0;
  if (nutritionH >= BAL.skillDecayThresholdHours) {
    events.push({ type: result.nutrition - nutritionBefore > 0.5 ? "good" : "neutral", text: `🥗 Food (${nutritionH}h): Nutrition ${fmtSigned(result.nutrition - nutritionBefore)}` });
  } else {
    events.push({ type: "bad", text: `🥗 Skipped Food: Nutrition fell to ${fmt(s.nutrition)} — tomorrow's hours may shrink` });
  }

  // ---- Rest ----
  s.rest = result.rest;
  if (result.sleepH < 6) {
    events.push({ type: "bad", text: `🌙 Only slept ${result.sleepH}h: Rest dropping fast (${fmtSigned(result.restDelta)})` });
  } else if (result.sleepH >= BAL.idealSleep) {
    events.push({ type: "good", text: `🌙 Slept ${result.sleepH}h: well rested (Rest ${fmtSigned(result.restDelta)})` });
  }

  // ---- Composure ----
  s.composure = result.composure;
  if (result.relaxH < BAL.relaxComposureThreshold && result.composureLoad > 5) {
    events.push({ type: "bad", text: `🔥 Under ${BAL.relaxComposureThreshold}h Relax: Composure fell to ${fmt(s.composure)}` });
  } else if (result.relaxH >= BAL.relaxComposureThreshold) {
    events.push({ type: "good", text: `🎮 Relaxed ${result.relaxH}h: Composure ${fmtSigned(result.composureDelta)}` });
  }

  // ---- Burnout state transitions ----
  if (!state.burnout.active && s.composure <= BAL.burnoutComposureThreshold) {
    state.burnout = { active: true, daysLeft: 1 };
    events.push({ type: "bad", text: `⚠️ BURNOUT! You've pushed too hard with too little rest. Training and the Gym are far less effective until your Composure recovers — Relax more.` });
  } else if (state.burnout.active && s.composure >= BAL.burnoutRecoverThreshold) {
    state.burnout = { active: false, daysLeft: 0 };
    events.push({ type: "good", text: `✅ Recovered from burnout. You're focused again.` });
  }

  // ---- Injury countdown ----
  if (wasInjured && state.injury.active) {
    state.injury.daysLeft -= 1;
    if (state.injury.daysLeft <= 0) {
      state.injury = { active: false, daysLeft: 0 };
      events.push({ type: "good", text: `✅ Injury healed. The Gym is open again.` });
    }
  }

  return events;
}

// This job search's hour target, rolled when the job was lost. Saves that
// were already unemployed before targets were randomised keep the fixed
// 30h they were promised.
function jobSearchTarget(emp = state.employment) {
  return emp.jobSearchHoursNeeded || 30;
}

// Resolves one day of Work / Job Search / Pro Duties. Kept separate from
// resolveDay() since it mutates state.employment (a whole status machine,
// not a decaying stat) and generates its own log events.
function resolveEmploymentDay() {
  const emp = state.employment;
  const workH = state.allocation.work || 0;
  const events = [];

  // Cost of living — charged every day no matter what, employed or not, so
  // losing your job actually drains cash instead of just stalling income.
  state.cash -= BAL.dailyExpenses;
  state.yearCashFlow.expenses += BAL.dailyExpenses;

  emp.strikes = emp.strikes.filter((s) => state.day - s.day < BAL.strikeWindowDays);

  if (emp.status === "unemployed") {
    emp.jobSearchHours += workH;
    if (emp.jobSearchHours >= jobSearchTarget()) {
      emp.status = "employed";
      emp.jobSearchHours = 0;
      emp.jobSearchHoursNeeded = null;
      // Job Search allows up to 16h but Work caps at its 9h requirement, so
      // reset to exactly that rather than leaving extra hours stranded.
      state.allocation.work = BAL.workHoursRequired;
      events.push({ type: "good", text: `💼 Found a new job! Work resumes at ${BAL.workHoursRequired}h/day.` });
    } else if (workH > 0) {
      events.push({ type: "neutral", text: `🔍 Job search: ${fmt(emp.jobSearchHours)}/${jobSearchTarget()}h` });
    }
  } else {
    const isPro = emp.status === "pro";
    const required = isPro ? BAL.proDutyHoursRequired : BAL.workHoursRequired;

    // Pay is tied to still HAVING the job, not to hitting the exact hour
    // target every single day — falling short costs a chance (below), not
    // income. Only actually losing the job/Pro role costs you pay.
    const pay = isPro ? emp.proPay : emp.workPay;
    state.cash += pay;
    state.yearCashFlow.workPay += pay;

    if (workH >= required) {
      if (isPro) {
        const extra = workH - required;
        if (extra > 0 && emp.techniqueQueue.length > 0) {
          const current = emp.techniqueQueue[0];
          current.hoursDone = Math.min(current.hoursNeeded, current.hoursDone + extra);
          if (current.hoursDone >= current.hoursNeeded) {
            emp.techniqueQueue.shift();
            events.push({
              type: "good",
              text: `📘 Mastered ${current.name}!${emp.techniqueQueue.length ? " Next up: " + emp.techniqueQueue[0].name + "." : " You're fully caught up."}`,
            });
          }
        }
      }
    } else {
      const amount = strikeAmount(workH, isPro);
      emp.strikes.push({ day: state.day, amount });
      const label = isPro ? "Pro Duties" : "Work";
      const total = employmentStrikeCount();
      const lives = livesRemaining();
      events.push({ type: "bad", text: `💼 Missed your ${label} hours — lost ${fmt1(amount)} ${amount === 1 ? "chance" : "chances"} (${fmt1(lives)}/${fmt1(BAL.strikesToFire)} left).` });
      if (total >= BAL.strikesToFire) {
        emp.status = "unemployed";
        emp.strikes = [];
        emp.jobSearchHours = 0;
        emp.jobSearchHoursNeeded = randInt(BAL.jobSearchHoursRange[0], BAL.jobSearchHoursRange[1]);
        // Seniority is lost, not carried over — whichever role you were just
        // dropped from starts back at its minimum next time you're in it.
        if (isPro) emp.proPay = BAL.proPayMin;
        else emp.workPay = BAL.workPayMin;
        events.push({
          type: "bad",
          text: isPro
            ? `🔥 Out of chances — dropped by your sponsor after too many missed Pro Duties. Time to find a new job — this search will take ${emp.jobSearchHoursNeeded}h.`
            : `🔥 Out of chances — fired after too many missed shifts. Time to find a new job — this search will take ${emp.jobSearchHoursNeeded}h.`,
        });
      }
    }

    if (isPro && emp.status === "pro") {
      emp.techniqueDayCounter += 1;
      if (emp.techniqueDayCounter >= BAL.techniqueIntervalDays) {
        emp.techniqueDayCounter = 0;
        const name = pickTechniqueName();
        const hoursNeeded = randInt(BAL.techniqueMinHours, BAL.techniqueMaxHours);
        emp.techniqueQueue.push({ name, hoursNeeded, hoursDone: 0 });
        events.push({ type: "neutral", text: `📘 New technique to master: ${name} (${hoursNeeded}h).` });
      }
    }
  }

  if (emp.status === "employed" && checkGoProEligible()) {
    emp.status = "pro";
    emp.strikes = [];
    emp.techniqueDayCounter = 0;
    state.allocation.work = BAL.proDutyHoursRequired;
    events.push({
      type: "good",
      text: `🏆 You've gone PRO! Sponsorship replaces your day job — Pro Duties are just ${BAL.proDutyHoursRequired}h/day, but you'll need to keep up with new techniques.`,
    });
  }

  state.allocation.work = Math.min(state.allocation.work || 0, workMaxHours());
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

// Single source of truth for what a match performance score is made of, so
// the result modal can show the breakdown instead of just the final number
// (skill isn't the whole story — Physical Health, Composure and Rest all
// weigh in too).
const PERF_WEIGHTS = { skill: 0.55, phys: 0.25, composure: 0.15, rest: 0.05 };

function performanceBreakdown() {
  const s = state.stats;
  // Low Composure hits you hardest where it matters most: match day. This
  // stacks on top of (doesn't replace) Composure's own weighted contribution
  // below, same as before.
  const composureMult = composureMatchMultiplier(s.composure);
  const skillAvg = activeSkillAverage();
  const skillComponent = skillAvg * composureMult;
  const weighted = {
    skill: skillComponent * PERF_WEIGHTS.skill,
    phys: s.phys * PERF_WEIGHTS.phys,
    composure: s.composure * PERF_WEIGHTS.composure,
    rest: s.rest * PERF_WEIGHTS.rest,
  };
  // Raw inputs, snapshotted for the match result table.
  const raw = {
    skills: state.activeSkills.map((k) => ({ key: k, value: s.skills[k] })),
    skillAvg,
    skillComponent,
    phys: s.phys,
    composure: s.composure,
    rest: s.rest,
  };
  let score = clamp(weighted.skill + weighted.phys + weighted.composure + weighted.rest, 0, 100);
  // Pros who fall behind on current technique compete at a real disadvantage
  // — every not-yet-mastered entry in the queue costs match performance,
  // not training. Capped so a long backlog can't go negative.
  let techniquePenaltyMult = 1;
  if (state.employment.status === "pro" && state.employment.techniqueQueue.length > 0) {
    const penalty = Math.min(1, state.employment.techniqueQueue.length * BAL.techniquePenaltyPerUnmastered);
    techniquePenaltyMult = 1 - penalty;
    score *= techniquePenaltyMult;
  }
  return { weighted, raw, composureMult, techniquePenaltyMult, score };
}

function performanceScore() {
  return performanceBreakdown().score;
}

// Your effective strength for a match right now: base rank, adjusted by
// today's performance. The single source of truth for that adjustment, so
// win-probability and any "what fought" display can never drift apart.
function currentMatchRating() {
  return state.rank + (performanceScore() - 70) * 3;
}

// Resolves one real match for the player against a specific opponent rating.
// Used for both regular-season fixtures and playoff matches.
// Perfect performance (100) gives a strong but no longer overwhelming form
// bonus — enough to matter, not enough to guarantee a win against a tough
// opponent. Shared by the real match resolution and any "what are my
// chances" preview, so the two can never drift apart.
function winProbabilityAgainst(opponentRating) {
  const matchRating = currentMatchRating();
  return 1 / (1 + Math.pow(10, (opponentRating - matchRating) / 400));
}

// Fixtures and bracket slots copy a rival's rating when they're built, but
// the rival keeps playing (and moving) after that — always read it live so
// the odds, the topbar and the league table all agree.
function liveOpponentRating(entry) {
  const rival = entry && entry.rivalId != null ? findRival(entry.rivalId) : null;
  return rival ? rival.rating : entry.rating;
}

// The player's opponent is a real rival too: their rating and lifetime
// record move with this result, same as any rival-vs-rival match.
function recordRivalResultVsPlayer(rivalId, playerWon, playerRating) {
  const rival = findRival(rivalId);
  if (!rival) return;
  const change = eloChange(rival.rating, playerRating, !playerWon);
  rival.rating = clamp(rival.rating + change, BAL.ratingMin, BAL.ratingMax);
  if (playerWon) rival.losses += 1;
  else rival.wins += 1;
}

function resolveMatch(opponentRating) {
  const managerEff = upgradeEffect("manager");
  const rankLossMult = managerEff ? managerEff.rankLossMult : 1.0;
  const cashBonusMult = managerEff ? managerEff.cashBonusMult : 1.0;
  const rankBefore = state.rank;

  const breakdown = performanceBreakdown();
  const perf = breakdown.score;
  const matchRating = currentMatchRating();
  const winProb = winProbabilityAgainst(opponentRating);
  // Each side has its own luck on the day, and the higher match-day rating
  // wins. The two draws are Gumbel-distributed, whose difference is exactly
  // the logistic curve behind winProb — so the odds (and upsets) are the
  // same as a single "roll under your win chance", just told as two scores.
  const yourLuck = matchLuckDraw();
  const oppLuck = matchLuckDraw();
  const yourFinal = matchRating + yourLuck;
  const oppFinal = opponentRating + oppLuck;
  const win = yourFinal > oppFinal || (yourFinal === oppFinal && matchRating >= opponentRating);
  // Rivals have no stats, so their "performance" is cosmetic: a random
  // 55-85 on the same ×3 scale as yours, taken back out of their luck so
  // their match-day rating (and the result) is untouched.
  const oppPerf = randInt(BAL.oppPerfRange[0], BAL.oppPerfRange[1]);
  const K = 24;
  const actual = win ? 1 : 0;
  let ratingChange = Math.round(K * (actual - winProb));
  if (ratingChange < 0) ratingChange = Math.round(ratingChange * rankLossMult);
  state.rank = clamp(state.rank + ratingChange, BAL.ratingMin, BAL.ratingMax);
  state.peakRank = Math.max(state.peakRank, state.rank);

  const cashBefore = state.cash;
  const recordBefore = { wins: state.wins, losses: state.losses };
  const cashReward = Math.round((win ? 150 + state.rank / 10 : 40) * cashBonusMult);
  state.cash += cashReward;
  state.yearCashFlow.matchCash += cashReward;
  if (win) state.wins += 1;
  else state.losses += 1;

  // Display figures for the result screen. Each column's rows are rounded
  // so they add up exactly: luck absorbs the rounding, and the opponent's
  // luck also absorbs their cosmetic performance.
  const yourRatingShown = Math.round(rankBefore);
  const yourPerfShown = Math.round(matchRating) - yourRatingShown;
  const yourFinalShown = Math.round(yourFinal);
  const oppRatingShown = Math.round(opponentRating);
  const oppPerfShown = Math.round((oppPerf - 70) * 3);
  const oppFinalShown = Math.round(oppFinal);

  return {
    win,
    perf,
    breakdown,
    matchRating: Math.round(matchRating),
    opponentRating: oppRatingShown,
    winProb: Math.round(winProb * 100),
    ratingChange,
    ratingChangeRaw: Math.round(K * (actual - winProb)),
    rankLossMult,
    cashBonusMult,
    rankBefore,
    rankAfterMatch: state.rank,
    cashReward,
    cashBefore,
    recordBefore,
    sides: {
      you: { rating: yourRatingShown, perfScore: perf, perfAdj: yourPerfShown, luck: yourFinalShown - yourRatingShown - yourPerfShown, final: yourFinalShown },
      opp: { rating: oppRatingShown, perfScore: oppPerf, perfAdj: oppPerfShown, luck: oppFinalShown - oppRatingShown - oppPerfShown, final: oppFinalShown },
    },
    // From the displayed totals so the verdict always matches them; rounding
    // can make a sub-1 margin look like a tie, which reads as "less than 1".
    margin: Math.abs(yourFinalShown - oppFinalShown),
  };
}

// One side's luck on match day: a Gumbel draw with the Elo scale (400 /
// ln 10), centred on zero. The difference of two such draws is logistic, so
// P(you out-score them) is exactly the Elo win chance. Usually within about
// ±150; now and then a side has an inspired day of +400 or more.
const LUCK_SCALE = 400 / Math.LN10;
const EULER_GAMMA = 0.5772156649;
function matchLuckDraw() {
  const u = clamp(Math.random(), 1e-9, 1 - 1e-9);
  return LUCK_SCALE * (-Math.log(-Math.log(u)) - EULER_GAMMA);
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
// ---- Match history ----
// mine: every match you play, as the full result object (so its result
//   screen can be reopened later). Kept for the whole career.
// league: compact results for every match in all 5 leagues —
//   league[year][tier][roundKey] = [[winnerId, loserId, winnerRating,
//   loserRating], …], roundKey "1".."39" or a playoff stage; you are id -1.
//   Only this season and last are kept.
const HISTORY_PLAYER_ID = -1;
const HISTORY_ROUND_KEYS = Array.from({ length: BAL.seasonRounds }, (_, i) => String(i + 1)).concat(["r16", "qf", "sf", "f"]);
const BRACKET_STAGE_BY_SIZE = { 16: "r16", 8: "qf", 4: "sf", 2: "f" };

function ensureHistory() {
  if (!state.matchHistory) state.matchHistory = { mine: [], league: {} };
  return state.matchHistory;
}

function recordLeagueResult(tier, roundKey, winnerId, loserId, winnerRating, loserRating) {
  if (!tier || !roundKey) return;
  const h = ensureHistory();
  const year = state.year;
  Object.keys(h.league).forEach((y) => {
    if (Number(y) < year - 1) delete h.league[y];
  });
  const byTier = (h.league[year] = h.league[year] || {});
  const rounds = (byTier[tier] = byTier[tier] || {});
  (rounds[roundKey] = rounds[roundKey] || []).push([winnerId, loserId, Math.round(winnerRating), Math.round(loserRating)]);
}

function recordMyMatch(result) {
  const copy = JSON.parse(JSON.stringify(result));
  copy.year = state.year;
  ensureHistory().mine.push(copy);
}

function resolveNpcMatch(rivalA, rivalB) {
  const aWon = simulateNpcMatch(rivalA.rating, rivalB.rating);
  const change = eloChange(rivalA.rating, rivalB.rating, aWon);
  rivalA.rating = clamp(rivalA.rating + change, BAL.ratingMin, BAL.ratingMax);
  rivalB.rating = clamp(rivalB.rating - change, BAL.ratingMin, BAL.ratingMax);
  if (aWon) {
    rivalA.wins += 1;
    rivalB.losses += 1;
  } else {
    rivalB.wins += 1;
    rivalA.losses += 1;
  }
  return aWon;
}

// Resolves every NPC-vs-NPC pair in this tier's round-robin round (the
// player's own pairing, if this is their tier, is resolved separately via
// resolveMatch()) and tallies points — called every week for all 5 tiers in
// lockstep, so every league's table is live all season, never just a single
// bulk computation sprung at season's end.
function resolveLeagueRoundForTier(tier, roundIndex) {
  const round = state.leagueRoundRobins[tier] && state.leagueRoundRobins[tier][roundIndex];
  if (!round) return;
  round.forEach(([a, b]) => {
    if (a === "player" || b === "player") return;
    const rivalA = findRival(a);
    const rivalB = findRival(b);
    if (!rivalA || !rivalB) return;
    const ra = rivalA.rating, rb = rivalB.rating;
    const aWon = resolveNpcMatch(rivalA, rivalB);
    const winnerId = aWon ? a : b;
    recordLeagueResult(tier, String(roundIndex + 1), aWon ? a : b, aWon ? b : a, aWon ? ra : rb, aWon ? rb : ra);
    state.leaguePoints[tier][winnerId] = (state.leaguePoints[tier][winnerId] || 0) + 3;
  });
}

// Reads a tier's current standings straight from the live points table —
// valid at any time, whether mid-season (a true "right now" snapshot) or
// after round 39 (the final result), since points only reset when a new
// season's round-robin is built.
function buildStandingsFromPoints(tier) {
  const points = state.leaguePoints[tier] || {};
  const standings = state.rivals
    .filter((r) => r.league === tier)
    .map((r) => ({ name: r.name, rating: r.rating, points: points[r.id] || 0, isPlayer: false, rivalId: r.id }));
  if (tier === state.leagueTier) {
    standings.push({ name: escapeHtml(state.playerName || "You"), rating: state.rank, points: points.player || 0, isPlayer: true, rivalId: null });
  }
  return standings.sort((a, b) => b.points - a.points || b.rating - a.rating);
}

// The player's table position for the topbar's Leagues button: live during
// the regular season (once a round has been played), the final regular-
// season position through the playoffs and offseason, and nothing in
// preseason when no table exists yet.
function playerTablePosition() {
  if (state.seasonPhase === "regular") {
    if (state.roundIndex === 0) return null;
    const standings = buildStandingsFromPoints(state.leagueTier);
    return { pos: standings.findIndex((r) => r.isPlayer) + 1, size: standings.length, tier: state.leagueTier, final: false };
  }
  if ((state.seasonPhase === "playoffs" || state.seasonPhase === "offseason") && state.lastPlayerPosition && state.lastStandingsTier) {
    return { pos: state.lastPlayerPosition, size: (state.lastStandings && state.lastStandings.length) || BAL.seasonRounds + 1, tier: state.lastStandingsTier, final: true };
  }
  return null;
}

function isSameEntry(entry, champ) {
  return !!champ && (champ.isPlayer ? !!entry.isPlayer : entry.rivalId === champ.rivalId);
}

// Each tier promotes its playoff champion plus the top
// BAL.promotionTablePlaces of the table other than that champion, and
// relegates the bottom BAL.relegationCount — all 5 tiers simultaneously so
// each stays at exactly 40 members. Returns the player's own resulting tier.
function applyPromotionRelegation(allStandings, champions) {
  const promoted = {};
  const relegated = {};
  const stayed = {};

  for (let tier = 1; tier <= BAL.leagueCount; tier++) {
    const standings = allStandings[tier];
    const champ = champions[tier];
    const relCount = tier < BAL.leagueCount ? BAL.relegationCount : 0;
    promoted[tier] =
      tier > 1
        ? standings.filter((e) => isSameEntry(e, champ)).concat(standings.filter((e) => !isSameEntry(e, champ)).slice(0, BAL.promotionTablePlaces))
        : [];
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

// Called once the player's 39th regular-season match resolves. Every
// league's points have already accumulated live, round by round, all
// season — this snapshots the final tally. Promotion/relegation waits for
// the playoffs (see finalizeSeasonMovement), since the champion goes up too.
function processLeagueSeasonEnd() {
  const allStandings = {};
  for (let tier = 1; tier <= BAL.leagueCount; tier++) {
    allStandings[tier] = buildStandingsFromPoints(tier);
    state.leagueStandings[tier] = allStandings[tier];
  }
  state.playoffChampions = null;
  state.seasonMovementApplied = false;

  const playerStandings = allStandings[state.leagueTier];
  const playerPosition = playerStandings.findIndex((s) => s.isPlayer) + 1;
  const qualified = playerPosition <= BAL.playoffSize;
  return { standings: playerStandings, playerPosition, qualified };
}

// Standard 16-bracket seeding pairs so top seeds are spread across the draw
// (1 and 2 can only meet in the Final, etc).
const SEED_PAIRS_16 = [[1, 16], [8, 9], [5, 12], [4, 13], [6, 11], [3, 14], [7, 10], [2, 15]];

function seedBracket(standings) {
  const top = standings.slice(0, BAL.playoffSize);
  const round = [];
  SEED_PAIRS_16.forEach(([a, b]) => {
    round.push(top[a - 1], top[b - 1]);
  });
  return round;
}

function buildPlayoffBracket(standings, tier) {
  const playerSeed = standings.slice(0, BAL.playoffSize).findIndex((t) => t.isPlayer) + 1;
  return { stage: "r16", currentRound: seedBracket(standings), eliminated: false, champion: false, playerSeed, tier };
}

// Plays an all-rival knockout from the given round (teams in bracket order)
// down to one champion. Never called with the player still in it.
function finishNpcBracket(teams, tier) {
  let round = teams;
  while (round.length > 1) {
    const next = [];
    const stage = BRACKET_STAGE_BY_SIZE[round.length];
    for (let i = 0; i < round.length; i += 2) {
      const a = round[i];
      const b = round[i + 1];
      const ra = findRival(a.rivalId);
      const rb = findRival(b.rivalId);
      const ratingA = ra ? ra.rating : a.rating, ratingB = rb ? rb.rating : b.rating;
      const aWins = ra && rb ? resolveNpcMatch(ra, rb) : simulateNpcMatch(a.rating, b.rating);
      if (tier && stage) {
        const w = aWins ? a : b, l = aWins ? b : a;
        recordLeagueResult(tier, stage, w.isPlayer ? HISTORY_PLAYER_ID : w.rivalId, l.isPlayer ? HISTORY_PLAYER_ID : l.rivalId, aWins ? ratingA : ratingB, aWins ? ratingB : ratingA);
      }
      next.push(aWins ? a : b);
    }
    round = next;
  }
  return round[0];
}

// Runs once the player's season is fully over (straight after round 39 if
// they missed the playoffs, otherwise once they're eliminated or win the
// Final). Every other tier's knockout is simulated here so each has a
// champion, then promotion/relegation is applied across all 5 tiers.
// Returns log text describing the player's own league outcome.
function finalizeSeasonMovement(playerTierChampion) {
  const playedTier = state.lastStandingsTier;
  const champions = {};
  for (let tier = 1; tier <= BAL.leagueCount; tier++) {
    const standings = state.leagueStandings[tier];
    if (!standings) continue;
    const champ = tier === playedTier && playerTierChampion ? playerTierChampion : finishNpcBracket(seedBracket(standings), tier);
    champions[tier] = { name: champ.name, rivalId: champ.rivalId, isPlayer: !!champ.isPlayer };
  }
  state.playoffChampions = champions;

  // Saves already mid-playoffs/offseason when this rule shipped had their
  // movement applied under the old table-only rule — never apply it twice.
  if (state.seasonMovementApplied) return "";
  state.seasonMovementApplied = true;

  const newTier = applyPromotionRelegation(state.leagueStandings, champions);
  state.leagueTier = newTier;
  state.peakLeagueTier = Math.min(state.peakLeagueTier, newTier);

  const champ = champions[playedTier];
  const champText = champ && !champ.isPlayer ? ` ${champ.name} won the League ${playedTier} playoffs.` : "";
  const moveText = newTier < playedTier ? ` ⬆️ Promoted to League ${newTier}!` : newTier > playedTier ? ` ⬇️ Relegated to League ${newTier}.` : "";
  return champText + moveText;
}

const PLAYOFF_ROUND_NAMES = { r16: "Round of 16", qf: "Quarterfinal", sf: "Semifinal", f: "Final" };
const PLAYOFF_NEXT_STAGE = { r16: "qf", qf: "sf", sf: "f" };
const PLAYOFF_STAGES = ["r16", "qf", "sf", "f"];

// Knocked out in `stage`: camp covers the playoff rounds still to come.
// R16 exit → 21 days, QF → 14, SF and Final → the normal 7-day break.
function eliminatedCampDays(stage) {
  const roundsLeft = PLAYOFF_STAGES.length - 1 - Math.max(0, PLAYOFF_STAGES.indexOf(stage));
  return Math.max(BAL.offseasonDays, roundsLeft * BAL.roundIntervalDays);
}

// Training camp = the season ended without a title, by missing the
// playoffs or by being knocked out of them.
function isTrainingCamp(s = state) {
  return s.offseasonReason === "missed" || s.offseasonReason === "eliminated";
}

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
      matchResult = resolveMatch(liveOpponentRating(opp));
      matchResult.opponentName = opp.name;
      matchResult.roundLabel = roundName;
      // Playoffs have no live table: show where each finished the regular season.
      const finalTable = state.lastStandings || [];
      const youFinal = finalTable.findIndex((t) => t.isPlayer) + 1;
      const oppFinal = finalTable.findIndex((t) => !t.isPlayer && t.rivalId === opp.rivalId) + 1;
      if (youFinal > 0 && oppFinal > 0) matchResult.h2hTable = { you: youFinal, opp: oppFinal, label: "📊 Final table" };
      const win = matchResult.win;
      matchResult.tier = p.tier;
      recordLeagueResult(p.tier, p.stage, win ? HISTORY_PLAYER_ID : opp.rivalId, win ? opp.rivalId : HISTORY_PLAYER_ID, win ? matchResult.rankBefore : matchResult.opponentRating, win ? matchResult.opponentRating : matchResult.rankBefore);
      recordRivalResultVsPlayer(opp.rivalId, win, matchResult.matchRating);
      winners.push(win ? player : opp);
      summary = win ? `🏆 ${roundName} WIN vs ${opp.name}!` : `💔 Eliminated in the ${roundName} by ${opp.name}.`;
      if (!win) p.eliminated = true;
    } else {
      // Uses the persistent-rating-updating resolver so playoff results
      // also feed back into the rivals' ongoing world, same as the regular
      // season's round-robin does.
      const rivalA = findRival(teamA.rivalId);
      const rivalB = findRival(teamB.rivalId);
      const ratingA = rivalA ? rivalA.rating : teamA.rating, ratingB = rivalB ? rivalB.rating : teamB.rating;
      const aWins = rivalA && rivalB ? resolveNpcMatch(rivalA, rivalB) : simulateNpcMatch(teamA.rating, teamB.rating);
      recordLeagueResult(p.tier, p.stage, (aWins ? teamA : teamB).rivalId, (aWins ? teamB : teamA).rivalId, aWins ? ratingA : ratingB, aWins ? ratingB : ratingA);
      winners.push(aWins ? teamA : teamB);
    }
  }

  let seasonOver = false;
  let champion = null;
  if (p.eliminated) {
    seasonOver = true;
    // The bracket still needs a champion (they're promoted) — play out the
    // remaining rounds without the player.
    champion = finishNpcBracket(winners, p.tier);
  } else if (p.stage === "f") {
    p.champion = true;
    seasonOver = true;
    champion = winners[0];
    const bonus = { cash: 2000, rating: 40 };
    summary = `👑 CHAMPION! You won the Year ${state.year} League ${p.tier} Final! +$${bonus.cash} and +${bonus.rating} rating.`;
    state.cash += bonus.cash;
    state.yearCashFlow.matchCash += bonus.cash;
    state.rank = Math.min(BAL.ratingMax, state.rank + bonus.rating);
    state.peakRank = Math.max(state.peakRank, state.rank);
    matchResult.championBonus = bonus;
  } else {
    p.stage = PLAYOFF_NEXT_STAGE[p.stage];
    p.currentRound = winners;
  }

  return { matchResult, summary, seasonOver, champion };
}

/* ---------------------------------------------------------------------- */
/* Season/day phase engine                                                */
/* ---------------------------------------------------------------------- */
// Advances the season state machine by one day. Call after resolveDay().
// Returns { matchResult, phaseEvent, yearSummary } — any may be null.
function processDayEnd() {
  state.phaseDay += 1;
  let matchResult = null;
  let phaseEvent = null;
  let yearSummary = null;

  if (state.seasonPhase === "preseason") {
    if (state.phaseDay >= BAL.preseasonDays) {
      state.seasonPhase = "regular";
      state.phaseDay = 0;
      state.roundIndex = 0;
      state.skillCycleDay = 0;
      state.activeSkills = rollActiveSkills();
      SKILL_KEYS.forEach((k) => {
        state.allocation.skills[k] = 0;
      });
      const focus = state.activeSkills.map((k) => `${skillMeta(k).icon} ${skillMeta(k).name}`).join(", ");
      phaseEvent = `🏁 Preseason over — the ${BAL.seasonRounds}-round regular season begins. This week's focus: ${focus}.`;
    }
    return { matchResult, phaseEvent, yearSummary };
  }

  if (state.seasonPhase === "regular") {
    if (state.phaseDay >= BAL.roundIntervalDays) {
      state.phaseDay = 0;
      const fixture = state.schedule[state.roundIndex];
      // Table position before this round's results land, for the match
      // summary's "moved from X to Y" — rank/rating alone don't tell you
      // where you actually sit in the table.
      const standingsBefore = buildStandingsFromPoints(state.leagueTier);
      const positionBefore = standingsBefore.findIndex((r) => r.isPlayer) + 1;
      matchResult = resolveMatch(liveOpponentRating(fixture));
      matchResult.opponentName = fixture.name;
      matchResult.roundLabel = `Round ${state.roundIndex + 1}/${BAL.seasonRounds}`;
      matchResult.tier = state.leagueTier;
      recordLeagueResult(state.leagueTier, String(state.roundIndex + 1), matchResult.win ? HISTORY_PLAYER_ID : fixture.rivalId, matchResult.win ? fixture.rivalId : HISTORY_PLAYER_ID, matchResult.win ? matchResult.rankBefore : matchResult.opponentRating, matchResult.win ? matchResult.opponentRating : matchResult.rankBefore);
      const win = matchResult.win;
      recordRivalResultVsPlayer(fixture.rivalId, win, matchResult.matchRating);
      fixture.played = true;
      fixture.win = win;
      state.seasonResults.push({ opponent: fixture.name, win });
      // Credit the winner of the player's own match too — the rival still
      // earns their 3 points when they beat the player, same as any other
      // pairing, so their spot in the table reflects every result.
      const pointsTier = state.leaguePoints[state.leagueTier];
      if (win) pointsTier.player = (pointsTier.player || 0) + 3;
      else pointsTier[fixture.rivalId] = (pointsTier[fixture.rivalId] || 0) + 3;
      for (let tier = 1; tier <= BAL.leagueCount; tier++) {
        resolveLeagueRoundForTier(tier, state.roundIndex);
      }
      const standingsAfter = buildStandingsFromPoints(state.leagueTier);
      matchResult.positionBefore = positionBefore;
      matchResult.positionAfter = standingsAfter.findIndex((r) => r.isPlayer) + 1;
      // Both sides' table places going into the match, for Head to Head.
      const oppPos = standingsBefore.findIndex((r) => r.rivalId === fixture.rivalId) + 1;
      if (oppPos > 0) matchResult.h2hTable = { you: positionBefore, opp: oppPos, label: "📊 Table" };
      matchResult.leagueSize = standingsAfter.length;
      state.roundIndex += 1;

      if (state.roundIndex >= BAL.seasonRounds) {
        const playedTier = state.leagueTier;
        const result = processLeagueSeasonEnd();
        state.lastStandings = result.standings;
        state.lastPlayerPosition = result.playerPosition;
        state.lastStandingsTier = playedTier;
        const pos = result.playerPosition;
        const size = result.standings.length;
        if (result.qualified) {
          state.seasonPhase = "playoffs";
          state.playoff = buildPlayoffBracket(result.standings, playedTier);
          const stakes =
            playedTier > 1 && pos <= BAL.promotionTablePlaces
              ? ` Top ${BAL.promotionTablePlaces} — promotion secured!`
              : playedTier > 1
              ? " Win the playoffs to earn promotion."
              : "";
          phaseEvent = `🏆 Regular season complete! Finished #${pos} of ${size} in League ${playedTier} — through to the playoffs as seed ${state.playoff.playerSeed}.${stakes}`;
        } else {
          state.seasonPhase = "offseason";
          state.offseasonDays = BAL.trainingCampDays;
          state.offseasonReason = "missed";
          const outcome = finalizeSeasonMovement(null);
          phaseEvent = `📋 Regular season complete. Finished #${pos} of ${size} in League ${playedTier} — missed the top ${BAL.playoffSize} playoff cutoff.${outcome} ${BAL.trainingCampDays}-day training camp starts now to get ready for next season.`;
        }
      }
    }
    return { matchResult, phaseEvent, yearSummary };
  }

  if (state.seasonPhase === "playoffs") {
    if (state.phaseDay >= BAL.roundIntervalDays) {
      state.phaseDay = 0;
      const result = resolvePlayoffRound();
      matchResult = result.matchResult;
      phaseEvent = result.summary;
      if (result.seasonOver) {
        state.seasonPhase = "offseason";
        if (state.playoff.eliminated) {
          // Knocked out: training camp for the rest of the playoff window
          // (never shorter than the normal break) — an early exit gets its
          // time back as training, same as missing the playoffs entirely.
          state.offseasonDays = eliminatedCampDays(state.playoff.stage);
          state.offseasonReason = "eliminated";
        } else {
          state.offseasonDays = BAL.offseasonDays;
          state.offseasonReason = "playoffs";
        }
        phaseEvent += finalizeSeasonMovement(result.champion);
        if (state.offseasonReason === "eliminated") {
          phaseEvent += ` ${state.offseasonDays}-day training camp starts now to get ready for next season.`;
        }
      }
    }
    return { matchResult, phaseEvent, yearSummary };
  }

  if (state.seasonPhase === "offseason") {
    if (state.phaseDay >= (state.offseasonDays || BAL.offseasonDays)) {
      const flow = state.yearCashFlow;
      const income = flow.workPay + flow.matchCash;
      yearSummary = {
        year: state.year,
        workPay: flow.workPay,
        matchCash: flow.matchCash,
        income,
        expenses: flow.expenses,
        net: income - flow.expenses,
        cashNow: state.cash,
      };
      state.yearCashFlow = { workPay: 0, matchCash: 0, expenses: 0 };

      // Annual raise — rises $10/year for 5 years of unbroken tenure in
      // whichever role is currently held, then plateaus. Only the active
      // role's rate moves; the other sits untouched until it's relevant.
      const emp = state.employment;
      const maxRaise = BAL.payRaisePerYear * BAL.payRaiseMaxYears;
      if (emp.status === "employed") {
        emp.workPay = Math.min(emp.workPay + BAL.payRaisePerYear, BAL.workPayMin + maxRaise);
        yearSummary.newPayRate = emp.workPay;
      } else if (emp.status === "pro") {
        emp.proPay = Math.min(emp.proPay + BAL.payRaisePerYear, BAL.proPayMin + maxRaise);
        yearSummary.newPayRate = emp.proPay;
      }

      state.phaseDay = 0;
      state.year += 1;
      state.seasonPhase = "preseason";
      state.roundIndex = 0;
      state.seasonResults = [];
      const newLeagueData = buildSeasonLeagueData(state.rivals, state.leagueTier);
      state.schedule = newLeagueData.schedule;
      state.leagueRoundRobins = newLeagueData.roundRobins;
      state.leaguePoints = newLeagueData.points;
      state.playoff = null;
      // Offseason recovery — a clean slate for the new year.
      state.stats.composure = 100;
      state.stats.rest = 100;
      state.stats.nutrition = 100;
      phaseEvent = `🎉 Year ${state.year} begins! A fresh ${BAL.seasonRounds}-round season has been scheduled — good luck.`;
    }
    return { matchResult, phaseEvent, yearSummary };
  }

  return { matchResult, phaseEvent, yearSummary };
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

// Phrases a days-until countdown for the topbar/Career "next up" text.
// 1 day away means ending today's day is what triggers it, so "after
// today" reads clearer in context than "in 1d" — everything else is
// spelled out in full rather than abbreviated ("in 3 days", not "in 3d").
function daysUntilPhrase(daysUntil) {
  return daysUntil === 1 ? "after today" : `in ${daysUntil} days`;
}

// Topbar line 2. Kept short so the Unassigned hours pill fits beside it on
// a phone — a longer line pushes the pill down onto a row of its own.
function matchOddsLine(winPct, daysUntil) {
  return `${winPct}% to win · match ${daysUntil <= 0 ? "today" : daysUntilPhrase(daysUntil)}`;
}

function getNextMatchInfo() {
  const s = state;
  if (s.seasonPhase === "preseason") {
    const daysUntil = BAL.preseasonDays - s.phaseDay;
    return { kind: "preseason", label: `Season starts ${daysUntilPhrase(daysUntil)}` };
  }
  if (s.seasonPhase === "regular") {
    const daysUntil = BAL.roundIntervalDays - s.phaseDay;
    const fixture = s.schedule[s.roundIndex];
    if (!fixture) return { kind: "regular-end", label: "Season wrapping up…" };
    const roundNum = s.roundIndex + 1;
    const liveRating = liveOpponentRating(fixture);
    const diff = difficultyLabel(liveRating, s.rank);
    const winPct = Math.round(winProbabilityAgainst(liveRating) * 100);
    const oppRating = Math.round(liveRating);
    return {
      kind: "fixture",
      daysUntil,
      opponentName: fixture.name,
      opponentRating: oppRating,
      opponentRivalId: fixture.rivalId,
      difficulty: diff,
      winPct,
      label: daysUntil <= 0
        ? `Rd ${roundNum} TODAY vs ${fixture.name} · ${winPct}%`
        : `Rd ${roundNum}/${BAL.seasonRounds} vs ${fixture.name} · ${winPct}% · ${daysUntilPhrase(daysUntil)}`,
      labelLine1: daysUntil <= 0
        ? `Rd ${roundNum} TODAY vs ${fixture.name} (${oppRating})`
        : `Rd ${roundNum}/${BAL.seasonRounds} vs ${fixture.name} (${oppRating})`,
      labelLine2: matchOddsLine(winPct, daysUntil),
    };
  }
  if (s.seasonPhase === "playoffs") {
    const daysUntil = BAL.roundIntervalDays - s.phaseDay;
    const p = s.playoff;
    const idx = p.currentRound.findIndex((t) => t.isPlayer);
    const opp = idx >= 0 ? p.currentRound[idx % 2 === 0 ? idx + 1 : idx - 1] : null;
    const roundName = PLAYOFF_ROUND_NAMES[p.stage];
    const liveRating = opp ? liveOpponentRating(opp) : null;
    const diff = opp ? difficultyLabel(liveRating, s.rank) : null;
    const winPct = opp ? Math.round(winProbabilityAgainst(liveRating) * 100) : null;
    const oppRating = opp ? Math.round(liveRating) : null;
    const oppTag = opp ? `${opp.name} (${oppRating})` : "?";
    return {
      kind: "playoff",
      daysUntil,
      opponentName: opp ? opp.name : "?",
      opponentRating: oppRating,
      opponentRivalId: opp ? opp.rivalId : null,
      difficulty: diff,
      winPct,
      label: daysUntil <= 0
        ? `${roundName} TODAY vs ${opp ? opp.name : "?"}${winPct != null ? " · " + winPct + "%" : ""}`
        : `${roundName} vs ${opp ? opp.name : "?"}${winPct != null ? " · " + winPct + "%" : ""} · ${daysUntilPhrase(daysUntil)}`,
      labelLine1: daysUntil <= 0 ? `${roundName} TODAY vs ${oppTag}` : `${roundName} vs ${oppTag}`,
      labelLine2: winPct != null ? matchOddsLine(winPct, daysUntil) : daysUntilPhrase(daysUntil),
    };
  }
  if (s.seasonPhase === "offseason") {
    const daysUntil = (s.offseasonDays || BAL.offseasonDays) - s.phaseDay;
    let outcome = "Season over";
    if (isTrainingCamp(s)) outcome = "Training camp";
    else if (s.playoff && s.playoff.champion) outcome = "🏆 Champion!";
    else if (s.playoff && s.playoff.eliminated) outcome = "Eliminated";
    return { kind: "offseason", daysUntil, label: `${outcome} · Year ${s.year + 1} ${daysUntilPhrase(daysUntil)}` };
  }
  return { kind: "unknown", label: "" };
}

// Every 7 days, reroll which 1-3 skills are trainable for the coming week —
// revealed only now, at the end of the previous cycle, never in advance.
// Hours pending on the old active skills are cleared since they're about to
// become untrainable; the player reassigns under the new focus.
function advanceSkillCycle(wasInSeason) {
  // Outside the season everything is trainable (see trainableSkills()), so
  // the weekly reveal only runs on days that are in-season both before and
  // after processDayEnd() — which runs first and can change seasonPhase on
  // this exact call. The regular season's first week is rolled when it
  // starts (processDayEnd), and a season that just ended doesn't announce a
  // "week ahead" it no longer has.
  if (!wasInSeason || !inSeason()) return null;
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
  if (s.seasonPhase === "regular") return league + "Regular";
  if (s.seasonPhase === "playoffs") return league + "Playoffs";
  // The offseason label right after this already names it (Training camp,
  // 🏆 Champion!, …), so don't say "Camp" twice.
  if (s.seasonPhase === "offseason") return `League ${s.leagueTier}`;
  return "";
}

/* ---------------------------------------------------------------------- */
/* UI rendering                                                           */
/* ---------------------------------------------------------------------- */
const $ = (id) => document.getElementById(id);

const ACT_MAX = { exercise: BAL.gymMaxHours, sleep: 12, relax: 12, nutrition: 4 };
const SKILL_MAX_HOURS = 16;

function markerPct(threshold, max) {
  return clamp((threshold / max) * 100, 0, 100);
}

function renderTopbar() {
  $("cashVal").textContent = fmt(state.cash);
  $("rankVal").textContent = fmt(state.rank);
  // Leagues button carries your table position, tinted like the table's own
  // zones: promotion (green), relegation (red), outside the playoffs (dim).
  const tp = playerTablePosition();
  const lb = $("leaguesBtn");
  if (tp) {
    const zone =
      tp.tier > 1 && tp.pos <= BAL.promotionTablePlaces ? "pos-promo"
      : tp.tier < BAL.leagueCount && tp.pos > tp.size - BAL.relegationCount ? "pos-releg"
      : tp.pos > BAL.playoffSize ? "pos-out"
      : "";
    lb.innerHTML = `🏅<span class="leagues-pos ${zone}">#${tp.pos}</span>`;
    lb.classList.add("has-pos");
    lb.title = `${tp.final ? "Final regular-season" : "Current"} position: #${tp.pos} of ${tp.size} in League ${tp.tier}`;
  } else {
    lb.textContent = "🏅";
    lb.classList.remove("has-pos");
    lb.title = "Leagues";
  }
  const info = getNextMatchInfo();
  const phase = `Y${state.year} · ${phaseLabelText()}`;
  $("matchCounter").textContent = `${phase} · ${info.labelLine1 || info.label}`;
  $("matchCounter2").textContent = info.labelLine1 ? (info.labelLine2 || "") : "";
}

function renderStats() {
  const s = state.stats;
  const banner = $("warningBanner");
  const msgs = [];
  if (state.burnout.active) msgs.push("🔥 Burnout active — training & exercise are far less effective. Relax to recover.");
  else if (s.composure <= 10) msgs.push("🔥 Composure critical — burnout imminent. Schedule relaxation soon.");
  if (state.injury.active) msgs.push(`🤕 Injured — Gym locked for ${state.injury.daysLeft} more ${state.injury.daysLeft === 1 ? "day" : "days"}.`);
  if (s.rest <= 30) msgs.push("🌙 Severely low Rest — your body is breaking down. Sleep more.");
  if (s.phys <= 20) msgs.push("💪 Health critically low — it's capping your training gains.");
  if (s.nutrition <= BAL.nutritionHoursCapLow) msgs.push(`🥗 Nutrition critical — your day is capped at only ${dailyHoursCap(s.nutrition)}h. Eat better to earn more hours back.`);

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
  return skillSum + a.exercise + a.sleep + a.relax + a.nutrition + (a.work || 0);
}

// One row does double duty as both the stat readout (name, value/cap, bar)
// and the time-allocation control (slider + stepper) for a single lever —
// no more showing the same name twice in two separate sections.
// scaleMax: what a full bar means. Defaults to cap; skills and Health pass
// 100 so every bar shares one scale, with the stretch from the current cap
// to 100 hatched as locked (otherwise 39/50 and 47/60 both look "nearly full").
function lockedZoneHtml(cap, scaleMax) {
  if (cap >= scaleMax) return "";
  const capPct = clamp((cap / scaleMax) * 100, 0, 100);
  return `<div class="bar-locked" style="left:${capPct}%;width:${100 - capPct}%"></div>`;
}

function comboRowHtml(key, { icon, label, outcomeText, value, previewValue, cap, scaleMax = cap, hours, maxHours, markerHours, shopTag, disabled, barClass }) {
  const barPct = clamp((value / scaleMax) * 100, 0, 100);
  const previewPct = previewValue == null ? barPct : clamp((previewValue / scaleMax) * 100, 0, 100);
  const overlayLeft = Math.min(barPct, previewPct);
  const overlayWidth = Math.abs(previewPct - barPct);
  // Always green for a gain, always red for a loss — not tinted by the
  // bar's own color — so direction reads at a glance.
  const overlayCls = previewPct > barPct ? "bar-preview-gain" : "bar-preview-loss";
  const overlayHtml = overlayWidth > 0.3 ? `<div class="bar-preview ${overlayCls} bar-fill" style="left:${overlayLeft}%;width:${overlayWidth}%"></div>` : "";
  const mPct = markerPct(markerHours, maxHours);
  // Tint the slider track itself so meeting (or missing) a stat's decay
  // threshold reads at a glance, not just from the marker tick.
  const meetsMarker = markerHours > 0 ? hours >= markerHours : null;
  const sliderCls = meetsMarker === null ? "" : meetsMarker ? "slider-meets" : "slider-under";
  return `
  <div class="activity combo-row ${barClass === "skill" ? "skill-row" : ""}" data-act="${key}" style="${disabled ? "opacity:0.45" : ""}">
    <div class="activity-row">
      <span class="activity-icon">${icon}</span>
      <span class="activity-name">${label}</span>
      ${shopTag || ""}
      <span class="combo-outcome">${outcomeText}</span>
      <span class="activity-hours"><span id="hoursVal_${key}">${hours}</span>h</span>
    </div>
    <div class="bar combo-bar"><div class="bar-fill ${barClass}" style="width:${barPct}%"></div>${lockedZoneHtml(cap, scaleMax)}${overlayHtml}</div>
    <div class="stepper">
      <button class="step-btn" data-key="${key}" data-dir="-1" ${disabled ? "disabled" : ""}>−</button>
      <div class="slider-wrap">
        <input type="range" class="${sliderCls}" min="0" max="${maxHours}" step="1" value="${hours}" id="hours_${key}" data-key="${key}" ${disabled ? "disabled" : ""} />
        <div class="slider-marker" style="left:${mPct}%"></div>
      </div>
      <button class="step-btn" data-key="${key}" data-dir="1" ${disabled ? "disabled" : ""}>+</button>
    </div>
  </div>`;
}

// Net of cost of living, not gross pay — what actually lands in (or leaves)
// cash each day. The Career modal still breaks out gross pay and expenses.
function netPerDayTag(net) {
  return `<span class="skill-shop-tag${net < 0 ? " skill-shop-tag-negative" : ""}">Net ${fmtMoney(net)}/d</span>`;
}

// Work / Job Search / Pro Duties share one slider and one row — only what
// it's called, what it requires, and what its bar means change with status.
function workRowHtml() {
  const emp = state.employment;
  const a = state.allocation;
  const hours = a.work || 0;
  const maxHours = workMaxHours();

  if (emp.status === "unemployed") {
    return comboRowHtml("work", {
      icon: "🔍",
      label: "Job Search",
      outcomeText: `${fmt(emp.jobSearchHours)}/${jobSearchTarget()}h`,
      shopTag: netPerDayTag(-BAL.dailyExpenses),
      value: emp.jobSearchHours,
      previewValue: Math.min(emp.jobSearchHours + hours, jobSearchTarget()),
      cap: jobSearchTarget(),
      hours,
      maxHours,
      markerHours: BAL.jobSearchMinHours,
      barClass: "work",
    });
  }

  const isPro = emp.status === "pro";
  const required = isPro ? BAL.proDutyHoursRequired : BAL.workHoursRequired;
  const atRisk = hours < required;
  const livesText = `${fmt1(livesRemaining())}/${fmt1(BAL.strikesToFire)} left`;
  const outcomeText = (atRisk ? "⚠️ " : "") + (isPro ? `${livesText} · ${techniqueStatusText()}` : livesText);
  const statusTag = netPerDayTag((isPro ? emp.proPay : emp.workPay) - BAL.dailyExpenses);
  return comboRowHtml("work", {
    icon: isPro ? "📱" : "💼",
    label: isPro ? "Pro Duties" : "Work",
    outcomeText,
    shopTag: statusTag,
    value: jobSecurityPct(),
    previewValue: jobSecurityPreviewPct(hours, isPro),
    cap: 100,
    hours,
    maxHours,
    markerHours: required,
    barClass: "work",
  });
}

function renderPlannerRows() {
  const a = state.allocation;
  const s = state.stats;
  const preview = previewTomorrow();
  const rows = [workRowHtml()];

  trainableSkills().forEach((key) => {
    const meta = skillMeta(key);
    const cap = skillCap(key);
    const lvl = hiredLevel(coachKey(key));
    const shopTag = lvl > 0 ? `<span class="skill-shop-tag">🧑‍🏫 Lv${lvl}</span>` : `<span class="skill-shop-tag skill-shop-tag-none">no coach</span>`;
    rows.push(
      comboRowHtml(key, {
        icon: meta.icon,
        label: meta.name,
        outcomeText: `${fmt(s.skills[key])}/${cap}`,
        value: s.skills[key],
        previewValue: preview.skills[key],
        cap,
        scaleMax: 100,
        hours: a.skills[key] || 0,
        maxHours: SKILL_MAX_HOURS,
        markerHours: BAL.skillDecayThresholdHours,
        shopTag,
        barClass: "skill",
      })
    );
  });

  const pCap = physCap();
  // Room for one note plus the injury readout before the row overflows.
  // The note is "Rest −N" when poor Rest is draining Health (otherwise
  // Health can keep dropping with Gym hours in, which reads as a bug), else
  // the skill multiplier. While injured the readout shows days left instead.
  const gymNote = preview.physDecayFromRest > 1 ? `Rest −${fmt1(preview.physDecayFromRest)}` : `×${physSynergy(s.phys).toFixed(2)} skill`;
  const risk = injuryChance(a.exercise);
  const injuryText = state.injury.active ? ` · 🤕 ${state.injury.daysLeft}d` : risk > 0 ? ` · ${fmtPct(risk)} 🤕` : "";
  rows.push(
    comboRowHtml("exercise", {
      icon: "🏃",
      label: "Gym",
      outcomeText: `Health ${fmt(s.phys)}/${pCap} · ${gymNote}${injuryText}`,
      value: s.phys,
      previewValue: preview.phys,
      cap: pCap,
      scaleMax: 100,
      hours: a.exercise,
      maxHours: ACT_MAX.exercise,
      markerHours: BAL.skillDecayThresholdHours,
      disabled: state.injury.active,
      barClass: "phys",
    })
  );
  rows.push(
    comboRowHtml("sleep", {
      icon: "🌙",
      label: "Sleep",
      outcomeText: `Rest ${fmt(s.rest)}/100 · ×${restTrainingMultiplier(s.rest).toFixed(1)} training`,
      value: s.rest,
      previewValue: preview.rest,
      cap: 100,
      hours: a.sleep,
      maxHours: ACT_MAX.sleep,
      markerHours: BAL.idealSleep,
      barClass: "rest",
    })
  );
  // Composure moves from more than just Relax hours — a good night's Sleep
  // (>=idealSleep) grants a flat relief bonus on its own, and low Rest bleeds
  // it down independent of Relax too. Without this, Composure can visibly
  // rise (or fall) with hours that look "not enough" on this row alone.
  // Only room for one note before the row overflows, so show whichever is
  // the bigger factor today, and drop "match" to make space for it.
  const sleepBonus = preview.sleepH >= BAL.idealSleep ? BAL.restGoodSleepBonus : 0;
  const composureRestDrag = preview.composurePenaltyFromRest > 1 ? preview.composurePenaltyFromRest : 0;
  let composureNote = "";
  if (sleepBonus > 0 || composureRestDrag > 0) {
    composureNote = sleepBonus >= composureRestDrag ? ` · Sleep +${sleepBonus}` : ` · Rest −${fmt1(composureRestDrag)}`;
  }
  const matchSuffix = composureNote ? "" : " match";
  rows.push(
    comboRowHtml("relax", {
      icon: "🎮",
      label: "Relax",
      outcomeText: `Composure ${fmt(s.composure)}/100 · ×${composureMatchMultiplier(s.composure)}${matchSuffix}${composureNote}`,
      value: s.composure,
      previewValue: preview.composure,
      cap: 100,
      hours: a.relax,
      maxHours: ACT_MAX.relax,
      markerHours: BAL.relaxComposureThreshold,
      barClass: "composure",
    })
  );
  const capToday = dailyHoursCap(s.nutrition);
  const capTomorrow = dailyHoursCap(preview.nutrition);
  let nutritionOutcome = `${fmt(s.nutrition)}/100`;
  if (capToday < BAL.dailyHoursCeiling) nutritionOutcome += ` · ${capToday}h today`;
  if (capTomorrow !== capToday) nutritionOutcome += ` → ${capTomorrow}h tomorrow`;
  rows.push(
    comboRowHtml("nutrition", {
      icon: "🥗",
      label: "Food",
      outcomeText: nutritionOutcome,
      value: s.nutrition,
      previewValue: preview.nutrition,
      cap: 100,
      hours: a.nutrition,
      maxHours: ACT_MAX.nutrition,
      markerHours: BAL.skillDecayThresholdHours,
      barClass: "nutrition",
    })
  );

  $("plannerRows").innerHTML = rows.join("");
}

function renderPlanner() {
  renderPlannerRows();
  const cap = dailyHoursCap(state.stats.nutrition);
  const left = cap - totalAssigned();
  const hoursLeftEl = $("hoursLeft");
  hoursLeftEl.textContent = left;
  hoursLeftEl.classList.toggle("over", left < 0);
  hoursLeftEl.classList.toggle("unassigned", left > 0);
  // The "why" (Nutrition) lives on Nutrition's own row now — this just
  // states the number so the header stays one line.
  const capNoteEl = $("hoursCapNote");
  if (capNoteEl) capNoteEl.textContent = cap < BAL.dailyHoursCeiling ? ` of ${cap}h` : "";
  // Over-allocated (the cap can shrink overnight via Nutrition after hours
  // were already set against yesterday's higher cap) — block ending the day
  // until it's brought back down to the new, smaller budget.
  const blocker = endDayBlocker();
  const endDayBtn = $("endDayBtn");
  endDayBtn.disabled = !!blocker;
  endDayBtn.title = blocker || "";
  const hint = $("endDayHint");
  if (hint) hint.textContent = blocker || "";
  const endWeekBtn = $("endWeekBtn");
  if (endWeekBtn) {
    const days = daysLeftInWeek();
    const toMatch = state.seasonPhase === "regular" || state.seasonPhase === "playoffs";
    endWeekBtn.disabled = !!blocker;
    endWeekBtn.innerHTML = `${toMatch ? "To Match" : "End Week"} ▶▶ <span class="end-week-days">${days}d</span>`;
    endWeekBtn.title = toMatch
      ? `Repeat today's plan for ${days} day${days === 1 ? "" : "s"}, through the next match`
      : `Repeat today's plan for ${days} day${days === 1 ? "" : "s"}, to the end of the week`;
  }
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
// A title row pinned to the top of the scrolling sheet, with its own ✕
// (pass { ownClose: true } to openModal). extraHtml sits between the two —
// it and the ✕ never shrink; the title truncates first.
function stickyHeadHtml(title, extraHtml = "") {
  return `
    <div class="modal-sticky-head">
      <h2>${title}</h2>
      ${extraHtml}
      <button class="modal-close modal-close-inline" data-modal-close aria-label="Close">✕</button>
    </div>`;
}

function openModal(html, { ownClose = false, keepScroll = false } = {}) {
  $("modalBody").innerHTML = html;
  // Modals with their own sticky header carry their own ✕; hide the
  // floating one, which scrolls away with the content.
  $("modal").classList.toggle("modal-own-close", ownClose);
  $("modalOverlay").classList.remove("hidden");
  // The sheet is reused for every pop-up, so it remembers how far the last
  // one was scrolled — start each new one at the top. Must happen after it
  // is shown: browsers ignore scrolling an element that's display:none. A
  // pop-up redrawing itself in place (the Shop after a purchase) keeps its
  // position.
  if (!keepScroll) $("modal").scrollTop = 0;
  const inlineClose = $("modalBody").querySelector("[data-modal-close]");
  if (inlineClose) inlineClose.addEventListener("click", closeModal);
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
// Match result: what the performance score was actually made of — the raw
// value of each input, its weight, and the points it contributed. The Skill
// row's sub-rows show where its value came from.
function performanceTableHtml(b) {
  const one = (n) => (Math.round(n * 10) / 10).toFixed(1);
  const pct = (w) => `${Math.round(w * 100)}%`;
  const raw = b.raw;
  const sub = (label, value) => `<tr class="perf-sub"><td>${label}</td><td>${value}</td><td></td><td></td></tr>`;
  const row = (label, value, w, points) =>
    `<tr class="perf-main"><td>${label}</td><td>${value}</td><td>${pct(w)}</td><td>${one(points)}</td></tr>`;
  let skillDetail = "";
  if (raw) {
    skillDetail = raw.skills
      .map((sk) => {
        const m = skillMeta(sk.key);
        return sub(`${m.icon} ${m.name}`, fmt(sk.value));
      })
      .join("");
    if (raw.skills.length > 1) skillDetail += sub("Average", one(raw.skillAvg));
    if (b.composureMult < 1) skillDetail += sub(`😰 Low Composure (under ${BAL.composureMatchMid})`, `×${b.composureMult.toFixed(2)}`).replace("perf-sub", "perf-sub perf-warn");
  }
  const subtotal = b.weighted.skill + b.weighted.phys + b.weighted.composure + b.weighted.rest;
  const techRow =
    b.techniquePenaltyMult < 1
      ? `<tr class="perf-main perf-warn"><td>📘 Unmastered techniques</td><td>×${b.techniquePenaltyMult.toFixed(2)}</td><td></td><td>−${one(subtotal - b.score)}</td></tr>`
      : "";
  return `
    <table class="perf-table">
      <thead><tr><th>Performance</th><th>Value</th><th>Weight</th><th>Points</th></tr></thead>
      <tbody>
        ${row("🧠 Skill", raw ? one(raw.skillComponent) : "", PERF_WEIGHTS.skill, b.weighted.skill)}
        ${skillDetail}
        ${row("🏃 Health", raw ? fmt(raw.phys) : "", PERF_WEIGHTS.phys, b.weighted.phys)}
        ${row("🎮 Composure", raw ? fmt(raw.composure) : "", PERF_WEIGHTS.composure, b.weighted.composure)}
        ${row("🌙 Rest", raw ? fmt(raw.rest) : "", PERF_WEIGHTS.rest, b.weighted.rest)}
        ${techRow}
      </tbody>
      <tfoot><tr><td>Total</td><td></td><td></td><td>${one(b.score)}<span class="perf-of">/100</span></td></tr></tfoot>
    </table>
    <div class="match-sub perf-note">Values are out of 100; points = value × weight. Skill only counts this round's active skill${
      raw && raw.skills.length > 1 ? "s, averaged" : ""
    }.</div>`;
}

// Match result screen, top to bottom: what changed (Result), why (You vs
// Opponent, every row meaning the same thing in both columns), then two
// collapsed explanations (your performance breakdown; luck and win chance).
const signedNum = (n) => (n === 0 ? "±0" : n < 0 ? `−${Math.abs(n)}` : `+${n}`);
const deltaClass = (n) => (n > 0 ? "wk-up" : n < 0 ? "wk-down" : "");

function gapLabel(gap) {
  if (gap === 0) return "Level";
  return `${Math.abs(gap)} ${gap > 0 ? "ahead" : "behind"}`;
}

function matchResultTableHtml(result) {
  const bonus = result.championBonus || { cash: 0, rating: 0 };
  const ratingAfter = result.ratingAfter != null ? result.ratingAfter : state.rank;
  const cashAfter = result.cashBefore + result.cashReward + bonus.cash;
  const rec = result.recordBefore;
  const recAfter = { wins: rec.wins + (result.win ? 1 : 0), losses: rec.losses + (result.win ? 0 : 1) };
  const row = (label, before, after, change, cls) =>
    `<tr><td>${label}</td><td>${before}</td><td>${after}</td><td class="${cls}">${change}</td></tr>`;
  const rows = [row("🏆 Rating", fmt(result.rankBefore), fmt(ratingAfter), signedNum(ratingAfter - result.rankBefore), deltaClass(ratingAfter - result.rankBefore))];
  if (typeof result.positionAfter === "number") {
    const moved = result.positionBefore - result.positionAfter; // + = climbed
    rows.push(row("📊 Table", `#${result.positionBefore}`, `#${result.positionAfter}`, moved === 0 ? "–" : `${moved > 0 ? "↑" : "↓"}${Math.abs(moved)}`, deltaClass(moved)));
  }
  const cashDelta = cashAfter - result.cashBefore;
  rows.push(row("💰 Cash", fmtMoney(result.cashBefore), fmtMoney(cashAfter), `+${fmtMoney(cashDelta)}`, "wk-up"));
  rows.push(row("📋 Record", `${rec.wins}–${rec.losses}`, `${recAfter.wins}–${recAfter.losses}`, result.win ? "W" : "L", result.win ? "wk-up" : "wk-down"));

  return `
    <table class="perf-table result-table">
      <thead><tr><th>Your Result</th><th>Before</th><th>After</th><th>Change</th></tr></thead>
      <tbody>${rows.join("")}</tbody>
    </table>`;
}

// Collapsed "How were rating and cash worked out?": the two sums behind the
// Your Result table's Rating and Cash changes.
function ratingCashExplainerHtml(result) {
  const bonus = result.championBonus || { cash: 0, rating: 0 };
  const row = (label, value, cls = "") => `<tr class="${cls}"><td>${label}</td><td>${value}</td></tr>`;
  const raw = result.ratingChangeRaw;
  const ratingRows = [
    result.win
      ? row(`Win: 24 × your ${100 - result.winProb}% chance to lose`, signedNum(raw))
      : row(`Loss: 24 × your ${result.winProb}% chance to win`, signedNum(raw)),
  ];
  if (result.ratingChange !== raw) {
    ratingRows.push(row(`🧑‍💼 Team Manager (−${Math.round((1 - result.rankLossMult) * 100)}% on losses)`, signedNum(result.ratingChange - raw), "perf-sub"));
  }
  if (bonus.rating) ratingRows.push(row("👑 Champion bonus", signedNum(bonus.rating), "perf-sub"));
  const ratingTotal = (result.ratingAfter != null ? result.ratingAfter : state.rank) - result.rankBefore;

  const basePrize = result.win ? 150 + result.rankAfterMatch / 10 : 40;
  const cashRows = [
    result.win
      ? row(`Win: $150 + your new rating (${fmt(result.rankAfterMatch)}) ÷ 10`, `+${fmtMoney(Math.round(basePrize))}`)
      : row("Loss: flat prize", `+${fmtMoney(40)}`),
  ];
  if (result.cashBonusMult > 1) {
    cashRows.push(row(`🧑‍💼 Team Manager (+${Math.round((result.cashBonusMult - 1) * 100)}% prize money)`, `+${fmtMoney(result.cashReward - Math.round(basePrize))}`, "perf-sub"));
  }
  if (bonus.cash) cashRows.push(row("👑 Champion bonus", `+${fmtMoney(bonus.cash)}`, "perf-sub"));
  const cashTotal = result.cashReward + bonus.cash;

  return `
    <table class="perf-table calc-table">
      <thead><tr><th>🏆 Rating</th><th></th></tr></thead>
      <tbody>${ratingRows.join("")}</tbody>
      <tfoot><tr><td>Rating change</td><td>${signedNum(ratingTotal)}</td></tr></tfoot>
    </table>
    <div class="match-sub perf-note">A win is worth up to 24 points and a loss costs up to 24. The less likely the result, the bigger the change — beating a favourite earns a lot, beating an underdog very little.</div>
    <table class="perf-table calc-table">
      <thead><tr><th>💰 Cash</th><th></th></tr></thead>
      <tbody>${cashRows.join("")}</tbody>
      <tfoot><tr><td>Prize money</td><td>+${fmtMoney(cashTotal)}</td></tr></tfoot>
    </table>
    <div class="match-sub perf-note">Your cash "Before" already includes today's pay and living costs, so the change is just this match's prize.</div>`;
}

function headToHeadHtml(result) {
  const y = result.sides.you;
  const o = result.sides.opp;
  const row = (label, a, b, cls = "") => `<tr class="${cls}"><td>${label}</td><td>${a}</td><td>${b}</td></tr>`;
  const marginText = result.margin === 0 ? "by less than 1" : `by ${result.margin}`;
  return `
    <table class="perf-table h2h-table">
      <thead><tr><th>Head to Head</th><th>You</th><th class="h2h-opp">${result.opponentName || "Opponent"}</th></tr></thead>
      <tbody>
        ${row("🏆 Rating", y.rating, o.rating)}
        ${result.h2hTable ? row(result.h2hTable.label, `#${result.h2hTable.you}`, `#${result.h2hTable.opp}`) : ""}
        ${row("📈 Performance", `${y.perfScore.toFixed(1)} → ${signedNum(y.perfAdj)}`, `${o.perfScore} → ${signedNum(o.perfAdj)}`)}
        ${row("🎲 Luck on the day", signedNum(y.luck), signedNum(o.luck))}
      </tbody>
      <tfoot>${row("Match-day rating", y.final, o.final)}</tfoot>
    </table>
    <div class="match-verdict ${result.win ? "win" : "loss"}">${result.win ? "Won" : "Lost"} ${marginText}</div>
    <div class="match-sub perf-note">Before the match you had a <b>${result.winProb}%</b> win chance.</div>`;
}

function luckExplainerHtml(result) {
  const gap = result.matchRating - result.opponentRating; // + = you're ahead, before luck
  const chanceFor = (g) => Math.round(100 / (1 + Math.pow(10, -g / 400)));
  const ladderGaps = [-400, -200, -100, 0, 100, 200, 400];
  const rows = ladderGaps.filter((g) => g !== gap).map((g) => ({ gap: g, chance: chanceFor(g), you: false }));
  rows.push({ gap, chance: result.winProb, you: true });
  rows.sort((x, y) => x.gap - y.gap || (x.you ? 1 : -1));
  const ladder = rows
    .map((r) => `<tr class="${r.you ? "ladder-you" : "perf-sub"}"><td>${r.you ? `👉 You: ${gapLabel(r.gap)}` : gapLabel(r.gap)}</td><td>${r.chance}%</td></tr>`)
    .join("");
  return `
    <div class="match-sub perf-note">Both sides get random luck every match — usually somewhere between about −150 and +120, with the occasional inspired day of +400 or more. The highest match-day rating wins.</div>
    <div class="match-sub perf-note">So your win chance comes down to the gap before luck: your rating + performance (${result.matchRating}) against their rating (${result.opponentRating}) — <b>${gapLabel(gap)}</b> this time. The bigger the gap, the more luck the underdog needs, but upsets always stay possible:</div>
    <table class="perf-table win-ladder">
      <thead><tr><th>Rating gap</th><th>Win chance</th></tr></thead>
      <tbody>${ladder}</tbody>
    </table>`;
}

function showMatchModal(result, extraHtml = "", { onContinue = null, continueLabel = "Continue" } = {}) {
  const context = result.roundLabel ? `${result.roundLabel}${result.opponentName ? " vs " + result.opponentName : ""}` : "";
  const title = `<span class="match-head ${result.win ? "win" : "loss"}">${result.win ? "VICTORY" : "DEFEAT"}</span>${
    context ? `<span class="match-head-sub">${context}</span>` : ""
  }`;
  const html = `
      ${stickyHeadHtml(title)}
      <div class="match-card">
        ${result.championBonus ? `<div class="match-sub champion-line">👑 Champion! +$${fmt(result.championBonus.cash)} · ${signedNum(result.championBonus.rating)} rating</div>` : ""}
        ${matchResultTableHtml(result)}
        ${headToHeadHtml(result)}
        <details class="match-more">
          <summary>How was my performance calculated?</summary>
          ${performanceTableHtml(result.breakdown)}
          <div class="match-sub perf-note">Each point of performance above 70 adds 3 to your match-day rating; below 70 it costs 3.</div>
        </details>
        <details class="match-more">
          <summary>How do luck and win chance work?</summary>
          ${luckExplainerHtml(result)}
        </details>
        <details class="match-more">
          <summary>How were rating and cash worked out?</summary>
          ${ratingCashExplainerHtml(result)}
        </details>
        ${extraHtml}
        <button class="primary-btn" id="matchOk">${continueLabel}</button>
      </div>`;
  openModal(html, { ownClose: true });
  $("matchOk").addEventListener("click", onContinue || closeModal);
}

/* ---------------------------------------------------------------------- */
/* Year-end cash flow summary — shown right as a new year begins          */
/* ---------------------------------------------------------------------- */
function showYearSummaryModal(summary, extraHtml = "") {
  const html = `
    <div class="match-card">
      <div class="match-result neutral">YEAR ${summary.year} COMPLETE</div>
      <div class="match-sub">Cash flow for the year — this is what a new year starting means</div>
      <div class="match-stats">
        <div><b>$${fmt(summary.workPay)}</b>Work/Pro pay</div>
        <div><b>$${fmt(summary.matchCash)}</b>Match winnings</div>
        <div><b>-$${fmt(summary.expenses)}</b>Expenses</div>
      </div>
      <div class="match-sub">Net for the year: <b style="color:${summary.net >= 0 ? "var(--accent)" : "var(--danger)"}">${fmtSigned(summary.net, 0)}</b></div>
      <div class="match-sub">Cash now: ${fmtMoney(summary.cashNow)}</div>
      ${summary.newPayRate != null ? `<div class="match-sub">📈 Annual raise: pay is now $${fmt(summary.newPayRate)}/day</div>` : ""}
      ${extraHtml}
      <button class="primary-btn" id="yearSummaryOk">Continue</button>
    </div>`;
  openModal(html);
  $("yearSummaryOk").addEventListener("click", closeModal);
}

/* ---------------------------------------------------------------------- */
/* Shop / Menu                                                            */
/* ---------------------------------------------------------------------- */
// ---- Staff screen ----
// Staff are hired one week at a time. Wages are paid up front (pro-rated
// if you hire mid-week) and nothing renews: each new week you choose again,
// usually around that week's focus skills.
function weekEndPhrase() {
  const d = daysLeftInWeek();
  const inSeasonWeek = state.seasonPhase === "regular" || state.seasonPhase === "playoffs";
  const end = inSeasonWeek ? "after the next match" : "at the end of the week";
  return `${d} day${d === 1 ? "" : "s"} left — contracts end ${end}`;
}

function staffCardHtml(key, { focus = false, statNote = "" } = {}) {
  const u = UPGRADES[key];
  const owned = upgradeLevel(key);
  const hired = hiredLevel(key);
  const max = u.levels.length;
  const isCoach = key.startsWith("coach_");
  const days = daysLeftInWeek();
  const hireChips = hired
    ? `<span class="staff-hired">✓ Hired Lv${hired} this week</span>`
    : Array.from({ length: owned }, (_, i) => i + 1)
        .reverse()
        .map((lvl) => {
          const cost = staffCost(key, lvl);
          return `<button class="staff-chip" data-hire="${key}" data-level="${lvl}" ${state.cash < cost ? "disabled" : ""}>Hire Lv${lvl} · $${fmt(cost)}</button>`;
        })
        .join("");
  let unlockHtml = "";
  if (owned < max) {
    const next = u.levels[owned];
    unlockHtml = canUnlockLevel(key, owned + 1)
      ? `<button class="staff-chip staff-unlock" data-unlock="${key}" ${state.cash < next.cost ? "disabled" : ""}>Unlock Lv${owned + 1} · $${fmt(next.cost)}</button>`
      : `<span class="staff-locked">🔒 Lv${owned + 1} unlocks in League ${next.unlock}</span>`;
  }
  const lvlInfo = (lvl) => `Lv${lvl}: ${u.levels[lvl - 1].desc} · $${fmt(u.levels[lvl - 1].wage)}/wk`;
  const showLevels = Array.from({ length: Math.min(owned + 1, max) }, (_, i) => i + 1);
  const unhired = isCoach ? `<div class="staff-line">No coach: trains up to ${BAL.skillShopCapBase}, no bonus</div>` : "";
  return `
  <div class="shop-item staff-card${focus ? " shop-item-match" : ""}${hired ? " staff-card-hired" : ""}">
    <div class="shop-item-icon">${u.icon}</div>
    <div class="shop-item-info">
      <div class="shop-item-name">${u.name}${focus ? ` <span class="shop-item-match-tag">This week</span>` : ""}</div>
      ${statNote ? `<div class="staff-line staff-stat">${statNote}</div>` : ""}
      ${unhired}
      ${showLevels.map((lvl) => `<div class="staff-line ${lvl > owned ? "staff-line-locked" : ""}">${lvlInfo(lvl)}</div>`).join("")}
      <div class="staff-actions">${hireChips}${unlockHtml}</div>
      ${!hired && days < 7 ? `<div class="staff-line staff-prorate">Pro-rated: ${days}/7 of the weekly wage</div>` : ""}
    </div>
  </div>`;
}

function shopHtml() {
  const focusKeys = inSeason() ? state.activeSkills : [];
  const order = SKILLS.map((sk) => sk.key).sort((x, y) => focusKeys.includes(y) - focusKeys.includes(x));
  const skillCards = order
    .map((k) => staffCardHtml(coachKey(k), { focus: focusKeys.includes(k), statNote: `${skillMeta(k).icon} ${skillMeta(k).name} now ${fmt(state.stats.skills[k])} · league cap ${leagueSkillCap()}` }))
    .join("");
  const supportKeys = ["physio", "nutritionist", "manager", "sleepApp", "meditation", "recovery"];
  const supportCards = supportKeys
    .map((k) => staffCardHtml(k, { statNote: k === "physio" ? `🏃 Health now ${fmt(state.stats.phys)} · ceiling ${BAL.statCapBase} without a physio` : "" }))
    .join("");
  const weeklyTotal = Object.entries(state.staff.hired).reduce((a, [k, l]) => a + UPGRADES[k].levels[l - 1].wage, 0);
  return `
    ${stickyHeadHtml("Staff", `<span class="sticky-cash">💰 ${fmtMoney(state.cash)}</span>`)}
    <div class="callout staff-week">
      <div class="callout-label">This week</div>
      <div>${weekEndPhrase()}. Pay up front; nothing renews — hire again each week.</div>
      ${weeklyTotal ? `<div class="staff-line">Hired staff cost $${fmt(weeklyTotal)}/wk at full rate.</div>` : ""}
    </div>
    <div class="modal-section">
      <h3>Skill Coaches</h3>
      <p class="modal-sub">A hired coach lifts that skill's ceiling (never past your league cap of ${leagueSkillCap()}) and speeds up its training. With no coach a skill trains up to ${BAL.skillShopCapBase}; above that, an hour a day holds it and less lets it slip.</p>
      ${skillCards}
    </div>
    <div class="modal-section">
      <h3>Support Team</h3>
      <p class="modal-sub">Support staff only help in the weeks they're hired. Higher levels unlock in League 3 and League 1.</p>
      ${supportCards}
    </div>`;
}

function openShop(opts) {
  // Called directly as a click handler too, so opts may be an Event.
  const keepScroll = !!(opts && opts.keepScroll === true);
  openModal(shopHtml(), { ownClose: true, keepScroll });
  const refresh = () => {
    saveState();
    openShop({ keepScroll: true });
    // Coaches change skill ceilings and the physio the Health ceiling —
    // the planner behind the sheet shows both, so refresh it too.
    renderTopbar();
    renderStats();
    renderPlanner();
  };
  document.querySelectorAll("[data-hire]").forEach((btn) =>
    btn.addEventListener("click", () => {
      if (hireStaff(btn.dataset.hire, Number(btn.dataset.level))) refresh();
    })
  );
  document.querySelectorAll("[data-unlock]").forEach((btn) =>
    btn.addEventListener("click", () => {
      if (unlockStaffLevel(btn.dataset.unlock)) refresh();
    })
  );
}

function openMenu() {
  const html = `
    <h2>Menu</h2>
    <div class="menu-row" id="menuShop"><span>🧑‍🏫 Staff</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuCareer"><span>📈 Career &amp; Season</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuLeagues"><span>🏅 Leagues</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuHistory"><span>📜 Match History</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuRename"><span>✏️ Rename Player</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuHow"><span>❓ How to Play</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuInstall"><span>📲 Add to Home Screen</span><span class="arrow">›</span></div>
    <div class="menu-row" id="menuSaveTransfer"><span>💾 Export / Import Save</span><span class="arrow">›</span></div>
    <button class="ghost-btn" id="menuReset">Reset Career</button>
    <div class="version-tag">Cell Grind v${APP_VERSION}</div>
  `;
  openModal(html);
  $("menuShop").addEventListener("click", openShop);
  $("menuCareer").addEventListener("click", openCareer);
  $("menuLeagues").addEventListener("click", () => openLeagues(state.leagueTier));
  $("menuHistory").addEventListener("click", () => openHistory());
  $("menuRename").addEventListener("click", () => openNameModal(false));
  $("menuHow").addEventListener("click", openHowTo);
  $("menuInstall").addEventListener("click", openInstall);
  $("menuSaveTransfer").addEventListener("click", openSaveTransfer);
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
      <input type="text" id="playerNameInput" maxlength="20" placeholder="Your name" value="${state.playerName ? escapeHtml(state.playerName) : ""}" class="name-input" />
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

function employmentSectionHtml() {
  const emp = state.employment;
  const lives = livesRemaining();
  // Chances lost come back strikeWindowDays later — detail tucked away in a
  // collapsed row so the section stays short.
  const returning = emp.strikes
    .map((s) => ({ amount: s.amount != null ? s.amount : 1, lost: s.day, back: s.day + BAL.strikeWindowDays }))
    .sort((a, b) => a.back - b.back);
  const strikeLines = returning.length
    ? `<details class="match-more emp-returns"><summary>Chances coming back (${fmt1(returning.reduce((a, r) => a + r.amount, 0))})</summary>${returning
        .map((r) => `<div class="kv-row"><span>+${fmt1(r.amount)} on Day ${r.back}</span><span class="kv-dim">lost Day ${r.lost} · in ${r.back - state.day} day${r.back - state.day === 1 ? "" : "s"}</span></div>`)
        .join("")}</details>`
    : "";
  // Shown regardless of status — it's charged every day no matter what, so
  // it belongs here rather than cluttering the main planner row, which
  // already shows the pay side of the ledger.
  const expensesLine = `<p class="modal-sub">💸 Cost of living: $${fmt(BAL.dailyExpenses)}/day, every day, regardless of employment status.</p>`;

  if (emp.status === "unemployed") {
    return `${expensesLine}<p>🔍 <b>Unemployed</b> — job searching: ${fmt(emp.jobSearchHours)} / ${jobSearchTarget()}h accumulated (each search needs a random ${BAL.jobSearchHoursRange[0]}–${BAL.jobSearchHoursRange[1]}h). Every hour counts toward it, and you need at least ${BAL.jobSearchMinHours}h a day to end the day.</p>`;
  }

  if (emp.status === "pro") {
    const queue = emp.techniqueQueue;
    const penalty = Math.round(Math.min(1, queue.length * BAL.techniquePenaltyPerUnmastered) * 100);
    const queueHtml = queue.length
      ? queue.map((t, i) => `${i === 0 ? "▶" : "⏸"} ${t.name}: ${fmt(t.hoursDone)}/${t.hoursNeeded}h`).join("<br>")
      : "Fully caught up — no match penalty.";
    return `${expensesLine}<p>🏆 <b>Pro</b> · Pro Duties ${BAL.proDutyHoursRequired}h/day required ($${fmt(emp.proPay)}/day) · ${fmt1(lives)}/${fmt1(BAL.strikesToFire)} chances</p>${strikeLines}
      <p><b>Technique queue</b> (hours above ${BAL.proDutyHoursRequired}h/day go here)${penalty > 0 ? ` — currently <b>-${penalty}%</b> match performance` : ""}:<br>${queueHtml}</p>`;
  }

  const goProHint = checkGoProEligible()
    ? "Thresholds met — going pro next time a day resolves."
    : `Go pro at League ${BAL.goProLeagueTier} or higher with $${BAL.goProCash}+ banked (currently League ${state.leagueTier}, ${fmtMoney(state.cash)}).`;
  return `${expensesLine}<p>💼 <b>Employed</b> · Work ${BAL.workHoursRequired}h/day required ($${fmt(emp.workPay)}/day) · ${fmt1(lives)}/${fmt1(BAL.strikesToFire)} chances</p>${strikeLines}
    <p class="modal-sub">${goProHint}</p>`;
}

// Same 0-100 bar layout as the Career skill list, so stats and skills read
// alike. Only Health has a cap below 100 (Sports Physio), hatched and named.
function currentStatsHtml() {
  const s = state.stats;
  const physioLvl = hiredLevel("physio");
  const pCap = physCap();
  const rows = [
    { icon: "🏃", name: "Physical Health", val: s.phys, cap: pCap, cls: "phys", note: pCap < 100 ? `cap: ${physioLvl > 0 ? `Physio Lv${physioLvl}` : "no Physio"}` : "" },
    { icon: "🌙", name: "Rest", val: s.rest, cap: 100, cls: "rest", note: "" },
    { icon: "🎮", name: "Composure", val: s.composure, cap: 100, cls: "composure", note: "" },
    { icon: "🥗", name: "Nutrition", val: s.nutrition, cap: 100, cls: "nutrition", note: `${dailyHoursCap(s.nutrition)}h today` },
  ];
  return rows
    .map(
      (r) => `
        <div class="skill-row-detail">
          <div class="skill-row-detail-label"><span>${r.icon} ${r.name}</span><span>${fmt(r.val)}/${r.cap}${r.note ? ` · ${r.note}` : ""}</span></div>
          <div class="bar"><div class="bar-fill ${r.cls}" style="width:${clamp(r.val, 0, 100)}%"></div>${lockedZoneHtml(r.cap, 100)}</div>
        </div>`
    )
    .join("");
}

function openCareer() {
  const info = getNextMatchInfo();
  let nextSection;
  if (info.kind === "fixture" || info.kind === "playoff") {
    const rival = info.opponentRivalId != null ? findRival(info.opponentRivalId) : null;
    // Their league is only worth saying if it isn't yours (it always is,
    // outside of edge cases) — the record is what tells you something.
    const opponentExtra = rival
      ? `${rival.league !== state.leagueTier ? ` · League ${rival.league}` : ""} · ${rival.wins}–${rival.losses}`
      : "";
    const skillsForMatch = state.activeSkills
      .map((k) => {
        const meta = skillMeta(k);
        return `<div class="kv-row"><span>${meta.icon} ${meta.name}</span><span>${fmt(state.stats.skills[k])}/${skillCap(k)}</span></div>`;
      })
      .join("");
    const matchName = info.kind === "playoff" ? PLAYOFF_ROUND_NAMES[state.playoff.stage] : `Round ${state.roundIndex + 1}/${BAL.seasonRounds}`;
    const when = info.daysUntil <= 0 ? "today" : daysUntilPhrase(info.daysUntil);
    nextSection = `
      <div class="kv-row"><span>Match</span><span>${matchName} · ${when}</span></div>
      <div class="kv-row"><span>Opponent</span><span>${info.opponentName} · ${info.opponentRating}${opponentExtra}</span></div>
      <div class="kv-row"><span>Win chance</span><span><b>${info.winPct}%</b> · performance ${fmt(performanceScore())}/100</span></div>
      <div class="callout">
        <div class="callout-label">Tested this match</div>
        ${skillsForMatch}
      </div>`;
  } else {
    nextSection = `<p>${info.label}</p>`;
  }

  const seasonWins = state.seasonResults.filter((r) => r.win).length;
  const seasonPlayed = state.seasonResults.length;
  const tablePos = playerTablePosition();

  let standingsSection = "";
  if (state.lastStandings) {
    const top5 = state.lastStandings.slice(0, 5);
    standingsSection = `
    <div class="modal-section">
      <h3>${state.seasonPhase === "playoffs" || state.seasonPhase === "offseason" ? "This Season" : "Last Season"} — League ${state.lastStandingsTier} Final Standings</h3>
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
    ${stickyHeadHtml("Career &amp; Season")}
    <div class="modal-section">
      <h3>Next Up</h3>
      ${nextSection}
    </div>
    <div class="modal-section">
      <h3>This Season</h3>
      <div class="kv-row"><span>Year ${state.year} · League ${state.leagueTier}</span><span>Played ${seasonPlayed}/${BAL.seasonRounds}</span></div>
      <div class="kv-row"><span>Record</span><span>${seasonWins}–${seasonPlayed - seasonWins}${tablePos ? ` · #${tablePos.pos} of ${tablePos.size}` : ""}</span></div>
      ${state.peakLeagueTier !== state.leagueTier ? `<div class="kv-row"><span>Highest league reached</span><span>League ${state.peakLeagueTier}</span></div>` : ""}
      <div class="menu-row" id="viewLeaguesLink"><span>📊 View League Standings</span><span class="arrow">›</span></div>
      <div class="menu-row" id="viewHistoryLink"><span>📜 Match History</span><span class="arrow">›</span></div>
    </div>
    ${playoffSection}
    ${standingsSection}
    <div class="modal-section">
      <h3>Career — ${escapeHtml(state.playerName || "Player")}</h3>
      <div class="kv-row"><span>Rating</span><span>${fmt(state.rank)} · peak ${fmt(state.peakRank)}</span></div>
      ${state.wins !== seasonWins || state.losses !== seasonPlayed - seasonWins ? `<div class="kv-row"><span>All-time record</span><span>${state.wins}–${state.losses}</span></div>` : ""}
    </div>
    <div class="modal-section">
      <h3>Employment</h3>
      ${employmentSectionHtml()}
    </div>
    <div class="modal-section">
      <h3>Skills${!inSeason() ? ` — ${isPreseason() ? "preseason" : "off-season"}: train anything` : ""}</h3>
      ${inSeason() ? `<p class="modal-sub skills-legend">🟢 = this week's focus, the only skills you can train and the ones tested in the next match.</p>` : ""}
      ${SKILLS.map((sk) => {
        const cap = skillCap(sk.key);
        const val = state.stats.skills[sk.key];
        const pct = clamp(val, 0, 100);
        const active = trainableSkills().includes(sk.key);
        const cause = skillCapCause(sk.key);
        return `
        <div class="skill-row-detail ${active ? "skill-row-detail-active" : ""}">
          <div class="skill-row-detail-label"><span>${sk.icon} ${sk.name}${active ? " 🟢" : ""}</span><span>${fmt(val)}/${cap}${cause ? ` · cap: ${cause}` : ""}</span></div>
          <div class="bar"><div class="bar-fill skill" style="width:${pct}%"></div>${lockedZoneHtml(cap, 100)}</div>
        </div>`;
      }).join("")}
    </div>
    <div class="modal-section">
      <h3>Current Stats</h3>
      ${currentStatsHtml()}
    </div>`;
  openModal(html, { ownClose: true });
  $("viewLeaguesLink").addEventListener("click", () => openLeagues(state.leagueTier));
  $("viewHistoryLink").addEventListener("click", () => openHistory());
}

function leagueTableHtml(tier) {
  let standings;
  let noteText;
  let isSnapshot = false;
  if (state.roundIndex >= BAL.seasonRounds && state.leagueStandings[tier]) {
    // Season over: read the snapshot taken at round 39. Rebuilding from the
    // rivals' live league field would be wrong once promotion has moved them.
    standings = state.leagueStandings[tier];
    noteText = "Final standings";
    isSnapshot = true;
  } else if (state.roundIndex > 0) {
    // Live: every league's round-robin resolves in lockstep every week, so
    // this is a real "right now" snapshot, not just a post-season report.
    standings = buildStandingsFromPoints(tier);
    noteText = `Live — through round ${state.roundIndex}/${BAL.seasonRounds}`;
  } else if (state.leagueStandings[tier]) {
    standings = state.leagueStandings[tier];
    noteText = "Last season's final standings";
    isSnapshot = true;
  } else {
    return `<p class="modal-sub">Not yet available — this table fills in once this league's first round resolves.</p>`;
  }
  // League 1 is already the top — no promotion zone. League 5 is already
  // the bottom — no relegation zone. Matches applyPromotionRelegation()'s
  // own tier > 1 / tier < leagueCount gating exactly. Only the top 3 are
  // guaranteed; the 4th spot goes to the playoff champion, marked once known.
  const showPromo = tier > 1;
  const showReleg = tier < BAL.leagueCount;
  const champ = isSnapshot && state.playoffChampions ? state.playoffChampions[tier] : null;
  const rows = standings
    .map((t, i) => {
      const pos = i + 1;
      const isChamp = isSameEntry(t, champ);
      const zone =
        showPromo && (pos <= BAL.promotionTablePlaces || isChamp)
          ? "zone-promo"
          : showReleg && pos > standings.length - BAL.relegationCount
          ? "zone-releg"
          : "";
      const playoff = pos <= BAL.playoffSize ? "zone-playoff" : "";
      // Dashed divider straight after the last qualifying place.
      const cutLine = pos === BAL.playoffSize ? `<div class="league-cutline">Playoff line — top ${BAL.playoffSize} qualify</div>` : "";
      return `
      <div class="league-row ${zone} ${playoff} ${t.isPlayer ? "league-row-you" : ""}" ${!t.isPlayer && t.rivalId != null ? `data-rival="${t.rivalId}"` : ""}>
        <span class="league-pos">${pos}</span>
        <span class="league-name">${isChamp ? "🏆 " : ""}${t.name}${t.isPlayer ? " (You)" : ""}</span>
        <span class="league-rating">${fmt(t.rating)}</span>
        <span class="league-points">${t.points}</span>
      </div>${cutLine}`;
    })
    .join("");
  const legendParts = [`<span class="legend-dot legend-playoff"></span> Playoffs (top ${BAL.playoffSize})`];
  if (showPromo) legendParts.push(`<span class="legend-dot legend-promo"></span> Promoted (top ${BAL.promotionTablePlaces} + 🏆 playoff champion)`);
  if (showReleg) legendParts.push(`<span class="legend-dot legend-releg"></span> Relegation zone`);
  const legend = legendParts.length ? `<p class="modal-sub league-legend">${legendParts.join(" · ")}</p>` : "";
  return `
    <p class="modal-sub">${noteText}</p>
    <div class="league-table-header">
      <span class="league-pos">#</span>
      <span class="league-name">Name</span>
      <span class="league-rating">Rating</span>
      <span class="league-points">Pts</span>
    </div>
    <div class="league-table">${rows}</div>
    ${legend}`;
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
    <div class="menu-row" id="leaguesHistoryLink"><span>📜 Results by round</span><span class="arrow">›</span></div>
    <div id="leagueTableContainer">${leagueTableHtml(startTier)}</div>`;
  openModal(html);
  let shownTier = startTier;
  $("leaguesHistoryLink").addEventListener("click", () => openHistory({ tab: "league", view: { tier: shownTier } }));
  // Tap a rival in the table to see their season (delegated: the table is
  // rebuilt whenever the tab changes).
  $("leagueTableContainer").addEventListener("click", (e) => {
    const row = e.target.closest("[data-rival]");
    if (row) openRival(Number(row.dataset.rival), () => openLeagues(shownTier));
  });
  document.querySelectorAll(".league-tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".league-tab").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      const tier = Number(btn.getAttribute("data-tier"));
      shownTier = tier;
      $("leagueTableContainer").innerHTML = leagueTableHtml(tier);
    });
  });
}


/* ---------------------------------------------------------------------- */
/* Match History + rival pages                                            */
/* ---------------------------------------------------------------------- */
function historyName(id) {
  if (id === HISTORY_PLAYER_ID) return `${escapeHtml(state.playerName || "You")} (You)`;
  const r = findRival(id);
  return r ? r.name : "Unknown";
}
function historyNameLink(id) {
  if (id === HISTORY_PLAYER_ID) return `<b>${historyName(id)}</b>`;
  return `<button class="link-name" data-rival="${id}">${historyName(id)}</button>`;
}
function roundKeyLabel(key) {
  return PLAYOFF_ROUND_NAMES[key] || `Round ${key}`;
}
function roundShort(label) {
  if (!label) return "";
  const m = label.match(/^Round (\d+)/);
  if (m) return `R${m[1]}`;
  return { "Round of 16": "R16", Quarterfinal: "QF", Semifinal: "SF", Final: "F" }[label] || label;
}

function myMatchesHtml() {
  const mine = ensureHistory().mine;
  if (!mine.length) return `<p class="modal-sub">No matches played yet. Every match you play from now on is kept here — tap one to see its full result again.</p>`;
  const years = [...new Set(mine.map((m) => m.year))].sort((a, b) => b - a);
  return years
    .map((y) => {
      const list = mine.map((m, i) => ({ m, i })).filter((x) => x.m.year === y).reverse();
      const w = list.filter((x) => x.m.win).length;
      const tiers = [...new Set(list.map((x) => x.m.tier).filter(Boolean))];
      const rows = list
        .map(({ m, i }) => {
          const n = (x) => (x < 0 ? `−${-x}` : `${x}`);
          const score = m.sides ? `${n(m.sides.you.final)}–${n(m.sides.opp.final)}` : "—";
          const delta = m.ratingAfter != null ? signedNum(m.ratingAfter - m.rankBefore) : "";
          return `<div class="hist-row ${m.legacy ? "hist-legacy" : ""}" ${m.legacy ? "" : `data-mi="${i}"`}>
            <span class="hist-round">${roundShort(m.roundLabel)}</span>
            <span class="hist-wl ${m.win ? "win" : "loss"}">${m.win ? "W" : "L"}</span>
            <span class="hist-opp">${m.opponentName || "?"}</span>
            <span class="hist-score">${score}</span>
            <span class="hist-delta ${deltaClass(m.ratingAfter != null ? m.ratingAfter - m.rankBefore : 0)}">${delta}</span>
          </div>`;
        })
        .join("");
      const legacyNote = list.some((x) => x.m.legacy) ? `<p class="modal-sub">Matches without a score were played before full results were saved.</p>` : "";
      return `<div class="modal-section">
        <h3>Year ${y}${tiers.length ? ` · League ${tiers.join(" → ")}` : ""} · ${w}–${list.length - w}</h3>
        <div class="hist-list">${rows}</div>${legacyNote}
      </div>`;
    })
    .join("");
}

// Which season / league / round the League results tab shows by default:
// your league, the latest round with results.
function defaultLeagueView(view = {}) {
  const league = ensureHistory().league;
  const years = Object.keys(league).map(Number).sort((a, b) => b - a);
  const year = view.year != null && league[view.year] ? view.year : years[0];
  if (year == null) return null;
  const tiers = league[year];
  // Default to the league you played in that season (you may have been
  // promoted or relegated since), else your current one.
  const playedIn = Object.keys(tiers).map(Number).find((t) => Object.values(tiers[t]).some((rs) => rs.some((x) => x[0] === HISTORY_PLAYER_ID || x[1] === HISTORY_PLAYER_ID)));
  const tier = view.tier && tiers[view.tier] ? view.tier : playedIn || (tiers[state.leagueTier] ? state.leagueTier : Number(Object.keys(tiers)[0]));
  const keys = HISTORY_ROUND_KEYS.filter((k) => tiers[tier] && tiers[tier][k]);
  const round = view.round && keys.includes(view.round) ? view.round : keys[keys.length - 1];
  return { year, tier, round, years, keys };
}

function leagueResultsHtml(view) {
  const v = defaultLeagueView(view);
  if (!v) return `<p class="modal-sub">No results yet. Every match in all 5 leagues is recorded from now on — this season and last are kept.</p>`;
  const league = ensureHistory().league;
  const results = (league[v.year][v.tier] && league[v.year][v.tier][v.round]) || [];
  const idx = v.keys.indexOf(v.round);
  const yearPills = v.years.length > 1
    ? `<div class="league-tabs">${v.years.map((y) => `<button class="league-tab ${y === v.year ? "active" : ""}" data-hyear="${y}">Year ${y}${y === state.year ? " (now)" : ""}</button>`).join("")}</div>`
    : "";
  const tierTabs = `<div class="league-tabs">${[1, 2, 3, 4, 5]
    .map((t) => `<button class="league-tab ${t === v.tier ? "active" : ""}" data-htier="${t}" ${league[v.year][t] ? "" : "disabled"}>L${t}</button>`)
    .join("")}</div>`;
  const stepper = `<div class="round-stepper">
      <button class="icon-btn" data-hround="${idx > 0 ? v.keys[idx - 1] : ""}" ${idx > 0 ? "" : "disabled"} aria-label="Previous round">◀</button>
      <span>${roundKeyLabel(v.round)}</span>
      <button class="icon-btn" data-hround="${idx < v.keys.length - 1 ? v.keys[idx + 1] : ""}" ${idx < v.keys.length - 1 ? "" : "disabled"} aria-label="Next round">▶</button>
    </div>`;
  const rows = results
    .map(([w, l, wr, lr]) => `<div class="res-row ${w === HISTORY_PLAYER_ID || l === HISTORY_PLAYER_ID ? "res-you" : ""}">
        <span class="res-side">${historyNameLink(w)} <span class="res-rt">${wr}</span></span>
        <span class="res-beat">beat</span>
        <span class="res-side">${historyNameLink(l)} <span class="res-rt">${lr}</span></span>
      </div>`)
    .join("");
  return `${yearPills}${tierTabs}${stepper}
    <div class="res-list">${rows}</div>
    <p class="modal-sub">Ratings are as they stood going into the match. Tap a name to see that rival's season.</p>`;
}

function openHistory(opts = {}) {
  const tab = opts.tab || "mine";
  const view = opts.view || {};
  const html = `
    ${stickyHeadHtml("Match History")}
    <div class="league-tabs history-tabs">
      <button class="league-tab ${tab === "mine" ? "active" : ""}" data-htab="mine">My matches</button>
      <button class="league-tab ${tab === "league" ? "active" : ""}" data-htab="league">League results</button>
    </div>
    <div id="historyBody">${tab === "mine" ? myMatchesHtml() : leagueResultsHtml(view)}</div>`;
  openModal(html, { ownClose: true, keepScroll: !!opts.keepScroll });
  const body = $("modalBody");
  body.querySelectorAll("[data-htab]").forEach((b) => b.addEventListener("click", () => openHistory({ tab: b.dataset.htab })));
  const cur = tab === "league" ? defaultLeagueView(view) : null;
  body.querySelectorAll("[data-hyear]").forEach((b) => b.addEventListener("click", () => openHistory({ tab, view: { year: Number(b.dataset.hyear), tier: cur && cur.tier }, keepScroll: true })));
  body.querySelectorAll("[data-htier]").forEach((b) => b.addEventListener("click", () => openHistory({ tab, view: { year: cur.year, tier: Number(b.dataset.htier) }, keepScroll: true })));
  body.querySelectorAll("[data-hround]").forEach((b) => b.addEventListener("click", () => b.dataset.hround && openHistory({ tab, view: { year: cur.year, tier: cur.tier, round: b.dataset.hround }, keepScroll: true })));
  body.querySelectorAll("[data-rival]").forEach((b) =>
    b.addEventListener("click", () => openRival(Number(b.dataset.rival), () => openHistory({ tab, view: cur ? { year: cur.year, tier: cur.tier, round: cur.round } : view })))
  );
  body.querySelectorAll("[data-mi]").forEach((row) =>
    row.addEventListener("click", () => {
      const m = ensureHistory().mine[Number(row.dataset.mi)];
      showMatchModal(m, "", { onContinue: () => openHistory({ tab: "mine" }), continueLabel: "◀ Back to Match History" });
    })
  );
}

// A rival's season so far, read straight from the stored league results —
// nothing extra is saved for this page.
function openRival(id, back) {
  const r = findRival(id);
  if (!r) return;
  const yearData = ensureHistory().league[state.year] || {};
  const games = [];
  Object.keys(yearData).forEach((tier) =>
    HISTORY_ROUND_KEYS.forEach((key) =>
      (yearData[tier][key] || []).forEach(([w, l, wr, lr]) => {
        if (w === id) games.push({ key, win: true, opp: l, oppRating: lr, rating: wr });
        else if (l === id) games.push({ key, win: false, opp: w, oppRating: wr, rating: lr });
      })
    )
  );
  const w = games.filter((g) => g.win).length;
  const rows = games
    .slice()
    .reverse()
    .map((g) => `<div class="hist-row">
        <span class="hist-round">${roundShort(roundKeyLabel(g.key))}</span>
        <span class="hist-wl ${g.win ? "win" : "loss"}">${g.win ? "W" : "L"}</span>
        <span class="hist-opp">${g.win ? "beat" : "lost to"} ${historyName(g.opp)}</span>
        <span class="hist-score">${g.oppRating}</span>
      </div>`)
    .join("");
  const html = `
    ${stickyHeadHtml(r.name)}
    <div class="modal-section">
      <div class="kv-row"><span>League</span><span>League ${r.league}</span></div>
      <div class="kv-row"><span>Rating</span><span>${fmt(r.rating)}</span></div>
      <div class="kv-row"><span>This season</span><span>${w}–${games.length - w}</span></div>
      <div class="kv-row"><span>All-time record</span><span>${r.wins}–${r.losses}</span></div>
    </div>
    <div class="modal-section">
      <h3>Year ${state.year} results</h3>
      ${games.length ? `<div class="hist-list">${rows}</div><p class="modal-sub">Number = the opponent's rating going into the match.</p>` : `<p class="modal-sub">No results recorded for them this season yet.</p>`}
    </div>
    ${back ? `<button class="ghost-btn" id="rivalBack">◀ Back</button>` : ""}`;
  openModal(html, { ownClose: true });
  if (back) $("rivalBack").addEventListener("click", back);
}

function openHowTo() {
  const html = `
    ${stickyHeadHtml("How to Play")}
    <div class="modal-section">
      <p>You manage a rising Excel esports competitor. Every day has up to 24 hours — split them across:</p>
      <p>
      📈🗺️📝🎲🔢⏱️🃏 <b>Skill Training</b> — 7 case specialties (Data, Mapping, Text, Game Logic, Math, Time, Cards). During preseason and the off-season, all 7 are open for training. Once the regular season starts, only 1-3 are "active" each round, revealed at the start of that round's week — the rest can't be trained until they come up again.<br>
      🏃 <b>Gym</b> — raises Physical Health.<br>
      🌙 <b>Sleep</b> — builds Rest.<br>
      🎮 <b>Relax</b> — builds Composure and prevents burnout.<br>
      🥗 <b>Food</b> — builds Nutrition, which keeps tomorrow's day at full length.<br>
      💼 <b>Work</b> — pays the bills and keeps you employed.
      </p>
      <p><b>It's all connected:</b> low Rest wears down Physical Health even if you train well, and low Physical Health caps how much your skill training actually helps. Training hard without Relax drains Composure — hit 0 and you burn out, tanking your effectiveness until it recovers.</p>
      <p><b>Decay:</b> every stat needs upkeep or it slips. Any skill that isn't active this round rusts; an active skill still rusts below ${BAL.skillDecayThresholdHours}h of training. Gym below ${BAL.skillDecayThresholdHours}h detrains Physical Health. Sleep below ${BAL.idealSleep}h drains Rest. Relax below ${BAL.relaxComposureThreshold}h drains Composure. Food below ${BAL.skillDecayThresholdHours}h drains Nutrition. Each slider shows a marker at its threshold, and each bar previews tomorrow's value based on your current plan — green for a gain, red for a loss.</p>
      <p><b>Rest</b> swings training itself: above ${BAL.restTrainingBoostThreshold} it's 150% effective, above ${BAL.restTrainingBoostHigh} it's 200% effective. <b>Composure</b> hits match day specifically — below ${BAL.composureMatchMid} your active skills count for only 75%, below ${BAL.composureMatchLow} just 50%. <b>Nutrition</b> sets how many hours you get at all: below ${BAL.nutritionHoursCapLow} your day shrinks to just ${BAL.dailyHoursFloor}h, sliding up to the full ${BAL.dailyHoursCeiling}h at ${BAL.nutritionHoursCapHigh}+.</p>
      <p><b>Gym injuries:</b> every Gym hour adds a ${+(BAL.injuryChancePerHour * 100).toFixed(1)}% chance of injury that day, so only a 0h day is risk-free — ${BAL.gymMaxHours}h (the most you can do) is a ${+(BAL.gymMaxHours * BAL.injuryChancePerHour * 100).toFixed(1)}% chance. The Gym row shows today's risk as <b>#% 🤕</b>. Sports Physio cuts that risk by 20%, 35% or 50%. An injury costs ${BAL.injuryPhysLoss[0]}–${BAL.injuryPhysLoss[1]} Health and locks the Gym for ${BAL.injuryDaysRange[0]}–${BAL.injuryDaysRange[1]} days (Recovery Program takes 1–3 days off, minimum 1); the row shows <b>🤕 #d</b> while it heals. Injuries never stop you playing matches — they only shut the Gym.</p>
      <p><b>Staff &amp; ceilings:</b> skills train up to ${BAL.skillShopCapBase} on your own. To go higher, hire that skill's <b>Coach</b> in 🧑‍🏫 Staff: a hired coach lifts the ceiling (Lv1 60 … Lv5 100) and speeds up training — but never past your league cap (the highest league you've reached: 60 in League 5 up to 100 in League 1). Above the ceiling a skill isn't cut down: an hour a day holds it, less lets it slip. Physical Health caps at ${BAL.statCapBase} without a hired Sports Physio. Staff are hired <b>a week at a time</b> (a week ends after each match): you pay the weekly wage up front (pro-rated if you hire mid-week) and nothing renews, so each week you choose who's worth it — usually that week's focus skills. Higher levels cost a one-off fee to unlock, only once your league allows it (coaches: Lv2 in League 4 … Lv5 in League 1; support team: Lv2 in League 3, Lv3 in League 1), and cost more per week. The hatched end of a bar is the part this week's ceiling locks off.</p>
      <p><b>End Day / To Match:</b> <b>End Day ▶</b> plays one day. <b>To Match ▶▶</b> repeats today's plan every day up to and including the next match, then shows the result with a summary of how your stats moved over the week (outside the season it's <b>End Week ▶▶</b>, up to 7 days, to the end of the week). It stops early so you can re-plan if you get injured, lose or find a job (or go pro), burn out, get a new technique to master, or Nutrition drops so far that your plan no longer fits in the day.</p>
      <p><b>The season:</b> a ${BAL.preseasonDays}-day preseason to train, then a ${BAL.seasonRounds}-round regular season — one match a week against a named rival, all scheduled in advance, each testing that week's active skills. Finish in the top ${BAL.playoffSize} of your ${BAL.seasonRounds + 1}-competitor league to reach the knockout playoffs. Lose a playoff match and you're out; win the Final and you're champion.</p>
      <p>Miss the playoffs and your season ends early — but training never stops. You get a ${BAL.trainingCampDays}-day training camp to prepare for next year, the same amount of time a full playoff run would have taken. Get knocked out of the playoffs and you go to training camp too, for the rest of the playoff window (at least ${BAL.offseasonDays} days) — so an early exit gets its time back as training, just like missing the cut. Only the champion gets a plain ${BAL.offseasonDays}-day break.</p>
      <p><b>Match day:</b> both players get a <b>match-day rating</b> = rating + performance + luck. Your performance comes from your stats (each point above 70 adds 3, below 70 costs 3); every rival has a performance on the same scale. Luck is random for both sides every match — usually between about −150 and +120, occasionally +400 or more on an inspired day. The higher match-day rating wins, so the bigger your rating gap the likelier you are to win, but upsets always stay possible. The result screen shows every number side by side.</p>
      <p><b>Leagues:</b> there are ${BAL.leagueCount} leagues, League 1 at the top and League 5 at the bottom — you start in League 5. Every league has a persistent roster of named rivals whose ratings evolve from real simulated results every week, same as yours — every tier's table is live from round 1, not just visible once the season ends. Four go up from every league below League 1: the playoff champion, plus the top ${BAL.promotionTablePlaces} of the table other than the champion — so a top-${BAL.promotionTablePlaces} finish is always promoted, and anyone in the playoffs can still win their way up. Finish bottom ${BAL.relegationCount} and you're relegated. This applies to every competitor in every league, not just you — every league plays out its own knockout too — so the standings you see are a living world, not scenery. Check the Leagues screen any time (Menu, or the shortcut in Career) to see all ${BAL.leagueCount} tables. <b>Match History</b> keeps every match you play (tap one to see its full result again) and every result in all ${BAL.leagueCount} leagues for this season and last — tap a rival's name, there or in a league table, to see their season.</p>
      <p><b>Rating</b> is your skill score (the 🏆 number), the same scale every rival is measured on — it rises and falls with each result, and it's what your win chance is worked out from. Your <b>table position</b> (#1–#40) is separate: it comes from league points, 3 per win. Cash and Rating carry across seasons and leagues — spend cash on Staff each week.</p>
      <p><b>Employment:</b> your day job funds everything else, every phase, no exceptions. Work ${BAL.workHoursRequired}h/day (every phase, preseason included), starting at $${BAL.workPayMin}/day — pay is tied to still <i>having</i> the job, not to hitting the exact hour target every day, so falling short doesn't cost you income, only a chunk of a chance scaled to the shortfall (regained ${BAL.strikeWindowDays} days later). Run out of your ${fmt1(BAL.strikesToFire)} chances and <i>that's</i> when pay actually stops — you're fired: the same slider becomes a Job Search, needing a random ${BAL.jobSearchHoursRange[0]}–${BAL.jobSearchHoursRange[1]} cumulative hours (rolled when you lose the job and shown on the slider) to get rehired — at least ${BAL.jobSearchMinHours}h of searching a day, or the day can't end. Reach League ${BAL.goProLeagueTier} or higher with $${BAL.goProCash}+ banked while employed and you go Pro automatically — Work drops to just ${BAL.proDutyHoursRequired}h/day of Pro Duties, starting at $${BAL.proPayMin}/day, with the same chances rule and the same fallback to Job Search if you're dropped.</p>
      <p><b>Pay &amp; seniority:</b> pay rises $${BAL.payRaisePerYear}/year for your first ${BAL.payRaiseMaxYears} years in a role, then holds — Work tops out at $${BAL.workPayMin + BAL.payRaisePerYear * BAL.payRaiseMaxYears}/day, Pro Duties at $${BAL.proPayMin + BAL.payRaisePerYear * BAL.payRaiseMaxYears}/day. Lose the job or get dropped from Pro and that role's pay resets to its minimum for next time — seniority isn't carried over.</p>
      <p><b>Cost of living:</b> $${BAL.dailyExpenses}/day, charged every single day no matter your employment status — stay employed and you net a profit, but lose your job and the bills don't stop, so cash actively drains while you're out of work. Each new year opens with a summary of that year's full cash flow: pay earned, match winnings, and expenses paid.</p>
      <p>Pros have one more thing to manage: staying current. Roughly every ${BAL.techniqueIntervalDays} days a new Excel technique appears that needs ${BAL.techniqueMinHours}-${BAL.techniqueMaxHours}h to master — any Pro Duties hours beyond the ${BAL.proDutyHoursRequired}h minimum go toward it. Falling behind never costs you progress (new ones just queue up), but every technique still unmastered costs ${Math.round(BAL.techniquePenaltyPerUnmastered * 100)}% match performance, stacking.</p>
    </div>`;
  openModal(html, { ownClose: true });
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
      <p>Already have a career in Safari? The Home Screen app keeps its own separate save — use <b>💾 Export / Import Save</b> in the menu to move it across.</p>
    </div>`;
  openModal(html);
}

/* ---------------------------------------------------------------------- */
/* Export / import save                                                   */
/* ---------------------------------------------------------------------- */
// A save code is the JSON state, gzipped where the browser supports it
// (~85KB -> ~16KB), then base64'd so it survives copy/paste and Notes.
// "CG1z:" = gzipped, "CG1j:" = plain JSON fallback. Raw JSON is accepted too.
const SAVE_CODE_GZ = "CG1z:";
const SAVE_CODE_PLAIN = "CG1j:";

function bytesToBase64(bytes) {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}
function base64ToBytes(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

async function encodeSaveCode() {
  if (state.logEntries.length > LOG_KEEP) state.logEntries = state.logEntries.slice(-LOG_KEEP);
  const json = JSON.stringify(state);
  if (typeof CompressionStream === "function") {
    try {
      const stream = new Blob([json]).stream().pipeThrough(new CompressionStream("gzip"));
      const gz = new Uint8Array(await new Response(stream).arrayBuffer());
      return SAVE_CODE_GZ + bytesToBase64(gz);
    } catch (e) {
      // fall through to the uncompressed form
    }
  }
  return SAVE_CODE_PLAIN + bytesToBase64(new TextEncoder().encode(json));
}

// Errors whose message is written for the player; anything else thrown while
// decoding (bad base64, cut-off gzip, bad JSON) means a damaged code.
class SaveCodeError extends Error {}

// Pasted text on iOS often isn't just the code: the share sheet / Notes can
// put a "Cell Grind save" title in front, a URL or line breaks can ride
// along, and invisible characters sneak in. So find the code wherever it
// is rather than requiring the paste to start with it. Every gzip stream
// begins with the same bytes, which base64 always renders as "H4sI" — that
// finds a compressed code even if its "CG1z:" prefix got lost.
function extractSaveCode(text) {
  const raw = (text || "").replace(/[\u200B-\u200D\u2060\uFEFF]/g, "");
  const trimmed = raw.trim();
  if (trimmed.startsWith("{")) return { kind: "json", body: trimmed };
  // Rejoin a code wrapped over several lines: from where it starts, keep
  // taking whitespace-separated pieces while they're pure base64, so a
  // title before it or a link after it is left out.
  const tokens = raw.split(/\s+/).filter(Boolean);
  const isB64 = (t) => /^[A-Za-z0-9+/=_-]+$/.test(t);
  const runFrom = (i, offset) => {
    let body = tokens[i].slice(offset);
    for (let j = i + 1; j < tokens.length && isB64(tokens[j]); j++) body += tokens[j];
    return body;
  };
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    const gz = t.indexOf("H4sI");
    if (gz !== -1 && isB64(t.slice(gz))) return { kind: "gz", body: runFrom(i, gz) };
    const plain = t.toUpperCase().indexOf(SAVE_CODE_PLAIN.toUpperCase());
    if (plain !== -1) return { kind: "plain", body: runFrom(i, plain + SAVE_CODE_PLAIN.length) };
  }
  const flat = tokens.join("");
  if (/^[A-Za-z0-9+/=_-]{200,}$/.test(flat)) return { kind: "plain", body: flat };
  return null;
}

async function decodeSaveCode(text) {
  if (!(text || "").trim()) throw new SaveCodeError("Paste a save code first.");
  const found = extractSaveCode(text);
  if (!found) {
    const peek = text.replace(/\s+/g, " ").trim().slice(0, 16);
    throw new SaveCodeError(`That doesn't look like a Cell Grind save code (it starts “${peek}…”). Make sure you copied the whole code, from the very start.`);
  }
  // Tolerate URL-safe base64 and lost "=" padding.
  const b64 = (body) => {
    const std = body.replace(/=+$/, "").replace(/-/g, "+").replace(/_/g, "/");
    return base64ToBytes(std + "===".slice((std.length + 3) % 4));
  };
  let json;
  if (found.kind === "json") {
    json = found.body;
  } else if (found.kind === "gz") {
    if (typeof DecompressionStream !== "function") throw new SaveCodeError("This browser can't read compressed save codes — update iOS / your browser and try again.");
    const stream = new Blob([b64(found.body)]).stream().pipeThrough(new DecompressionStream("gzip"));
    json = await new Response(stream).text();
  } else {
    json = new TextDecoder().decode(b64(found.body));
  }
  const parsed = JSON.parse(json);
  const looksValid =
    parsed && typeof parsed === "object" &&
    typeof parsed.day === "number" &&
    parsed.stats && typeof parsed.stats === "object" &&
    Array.isArray(parsed.rivals);
  if (!looksValid) throw new SaveCodeError("That save code is incomplete or from something else.");
  return parsed;
}

function saveFileName() {
  const who = (state.playerName || "career").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-|-$/g, "") || "career";
  return `cellgrind-${who}-day${state.day}.txt`;
}

function openSaveTransfer() {
  const canShare = typeof navigator.share === "function";
  const html = `
    <h2>Export / Import Save</h2>
    <div class="modal-section">
      <p class="save-transfer-note">Your career is stored on this device only, and Safari and the Home Screen app each keep their own save. To move a career, export it in one place and import it in the other.</p>
    </div>
    <div class="modal-section">
      <h3>Export</h3>
      <p class="save-transfer-note">${state.playerName ? escapeHtml(state.playerName) + " · " : ""}Day ${state.day} · ${fmtMoney(state.cash)} · Rating ${state.rank}</p>
      <button class="primary-btn" id="saveExportCopy">📋 Copy save code</button>
      ${canShare ? `<button class="ghost-btn" id="saveExportShare">📤 Share / Save to Files</button>` : `<button class="ghost-btn" id="saveExportDownload">⬇️ Download save file</button>`}
      <div class="save-transfer-status" id="saveExportStatus"></div>
    </div>
    <div class="modal-section">
      <h3>Import</h3>
      <textarea id="saveImportText" class="name-input save-code-input" rows="4" placeholder="Paste a save code here" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false"></textarea>
      <button class="primary-btn" id="saveImportBtn">Load pasted code</button>
      <label class="ghost-btn save-file-btn">📂 Load from file<input type="file" id="saveImportFile" accept=".txt,.json,text/plain,application/json" hidden /></label>
      <div class="save-transfer-status" id="saveImportStatus"></div>
      <p class="save-transfer-note">Importing replaces the career on this device.</p>
    </div>`;
  openModal(html);

  const exportStatus = (msg, bad) => {
    const el = $("saveExportStatus");
    el.textContent = msg;
    el.classList.toggle("bad", !!bad);
  };
  const importStatus = (msg) => {
    const el = $("saveImportStatus");
    el.textContent = msg;
    el.classList.add("bad");
  };

  // Encode up front so the copy/share happens inside the tap itself —
  // iOS refuses clipboard writes and share sheets that come after an await.
  let code = null;
  encodeSaveCode().then((c) => {
    code = c;
  });
  const ready = () => {
    if (code) return true;
    exportStatus("Still preparing the save code — tap again in a moment.", true);
    return false;
  };

  $("saveExportCopy").addEventListener("click", () => {
    if (!ready()) return;
    const fallback = () => {
      // Clipboard API blocked: drop the code into the box so it can be
      // selected and copied by hand.
      const box = $("saveImportText");
      box.value = code;
      box.focus();
      box.select();
      box.setSelectionRange(0, code.length); // iOS ignores select() alone
      exportStatus("Couldn't copy automatically — the code is selected in the box below; copy it from there.", true);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(code).then(() => exportStatus(`✓ Copied (${Math.ceil(code.length / 1024)} KB). Paste it into Notes or straight into Import on your other device.`), fallback);
    } else {
      fallback();
    }
  });

  if (canShare) {
    $("saveExportShare").addEventListener("click", () => {
      if (!ready()) return;
      const file = new File([code], saveFileName(), { type: "text/plain" });
      const data = navigator.canShare && navigator.canShare({ files: [file] }) ? { files: [file], title: "Cell Grind save" } : { text: code, title: "Cell Grind save" };
      navigator.share(data).catch((e) => {
        if (e && e.name !== "AbortError") exportStatus("Sharing didn't work here — use Copy save code instead.", true);
      });
    });
  } else {
    $("saveExportDownload").addEventListener("click", () => {
      if (!ready()) return;
      const url = URL.createObjectURL(new Blob([code], { type: "text/plain" }));
      const a = document.createElement("a");
      a.href = url;
      a.download = saveFileName();
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      exportStatus("✓ Save file downloaded.");
    });
  }

  const importFrom = async (text) => {
    let parsed;
    try {
      parsed = await decodeSaveCode(text);
    } catch (e) {
      importStatus(e instanceof SaveCodeError ? e.message : "That save code is damaged or cut off — make sure you copied all of it, from the very first character to the last.");
      return;
    }
    const who = parsed.playerName ? `${parsed.playerName}, ` : "";
    if (!confirm(`Load ${who}Day ${parsed.day}? This replaces the career on this device.`)) return;
    try {
      state = Object.assign(freshState(), migrateSave(parsed));
    } catch (e) {
      importStatus("That save couldn't be loaded — it may be from a much older version.");
      return;
    }
    lastLoadedFromSave = false;
    const ok = saveState();
    closeModal();
    renderAll();
    const entry = { html: `💾 Save imported — resumed from Day ${state.day}.`, cls: "event-good" };
    state.logEntries.push(entry);
    appendLog(entry.html, entry.cls);
    saveState();
    showSaveToast(ok);
    if (!state.playerName) openNameModal(true);
  };

  $("saveImportBtn").addEventListener("click", () => importFrom($("saveImportText").value));
  $("saveImportFile").addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    file.text().then(importFrom, () => importStatus("Couldn't read that file."));
    e.target.value = "";
  });
}

/* ---------------------------------------------------------------------- */
/* End day flow                                                           */
/* ---------------------------------------------------------------------- */
function isOverAllocated() {
  return dailyHoursCap(state.stats.nutrition) - totalAssigned() < 0;
}

// Why today's plan can't be played yet (null = it can). Shared by both
// buttons and both actions, and shown as a hint under the buttons.
function endDayBlocker() {
  if (isOverAllocated()) return "Over today's hours budget — trim your plan to end the day.";
  if (state.employment.status === "unemployed" && (state.allocation.work || 0) < BAL.jobSearchMinHours) {
    return `🔍 Put at least ${BAL.jobSearchMinHours}h into Job Search to end the day.`;
  }
  return null;
}

// Resolves one day with the current plan and logs it. Shared by End Day and
// End Week; saving, re-rendering and modals are left to the caller.
function runDay() {
  // Staff contracts run to the end of the week (match day in season).
  const weekEndsToday = daysLeftInWeek() === 1;
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

  const employmentEvents = resolveEmploymentDay();
  employmentEvents.forEach((ev) => {
    const cls = ev.type === "good" ? "event-good" : ev.type === "bad" ? "event-bad" : "";
    const e = { html: ev.text, cls };
    state.logEntries.push(e);
    appendLog(e.html, e.cls);
  });

  // Snapshot before processDayEnd() can flip seasonPhase (e.g. preseason ->
  // regular right on this call) out from under advanceSkillCycle()'s own
  // preseason check below.
  const wasInSeason = inSeason();

  // Resolve this week's match (if today's the day) before the skill focus
  // rerolls — the match grades the skills actually trained this week, not
  // whatever gets revealed for the week ahead.
  const { matchResult, phaseEvent, yearSummary } = processDayEnd();
  if (matchResult) {
    // Rating after everything this match did (incl. a champion bonus), so a
    // replay of the result screen shows the same numbers later on.
    matchResult.ratingAfter = state.rank;
    recordMyMatch(matchResult);
  }

  if (matchResult) {
    const oppText = matchResult.opponentName ? ` vs ${matchResult.opponentName}` : "";
    const label = matchResult.roundLabel || "Match";
    const summary = `🏆 ${label}${oppText} (${matchResult.opponentRating}) — ${matchResult.win ? "WON" : "LOST"}. Rating ${fmt(matchResult.rankBefore)} → ${fmt(state.rank)}. +$${matchResult.cashReward}.`;
    const e = { html: summary, cls: "event-match" };
    state.logEntries.push(e);
    appendLog(e.html, e.cls);
  }

  if (phaseEvent) {
    const e = { html: phaseEvent, cls: "event-season" };
    state.logEntries.push(e);
    appendLog(e.html, e.cls);
  }

  const skillCycleEvent = advanceSkillCycle(wasInSeason);
  if (skillCycleEvent) {
    const e = { html: skillCycleEvent, cls: "event-season" };
    state.logEntries.push(e);
    appendLog(e.html, e.cls);
  }

  if (yearSummary) {
    const raiseText = yearSummary.newPayRate != null ? ` 📈 Pay is now $${fmt(yearSummary.newPayRate)}/day.` : "";
    const summary = `💰 Year ${yearSummary.year} cash flow: +$${fmt(yearSummary.workPay)} work, +$${fmt(yearSummary.matchCash)} matches, -$${fmt(yearSummary.expenses)} expenses → net ${fmtSigned(yearSummary.net, 0)}.${raiseText}`;
    const e = { html: summary, cls: "event-season" };
    state.logEntries.push(e);
    appendLog(e.html, e.cls);
  }

  if (weekEndsToday && state.staff && Object.keys(state.staff.hired).length) {
    state.staff.hired = {};
    const e = { html: "📋 Staff contracts have ended — hire for the new week in 🧑‍🏫 Staff.", cls: "event-season" };
    state.logEntries.push(e);
    appendLog(e.html, e.cls);
  }

  state.day += 1;
  return { matchResult, phaseEvent, yearSummary };
}

function finishTurn() {
  const saved = saveState();
  showSaveToast(saved);
  renderTopbar();
  renderStats();
  renderPlanner();
  $("log").scrollTop = $("log").scrollHeight;
}

function endDay() {
  // Defensive: the button is disabled whenever this is true, but guard the
  // action itself too in case it's ever reachable another way.
  if (endDayBlocker()) return;
  const { matchResult, yearSummary } = runDay();
  finishTurn();
  if (matchResult) {
    showMatchModal(matchResult);
  } else if (yearSummary) {
    showYearSummaryModal(yearSummary);
  }
}

/* ---------------------------------------------------------------------- */
/* End Week — repeat today's plan up to the end of the week               */
/* ---------------------------------------------------------------------- */
// Days left in the current week, today included: in season that's up to and
// including match day; otherwise the end of this 7-day block of the phase
// (never past the phase's own end).
function daysLeftInWeek() {
  const week = BAL.roundIntervalDays;
  if (state.seasonPhase === "regular" || state.seasonPhase === "playoffs") return Math.max(1, week - state.phaseDay);
  const phaseLength = state.seasonPhase === "preseason" ? BAL.preseasonDays : state.offseasonDays || BAL.offseasonDays;
  return Math.max(1, Math.min(week - (state.phaseDay % week), phaseLength - state.phaseDay));
}

function weekSnapshot() {
  const s = state.stats;
  return {
    skills: { ...s.skills },
    phys: s.phys,
    rest: s.rest,
    composure: s.composure,
    nutrition: s.nutrition,
    cash: state.cash,
    rank: state.rank,
    injured: state.injury.active,
    burnout: state.burnout.active,
    employment: state.employment.status,
    techniques: state.employment.techniqueQueue.length,
  };
}

// Anything that changes what today's plan means stops End Week so the
// player can re-plan before the next day runs. Several can land on the same
// day (an injury and a shrinking day, say) — report them all.
function weekStopReason(before) {
  const emp = state.employment;
  const reasons = [];
  if (!before.injured && state.injury.active) {
    reasons.push(`🤕 Injured in the Gym — it's locked for ${state.injury.daysLeft} day${state.injury.daysLeft === 1 ? "" : "s"} and its hours are free to reassign.`);
  }
  if (before.employment !== emp.status) {
    if (emp.status === "unemployed") reasons.push(`🔥 Out of chances — you lost your job. The Work slider is now a Job Search (${jobSearchTarget()}h needed).`);
    else if (emp.status === "pro") reasons.push(`🏆 You went pro — Work is now just ${BAL.proDutyHoursRequired}h/day of Pro Duties.`);
    else reasons.push(`💼 Found a new job — Work is back to ${BAL.workHoursRequired}h/day.`);
  }
  if (!before.burnout && state.burnout.active) reasons.push("😵 Burnout — training is far less effective until Composure recovers. Schedule more Relax.");
  if (emp.techniqueQueue.length > before.techniques) reasons.push(`📘 A new technique to master: ${emp.techniqueQueue[emp.techniqueQueue.length - 1].name}.`);
  if (isOverAllocated()) {
    reasons.push(`🥗 Nutrition dropped — today only has ${dailyHoursCap(state.stats.nutrition)}h, but your plan uses ${totalAssigned()}h. Trim it to carry on.`);
  }
  return reasons.length ? reasons.join(" ") : null;
}

function endWeek() {
  if (endDayBlocker()) return;
  const planned = daysLeftInWeek();
  const start = weekSnapshot();
  // This week's match skills always; otherwise only skills you trained
  // (plus, at the end, any that moved — e.g. rust).
  const shownSkills = SKILL_KEYS.filter((k) => (inSeason() && state.activeSkills.includes(k)) || (state.allocation.skills[k] || 0) > 0);
  const startPhase = state.seasonPhase;
  let daysRun = 0;
  let stopReason = null;
  let matchResult = null;
  let yearSummary = null;
  while (daysRun < planned) {
    const before = weekSnapshot();
    const result = runDay();
    daysRun += 1;
    matchResult = result.matchResult;
    yearSummary = result.yearSummary;
    stopReason = weekStopReason(before);
    if (stopReason || matchResult || yearSummary || state.seasonPhase !== startPhase) break;
  }
  if (stopReason) {
    const e = { html: `⏸️ End Week stopped after ${daysRun} of ${planned} day${planned === 1 ? "" : "s"}: ${stopReason}`, cls: "event-bad" };
    state.logEntries.push(e);
    appendLog(e.html, e.cls);
  }
  finishTurn();
  const summaryHtml = weekSummaryHtml({ start, shownSkills, daysRun, planned, stopReason });
  if (matchResult) showMatchModal(matchResult, summaryHtml);
  else if (yearSummary) showYearSummaryModal(yearSummary, summaryHtml);
  else showWeekSummaryModal(summaryHtml, daysRun);
}

function weekSummaryHtml({ start, shownSkills, daysRun, planned, stopReason }) {
  const s = state.stats;
  const signed = (d) => (Math.abs(d) < 0.5 ? "±0" : d > 0 ? `+${fmt(d)}` : `−${fmt(-d)}`);
  const cls = (d) => (Math.abs(d) < 0.5 ? "" : d > 0 ? "wk-up" : "wk-down");
  const row = (label, a, b) => `<tr><td>${label}</td><td>${fmt(a)}</td><td>${fmt(b)}</td><td class="${cls(b - a)}">${signed(b - a)}</td></tr>`;
  const skillRows = SKILL_KEYS.filter((k) => shownSkills.includes(k) || Math.abs(s.skills[k] - start.skills[k]) >= 0.5)
    .map((k) => {
      const m = skillMeta(k);
      return row(`${m.icon} ${m.name}`, start.skills[k], s.skills[k]);
    })
    .join("");
  const cashDelta = state.cash - start.cash;
  return `
    <div class="week-summary">
      <h3 class="week-summary-title">${daysRun} day skip summary for this week</h3>
      ${stopReason ? `<div class="week-stop">⏸️ Stopped after ${daysRun} of ${planned} days — ${stopReason}</div>` : ""}
      <table class="perf-table week-table">
        <thead><tr><th>Stat</th><th>Start</th><th>Now</th><th>Change</th></tr></thead>
        <tbody>
          ${skillRows}
          ${row("🏃 Health", start.phys, s.phys)}
          ${row("🌙 Rest", start.rest, s.rest)}
          ${row("🎮 Composure", start.composure, s.composure)}
          ${row("🥗 Nutrition", start.nutrition, s.nutrition)}
          <tr><td>💰 Cash</td><td>${fmtMoney(start.cash)}</td><td>${fmtMoney(state.cash)}</td><td class="${cls(cashDelta)}">${cashDelta < 0 ? "−" : "+"}${fmtMoney(Math.abs(cashDelta))}</td></tr>
          ${row("🏆 Rating", start.rank, state.rank)}
        </tbody>
      </table>
      <div class="match-sub perf-note">Day-by-day detail is in the log.</div>
    </div>`;
}

function showWeekSummaryModal(summaryHtml, daysRun) {
  const html = `
    <div class="match-card">
      <div class="match-result neutral">${daysRun === 1 ? "DAY COMPLETE" : "WEEK COMPLETE"}</div>
      ${summaryHtml}
      <button class="primary-btn" id="weekOk">Continue</button>
    </div>`;
  openModal(html);
  $("weekOk").addEventListener("click", closeModal);
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
  const maxForKey = isSkillKey(key) ? SKILL_MAX_HOURS : key === "work" ? workMaxHours() : ACT_MAX[key];
  const others = totalAssigned() - getAllocHours(key);
  const maxAllowed = Math.min(maxForKey, dailyHoursCap(state.stats.nutrition) - others);
  setAllocHoursRaw(key, clamp(val, 0, Math.max(0, maxAllowed)));
  renderPlanner();
}

function wireInputs() {
  // Every row (skills and exercise/sleep/relax alike) is rebuilt on every
  // renderPlanner() call, so all inputs are wired via delegation on the
  // stable container rather than direct listeners that would go stale.
  const container = $("plannerRows");
  container.addEventListener("input", (e) => {
    if (!e.target.matches("input[type=range]")) return;
    const key = e.target.getAttribute("data-key");
    if (key) setAllocation(key, Number(e.target.value));
  });
  container.addEventListener("click", (e) => {
    const btn = e.target.closest(".step-btn");
    if (!btn) return;
    const key = btn.getAttribute("data-key");
    if (key === "exercise" && state.injury.active) return;
    const dir = Number(btn.getAttribute("data-dir"));
    setAllocation(key, getAllocHours(key) + dir);
  });

  $("endDayBtn").addEventListener("click", endDay);
  $("endWeekBtn").addEventListener("click", endWeek);
  $("menuBtn").addEventListener("click", openMenu);
  $("cashChip").addEventListener("click", openShop);
  $("rankChip").addEventListener("click", openCareer);
  $("leaguesBtn").addEventListener("click", () => openLeagues(state.leagueTier));
  $("historyLink").addEventListener("click", () => openHistory());
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
      .then((reg) => {
        reg.update();
        // A Home Screen app is often just resumed, not relaunched, so the
        // launch-time check above can be days old. Check again whenever the
        // app comes back to the foreground (at most every 10 minutes); a new
        // version then takes over and reloads via controllerchange below.
        let lastCheck = Date.now();
        document.addEventListener("visibilitychange", () => {
          if (document.visibilityState !== "visible" || Date.now() - lastCheck < 10 * 60 * 1000) return;
          lastCheck = Date.now();
          reg.update().catch(() => {});
        });
      })
      .catch(() => {});
    // If a newer service worker takes over (a fresh deploy was installed),
    // reload once so the page's own HTML/JS is the new version too, instead
    // of new cached assets running against this tab's already-loaded code.
    // Only when one was already in control: on a very first visit (e.g. a
    // fresh Home Screen install) clients.claim() also fires controllerchange,
    // and reloading then wiped the name prompt mid-entry and could leave the
    // page blank on iOS.
    const hadController = !!navigator.serviceWorker.controller;
    let refreshedForUpdate = false;
    navigator.serviceWorker.addEventListener("controllerchange", () => {
      if (!hadController || refreshedForUpdate) return;
      refreshedForUpdate = true;
      window.location.reload();
    });
  }
}

// Start even if the document finished parsing before this script ran
// (possible on a reload served from the service worker) — otherwise
// DOMContentLoaded has already fired and the game never starts.
if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
else init();
