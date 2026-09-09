// Sandner + Bishop Pick 'Em — one shared Google Sheet backs both leagues.
// Deploy this as a Web App (Deploy > New deployment > Web app).
// Execute as: Me. Who has access: Anyone.
//
// Data model: three sheet tabs, all holding a single JSON blob in cell A1.
//   - "State_Shared"   — universal data both leagues need: the schedule-
//                        derived deadlines, NFL game results, and the four
//                        players who compete in both leagues at once
//                        (SHARED_NAMES below) with their one shared pick
//                        per week.
//   - "State_Sandner"  — the Sandner league's own name, wildcard state, and
//                        its own local (non-shared) players.
//   - "State_Bishop"   — same, for the Bishop league.
// A request's `league` URL parameter (?league=sandner / ?league=bishop,
// defaults to sandner for the original link) picks which local tab to
// merge with the shared one. The front-end (index.html) never knows about
// this split — it always sends/receives one flat state object, identical
// in shape to before; this script does the merging and splitting.

const CELL = 'A1';
const SHARED_SHEET = 'State_Shared';
const LEAGUE_SHEETS = { sandner: 'State_Sandner', bishop: 'State_Bishop' };
const SHARED_NAMES = ['Tbone', 'JP', 'Isla', 'Jaxon'];

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function readJson_(sheetName, fallback) {
  const value = getOrCreateSheet_(sheetName).getRange(CELL).getValue();
  return value ? JSON.parse(value) : fallback;
}

function writeJson_(sheetName, obj) {
  getOrCreateSheet_(sheetName).getRange(CELL).setValue(JSON.stringify(obj));
}

function normalizeLeague_(e) {
  const raw = ((e && e.parameter && e.parameter.league) || 'sandner').toLowerCase();
  return LEAGUE_SHEETS[raw] ? raw : 'sandner';
}

function defaultWeekResults_() {
  const wr = {};
  WEEKS.forEach(function (w) { wr[w] = { entered: false, winners: [] }; });
  return wr;
}

function defaultShared_() {
  return {
    season: 2026, week1Deadline: null, weekOverrides: {}, adminPin: null,
    weekResults: defaultWeekResults_(),
    sharedPlayers: SHARED_NAMES.map(function (name) { return { id: Utilities.getUuid(), name: name, pin: null, picks: {} }; })
  };
}

function defaultLocal_(league) {
  return {
    league: league === 'bishop' ? "Bishop Pick 'Em League" : "Sandner Pick 'Em League",
    wildcard: { active: false, games: [], picks: {}, results: {} },
    localPlayers: []
  };
}

function buildMergedState_(league) {
  const shared = readJson_(SHARED_SHEET, null) || defaultShared_();
  const local = readJson_(LEAGUE_SHEETS[league], null) || defaultLocal_(league);
  return {
    league: local.league,
    season: shared.season,
    week1Deadline: shared.week1Deadline,
    weekOverrides: shared.weekOverrides,
    adminPin: shared.adminPin,
    weekResults: shared.weekResults,
    wildcard: local.wildcard,
    players: shared.sharedPlayers.concat(local.localPlayers)
  };
}

function splitAndSave_(league, state) {
  const sharedSet = {};
  SHARED_NAMES.forEach(function (n) { sharedSet[n] = true; });
  const sharedPlayers = state.players.filter(function (p) { return sharedSet[p.name]; });
  const localPlayers = state.players.filter(function (p) { return !sharedSet[p.name]; });

  writeJson_(SHARED_SHEET, {
    season: state.season,
    week1Deadline: state.week1Deadline,
    weekOverrides: state.weekOverrides,
    adminPin: state.adminPin,
    weekResults: state.weekResults,
    sharedPlayers: sharedPlayers
  });
  writeJson_(LEAGUE_SHEETS[league], {
    league: state.league,
    wildcard: state.wildcard,
    localPlayers: localPlayers
  });
}

/* Auto-migrates the old single-tab format the first time it's needed, so
   redeploying this script can never momentarily serve an empty state to a
   real visitor — the very first request after redeploy migrates in place,
   idempotently (a no-op on every request after the first). */
function ensureMigrated_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sharedSheet = ss.getSheetByName(SHARED_SHEET);
  if (sharedSheet && sharedSheet.getRange(CELL).getValue()) return; // already migrated
  const oldSheet = ss.getSheetByName('State');
  if (!oldSheet || !oldSheet.getRange(CELL).getValue()) return; // fresh install, nothing to migrate
  migrateToMultiLeague();
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureMigrated_();
    const league = normalizeLeague_(e);
    const state = buildMergedState_(league);
    return ContentService.createTextOutput(JSON.stringify(state))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureMigrated_();
    const league = normalizeLeague_(e);
    const body = (e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const state = JSON.parse(body); // throws if invalid — never save corrupt data
    splitAndSave_(league, state);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/* ==========================================================================
   ONE-TIME MIGRATION — run this once manually (select migrateToMultiLeague
   in the function dropdown, click Run) to split the original single-league
   "State" tab into the new State_Shared / State_Sandner / State_Bishop
   tabs. Safe to leave the old "State" tab in place afterward as a backup —
   nothing reads from it anymore once this has run.
   ========================================================================== */
function migrateToMultiLeague() {
  const old = readJson_('State', null);
  if (!old) throw new Error('No existing "State" tab found to migrate from.');

  const sharedSet = {};
  SHARED_NAMES.forEach(function (n) { sharedSet[n] = true; });
  const sharedPlayers = old.players.filter(function (p) { return sharedSet[p.name]; });
  const sandnerLocal = old.players.filter(function (p) { return !sharedSet[p.name]; });

  SHARED_NAMES.forEach(function (name) {
    if (!sharedPlayers.some(function (p) { return p.name === name; })) {
      sharedPlayers.push({ id: Utilities.getUuid(), name: name, pin: null, picks: {} });
    }
  });

  writeJson_(SHARED_SHEET, {
    season: old.season,
    week1Deadline: old.week1Deadline,
    weekOverrides: old.weekOverrides,
    adminPin: old.adminPin,
    weekResults: old.weekResults,
    sharedPlayers: sharedPlayers
  });

  writeJson_(LEAGUE_SHEETS.sandner, {
    league: old.league,
    wildcard: old.wildcard,
    localPlayers: sandnerLocal
  });

  const bishopNames = ['MayMay', 'Grandpa', 'Uncle Booch', 'Aunt Sarah', 'Uncle YaYa', 'Uncle Jace'];
  writeJson_(LEAGUE_SHEETS.bishop, {
    league: "Bishop Pick 'Em League",
    wildcard: { active: false, games: [], picks: {}, results: {} },
    localPlayers: bishopNames.map(function (name) {
      return { id: Utilities.getUuid(), name: name, pin: null, picks: {} };
    })
  });

  Logger.log('Migration complete: %s shared, %s Sandner-local, %s Bishop-local players.',
    sharedPlayers.length, sandnerLocal.length, bishopNames.length);
}

/* ==========================================================================
   Shared schedule data — kept in sync with the front-end (index.html).
   If the schedule ever changes there (flex scheduling, corrections), copy
   the updated TEAMS/SCHEDULE constants over here too.
   ========================================================================== */
const WEEKS = Array.from({ length: 16 }, function (_, i) { return i + 1; });

const TEAMS = [
  ["ARI","Arizona Cardinals","NFC West"],["ATL","Atlanta Falcons","NFC South"],
  ["BAL","Baltimore Ravens","AFC North"],["BUF","Buffalo Bills","AFC East"],
  ["CAR","Carolina Panthers","NFC South"],["CHI","Chicago Bears","NFC North"],
  ["CIN","Cincinnati Bengals","AFC North"],["CLE","Cleveland Browns","AFC North"],
  ["DAL","Dallas Cowboys","NFC East"],["DEN","Denver Broncos","AFC West"],
  ["DET","Detroit Lions","NFC North"],["GB","Green Bay Packers","NFC North"],
  ["HOU","Houston Texans","AFC South"],["IND","Indianapolis Colts","AFC South"],
  ["JAX","Jacksonville Jaguars","AFC South"],["KC","Kansas City Chiefs","AFC West"],
  ["LAC","Los Angeles Chargers","AFC West"],["LAR","Los Angeles Rams","NFC West"],
  ["LV","Las Vegas Raiders","AFC West"],["MIA","Miami Dolphins","AFC East"],
  ["MIN","Minnesota Vikings","NFC North"],["NE","New England Patriots","AFC East"],
  ["NO","New Orleans Saints","NFC South"],["NYG","New York Giants","NFC East"],
  ["NYJ","New York Jets","AFC East"],["PHI","Philadelphia Eagles","NFC East"],
  ["PIT","Pittsburgh Steelers","AFC North"],["SEA","Seattle Seahawks","NFC West"],
  ["SF","San Francisco 49ers","NFC West"],["TB","Tampa Bay Buccaneers","NFC South"],
  ["TEN","Tennessee Titans","AFC South"],["WAS","Washington Commanders","NFC East"]
]
  .map(function (t) { return { code: t[0], name: t[1], div: t[2] }; })
  .sort(function (a, b) { return a.name.localeCompare(b.name); });

const SCHEDULE = {"1":[{"away":"NE","home":"SEA","kickoff":"2026-09-09T20:20:00-04:00"},{"away":"SF","home":"LAR","kickoff":"2026-09-10T20:35:00-04:00"},{"away":"TB","home":"CIN","kickoff":"2026-09-13T13:00:00-04:00"},{"away":"NO","home":"DET","kickoff":"2026-09-13T13:00:00-04:00"},{"away":"NYJ","home":"TEN","kickoff":"2026-09-13T13:00:00-04:00"},{"away":"BAL","home":"IND","kickoff":"2026-09-13T13:00:00-04:00"},{"away":"ATL","home":"PIT","kickoff":"2026-09-13T13:00:00-04:00"},{"away":"CHI","home":"CAR","kickoff":"2026-09-13T13:00:00-04:00"},{"away":"CLE","home":"JAX","kickoff":"2026-09-13T13:00:00-04:00"},{"away":"BUF","home":"HOU","kickoff":"2026-09-13T13:00:00-04:00"},{"away":"MIA","home":"LV","kickoff":"2026-09-13T16:25:00-04:00"},{"away":"GB","home":"MIN","kickoff":"2026-09-13T16:25:00-04:00"},{"away":"WAS","home":"PHI","kickoff":"2026-09-13T16:25:00-04:00"},{"away":"ARI","home":"LAC","kickoff":"2026-09-13T16:25:00-04:00"},{"away":"DAL","home":"NYG","kickoff":"2026-09-13T20:20:00-04:00"},{"away":"DEN","home":"KC","kickoff":"2026-09-14T20:15:00-04:00"}],
"2":[{"away":"DET","home":"BUF","kickoff":"2026-09-17T20:15:00-04:00"},{"away":"CAR","home":"ATL","kickoff":"2026-09-20T13:00:00-04:00"},{"away":"MIN","home":"CHI","kickoff":"2026-09-20T13:00:00-04:00"},{"away":"PHI","home":"TEN","kickoff":"2026-09-20T13:00:00-04:00"},{"away":"PIT","home":"NE","kickoff":"2026-09-20T13:00:00-04:00"},{"away":"GB","home":"NYJ","kickoff":"2026-09-20T13:00:00-04:00"},{"away":"CLE","home":"TB","kickoff":"2026-09-20T13:00:00-04:00"},{"away":"NO","home":"BAL","kickoff":"2026-09-20T13:00:00-04:00"},{"away":"CIN","home":"HOU","kickoff":"2026-09-20T13:00:00-04:00"},{"away":"JAX","home":"DEN","kickoff":"2026-09-20T16:05:00-04:00"},{"away":"LV","home":"LAC","kickoff":"2026-09-20T16:05:00-04:00"},{"away":"WAS","home":"DAL","kickoff":"2026-09-20T16:25:00-04:00"},{"away":"SEA","home":"ARI","kickoff":"2026-09-20T16:25:00-04:00"},{"away":"MIA","home":"SF","kickoff":"2026-09-20T16:25:00-04:00"},{"away":"IND","home":"KC","kickoff":"2026-09-20T20:20:00-04:00"},{"away":"NYG","home":"LAR","kickoff":"2026-09-21T20:15:00-04:00"}],
"3":[{"away":"ATL","home":"GB","kickoff":"2026-09-24T20:15:00-04:00"},{"away":"LAC","home":"BUF","kickoff":"2026-09-27T13:00:00-04:00"},{"away":"CAR","home":"CLE","kickoff":"2026-09-27T13:00:00-04:00"},{"away":"NYJ","home":"DET","kickoff":"2026-09-27T13:00:00-04:00"},{"away":"HOU","home":"IND","kickoff":"2026-09-27T13:00:00-04:00"},{"away":"KC","home":"MIA","kickoff":"2026-09-27T13:00:00-04:00"},{"away":"TEN","home":"NYG","kickoff":"2026-09-27T13:00:00-04:00"},{"away":"CIN","home":"PIT","kickoff":"2026-09-27T13:00:00-04:00"},{"away":"SEA","home":"WAS","kickoff":"2026-09-27T13:00:00-04:00"},{"away":"NE","home":"JAX","kickoff":"2026-09-27T13:00:00-04:00"},{"away":"ARI","home":"SF","kickoff":"2026-09-27T16:05:00-04:00"},{"away":"MIN","home":"TB","kickoff":"2026-09-27T16:05:00-04:00"},{"away":"BAL","home":"DAL","kickoff":"2026-09-27T16:25:00-04:00"},{"away":"LV","home":"NO","kickoff":"2026-09-27T16:25:00-04:00"},{"away":"LAR","home":"DEN","kickoff":"2026-09-27T20:20:00-04:00"},{"away":"PHI","home":"CHI","kickoff":"2026-09-28T20:15:00-04:00"}],
"4":[{"away":"PIT","home":"CLE","kickoff":"2026-10-01T20:15:00-04:00"},{"away":"IND","home":"WAS","kickoff":"2026-10-04T09:30:00-04:00"},{"away":"NE","home":"BUF","kickoff":"2026-10-04T13:00:00-04:00"},{"away":"NYJ","home":"CHI","kickoff":"2026-10-04T13:00:00-04:00"},{"away":"JAX","home":"CIN","kickoff":"2026-10-04T13:00:00-04:00"},{"away":"ARI","home":"NYG","kickoff":"2026-10-04T13:00:00-04:00"},{"away":"LAR","home":"PHI","kickoff":"2026-10-04T13:00:00-04:00"},{"away":"GB","home":"TB","kickoff":"2026-10-04T13:00:00-04:00"},{"away":"TEN","home":"BAL","kickoff":"2026-10-04T13:00:00-04:00"},{"away":"DAL","home":"HOU","kickoff":"2026-10-04T13:00:00-04:00"},{"away":"MIA","home":"MIN","kickoff":"2026-10-04T16:05:00-04:00"},{"away":"KC","home":"LV","kickoff":"2026-10-04T16:25:00-04:00"},{"away":"DEN","home":"SF","kickoff":"2026-10-04T16:25:00-04:00"},{"away":"LAC","home":"SEA","kickoff":"2026-10-04T16:25:00-04:00"},{"away":"DET","home":"CAR","kickoff":"2026-10-04T20:20:00-04:00"},{"away":"ATL","home":"NO","kickoff":"2026-10-05T20:15:00-04:00"}],
"5":[{"away":"TB","home":"DAL","kickoff":"2026-10-08T20:15:00-04:00"},{"away":"PHI","home":"JAX","kickoff":"2026-10-11T09:30:00-04:00"},{"away":"HOU","home":"TEN","kickoff":"2026-10-11T13:00:00-04:00"},{"away":"CIN","home":"MIA","kickoff":"2026-10-11T13:00:00-04:00"},{"away":"LV","home":"NE","kickoff":"2026-10-11T13:00:00-04:00"},{"away":"MIN","home":"NO","kickoff":"2026-10-11T13:00:00-04:00"},{"away":"CLE","home":"NYJ","kickoff":"2026-10-11T13:00:00-04:00"},{"away":"IND","home":"PIT","kickoff":"2026-10-11T13:00:00-04:00"},{"away":"NYG","home":"WAS","kickoff":"2026-10-11T13:00:00-04:00"},{"away":"DEN","home":"LAC","kickoff":"2026-10-11T16:05:00-04:00"},{"away":"CHI","home":"GB","kickoff":"2026-10-11T16:25:00-04:00"},{"away":"DET","home":"ARI","kickoff":"2026-10-11T16:25:00-04:00"},{"away":"SF","home":"SEA","kickoff":"2026-10-11T16:25:00-04:00"},{"away":"BAL","home":"ATL","kickoff":"2026-10-11T20:20:00-04:00"},{"away":"BUF","home":"LAR","kickoff":"2026-10-12T20:15:00-04:00"}],
"6":[{"away":"SEA","home":"DEN","kickoff":"2026-10-15T20:15:00-04:00"},{"away":"HOU","home":"JAX","kickoff":"2026-10-18T09:30:00-04:00"},{"away":"CHI","home":"ATL","kickoff":"2026-10-18T13:00:00-04:00"},{"away":"BAL","home":"CLE","kickoff":"2026-10-18T13:00:00-04:00"},{"away":"TEN","home":"IND","kickoff":"2026-10-18T13:00:00-04:00"},{"away":"NYJ","home":"NE","kickoff":"2026-10-18T13:00:00-04:00"},{"away":"NO","home":"NYG","kickoff":"2026-10-18T13:00:00-04:00"},{"away":"CAR","home":"PHI","kickoff":"2026-10-18T13:00:00-04:00"},{"away":"PIT","home":"TB","kickoff":"2026-10-18T13:00:00-04:00"},{"away":"ARI","home":"LAR","kickoff":"2026-10-18T16:05:00-04:00"},{"away":"LAC","home":"KC","kickoff":"2026-10-18T16:25:00-04:00"},{"away":"BUF","home":"LV","kickoff":"2026-10-18T16:25:00-04:00"},{"away":"DAL","home":"GB","kickoff":"2026-10-18T20:20:00-04:00"},{"away":"WAS","home":"SF","kickoff":"2026-10-19T20:15:00-04:00"}],
"7":[{"away":"NE","home":"CHI","kickoff":"2026-10-22T20:15:00-04:00"},{"away":"PIT","home":"NO","kickoff":"2026-10-25T09:30:00-04:00"},{"away":"SF","home":"ATL","kickoff":"2026-10-25T13:00:00-04:00"},{"away":"CLE","home":"TEN","kickoff":"2026-10-25T13:00:00-04:00"},{"away":"IND","home":"MIN","kickoff":"2026-10-25T13:00:00-04:00"},{"away":"MIA","home":"NYJ","kickoff":"2026-10-25T13:00:00-04:00"},{"away":"TB","home":"CAR","kickoff":"2026-10-25T13:00:00-04:00"},{"away":"CIN","home":"BAL","kickoff":"2026-10-25T13:00:00-04:00"},{"away":"NYG","home":"HOU","kickoff":"2026-10-25T13:00:00-04:00"},{"away":"DEN","home":"ARI","kickoff":"2026-10-25T16:05:00-04:00"},{"away":"GB","home":"DET","kickoff":"2026-10-25T16:25:00-04:00"},{"away":"LAR","home":"LV","kickoff":"2026-10-25T16:25:00-04:00"},{"away":"KC","home":"SEA","kickoff":"2026-10-25T20:20:00-04:00"},{"away":"DAL","home":"PHI","kickoff":"2026-10-26T20:15:00-04:00"}],
"8":[{"away":"CAR","home":"GB","kickoff":"2026-10-29T20:15:00-04:00"},{"away":"BAL","home":"BUF","kickoff":"2026-11-01T13:00:00-05:00"},{"away":"TEN","home":"CIN","kickoff":"2026-11-01T13:00:00-05:00"},{"away":"ARI","home":"DAL","kickoff":"2026-11-01T13:00:00-05:00"},{"away":"MIN","home":"DET","kickoff":"2026-11-01T13:00:00-05:00"},{"away":"LV","home":"NYJ","kickoff":"2026-11-01T13:00:00-05:00"},{"away":"CLE","home":"PIT","kickoff":"2026-11-01T13:00:00-05:00"},{"away":"ATL","home":"TB","kickoff":"2026-11-01T13:00:00-05:00"},{"away":"IND","home":"JAX","kickoff":"2026-11-01T13:00:00-05:00"},{"away":"LAC","home":"LAR","kickoff":"2026-11-01T16:05:00-05:00"},{"away":"KC","home":"DEN","kickoff":"2026-11-01T16:25:00-05:00"},{"away":"NE","home":"MIA","kickoff":"2026-11-01T16:25:00-05:00"},{"away":"PHI","home":"WAS","kickoff":"2026-11-01T20:20:00-05:00"},{"away":"CHI","home":"SEA","kickoff":"2026-11-02T20:15:00-05:00"}],
"9":[{"away":"JAX","home":"BAL","kickoff":"2026-11-05T20:15:00-05:00"},{"away":"CIN","home":"ATL","kickoff":"2026-11-08T09:30:00-05:00"},{"away":"DEN","home":"CAR","kickoff":"2026-11-08T13:00:00-05:00"},{"away":"CLE","home":"NO","kickoff":"2026-11-08T13:00:00-05:00"},{"away":"DAL","home":"IND","kickoff":"2026-11-08T13:00:00-05:00"},{"away":"NYG","home":"PHI","kickoff":"2026-11-08T13:00:00-05:00"},{"away":"LAR","home":"WAS","kickoff":"2026-11-08T13:00:00-05:00"},{"away":"DET","home":"MIA","kickoff":"2026-11-08T13:00:00-05:00"},{"away":"NYJ","home":"KC","kickoff":"2026-11-08T13:00:00-05:00"},{"away":"HOU","home":"LAC","kickoff":"2026-11-08T16:05:00-05:00"},{"away":"LV","home":"SF","kickoff":"2026-11-08T16:05:00-05:00"},{"away":"ARI","home":"SEA","kickoff":"2026-11-08T16:25:00-05:00"},{"away":"GB","home":"NE","kickoff":"2026-11-08T16:25:00-05:00"},{"away":"TB","home":"CHI","kickoff":"2026-11-08T20:20:00-05:00"},{"away":"BUF","home":"MIN","kickoff":"2026-11-09T20:15:00-05:00"}],
"10":[{"away":"WAS","home":"NYG","kickoff":"2026-11-12T20:15:00-05:00"},{"away":"NE","home":"DET","kickoff":"2026-11-15T09:30:00-05:00"},{"away":"KC","home":"ATL","kickoff":"2026-11-15T13:00:00-05:00"},{"away":"HOU","home":"CLE","kickoff":"2026-11-15T13:00:00-05:00"},{"away":"MIN","home":"GB","kickoff":"2026-11-15T13:00:00-05:00"},{"away":"JAX","home":"TEN","kickoff":"2026-11-15T13:00:00-05:00"},{"away":"MIA","home":"IND","kickoff":"2026-11-15T13:00:00-05:00"},{"away":"CAR","home":"NO","kickoff":"2026-11-15T13:00:00-05:00"},{"away":"BUF","home":"NYJ","kickoff":"2026-11-15T13:00:00-05:00"},{"away":"SEA","home":"LV","kickoff":"2026-11-15T16:05:00-05:00"},{"away":"LAR","home":"ARI","kickoff":"2026-11-15T16:05:00-05:00"},{"away":"SF","home":"DAL","kickoff":"2026-11-15T16:25:00-05:00"},{"away":"PIT","home":"CIN","kickoff":"2026-11-15T20:20:00-05:00"},{"away":"LAC","home":"BAL","kickoff":"2026-11-16T20:15:00-05:00"}],
"11":[{"away":"IND","home":"HOU","kickoff":"2026-11-19T20:15:00-05:00"},{"away":"MIA","home":"BUF","kickoff":"2026-11-22T13:00:00-05:00"},{"away":"NO","home":"CHI","kickoff":"2026-11-22T13:00:00-05:00"},{"away":"TEN","home":"DAL","kickoff":"2026-11-22T13:00:00-05:00"},{"away":"TB","home":"DET","kickoff":"2026-11-22T13:00:00-05:00"},{"away":"ARI","home":"KC","kickoff":"2026-11-22T13:00:00-05:00"},{"away":"JAX","home":"NYG","kickoff":"2026-11-22T13:00:00-05:00"},{"away":"BAL","home":"CAR","kickoff":"2026-11-22T13:00:00-05:00"},{"away":"NYJ","home":"LAC","kickoff":"2026-11-22T16:05:00-05:00"},{"away":"LV","home":"DEN","kickoff":"2026-11-22T16:25:00-05:00"},{"away":"PIT","home":"PHI","kickoff":"2026-11-22T16:25:00-05:00"},{"away":"MIN","home":"SF","kickoff":"2026-11-22T20:20:00-05:00"},{"away":"CIN","home":"WAS","kickoff":"2026-11-23T20:15:00-05:00"}],
"12":[{"away":"GB","home":"LAR","kickoff":"2026-11-25T20:00:00-05:00"},{"away":"CHI","home":"DET","kickoff":"2026-11-26T13:00:00-05:00"},{"away":"PHI","home":"DAL","kickoff":"2026-11-26T16:30:00-05:00"},{"away":"KC","home":"BUF","kickoff":"2026-11-26T20:20:00-05:00"},{"away":"DEN","home":"PIT","kickoff":"2026-11-27T15:00:00-05:00"},{"away":"ATL","home":"MIN","kickoff":"2026-11-29T13:00:00-05:00"},{"away":"BAL","home":"HOU","kickoff":"2026-11-29T13:00:00-05:00"},{"away":"LV","home":"CLE","kickoff":"2026-11-29T13:00:00-05:00"},{"away":"NYG","home":"IND","kickoff":"2026-11-29T13:00:00-05:00"},{"away":"NYJ","home":"MIA","kickoff":"2026-11-29T13:00:00-05:00"},{"away":"NO","home":"CIN","kickoff":"2026-11-29T13:00:00-05:00"},{"away":"TEN","home":"JAX","kickoff":"2026-11-29T16:05:00-05:00"},{"away":"WAS","home":"ARI","kickoff":"2026-11-29T16:25:00-05:00"},{"away":"SEA","home":"SF","kickoff":"2026-11-29T16:25:00-05:00"},{"away":"NE","home":"LAC","kickoff":"2026-11-29T20:20:00-05:00"},{"away":"CAR","home":"TB","kickoff":"2026-11-30T20:15:00-05:00"}],
"13":[{"away":"KC","home":"LAR","kickoff":"2026-12-03T20:15:00-05:00"},{"away":"DET","home":"ATL","kickoff":"2026-12-06T13:00:00-05:00"},{"away":"JAX","home":"CHI","kickoff":"2026-12-06T13:00:00-05:00"},{"away":"CIN","home":"CLE","kickoff":"2026-12-06T13:00:00-05:00"},{"away":"WAS","home":"TEN","kickoff":"2026-12-06T13:00:00-05:00"},{"away":"GB","home":"NO","kickoff":"2026-12-06T13:00:00-05:00"},{"away":"SF","home":"NYG","kickoff":"2026-12-06T13:00:00-05:00"},{"away":"LAC","home":"TB","kickoff":"2026-12-06T13:00:00-05:00"},{"away":"MIA","home":"DEN","kickoff":"2026-12-06T16:05:00-05:00"},{"away":"PHI","home":"ARI","kickoff":"2026-12-06T16:05:00-05:00"},{"away":"CAR","home":"MIN","kickoff":"2026-12-06T16:25:00-05:00"},{"away":"BUF","home":"NE","kickoff":"2026-12-06T16:25:00-05:00"},{"away":"HOU","home":"PIT","kickoff":"2026-12-06T20:20:00-05:00"},{"away":"DAL","home":"SEA","kickoff":"2026-12-07T20:15:00-05:00"}],
"14":[{"away":"MIN","home":"NE","kickoff":"2026-12-10T20:15:00-05:00"},{"away":"ATL","home":"CLE","kickoff":"2026-12-13T13:00:00-05:00"},{"away":"TEN","home":"DET","kickoff":"2026-12-13T13:00:00-05:00"},{"away":"CHI","home":"MIA","kickoff":"2026-12-13T13:00:00-05:00"},{"away":"DEN","home":"NYJ","kickoff":"2026-12-13T13:00:00-05:00"},{"away":"IND","home":"PHI","kickoff":"2026-12-13T13:00:00-05:00"},{"away":"HOU","home":"WAS","kickoff":"2026-12-13T13:00:00-05:00"},{"away":"NO","home":"CAR","kickoff":"2026-12-13T13:00:00-05:00"},{"away":"TB","home":"BAL","kickoff":"2026-12-13T13:00:00-05:00"},{"away":"LAC","home":"LV","kickoff":"2026-12-13T16:05:00-05:00"},{"away":"KC","home":"CIN","kickoff":"2026-12-13T16:25:00-05:00"},{"away":"LAR","home":"SF","kickoff":"2026-12-13T16:25:00-05:00"},{"away":"NYG","home":"SEA","kickoff":"2026-12-13T16:25:00-05:00"},{"away":"BUF","home":"GB","kickoff":"2026-12-13T20:20:00-05:00"},{"away":"PIT","home":"JAX","kickoff":"2026-12-14T20:15:00-05:00"}],
"15":[{"away":"SF","home":"LAC","kickoff":"2026-12-17T20:15:00-05:00"},{"away":"SEA","home":"PHI","kickoff":"2026-12-19T17:00:00-05:00"},{"away":"CHI","home":"BUF","kickoff":"2026-12-19T20:20:00-05:00"},{"away":"MIA","home":"GB","kickoff":"2026-12-20T13:00:00-05:00"},{"away":"IND","home":"TEN","kickoff":"2026-12-20T13:00:00-05:00"},{"away":"CLE","home":"NYG","kickoff":"2026-12-20T13:00:00-05:00"},{"away":"BAL","home":"PIT","kickoff":"2026-12-20T13:00:00-05:00"},{"away":"NO","home":"TB","kickoff":"2026-12-20T13:00:00-05:00"},{"away":"ATL","home":"WAS","kickoff":"2026-12-20T13:00:00-05:00"},{"away":"CIN","home":"CAR","kickoff":"2026-12-20T13:00:00-05:00"},{"away":"JAX","home":"HOU","kickoff":"2026-12-20T13:00:00-05:00"},{"away":"NYJ","home":"ARI","kickoff":"2026-12-20T16:05:00-05:00"},{"away":"DEN","home":"LV","kickoff":"2026-12-20T16:25:00-05:00"},{"away":"DAL","home":"LAR","kickoff":"2026-12-20T16:25:00-05:00"},{"away":"DET","home":"MIN","kickoff":"2026-12-20T20:20:00-05:00"},{"away":"NE","home":"KC","kickoff":"2026-12-21T20:15:00-05:00"}],
"16":[{"away":"HOU","home":"PHI","kickoff":"2026-12-24T20:15:00-05:00"},{"away":"GB","home":"CHI","kickoff":"2026-12-25T13:00:00-05:00"},{"away":"BUF","home":"DEN","kickoff":"2026-12-25T16:30:00-05:00"},{"away":"LAR","home":"SEA","kickoff":"2026-12-25T20:15:00-05:00"},{"away":"LAC","home":"MIA","kickoff":"2026-12-27T13:00:00-05:00"},{"away":"ARI","home":"NO","kickoff":"2026-12-27T13:00:00-05:00"},{"away":"NE","home":"NYJ","kickoff":"2026-12-27T13:00:00-05:00"},{"away":"CLE","home":"BAL","kickoff":"2026-12-27T13:00:00-05:00"},{"away":"TEN","home":"LV","kickoff":"2026-12-27T16:05:00-05:00"},{"away":"SF","home":"KC","kickoff":"2026-12-27T16:25:00-05:00"},{"away":"JAX","home":"DAL","kickoff":"2026-12-27T20:20:00-05:00"},{"away":"NYG","home":"DET","kickoff":"2026-12-28T20:15:00-05:00"},{"away":"TB","home":"ATL","kickoff":null},{"away":"CIN","home":"IND","kickoff":null},{"away":"WAS","home":"MIN","kickoff":null},{"away":"CAR","home":"PIT","kickoff":null}]};

function weekSundayInfo_(w) {
  const games = SCHEDULE[w];
  if (!games) return null;
  for (let i = 0; i < games.length; i++) {
    const g = games[i];
    if (!g.kickoff) continue;
    const datePart = g.kickoff.slice(0, 10);
    const offset = g.kickoff.slice(19);
    const dow = new Date(datePart + 'T12:00:00Z').getUTCDay();
    if (dow === 0) return { date: datePart, offset: offset };
  }
  return null;
}

function weekDeadline_(state, w) {
  if (state.weekOverrides && state.weekOverrides[w]) return new Date(state.weekOverrides[w]);
  const info = weekSundayInfo_(String(w));
  if (info) return new Date(info.date + 'T12:30:00' + info.offset);
  if (!state.week1Deadline) return null;
  const d = new Date(state.week1Deadline);
  d.setDate(d.getDate() + (w - 1) * 7);
  return d;
}

function isPast_(d) {
  return d && d.getTime() <= Date.now();
}

function currentWeek_(state) {
  for (let i = 0; i < WEEKS.length; i++) {
    const w = WEEKS[i];
    const dl = weekDeadline_(state, w);
    if (!dl || !isPast_(dl)) return w;
  }
  return 17;
}

function usedTeams_(player) {
  const used = {};
  WEEKS.forEach(function (w) {
    const pk = player.picks[w];
    if (pk) { used[pk.a] = true; used[pk.b] = true; }
  });
  return used;
}

function alphaAutoPick_(player) {
  const used = usedTeams_(player);
  const avail = TEAMS.filter(function (t) { return !used[t.code]; });
  return [avail[0] && avail[0].code, avail[1] && avail[1].code];
}

/* Assigns the alphabetically-first available teams to anyone who missed a
   week's deadline without submitting — matches the front-end's own
   reconcileMissedWeeks(). Takes an explicit players list since shared and
   per-league players now live in different sheet tabs; `dates` is any
   object carrying weekOverrides/week1Deadline (the shared blob). */
function reconcileMissedWeeks_(dates, players) {
  let changed = false;
  const cw = currentWeek_(dates);
  players.forEach(function (player) {
    for (let i = 0; i < WEEKS.length; i++) {
      const w = WEEKS[i];
      if (w >= cw) break;
      const dl = weekDeadline_(dates, w);
      if (!isPast_(dl)) continue;
      if (!player.picks[w]) {
        const pick = alphaAutoPick_(player);
        if (pick[0] && pick[1]) {
          player.picks[w] = { a: pick[0], b: pick[1], source: 'auto', ts: new Date().toISOString() };
          changed = true;
        }
      }
    }
  });
  return changed;
}

/* ==========================================================================
   Daily maintenance: auto-assign missed picks (shared players once, then
   each league's own local players), then pull any newly-finished game
   results from ESPN's public scoreboard feed into the shared blob. Games
   happen Thu/Sun/Mon (and occasionally Sat/Fri), so this checks every week
   1-16 each run rather than assuming "the current week" — cheap, and
   catches any week whose games just finished. Never removes a result
   that's already on file; only adds/corrects winners for games that are
   now final.
   ========================================================================== */
const SEASON_YEAR = 2026;
const ESPN_ABBR_TO_OURS = { WSH: 'WAS' };

function espnAbbr_(abbr) {
  return ESPN_ABBR_TO_OURS[abbr] || abbr;
}

function pullScores() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureMigrated_();
    const shared = readJson_(SHARED_SHEET, null);
    if (!shared) return; // fresh install, nothing to do yet
    let changed = reconcileMissedWeeks_(shared, shared.sharedPlayers);

    Object.keys(LEAGUE_SHEETS).forEach(function (league) {
      const local = readJson_(LEAGUE_SHEETS[league], null);
      if (!local) return;
      if (reconcileMissedWeeks_(shared, local.localPlayers)) {
        writeJson_(LEAGUE_SHEETS[league], local);
      }
    });

    for (let week = 1; week <= 16; week++) {
      const url = 'https://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard'
        + '?seasontype=2&week=' + week + '&dates=' + SEASON_YEAR;
      let data;
      try {
        data = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
      } catch (err) {
        continue; // transient network issue — try again on tomorrow's run
      }
      const events = data.events || [];
      if (!events.length) continue;

      const wk = String(week);
      const existing = (shared.weekResults[wk] && shared.weekResults[wk].winners) || [];
      const winners = new Set(existing);

      events.forEach(function (ev) {
        const comp = ev.competitions && ev.competitions[0];
        if (!comp || !comp.status || !comp.status.type || !comp.status.type.completed) return;
        const competitors = comp.competitors || [];
        const winnerComp = competitors.find(function (c) { return c.winner === true; });
        const loserComp = competitors.find(function (c) { return c.winner === false; });
        if (!winnerComp) return; // tie, or not yet finalized
        const winCode = espnAbbr_(winnerComp.team.abbreviation);
        const loseCode = loserComp ? espnAbbr_(loserComp.team.abbreviation) : null;
        if (loseCode) winners.delete(loseCode);
        if (!winners.has(winCode)) winners.add(winCode);
      });

      const newWinners = Array.from(winners).sort();
      const oldWinners = existing.slice().sort();
      if (newWinners.length && JSON.stringify(newWinners) !== JSON.stringify(oldWinners)) {
        shared.weekResults[wk] = { entered: true, winners: newWinners };
        changed = true;
      }
    }

    if (changed) {
      writeJson_(SHARED_SHEET, shared);
    }
  } finally {
    lock.releaseLock();
  }
}

/* Run this ONCE manually from the Apps Script editor (select it in the
   function dropdown, click Run) to schedule pullScores() every morning. */
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'pullScores') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pullScores')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .inTimezone('America/Detroit')
    .create();
}
