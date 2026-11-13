/**
 * Dublin Shiai League 2027 - backend
 * Google Apps Script web app. All data lives in a Google Sheet.
 *
 * Deploy as a web app: Execute as Me, access Anyone.
 *
 * EDITOR PASSWORD: replace the placeholder below with the real editor
 * password before deploying. The password lives ONLY here, server side.
 * It must never appear in index.html, in the sheet, or in any other file.
 *
 * Model: Players have a first and last name (no club). Turns ARE the
 * schedule: each turn takes its name from its date, and the Schedule
 * tab in the page is rendered from the Turns sheet. A turn also carries
 * a player count (derived from TurnPlayers) and a Fighting Spirit
 * award winner (a player id, set manually by an editor).
 * Ippon types (Men, Kote, Do, Tsuki, Hansoku) are recorded per ippon
 * scored and stored in the Matches sheet (columns P and Q).
 * No two fighters may share a name (checked on creation,
 * case-insensitive). A fighter can be deleted only when they have
 * no matches; deleting also removes them from turn player lists and
 * clears any Fighting Spirit award naming them.
 */

var SHEET_ID = '1ylmdjB724j_dQe1RPMdxhi08jqXnMnAdK-EPc-uOT3Y';
var EDITOR_PASSWORD = '___EDITOR_PASSWORD___';

/* ---------- generic helpers ---------- */

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

// Create a tab with headers if it does not exist yet, and return it.
function ensureSheet_(name, headers) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

// All data rows of a tab (row 1 is the header). Returns [] when empty.
function rows_(name) {
  var sh = sheet_(name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}

function num_(v, dflt) {
  var n = Number(v);
  return isNaN(n) ? dflt : n;
}

function pad3_(n) {
  var s = String(n);
  while (s.length < 3) s = '0' + s;
  return s;
}

function clampInt_(v, lo, hi) {
  var n = Math.round(Number(v));
  if (isNaN(n)) n = lo;
  if (n < lo) n = lo;
  if (n > hi) n = hi;
  return n;
}

// Normalise a sheet value to a 'YYYY-MM-DD' string (or '' when empty).
function fmtDate_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return s;
}

// A comma-separated ippon-type cell -> array of types (empty if none).
function parseTypes_(v) {
  return String(v || '').split(',').map(function (t) { return t.trim(); })
    .filter(function (t) { return !!t; });
}

function settings_() {
  var out = {};
  rows_('Settings').forEach(function (r) {
    if (r[0]) out[String(r[0])] = r[1];
  });
  return {
    startingElo: num_(out.startingElo, 1000),
    kEarly: num_(out.kEarly, 32),
    kLate: num_(out.kLate, 16),
    kThresholdMatches: num_(out.kThresholdMatches, 20),
    winPoints: num_(out.winPoints, 3),
    drawPoints: num_(out.drawPoints, 1),
    lossPoints: num_(out.lossPoints, 0),
    editorUsername: String(out.editorUsername || 'dublinshiaileague'),
    extraSeasons: String(out.extraSeasons || '')
  };
}

function setSetting_(key, value) {
  var sh = sheet_('Settings');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

// The seasons offered in the season dropdown: only 2027 is pre-loaded;
// any later year is added by an editor through the Add season form and
// stored in Settings, so the list stays short.
function seasonsList_() {
  var out = [2027];
  var settings = settings_();
  String(settings.extraSeasons || '').split(',').forEach(function (s) {
    var y2 = parseInt(String(s).trim(), 10);
    if (y2 >= 2000 && y2 <= 2100 && out.indexOf(y2) === -1) out.push(y2);
  });
  out.sort(function (a, b) { return a - b; });
  return out;
}

function addSeason_(body) {
  var year = parseInt(body.year, 10);
  if (isNaN(year) || year < 2000 || year > 2100) {
    throw new Error('Enter a valid year (2000-2100).');
  }
  var list = seasonsList_();
  if (list.indexOf(year) !== -1) throw new Error('That season already exists.');
  var settings = settings_();
  var extra = String(settings.extraSeasons || '').split(',')
    .map(function (s) { return String(s).trim(); })
    .filter(function (s) { return !!s; });
  extra.push(String(year));
  setSetting_('extraSeasons', extra.join(','));
  return { ok: true, seasons: seasonsList_() };
}

/* ---------- users, logins and roles ---------- */

// The app is hidden behind a login, even for viewers. Logins live in a
// Users tab (Username, Password, Role, PlayerID, FirstName, LastName,
// Created). Roles: 'admin' (only dublinshiaileague, the only account
// that sees the Admin tab), 'editor' (can make changes), 'viewer'
// (read-only). Passwords are stored as plain text in the sheet, so the
// Admin can find and share them with fighters; this is a small club
// tool, not a public service, and the trade-off is deliberate.
// A session is a random token in a Sessions tab (Token, Username,
// Created); the page stores it and sends it with every request.

// A random 10-digit password (first digit 1-9, so it is always 10 digits).
function randomPassword_() {
  return String(1000000000 + Math.floor(Math.random() * 9000000000));
}

// A login username is firstnamelastname, lowercased, with anything that
// is not a letter or digit removed (so "Patricia D'Vatz" is
// "patriciad vatz" without the apostrophe, "patriciad vatz" as one word).
function usernameFor_(firstName, lastName) {
  return (String(firstName || '') + String(lastName || '')).toLowerCase().replace(/[^a-z0-9]/g, '');
}

// Make a username unique: if it is taken, append 2, 3, ...
function uniqueUsername_(base) {
  var taken = {};
  rows_('Users').forEach(function (r) { taken[String(r[0]).toLowerCase()] = true; });
  if (!taken[base]) return base;
  var n = 2;
  while (taken[base + n]) n++;
  return base + n;
}

// The Users tab, seeded on first creation: the admin account
// (dublinshiaileague, with the deployed editor password) and a login
// for every fighter already in the system (viewer role, random
// 10-digit password). Seeding runs only while the tab is empty, so a
// later password change in the Admin tab is never overwritten.
function ensureUsers_() {
  var sh = ensureSheet_('Users', ['Username', 'Password', 'Role', 'PlayerID', 'FirstName', 'LastName', 'Created']);
  var data = sh.getDataRange().getValues();
  if (data.length > 1) return sh;
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var adminPass = (EDITOR_PASSWORD && EDITOR_PASSWORD !== '___EDITOR_PASSWORD___') ? EDITOR_PASSWORD : randomPassword_();
  sh.appendRow(['dublinshiaileague', adminPass, 'admin', '', '', '', today]);
  rows_('Players').forEach(function (r) {
    var first = String(r[1] || '').trim();
    var last = String(r[2] || '').trim();
    var base = usernameFor_(first, last);
    if (!base) return;
    var username = uniqueUsername_(base);
    sh.appendRow([username, randomPassword_(), 'viewer', String(r[0]), first, last, today]);
  });
  return sh;
}

function userRowToObj_(r) {
  return {
    username: String(r[0]),
    password: String(r[1]),
    role: String(r[2] || 'viewer'),
    playerId: String(r[3] || ''),
    firstName: String(r[4] || ''),
    lastName: String(r[5] || ''),
    name: (String(r[4] || '') + ' ' + String(r[5] || '')).trim(),
    created: fmtDate_(r[6])
  };
}

function findUser_(username) {
  var found = null;
  rows_('Users').forEach(function (r) {
    if (String(r[0]).toLowerCase() === String(username || '').toLowerCase()) found = userRowToObj_(r);
  });
  return found;
}

// Create a login row. Used by the Admin (manual users) and by
// createPlayer_ (automatic fighter logins).
function createUserRow_(username, password, role, playerId, firstName, lastName) {
  ensureUsers_();
  var sh = sheet_('Users');
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sh.appendRow([username, password, role, playerId || '', firstName || '', lastName || '', today]);
  return {
    username: username, password: password, role: role,
    playerId: playerId || '', firstName: firstName || '', lastName: lastName || '',
    name: ((firstName || '') + ' ' + (lastName || '')).trim(), created: today
  };
}

// A random session token (32 hex characters).
function randomToken_() {
  var chars = '0123456789abcdef';
  var out = '';
  for (var i = 0; i < 32; i++) out += chars[Math.floor(Math.random() * 16)];
  return out;
}

function login_(body) {
  var username = String(body.username || '').trim();
  var password = String(body.password || '');
  if (!username || !password) throw new Error('Enter a username and password.');
  ensureUsers_();
  var user = findUser_(username);
  if (!user || user.password !== password) throw new Error('Invalid username or password.');
  var token = randomToken_();
  var sSh = ensureSheet_('Sessions', ['Token', 'Username', 'Created']);
  sSh.appendRow([token, user.username, new Date().toISOString()]);
  return { ok: true, token: token, username: user.username, role: user.role, name: user.name || user.username };
}

function logout_(body) {
  var token = String(body.token || '');
  var sSh = sheet_('Sessions');
  if (sSh && token) {
    var data = sSh.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === token) sSh.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

// The user for a session token, or null when the token is missing or
// unknown. Sessions never expire on their own; logging out (or the
// Admin deleting the user) removes them.
function authUser_(token) {
  if (!token) return null;
  ensureUsers_();
  var sSh = sheet_('Sessions');
  if (!sSh) return null;
  var username = null;
  var data = sSh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(token)) { username = String(data[i][1]); break; }
  }
  if (!username) return null;
  return findUser_(username);
}

// Admin-only: list every login (passwords included, so the Admin can
// find and share them).
function listUsers_(body) {
  ensureUsers_();
  var users = rows_('Users').map(function (r) { return userRowToObj_(r); });
  users.sort(function (a, b) { return a.username.localeCompare(b.username); });
  return { ok: true, users: users };
}

// Admin-only: create a login. The password is optional: when blank, a
// random 10-digit password is generated and returned. The role must be
// viewer or editor; only dublinshiaileague is admin. A first and last
// name are required: they identify the login, and a fighter needs a
// name. When `fighter` is set, a Player row is created from the name
// and the login is linked to it, so the person shows up in the
// fighter list and can be added to rounds.
function createUser_(body) {
  var username = String(body.username || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
  if (!username) throw new Error('Enter a username (letters and digits only).');
  if (findUser_(username)) throw new Error('That username already exists.');
  var role = String(body.role || 'viewer');
  if (role !== 'viewer' && role !== 'editor') throw new Error('Role must be viewer or editor.');
  var password = String(body.password || '').trim();
  if (!password) password = randomPassword_();
  var firstName = String(body.firstName || '').trim();
  var lastName = String(body.lastName || '').trim();
  if (!firstName || !lastName) throw new Error('First name and last name are required.');
  var playerId = String(body.playerId || '');
  if (body.fighter === true || body.fighter === 'true') {
    playerId = createPlayerRow_(firstName, lastName);
  }
  var user = createUserRow_(username, password, role, playerId, firstName, lastName);
  return { ok: true, user: user, playerId: playerId };
}

// Admin-only: change a login's password and/or role. The password is
// optional (leave it blank to keep the current one); a blank role keeps
// the current role. The admin account stays admin and cannot be
// demoted.
function updateUser_(body) {
  var username = String(body.username || '').trim();
  var user = findUser_(username);
  if (!user) throw new Error('User not found.');
  var newPassword = String(body.password || '').trim();
  var newRole = String(body.role || '').trim();
  if (newRole && newRole !== 'viewer' && newRole !== 'editor' && newRole !== 'admin') {
    throw new Error('Role must be viewer, editor or admin.');
  }
  if (user.username === 'dublinshiaileague' && newRole && newRole !== 'admin') {
    throw new Error('The admin account cannot be demoted.');
  }
  var sh = sheet_('Users');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === username.toLowerCase()) {
      if (newPassword) sh.getRange(i + 1, 2).setValue(newPassword);
      if (newRole) sh.getRange(i + 1, 3).setValue(newRole);
      break;
    }
  }
  var updated = findUser_(username);
  return { ok: true, user: updated };
}

// Admin-only: delete a login (and its sessions). The admin account
// cannot be deleted.
function deleteUser_(body) {
  var username = String(body.username || '').trim();
  if (username.toLowerCase() === 'dublinshiaileague') throw new Error('The admin account cannot be deleted.');
  var user = findUser_(username);
  if (!user) throw new Error('User not found.');
  var sh = sheet_('Users');
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).toLowerCase() === username.toLowerCase()) sh.deleteRow(i + 1);
  }
  var sSh = sheet_('Sessions');
  if (sSh) {
    var sData = sSh.getDataRange().getValues();
    for (var j = sData.length - 1; j >= 1; j--) {
      if (String(sData[j][1]).toLowerCase() === username.toLowerCase()) sSh.deleteRow(j + 1);
    }
  }
  return { ok: true, username: username };
}

// Admin-only: change whether a login belongs to a fighter (the Fighter
// tick box in the Admin tab). Ticking it creates a Player row from the
// login's stored name and links it, so the person shows up in the
// fighter list and can be added to rounds. Unticking it removes them
// from the fighter list, which is only possible while they have no
// matches; the login itself stays, so they can still log in.
function setUserFighter_(body) {
  var username = String(body.username || '').trim();
  if (username.toLowerCase() === 'dublinshiaileague') throw new Error('The admin account cannot be made a fighter.');
  var user = findUser_(username);
  if (!user) throw new Error('User not found.');
  var makeFighter = body.fighter === true || body.fighter === 'true';
  var sh = sheet_('Users');
  if (makeFighter) {
    if (user.playerId) {
      var linked = rows_('Players').some(function (r) { return String(r[0]) === user.playerId; });
      if (linked) return { ok: true, user: user, playerId: user.playerId };
    }
    if (!user.firstName || !user.lastName) throw new Error('Add a first and last name for this login before marking them as a fighter.');
    var playerId = createPlayerRow_(user.firstName, user.lastName);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === username.toLowerCase()) { sh.getRange(i + 1, 4).setValue(playerId); break; }
    }
    return { ok: true, user: findUser_(username), playerId: playerId };
  }
  if (!user.playerId) return { ok: true, user: user, playerId: '' };
  var hasMatches = rows_('Matches').some(function (r) {
    return String(r[3]) === user.playerId || String(r[4]) === user.playerId;
  });
  if (hasMatches) throw new Error('This fighter has matches, so they cannot be removed from the fighter list.');
  deletePlayerRow_(user.playerId);
  var data2 = sh.getDataRange().getValues();
  for (var j = 1; j < data2.length; j++) {
    if (String(data2[j][0]).toLowerCase() === username.toLowerCase()) { sh.getRange(j + 1, 4).setValue(''); break; }
  }
  return { ok: true, user: findUser_(username), playerId: '' };
}

/* ---------- Hall of Fame ---------- */

// Award keys. The first five are auto-calculated (but can be overridden);
// the last two are manual only.
var HOF_AWARDS = ['gold', 'silver', 'bronze', 'mostGains', 'sharpestShooter', 'comeback', 'highestAttendance', 'fightingSpirit', 'bestShinpan'];
var HOF_AUTO_AWARDS = ['gold', 'silver', 'bronze', 'mostGains', 'sharpestShooter', 'comeback', 'highestAttendance'];

// The year of a turn date string ('YYYY-MM-DD').
function hofYear_(dateStr) {
  var y = parseInt(String(dateStr || '').substring(0, 4), 10);
  return isNaN(y) ? 0 : y;
}

// Read the HallOfFame tab into a map: season -> award -> [playerIds].
// The tab is created on first use. 'visible' rows carry PlayerID 'true'.
function hofRows_() {
  ensureSheet_('HallOfFame', ['Season', 'Award', 'PlayerID']);
  var out = {};
  rows_('HallOfFame').forEach(function (r) {
    var season = num_(r[0], 0);
    var award = String(r[1] || '');
    var pid = String(r[2] || '');
    if (!season || !award || !pid) return;
    if (!out[season]) out[season] = {};
    if (!out[season][award]) out[season][award] = [];
    if (out[season][award].indexOf(pid) === -1) out[season][award].push(pid);
  });
  return out;
}

// Compute the Hall of Fame for one season. players, turns, turnPlayers
// and matches are the getState_ mapped objects; seasonRows is
// hofRows()[season] (stored overrides, manual winners and visibility).
// Returns { visible, activePlayerIds, awards } where each award is
// { auto, override, winner, detail }.
function hallOfFameForSeason_(season, settings, players, turns, turnPlayers, matches, seasonRows) {
  seasonRows = seasonRows || {};
  var visible = (seasonRows.visible || []).indexOf('true') !== -1;
  // Turns in this season, in date order.
  var seasonTurns = turns.filter(function (t) { return hofYear_(t.date) === season; })
    .sort(function (a, b) { return (a.date || '').localeCompare(b.date || ''); });
  var seasonTurnIds = {};
  seasonTurns.forEach(function (t) { seasonTurnIds[t.id] = true; });
  // Active: attended one of the two latest completed turns in the
  // season, so a player can skip one turn and stay Active; only
  // missing the last two completed turns makes them Inactive. If no
  // turn is marked complete yet, everyone is Active.
  var completedTurns = seasonTurns.filter(function (t) { return t.status === 'Complete'; });
  var recentComplete = completedTurns.slice(-2);
  var activeSet = {};
  if (!recentComplete.length) {
    players.forEach(function (p) { activeSet[p.id] = true; });
  } else {
    var recentIds = {};
    recentComplete.forEach(function (t) { recentIds[t.id] = true; });
    turnPlayers.forEach(function (tp) {
      if (recentIds[tp.turnId]) activeSet[tp.playerId] = true;
    });
  }
  var activePlayerIds = players.filter(function (p) { return activeSet[p.id]; }).map(function (p) { return p.id; });
  // Matches that count in this season, in replay order (turn date, round,
  // match index). A match counts when its round or turn is marked
  // complete; an unrecorded match that counts is a 0-0 draw.
  var turnDate = {};
  turns.forEach(function (t) { turnDate[t.id] = t.date || ''; });
  var seasonMatches = matches.filter(function (m) {
    return seasonTurnIds[m.turnId] && m.counts;
  }).sort(function (a, b) {
    var da = turnDate[a.turnId] || '';
    var db = turnDate[b.turnId] || '';
    if (da !== db) return da < db ? -1 : 1;
    if (a.round !== b.round) return a.round - b.round;
    return matchIndex_(a.id) - matchIndex_(b.id);
  });
  // Per-player season stats. Elo history comes from the per-match
  // EloBefore/EloAfter columns (written by recomputeRatings_), so no
  // separate history store is needed.
  var stats = {};
  players.forEach(function (p) {
    stats[p.id] = {
      id: p.id, points: 0, ippons: 0, played: 0,
      eloStart: null, eloLow: null, eloEnd: null,
      turnsAttended: 0
    };
  });
  turnPlayers.forEach(function (tp) {
    if (seasonTurnIds[tp.turnId] && stats[tp.playerId]) stats[tp.playerId].turnsAttended++;
  });
  function noteElo(s, before, after) {
    // Elo trajectory: before the first match is the season-start Elo;
    // after each match is the Elo at that point. Lowest is the min.
    if (before !== null && before !== undefined) {
      if (s.eloStart === null) s.eloStart = before;
      if (s.eloLow === null || before < s.eloLow) s.eloLow = before;
    }
    if (after !== null && after !== undefined) {
      if (s.eloLow === null || after < s.eloLow) s.eloLow = after;
      s.eloEnd = after;
    }
  }
  seasonMatches.forEach(function (m) {
    var w = stats[m.whiteId];
    var d = stats[m.redId];
    // An unrecorded match that counts is a 0-0 draw: 1 point each, no
    // ippons, and the Elo history (written by the recompute for counting
    // matches) still applies.
    var res = m.completed ? m.result : 'Draw';
    if (w) {
      w.played++;
      w.ippons += m.whiteIppon;
      if (res === 'White') w.points += settings.winPoints;
      else if (res === 'Draw') w.points += settings.drawPoints;
      else w.points += settings.lossPoints;
      noteElo(w, m.eloWhiteBefore, m.eloWhiteAfter);
    }
    if (d) {
      d.played++;
      d.ippons += m.redIppon;
      if (res === 'Red') d.points += settings.winPoints;
      else if (res === 'Draw') d.points += settings.drawPoints;
      else d.points += settings.lossPoints;
      noteElo(d, m.eloRedBefore, m.eloRedAfter);
    }
  });
  // Fallbacks when Elo history is missing: season-start Elo defaults to
  // the starting Elo, year-end Elo to the player's current rating.
  var playerById = {};
  players.forEach(function (p) { playerById[p.id] = p; });
  Object.keys(stats).forEach(function (id) {
    var s = stats[id];
    if (s.eloStart === null) s.eloStart = settings.startingElo;
    if (s.eloLow === null) s.eloLow = s.eloStart;
    if (s.eloLow > s.eloStart) s.eloLow = s.eloStart;
    if (s.eloEnd === null) s.eloEnd = playerById[id] ? playerById[id].elo : settings.startingElo;
  });
  // Eligible for the performance awards: Active and played at least one
  // match that counts in the season. Highest Attendance needs Active only
  // (being Active already means attending at least one turn).
  var eligible = players.filter(function (p) {
    return activeSet[p.id] && stats[p.id].played >= 1;
  });
  var byPoints = eligible.slice().sort(function (a, b) {
    var sa = stats[a.id], sb = stats[b.id];
    if (sb.points !== sa.points) return sb.points - sa.points;
    if (sb.ippons !== sa.ippons) return sb.ippons - sa.ippons;
    return sb.eloEnd - sa.eloEnd;
  });
  var byElo = eligible.slice().sort(function (a, b) {
    var sa = stats[a.id], sb = stats[b.id];
    if (sb.eloEnd !== sa.eloEnd) return sb.eloEnd - sa.eloEnd;
    if (sb.points !== sa.points) return sb.points - sa.points;
    return 0;
  });
  var byShooter = eligible.slice().sort(function (a, b) {
    var sa = stats[a.id], sb = stats[b.id];
    var ra = sa.ippons / sa.played;
    var rb = sb.ippons / sb.played;
    if (rb !== ra) return rb - ra;
    if (sb.ippons !== sa.ippons) return sb.ippons - sa.ippons;
    return sb.eloEnd - sa.eloEnd;
  });
  var byComeback = eligible.slice().sort(function (a, b) {
    var sa = stats[a.id], sb = stats[b.id];
    var ca = sa.eloEnd - sa.eloLow;
    var cb = sb.eloEnd - sb.eloLow;
    if (cb !== ca) return cb - ca;
    return sb.eloEnd - sa.eloEnd;
  });
  var activeAtt = players.filter(function (p) { return activeSet[p.id]; });
  var maxTurns = 0;
  activeAtt.forEach(function (p) { if (stats[p.id].turnsAttended > maxTurns) maxTurns = stats[p.id].turnsAttended; });
  var attendanceWinners = maxTurns > 0
    ? activeAtt.filter(function (p) { return stats[p.id].turnsAttended === maxTurns; }).map(function (p) { return p.id; })
    : [];
  function oneOrEmpty(arr) { return arr.length ? [arr[0].id] : []; }
  var awards = {};
  awards.gold = {
    auto: oneOrEmpty(byPoints),
    override: null,
    winner: oneOrEmpty(byPoints),
    detail: byPoints.length ? stats[byPoints[0].id].points + ' pts' : null
  };
  awards.silver = {
    auto: byPoints.length > 1 ? [byPoints[1].id] : [],
    override: null,
    winner: byPoints.length > 1 ? [byPoints[1].id] : [],
    detail: byPoints.length > 1 ? stats[byPoints[1].id].points + ' pts' : null
  };
  awards.bronze = {
    auto: byPoints.length > 2 ? [byPoints[2].id] : [],
    override: null,
    winner: byPoints.length > 2 ? [byPoints[2].id] : [],
    detail: byPoints.length > 2 ? stats[byPoints[2].id].points + ' pts' : null
  };
  awards.mostGains = {
    auto: oneOrEmpty(byElo),
    override: null,
    winner: oneOrEmpty(byElo),
    detail: byElo.length ? 'Elo ' + stats[byElo[0].id].eloEnd : null
  };
  awards.sharpestShooter = {
    auto: oneOrEmpty(byShooter),
    override: null,
    winner: oneOrEmpty(byShooter),
    detail: byShooter.length ? (stats[byShooter[0].id].ippons / stats[byShooter[0].id].played).toFixed(2) + ' ippons/match' : null
  };
  awards.comeback = {
    auto: oneOrEmpty(byComeback),
    override: null,
    winner: oneOrEmpty(byComeback),
    detail: byComeback.length ? '+' + (stats[byComeback[0].id].eloEnd - stats[byComeback[0].id].eloLow) + ' Elo' : null
  };
  awards.highestAttendance = {
    auto: attendanceWinners,
    override: null,
    winner: attendanceWinners,
    detail: maxTurns > 0 ? maxTurns + (maxTurns === 1 ? ' turn' : ' turns') : null
  };
  awards.fightingSpirit = { auto: null, override: null, winner: [], detail: null };
  awards.bestShinpan = { auto: null, override: null, winner: [], detail: null };
  // Stored overrides and manual winners replace the auto result. Only
  // real player ids are kept, so a deleted player never wins an award.
  HOF_AWARDS.forEach(function (key) {
    var stored = seasonRows[key];
    if (stored && stored.length) {
      var valid = stored.filter(function (id) { return !!playerById[id]; });
      if (valid.length) {
        if (HOF_AUTO_AWARDS.indexOf(key) !== -1) {
          awards[key].override = valid;
          awards[key].winner = valid;
          // The detail belongs to the auto winner; clear it when
          // overridden so a stale stat is never shown next to a name.
          awards[key].detail = null;
        } else {
          // Manual award: the stored choice is the winner.
          awards[key].winner = valid;
        }
      }
    }
  });
  return { visible: visible, activePlayerIds: activePlayerIds, awards: awards };
}

// Set (or clear) the stored winner for one Hall of Fame award in a
// season. For an auto award this is an override; for a manual award it
// is the winner. An empty playerIds list clears the stored rows.
function setHallOfFameAward_(body) {
  var season = parseInt(body.season, 10);
  if (isNaN(season) || season < 2000 || season > 2100) throw new Error('Enter a valid season year (2000-2100).');
  var award = String(body.award || '');
  if (HOF_AWARDS.indexOf(award) === -1 || award === 'visible') throw new Error('Unknown award.');
  var ids = Array.isArray(body.playerIds)
    ? body.playerIds.map(function (x) { return String(x || ''); }).filter(function (x) { return !!x; })
    : [];
  var seen = {};
  ids = ids.filter(function (x) { if (seen[x]) return false; seen[x] = true; return true; });
  var known = {};
  rows_('Players').forEach(function (r) { known[String(r[0])] = true; });
  ids.forEach(function (id) { if (!known[id]) throw new Error('Player not found: ' + id); });
  var sh = ensureSheet_('HallOfFame', ['Season', 'Award', 'PlayerID']);
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (num_(data[i][0], 0) === season && String(data[i][1]) === award) sh.deleteRow(i + 1);
  }
  ids.forEach(function (id) { sh.appendRow([season, award, id]); });
  return { ok: true, season: season, award: award, playerIds: ids };
}

// Publish (or unpublish) one season's Hall of Fame to viewers.
function setHallOfFameVisible_(body) {
  var season = parseInt(body.season, 10);
  if (isNaN(season) || season < 2000 || season > 2100) throw new Error('Enter a valid season year (2000-2100).');
  var visible = body.visible === true || body.visible === 'true';
  var sh = ensureSheet_('HallOfFame', ['Season', 'Award', 'PlayerID']);
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (num_(data[i][0], 0) === season && String(data[i][1]) === 'visible') sh.deleteRow(i + 1);
  }
  if (visible) sh.appendRow([season, 'visible', 'true']);
  return { ok: true, season: season, visible: visible };
}

/* ---------- read state ---------- */

function getState_() {
  var settings = settings_();
  var players = rows_('Players').map(function (r) {
    var first = String(r[1] || '').trim();
    var last = String(r[2] || '').trim();
    return {
      id: String(r[0]),
      firstName: first,
      lastName: last,
      name: (first + ' ' + last).trim(),
      elo: num_(r[3], settings.startingElo),
      played: num_(r[4], 0),
      created: fmtDate_(r[5])
    };
  });
  var playerById = {};
  players.forEach(function (p) { playerById[p.id] = p; });
  var turnPlayers = rows_('TurnPlayers').map(function (r) {
    return { turnId: String(r[0]), playerId: String(r[1]) };
  });
  var playerCount = {};
  turnPlayers.forEach(function (tp) {
    playerCount[tp.turnId] = (playerCount[tp.turnId] || 0) + 1;
  });
  var turns = rows_('Turns').map(function (r) {
    // Fighting Spirit is a comma-separated list of player IDs, so a turn
    // can hold more than one award. A single ID is just a one-item list.
    var fsIds = parseTypes_(r[4]);
    var fsNames = fsIds.map(function (id) {
      var p = playerById[id];
      return p ? p.name : id;
    }).filter(function (n) { return !!n; });
    // The name IS the date: derive it from the date column every time,
    // instead of reading the stored Name cell. Sheets parses a date-like
    // name into a Date, whose default string carries a time
    // ("Fri Jan 29 2027 00:00:00 GMT+0000..."), so the stored cell
    // cannot be trusted for display.
    var turnDate = fmtDate_(r[2]);
    return {
      id: String(r[0]),
      name: turnNameFromDate_(turnDate),
      date: turnDate,
      status: String(r[3] || ''),
      fightingSpirit: fsIds,
      fightingSpiritNames: fsNames,
      fightingSpiritName: fsNames.join(', '),
      playerCount: playerCount[String(r[0])] || 0,
      notes: String(r[5] || '')
    };
  });
  // Round completion is stored per (turn, round) in its own tab, separate
  // from a match's Completed flag (which means "has a recorded result").
  // A match counts toward the Scoreboard, Elo, Fighters stats and Hall
  // of Fame only when its round is complete or its turn is Complete; an
  // unrecorded match that counts is a 0-0 draw. The tab is created on
  // first use, so older deployments without it still read fine.
  ensureSheet_('RoundStatus', ['TurnID', 'Round', 'Status']);
  var roundCompleteMap = {};
  rows_('RoundStatus').forEach(function (r) {
    if (String(r[2]) === 'Complete') roundCompleteMap[String(r[0]) + ':' + num_(r[1], 0)] = true;
  });
  var turnCompleteMap = {};
  turns.forEach(function (t) { if (t.status === 'Complete') turnCompleteMap[t.id] = true; });
  var matches = rows_('Matches').map(function (r) {
    var rc = !!roundCompleteMap[String(r[1]) + ':' + num_(r[2], 0)];
    return {
      id: String(r[0]),
      turnId: String(r[1]),
      round: num_(r[2], 0),
      whiteId: String(r[3]),
      redId: String(r[4]),
      whiteIppon: num_(r[5], 0),
      redIppon: num_(r[6], 0),
      whiteHansoku: num_(r[7], 0),
      redHansoku: num_(r[8], 0),
      whiteIpponTypes: parseTypes_(r[15]),
      redIpponTypes: parseTypes_(r[16]),
      result: String(r[9] || ''),
      completed: r[10] === true,
      roundComplete: rc,
      counts: rc || !!turnCompleteMap[String(r[1])],
      eloWhiteBefore: (r[11] === '' || r[11] === null) ? null : num_(r[11], 0),
      eloRedBefore: (r[12] === '' || r[12] === null) ? null : num_(r[12], 0),
      eloWhiteAfter: (r[13] === '' || r[13] === null) ? null : num_(r[13], 0),
      eloRedAfter: (r[14] === '' || r[14] === null) ? null : num_(r[14], 0)
    };
  });
  // Who sat out (bye) each round, recorded when the round was generated.
  // The tab is created on first use, so older deployments without it
  // still read fine.
  ensureSheet_('Byes', ['TurnID', 'Round', 'PlayerID']);
  var byes = rows_('Byes').map(function (r) {
    return { turnId: String(r[0]), round: num_(r[1], 0), playerId: String(r[2]) };
  });
  // Hall of Fame: auto-calculated awards, stored overrides/manual
  // winners and per-season visibility, for every season that has turns
  // or is in the season list.
  var hofRows = hofRows_();
  var hofSeasons = {};
  seasonsList_().forEach(function (y) { hofSeasons[y] = true; });
  turns.forEach(function (t) { var y = hofYear_(t.date); if (y) hofSeasons[y] = true; });
  var hallOfFame = {};
  Object.keys(hofSeasons).forEach(function (y) {
    var season = parseInt(y, 10);
    hallOfFame[y] = hallOfFameForSeason_(season, settings, players, turns, turnPlayers, matches, hofRows[season]);
  });
  return {
    ok: true,
    updatedAt: new Date().toISOString(),
    players: players,
    turns: turns,
    turnPlayers: turnPlayers,
    matches: matches,
    byes: byes,
    hallOfFame: hallOfFame,
    settings: settings,
    seasons: seasonsList_()
  };
}

/* ---------- Elo recompute ---------- */

function matchIndex_(matchId) {
  var m = /-M(\d+)$/.exec(matchId);
  return m ? parseInt(m[1], 10) : 0;
}

// Replay every completed match from scratch, in order (turn date, then
// round, then match index), starting every player at startingElo.
// Updates Players Elo/MatchesPlayed and the per-match Elo before/after
// columns. Call after any result is recorded or edited.
function recomputeRatings_() {
  var settings = settings_();
  var start = settings.startingElo;
  var turnDates = {};
  rows_('Turns').forEach(function (r) {
    turnDates[String(r[0])] = fmtDate_(r[2]);
  });
  var playersSh = sheet_('Players');
  var playerRows = rows_('Players');
  var rating = {};
  var played = {};
  playerRows.forEach(function (r) {
    var id = String(r[0]);
    rating[id] = start;
    played[id] = 0;
  });
  var matchesSh = sheet_('Matches');
  var matchRows = rows_('Matches');
  // Only matches that count are replayed: a match counts when its round
  // is marked complete or its turn is marked Complete. A match with no
  // recorded result in a completed round/turn counts as a 0-0 draw (its
  // blank Result already reads as a draw below, 0.5 points each).
  var roundCompleteMap = roundCompleteMap_();
  var turnCompleteMap = turnCompleteMap_();
  var ordered = [];
  matchRows.forEach(function (r, i) {
    if (matchCountsRow_(r, roundCompleteMap, turnCompleteMap)) ordered.push({ row: i + 2, data: r });
  });
  ordered.sort(function (a, b) {
    var da = turnDates[String(a.data[1])] || '';
    var db = turnDates[String(b.data[1])] || '';
    if (da !== db) return da < db ? -1 : 1;
    var ra = num_(a.data[2], 0);
    var rb = num_(b.data[2], 0);
    if (ra !== rb) return ra - rb;
    return matchIndex_(String(a.data[0])) - matchIndex_(String(b.data[0]));
  });
  var beforeAfter = {}; // sheet row -> [eloWhiteBefore, eloRedBefore, eloWhiteAfter, eloRedAfter]
  ordered.forEach(function (m) {
    var r = m.data;
    var w = String(r[3]);
    var d = String(r[4]);
    if (!(w in rating) || !(d in rating)) return;
    var rw = rating[w];
    var rd = rating[d];
    var kw = played[w] < settings.kThresholdMatches ? settings.kEarly : settings.kLate;
    var kd = played[d] < settings.kThresholdMatches ? settings.kEarly : settings.kLate;
    var ew = 1 / (1 + Math.pow(10, (rd - rw) / 400));
    var ed = 1 / (1 + Math.pow(10, (rw - rd) / 400));
    var res = String(r[9] || '');
    var sw = 0.5;
    var sd = 0.5;
    if (res === 'White') { sw = 1; sd = 0; }
    else if (res === 'Red') { sw = 0; sd = 1; }
    var nw = Math.round(rw + kw * (sw - ew));
    var nd = Math.round(rd + kd * (sd - ed));
    beforeAfter[m.row] = [rw, rd, nw, nd];
    rating[w] = nw;
    rating[d] = nd;
    played[w]++;
    played[d]++;
  });
  var pVals = playerRows.map(function (r) {
    var id = String(r[0]);
    return [rating[id], played[id]];
  });
  if (pVals.length) playersSh.getRange(2, 4, pVals.length, 2).setValues(pVals);
  var mVals = matchRows.map(function (r, i) {
    var row = i + 2;
    if (beforeAfter[row]) return beforeAfter[row];
    return ['', '', '', ''];
  });
  if (mVals.length) matchesSh.getRange(2, 12, mVals.length, 4).setValues(mVals);
}

/* ---------- write actions ---------- */

// A turn takes its name from its date, e.g. 'Friday, 29 January 2027'.
function turnNameFromDate_(iso) {
  var parts = String(iso || '').split('-');
  if (parts.length !== 3) return String(iso || '');
  var d = new Date(+parts[0], +parts[1] - 1, +parts[2]);
  if (isNaN(d.getTime())) return String(iso || '');
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEEE, d MMMM yyyy');
}

// Create a Player row (duplicate-name check included) and return the
// new PlayerID. Used by createPlayer_ and by the Admin when a login is
// marked as a fighter.
function createPlayerRow_(firstName, lastName) {
  var norm = (firstName + ' ' + lastName).replace(/\s+/g, ' ').toLowerCase();
  var dup = rows_('Players').some(function (r) {
    var n = (String(r[1] || '').trim() + ' ' + String(r[2] || '').trim())
      .replace(/\s+/g, ' ').toLowerCase();
    return n === norm;
  });
  if (dup) throw new Error('A fighter with that name already exists.');
  var settings = settings_();
  var sh = sheet_('Players');
  var n = 1;
  rows_('Players').forEach(function (r) {
    var m = /^P(\d+)$/.exec(String(r[0]));
    if (m) n = Math.max(n, parseInt(m[1], 10) + 1);
  });
  var id = 'P' + pad3_(n);
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sh.appendRow([id, firstName, lastName, settings.startingElo, 0, today]);
  return id;
}

function createPlayer_(body) {
  var firstName = String(body.firstName || '').trim();
  var lastName = String(body.lastName || '').trim();
  if (!firstName || !lastName) throw new Error('First name and last name are required.');
  var id = createPlayerRow_(firstName, lastName);
  // A fighter gets a login automatically: username firstnamelastname,
  // a random 10-digit password, viewer role. The credentials are
  // returned so the page can show them (and the Admin tab lists them).
  var login = createUserRow_(uniqueUsername_(usernameFor_(firstName, lastName)), randomPassword_(), 'viewer', id, firstName, lastName);
  return { ok: true, playerId: id, login: { username: login.username, password: login.password } };
}

// Delete a Player row and remove the fighter from every turn's player
// list and from any Fighting Spirit award naming them. Does NOT touch
// logins or sessions. The caller must have checked the fighter has no
// matches first: removing a fighter with matches would leave orphaned
// match rows and corrupt the Elo replay.
function deletePlayerRow_(playerId) {
  var sh = sheet_('Players');
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === playerId) sh.deleteRow(i + 1);
  }
  var tpSh = sheet_('TurnPlayers');
  var tpData = tpSh.getDataRange().getValues();
  for (var j = tpData.length - 1; j >= 1; j--) {
    if (String(tpData[j][1]) === playerId) tpSh.deleteRow(j + 1);
  }
  var tSh = sheet_('Turns');
  var tData = tSh.getDataRange().getValues();
  for (var k = 1; k < tData.length; k++) {
    var winners = String(tData[k][4] || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x && x !== playerId; });
    if (winners.join(',') !== String(tData[k][4] || '')) tSh.getRange(k + 1, 5).setValue(winners.join(','));
  }
}

// Delete a fighter. Only possible when they have no matches at all.
// Also removes them from turn player lists and clears any Fighting
// Spirit award naming them.
function deletePlayer_(body) {
  var playerId = String(body.playerId || '');
  var exists = rows_('Players').some(function (r) { return String(r[0]) === playerId; });
  if (!exists) throw new Error('Fighter not found.');
  var hasMatches = rows_('Matches').some(function (r) {
    return String(r[3]) === playerId || String(r[4]) === playerId;
  });
  if (hasMatches) throw new Error('This fighter has matches, so they cannot be deleted.');
  deletePlayerRow_(playerId);
  // The fighter's login goes with them (and any of their sessions).
  var uSh = sheet_('Users');
  if (uSh) {
    var uData = uSh.getDataRange().getValues();
    var uNames = [];
    for (var m = uData.length - 1; m >= 1; m--) {
      if (String(uData[m][3]) === playerId) {
        uNames.push(String(uData[m][0]));
        uSh.deleteRow(m + 1);
      }
    }
    var sSh = sheet_('Sessions');
    if (sSh && uNames.length) {
      var sData = sSh.getDataRange().getValues();
      for (var q = sData.length - 1; q >= 1; q--) {
        if (uNames.indexOf(String(sData[q][1])) !== -1) sSh.deleteRow(q + 1);
      }
    }
  }
  return { ok: true };
}

function createTurn_(body) {
  var date = fmtDate_(body.date || '');
  if (!date) throw new Error('Turn date is required.');
  var name = turnNameFromDate_(date);
  var sh = sheet_('Turns');
  var duplicate = rows_('Turns').some(function (r) { return fmtDate_(r[2]) === date; });
  if (duplicate) throw new Error('A turn already exists on that date.');
  var n = 1;
  rows_('Turns').forEach(function (r) {
    var m = /^T(\d+)$/.exec(String(r[0]));
    if (m) n = Math.max(n, parseInt(m[1], 10) + 1);
  });
  var id = 'T' + n;
  sh.appendRow([id, name, date, '', '', '']);
  return { ok: true, turnId: id };
}

function setTurnPlayers_(body) {
  var turnId = String(body.turnId || '');
  var playerIds = Array.isArray(body.playerIds) ? body.playerIds.map(String) : [];
  var turnExists = rows_('Turns').some(function (r) { return String(r[0]) === turnId; });
  if (!turnExists) throw new Error('Turn not found.');
  if (turnIsComplete_(turnId)) {
    throw new Error('This turn is marked complete. Unmark the turn to change its players.');
  }
  var known = {};
  rows_('Players').forEach(function (r) { known[String(r[0])] = true; });
  var seen = {};
  var valid = playerIds.filter(function (id) {
    if (!known[id] || seen[id]) return false;
    seen[id] = true;
    return true;
  });
  var sh = sheet_('TurnPlayers');
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === turnId) sh.deleteRow(i + 1);
  }
  valid.forEach(function (pid) { sh.appendRow([turnId, pid]); });
  return { ok: true, count: valid.length };
}

// Maximum matching over players who have not faced each other in this
// turn, exploring partners in Swiss order (turn points desc, then Elo
// desc), so the first maximum found is the most Swiss-preferred. A plain
// greedy pass can strand players (e.g. leave two byes when a full round
// is possible), so search with backtracking instead. A node budget keeps
// large late rounds fast; the first path explored is the greedy pairing,
// so a result is always found quickly.
function bestPairing_(order, faced) {
  var best = null;
  var nodes = 0;
  var NODE_LIMIT = 100000;
  var maxPairs = Math.floor(order.length / 2);
  function search(unpaired, pairs) {
    if (++nodes > NODE_LIMIT) return;
    if (best && best.pairs.length === maxPairs) return;
    if (unpaired.length === 0) {
      if (!best || pairs.length > best.pairs.length) best = { pairs: pairs.slice() };
      return;
    }
    if (best && pairs.length + Math.floor(unpaired.length / 2) <= best.pairs.length) return;
    var top = unpaired[0];
    var rest = unpaired.slice(1);
    for (var i = 0; i < rest.length; i++) {
      if (faced[top][rest[i]]) continue;
      var next = rest.slice();
      next.splice(i, 1);
      pairs.push([top, rest[i]]);
      search(next, pairs);
      pairs.pop();
      if (best && best.pairs.length === maxPairs) return;
    }
    // Leave the top player unpaired (bye) and pair the rest.
    search(rest, pairs);
  }
  search(order.slice(), []);
  return best ? best.pairs : [];
}

function generateRound_(body) {
  var turnId = String(body.turnId || '');
  var settings = settings_();
  var turnExists = rows_('Turns').some(function (r) { return String(r[0]) === turnId; });
  if (!turnExists) throw new Error('Turn not found.');
  if (turnIsComplete_(turnId)) {
    throw new Error('This turn is marked complete. Unmark the turn to generate another round.');
  }
  var selected = rows_('TurnPlayers')
    .filter(function (r) { return String(r[0]) === turnId; })
    .map(function (r) { return String(r[1]); });
  if (selected.length < 2) {
    throw new Error('Select at least two players for this turn before generating a round.');
  }
  var players = {};
  rows_('Players').forEach(function (r) {
    players[String(r[0])] = { id: String(r[0]), elo: num_(r[3], settings.startingElo) };
  });
  var turnMatches = rows_('Matches').filter(function (r) { return String(r[1]) === turnId; });
  // A new round cannot be generated until the previous round has been
  // marked complete (the first round in a turn has no previous round, so
  // it is always allowed). This keeps the standings used for pairing in
  // step with the matches that count.
  var maxRound = 0;
  turnMatches.forEach(function (r) { maxRound = Math.max(maxRound, num_(r[2], 0)); });
  if (maxRound >= 1 && !matchIsLocked_(turnId, maxRound)) {
    throw new Error('Mark Round ' + maxRound + ' complete before generating the next round.');
  }
  // Turn points from matches that count in this turn only (a match
  // counts when its round is marked complete or the turn is). An
  // unrecorded match in a completed round counts as a 0-0 draw.
  var roundCompleteMap = roundCompleteMap_();
  var turnCompleteMap = turnCompleteMap_();
  var points = {};
  var faced = {}; // playerId -> { opponentId: true } for every pairing made in this turn
  selected.forEach(function (id) { points[id] = 0; faced[id] = {}; });
  turnMatches.forEach(function (r) {
    var w = String(r[3]);
    var d = String(r[4]);
    if (faced[w]) faced[w][d] = true;
    if (faced[d]) faced[d][w] = true;
    if (matchCountsRow_(r, roundCompleteMap, turnCompleteMap)) {
      var res = String(r[9] || '');
      if (res === 'White') {
        points[w] = (points[w] || 0) + settings.winPoints;
        points[d] = (points[d] || 0) + settings.lossPoints;
      } else if (res === 'Red') {
        points[d] = (points[d] || 0) + settings.winPoints;
        points[w] = (points[w] || 0) + settings.lossPoints;
      } else {
        // Draw, including an unrecorded 0-0 in a completed round.
        points[w] = (points[w] || 0) + settings.drawPoints;
        points[d] = (points[d] || 0) + settings.drawPoints;
      }
    }
  });
  // Sort: turn points desc, then Elo desc, then id asc for determinism.
  var order = selected.filter(function (id) { return players[id]; }).sort(function (a, b) {
    if (points[b] !== points[a]) return points[b] - points[a];
    if (players[b].elo !== players[a].elo) return players[b].elo - players[a].elo;
    return a < b ? -1 : 1;
  });
  var unpaired = order.slice();
  var pairs = bestPairing_(unpaired, faced);
  // Any single player left over sits out (bye): no match row, no Elo change.
  if (pairs.length === 0) throw new Error('No new pairing is possible for this turn.');
  var round = 1;
  turnMatches.forEach(function (r) { round = Math.max(round, num_(r[2], 0) + 1); });
  var sh = sheet_('Matches');
  var created = [];
  pairs.forEach(function (pair, i) {
    var mid = turnId + '-R' + round + '-M' + (i + 1);
    sh.appendRow([mid, turnId, round, pair[0], pair[1], 0, 0, 0, 0, '', false, '', '', '', '']);
    created.push(mid);
  });
  // Record who sat out this round (usually one player when the count is
  // odd) so the page can show it. Stored, not derived, so it stays true
  // even if the competing list changes later.
  var paired = {};
  pairs.forEach(function (pair) { paired[pair[0]] = true; paired[pair[1]] = true; });
  var byes = order.filter(function (id) { return !paired[id]; });
  if (byes.length) {
    var bSh = ensureSheet_('Byes', ['TurnID', 'Round', 'PlayerID']);
    byes.forEach(function (id) { bSh.appendRow([turnId, round, id]); });
  }
  return { ok: true, round: round, matchIds: created, byes: byes };
}

// Ippon types, recorded per ippon scored. 'Hansoku' covers an ippon
// awarded because the opponent committed two hansoku. 'Ippon' is a
// generic ippon with no specific technique (e.g. awarded when the
// opponent is disqualified).
var IPPON_TYPES = ['Men', 'Kote', 'Do', 'Tsuki', 'Hansoku', 'Ippon'];

function cleanTypes_(arr, count) {
  var out = [];
  for (var i = 0; i < count; i++) {
    var t = (arr && arr[i]) ? String(arr[i]).trim() : '';
    out.push(IPPON_TYPES.indexOf(t) >= 0 ? t : '');
  }
  return out;
}

function recordMatch_(body) {
  var matchId = String(body.matchId || '');
  var wI = clampInt_(body.whiteIppon, 0, 2);
  var rI = clampInt_(body.redIppon, 0, 2);
  var wH = clampInt_(body.whiteHansoku, 0, 2);
  var rH = clampInt_(body.redHansoku, 0, 2);
  var wT = cleanTypes_(body.whiteIpponTypes, wI);
  var rT = cleanTypes_(body.redIpponTypes, rI);
  var sh = sheet_('Matches');
  var data = sh.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === matchId) { rowIdx = i + 1; break; }
  }
  if (rowIdx === -1) throw new Error('Match not found.');
  // A match in a completed round or turn is locked: it cannot be edited
  // until that round (or the turn) is unmarked complete.
  if (matchIsLocked_(String(data[rowIdx - 1][1]), num_(data[rowIdx - 1][2], 0))) {
    throw new Error('This match is in a completed round or turn. Unmark the round (or the turn) to edit it.');
  }
  // Result is derived from ippons only. Hansoku is recorded only and never
  // changes the result: the recorder adjusts ippons manually when a hansoku
  // earns the opponent an ippon.
  var result = 'Draw';
  if (wI > rI) result = 'White';
  else if (rI > wI) result = 'Red';
  sh.getRange(rowIdx, 6, 1, 6).setValues([[wI, rI, wH, rH, result, true]]);
  sh.getRange(rowIdx, 16, 1, 2).setValues([[wT.join(','), rT.join(',')]]);
  recomputeRatings_();
  return { ok: true, matchId: matchId, result: result };
}

function deleteMatch_(body) {
  var matchId = String(body.matchId || '');
  var sh = sheet_('Matches');
  var data = sh.getDataRange().getValues();
  var found = false;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === matchId) {
      if (matchIsLocked_(String(data[i][1]), num_(data[i][2], 0))) {
        throw new Error('This match is in a completed round or turn. Unmark the round (or the turn) to delete it.');
      }
      sh.deleteRow(i + 1); found = true;
    }
  }
  if (!found) throw new Error('Match not found.');
  recomputeRatings_();
  return { ok: true, matchId: matchId };
}

function deleteRound_(body) {
  var turnId = String(body.turnId || '');
  var round = num_(body.round, 0);
  if (matchIsLocked_(turnId, round)) {
    throw new Error('This round is complete (or its turn is). Unmark the round (or the turn) to delete it.');
  }
  var sh = sheet_('Matches');
  var data = sh.getDataRange().getValues();
  var deleted = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === turnId && num_(data[i][2], 0) === round) {
      sh.deleteRow(i + 1);
      deleted++;
    }
  }
  if (!deleted) throw new Error('No matches in that round.');
  // The round's bye goes with it.
  var bSh = sheet_('Byes');
  if (bSh) {
    var bData = bSh.getDataRange().getValues();
    for (var j = bData.length - 1; j >= 1; j--) {
      if (String(bData[j][0]) === turnId && num_(bData[j][1], 0) === round) bSh.deleteRow(j + 1);
    }
  }
  // Any completion row for the round goes with it too.
  var rsSh = sheet_('RoundStatus');
  if (rsSh) {
    var rsData = rsSh.getDataRange().getValues();
    for (var k = rsData.length - 1; k >= 1; k--) {
      if (String(rsData[k][0]) === turnId && num_(rsData[k][1], 0) === round) rsSh.deleteRow(k + 1);
    }
  }
  recomputeRatings_();
  return { ok: true, turnId: turnId, round: round, deleted: deleted };
}

function updateMatchPlayers_(body) {
  var matchId = String(body.matchId || '');
  var wId = String(body.whitePlayerId || '');
  var rId = String(body.redPlayerId || '');
  if (!wId || !rId) throw new Error('Both fighters are required.');
  if (wId === rId) throw new Error('The two fighters must be different.');
  var players = rows_('Players');
  var wExists = players.some(function (r) { return String(r[0]) === wId; });
  var rExists = players.some(function (r) { return String(r[0]) === rId; });
  if (!wExists || !rExists) throw new Error('Fighter not found.');
  var sh = sheet_('Matches');
  var data = sh.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === matchId) { rowIdx = i + 1; break; }
  }
  if (rowIdx === -1) throw new Error('Match not found.');
  if (matchIsLocked_(String(data[rowIdx - 1][1]), num_(data[rowIdx - 1][2], 0))) {
    throw new Error('This match is in a completed round or turn. Unmark the round (or the turn) to edit it.');
  }
  sh.getRange(rowIdx, 4, 1, 2).setValues([[wId, rId]]);
  // A recorded result now belongs to the new fighters, so replay every
  // completed match to keep Elo and the scoreboard correct.
  recomputeRatings_();
  return { ok: true, matchId: matchId };
}

function deleteTurn_(body) {
  var turnId = String(body.turnId || '');
  if (turnIsComplete_(turnId)) {
    throw new Error('This turn is marked complete. Unmark the turn to delete it.');
  }
  var tSh = sheet_('Turns');
  var tData = tSh.getDataRange().getValues();
  var found = false;
  for (var i = tData.length - 1; i >= 1; i--) {
    if (String(tData[i][0]) === turnId) { tSh.deleteRow(i + 1); found = true; }
  }
  if (!found) throw new Error('Turn not found.');
  // Remove the turn's matches and its player selections, then replay
  // the remaining completed matches so Elo stays correct.
  var mSh = sheet_('Matches');
  var mData = mSh.getDataRange().getValues();
  for (var j = mData.length - 1; j >= 1; j--) {
    if (String(mData[j][1]) === turnId) mSh.deleteRow(j + 1);
  }
  var tpSh = sheet_('TurnPlayers');
  var tpData = tpSh.getDataRange().getValues();
  for (var k = tpData.length - 1; k >= 1; k--) {
    if (String(tpData[k][0]) === turnId) tpSh.deleteRow(k + 1);
  }
  // The turn's byes go with it.
  var bSh = sheet_('Byes');
  if (bSh) {
    var bData = bSh.getDataRange().getValues();
    for (var b = bData.length - 1; b >= 1; b--) {
      if (String(bData[b][0]) === turnId) bSh.deleteRow(b + 1);
    }
  }
  // The turn's round-completion rows go with it.
  var rsSh = sheet_('RoundStatus');
  if (rsSh) {
    var rsData = rsSh.getDataRange().getValues();
    for (var c = rsData.length - 1; c >= 1; c--) {
      if (String(rsData[c][0]) === turnId) rsSh.deleteRow(c + 1);
    }
  }
  recomputeRatings_();
  return { ok: true, turnId: turnId };
}

// Completion locks (hierarchical): a match is locked when its round is
// marked complete or its turn is. A round is complete when its
// RoundStatus row says Complete; a turn is complete when its Status is
// 'Complete'. Unmarking a round while its turn is complete is refused:
// the turn must be unmarked first, so a completed turn stays fully
// locked until it is reopened.
function turnIsComplete_(turnId) {
  var found = false;
  rows_('Turns').forEach(function (r) {
    if (String(r[0]) === turnId && String(r[3] || '') === 'Complete') found = true;
  });
  return found;
}

function roundIsComplete_(turnId, round) {
  var sh = sheet_('RoundStatus');
  if (!sh) return false;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === turnId && num_(data[i][1], 0) === round && String(data[i][2]) === 'Complete') return true;
  }
  return false;
}

function matchIsLocked_(turnId, round) {
  return turnIsComplete_(turnId) || roundIsComplete_(turnId, round);
}

// A match feeds into the Scoreboard, Elo, Fighters stats and Hall of
// Fame only when it is marked complete: its round is marked complete
// in RoundStatus, or its turn is marked Complete. A match with no
// recorded result in a completed round/turn counts as a 0-0 draw.
// (Carlos, 11 Nov 2026: this replaces the earlier rule where a result
// counted as soon as the recorder set it.)
function roundCompleteMap_() {
  var map = {};
  var sh = sheet_('RoundStatus');
  if (!sh) return map;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === 'Complete') map[String(data[i][0]) + ':' + num_(data[i][1], 0)] = true;
  }
  return map;
}

function turnCompleteMap_() {
  var map = {};
  rows_('Turns').forEach(function (r) {
    if (String(r[3] || '') === 'Complete') map[String(r[0])] = true;
  });
  return map;
}

// Whether a Matches sheet row counts. Columns: [1]=TurnID, [2]=Round.
function matchCountsRow_(r, roundCompleteMap, turnCompleteMap) {
  var turnId = String(r[1]);
  var round = num_(r[2], 0);
  if (turnCompleteMap[turnId]) return true;
  return !!roundCompleteMap[turnId + ':' + round];
}

// Every match in a turn or a single round can be marked complete even
// when some matches have no recorded result: a match can finish 0-0
// with no hansoku and be left unrecorded. When the round or turn is
// marked complete, an unrecorded match counts as a 0-0 draw (1 point
// each) alongside the recorded matches in that round.
function setRoundComplete_(body) {
  var turnId = String(body.turnId || '');
  var round = num_(body.round, 0);
  var complete = body.complete === true || body.complete === 'true';
  var turnExists = rows_('Turns').some(function (r) { return String(r[0]) === turnId; });
  if (!turnExists) throw new Error('Turn not found.');
  var roundMatches = rows_('Matches').filter(function (r) {
    return String(r[1]) === turnId && num_(r[2], 0) === round;
  });
  if (!roundMatches.length) throw new Error('No matches in that round.');
  // A round can be marked complete even when some matches have no
  // recorded result: a match can finish 0-0 with no hansoku, and the
  // recorder may leave it unrecorded. When the round is marked complete
  // those matches count as 0-0 draws (1 point each) alongside the
  // recorded matches in the round, and the Elo is recomputed below.
  if (complete) {
    var rsSh = ensureSheet_('RoundStatus', ['TurnID', 'Round', 'Status']);
    var rsData = rsSh.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < rsData.length; i++) {
      if (String(rsData[i][0]) === turnId && num_(rsData[i][1], 0) === round) {
        rsSh.getRange(i + 1, 3).setValue('Complete');
        found = true;
      }
    }
    if (!found) rsSh.appendRow([turnId, round, 'Complete']);
  } else {
    if (turnIsComplete_(turnId)) {
      throw new Error('Unmark the turn first: this round is complete because the turn is complete.');
    }
    var rsSh2 = sheet_('RoundStatus');
    if (rsSh2) {
      var rsData2 = rsSh2.getDataRange().getValues();
      for (var j = rsData2.length - 1; j >= 1; j--) {
        if (String(rsData2[j][0]) === turnId && num_(rsData2[j][1], 0) === round) rsSh2.deleteRow(j + 1);
      }
    }
  }
  // Marking (or unmarking) a round changes which matches count, so the
  // Elo and MatchesPlayed are recomputed: an unrecorded match in a newly
  // completed round now counts as a 0-0 draw, and unmarking removes it.
  recomputeRatings_();
  return { ok: true, turnId: turnId, round: round, complete: complete };
}

function setTurnStatus_(body) {
  var turnId = String(body.turnId || '');
  var status = String(body.status || '');
  // The only status is Complete (or empty). A turn can be called back
  // from Complete by setting it to empty.
  if (['', 'Complete'].indexOf(status) === -1) {
    throw new Error('Invalid turn status.');
  }
  // Marking a turn complete locks the whole turn. It does not require
  // every match to have a recorded result: a match can finish 0-0 with
  // no hansoku and be left unrecorded. Unmarking needs no validation:
  // it reopens the turn and its rounds fall back to their own status.
  var sh = sheet_('Turns');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === turnId) {
      sh.getRange(i + 1, 4).setValue(status);
      // Marking (or unmarking) a turn changes which matches count, so the
      // Elo and MatchesPlayed are recomputed.
      recomputeRatings_();
      return { ok: true };
    }
  }
  throw new Error('Turn not found.');
}

function setFightingSpirit_(body) {
  var turnId = String(body.turnId || '');
  if (turnIsComplete_(turnId)) {
    throw new Error('This turn is marked complete. Unmark the turn to change its Fighting Spirit awards.');
  }
  // A turn can hold more than one Fighting Spirit award: the page sends
  // the full list each time, and this stores it comma-joined. An empty
  // list clears the awards.
  var ids = Array.isArray(body.playerIds)
    ? body.playerIds.map(function (x) { return String(x || ''); }).filter(function (x) { return !!x; })
    : [];
  var seen = {};
  ids = ids.filter(function (x) { if (seen[x]) return false; seen[x] = true; return true; });
  var sh = sheet_('Turns');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === turnId) {
      if (ids.length) {
        var competing = {};
        rows_('TurnPlayers').forEach(function (r) {
          if (String(r[0]) === turnId) competing[String(r[1])] = true;
        });
        ids.forEach(function (id) {
          if (!competing[id]) throw new Error('Fighting Spirit winners must be players competing in this turn.');
        });
      }
      sh.getRange(i + 1, 5).setValue(ids.join(','));
      return { ok: true };
    }
  }
  throw new Error('Turn not found.');
}

function updateTurnDate_(body) {
  var turnId = String(body.turnId || '');
  var date = fmtDate_(body.date || '');
  if (!date) throw new Error('Turn date is required.');
  if (turnIsComplete_(turnId)) {
    throw new Error('This turn is marked complete. Unmark the turn to change its date.');
  }
  var clash = rows_('Turns').some(function (r) {
    return String(r[0]) !== turnId && fmtDate_(r[2]) === date;
  });
  if (clash) throw new Error('Another turn already exists on that date.');
  var sh = sheet_('Turns');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === turnId) {
      sh.getRange(i + 1, 2, 1, 2).setValues([[turnNameFromDate_(date), date]]);
      return { ok: true };
    }
  }
  throw new Error('Turn not found.');
}

/* ---------- entry points ---------- */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'getState') {
    // The app is hidden behind a login, even for viewers: getState needs
    // a valid session token. The response carries the logged-in user so
    // the page knows their role.
    var token = (e.parameter && e.parameter.token) || '';
    var user = authUser_(token);
    if (!user) return jsonOut_({ ok: false, error: 'Login required.' });
    var st = getState_();
    st.currentUser = { username: user.username, role: user.role, name: user.name || user.username };
    return jsonOut_(st);
  }
  return jsonOut_({ ok: false, error: 'Unknown action.' });
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse((e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return jsonOut_({ ok: false, error: 'Invalid request body.' });
  }
  var action = body.action || '';
  // Login and logout are the only actions that do not need a token.
  // They run inside the try so a bad password comes back as a normal
  // {ok:false} reply instead of an uncaught error.
  try {
    if (action === 'login') return jsonOut_(login_(body));
    if (action === 'logout') return jsonOut_(logout_(body));
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
  // Every other action needs a valid session token. User management is
  // admin-only; all other writes need an editor or the admin. Viewers
  // are read-only.
  var user = authUser_(body.token || '');
  if (!user) return jsonOut_({ ok: false, error: 'Login required.' });
  var adminActions = ['listUsers', 'createUser', 'updateUser', 'deleteUser', 'setUserFighter'];
  if (adminActions.indexOf(action) !== -1) {
    if (user.role !== 'admin') return jsonOut_({ ok: false, error: 'Only the admin can manage users.' });
  } else if (user.role !== 'editor' && user.role !== 'admin') {
    return jsonOut_({ ok: false, error: 'Viewers cannot make changes. Log in as an editor.' });
  }
  var lock = LockService.getScriptLock();
  var acquired = false;
  try {
    lock.waitLock(15000);
    acquired = true;
    switch (action) {
      case 'createPlayer': return jsonOut_(createPlayer_(body));
      case 'createTurn': return jsonOut_(createTurn_(body));
      case 'setTurnPlayers': return jsonOut_(setTurnPlayers_(body));
      case 'generateRound': return jsonOut_(generateRound_(body));
      case 'recordMatch': return jsonOut_(recordMatch_(body));
      case 'deleteMatch': return jsonOut_(deleteMatch_(body));
      case 'deleteRound': return jsonOut_(deleteRound_(body));
      case 'updateMatchPlayers': return jsonOut_(updateMatchPlayers_(body));
      case 'addSeason': return jsonOut_(addSeason_(body));
      case 'deleteTurn': return jsonOut_(deleteTurn_(body));
      case 'setTurnStatus': return jsonOut_(setTurnStatus_(body));
      case 'setRoundComplete': return jsonOut_(setRoundComplete_(body));
      case 'setFightingSpirit': return jsonOut_(setFightingSpirit_(body));
      case 'setHallOfFameAward': return jsonOut_(setHallOfFameAward_(body));
      case 'setHallOfFameVisible': return jsonOut_(setHallOfFameVisible_(body));
      case 'updateTurnDate': return jsonOut_(updateTurnDate_(body));
      case 'deletePlayer': return jsonOut_(deletePlayer_(body));
      case 'listUsers': return jsonOut_(listUsers_(body));
      case 'createUser': return jsonOut_(createUser_(body));
      case 'updateUser': return jsonOut_(updateUser_(body));
      case 'deleteUser': return jsonOut_(deleteUser_(body));
      case 'setUserFighter': return jsonOut_(setUserFighter_(body));
      default: return jsonOut_({ ok: false, error: 'Unknown action.' });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    if (acquired) lock.releaseLock();
  }
}
