/** Single-file backend for the Peer Evaluation Apps Script app.
 * Required Apps Script files: Code.gs, Index.html, Faculty.html
 * Use a new Google Sheet or clear old tabs before first setup.
 */
const CONFIG = {
  PIN: "Admin2026",
  ALL: "ALL",
  S: {
    A: "Activities",
    R: "Rubrics",
    G: "Registrations",
    E: "Evaluations",
    L: "Audit Log",
    P: "Student Profiles"
  }
};

const DEFAULT_RUBRIC = [
  ["cooperation", "Cooperation", 10, 20, true],
  ["contribution", "Contribution", 10, 30, true],
  ["communication", "Communication", 10, 15, true],
  ["problemSolving", "Problem Solving", 10, 15, true],
  ["qualityOfWork", "Quality of Work", 10, 20, true]
];

const H = {
  A: ["Timestamp", "Activity Key", "Activity Name", "Class ID", "Group Key", "Status", "Allow Student Registration", "Registration Opens At", "Registration Closes At", "Evaluation Opens At", "Evaluation Closes At", "Last Updated"],
  R: ["Timestamp", "Activity Key", "Class ID", "Criterion ID", "Criterion Label", "Max Score", "Weight", "Required", "Status", "Last Updated"],
  G: ["Timestamp", "Activity Key", "Class ID", "Group Key", "Student ID", "Name", "Role", "Contribution", "Status", "Last Updated"],
  E: ["Timestamp", "Activity Key", "Class ID", "Group Key", "Evaluator Student ID", "Evaluator Name", "Target Student ID", "Target Name", "Raw Score", "Max Raw Score", "Weighted Score", "Weighted Percent", "Scores JSON", "Comments"],
  L: ["Timestamp", "Actor Type", "Actor ID", "Action", "Details"],
  P: ["Timestamp", "Student ID", "Name", "Passkey Hash", "Status", "Last Updated"]
};

function doGet(e) {
  ensureSetup();
  const v = String((e && e.parameter && e.parameter.view) || "").toLowerCase();
  return HtmlService.createHtmlOutputFromFile(v === "faculty" ? "Faculty" : "Index")
    .setTitle(v === "faculty" ? "Faculty Dashboard" : "Student Peer Evaluation")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function ensureSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [
    [CONFIG.S.A, H.A, "#dbeafe"],
    [CONFIG.S.R, H.R, "#e0e7ff"],
    [CONFIG.S.G, H.G, "#e2e8f0"],
    [CONFIG.S.E, H.E, "#e2e8f0"],
    [CONFIG.S.L, H.L, "#fef3c7"],
    [CONFIG.S.P, H.P, "#dcfce7"]
  ].forEach((x) => mkSheet_(ss, x[0], x[1], x[2]));
  const s = ss.getSheetByName("Sheet1");
  if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s);
}

function changeFacultyPin(oldPin, newPin) {
  if (String(oldPin || "") !== pin_()) throw Error("Current PIN is incorrect.");
  const p = text_(newPin, 100);
  if (p.length < 6) throw Error("PIN must be at least 6 characters.");
  PropertiesService.getScriptProperties().setProperty("FACULTY_PIN", p);
  return "Faculty PIN updated.";
}

function unlockFaculty(pin) {
  ensureSetup();
  if (String(pin || "") !== pin_()) return bad_("Incorrect faculty code.");
  return good_("Unlocked.", facultyState_());
}

function getStudentHintsV2() {
  ensureSetup();
  const a = activities_().filter((x) => x.status === "Active");
  return good_("Hints loaded.", {
    activities: uniq_(a.map((x) => x.activityKey)),
    classes: uniq_(a.map((x) => x.classId).filter((x) => x !== CONFIG.ALL)),
    cohorts: uniqCohorts_(a)
  });
}

function saveStudentProfile(f) {
  ensureSetup();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const studentId = sid_(f.studentId), name = text_(f.name, 120), passkey = validateStudentPasskey_(f.passkey);
    if (!studentId || !name) return bad_("ID number and name are required.");
    const s = sheet_(CONFIG.S.P), data = table_(s), hash = hashStudentPasskey_(studentId, passkey);
    const row = data.rows.find((r) => sid_(r.object["Student ID"]) === studentId);
    if (row) {
      if (status_(row.object.Status) !== "Active") return bad_("This student profile is inactive. Ask your faculty to reactivate it.");
      const stored = text_(row.object["Passkey Hash"], 300);
      if (stored && stored !== hash) return bad_("This ID number already has a profile. Enter the existing passkey.");
      setRow_(s, row.rowNumber, { Name: name, "Passkey Hash": hash, "Last Updated": new Date() });
      audit_("Student", studentId, "VERIFY_PROFILE", { studentId, name });
      return good_("Profile verified. Continue to activity login.", { studentId, name });
    }
    add_(CONFIG.S.P, H.P, { Timestamp: new Date(), "Student ID": studentId, Name: name, "Passkey Hash": hash, Status: "Active", "Last Updated": new Date() });
    audit_("Student", studentId, "CREATE_PROFILE", { studentId, name });
    return good_("Registration successful. Continue to activity login.", { studentId, name });
  } catch (e) {
    return bad_(e.message || String(e));
  } finally {
    lock.releaseLock();
  }
}

function loginStudentV2(d) {
  ensureSetup();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const studentId = sid_(d.studentId), passkey = validateStudentPasskey_(d.passkey);
    const activityKey = code_(d.activityKey, "Activity Key"), classId = code_(d.classCode || d.classId, "Class Code");
    const role = text_(d.groupRole || d.role, 120), contribution = text_(d.contribution, 1600);
    if (!role || !contribution) return bad_("Activity Key, Class Code, Group Role, and Contribution are required.");
    const profile = verifyStudentProfile_(studentId, passkey);
    if (!profile.success) return profile;
    const found = requireFacultyActivity_(activityKey, classId);
    if (!found.success) return found;
    const w = windowCheck_(found.activity, "reg");
    if (!w.success) return w;
    ensureRubric_(activityKey, found.activity.classId || classId);
    const me = upsertStudentActivityRegistration_({ activityKey, classId, groupKey: found.activity.groupKey, studentId, name: profile.name, role, contribution });
    audit_("Student", studentId, "LOGIN_ACTIVITY", { activityKey, classId, groupKey: me.groupKey, role });
    return loadStudentEvaluationSession_(me, passkey, "Activity loaded. You may now evaluate your teammates.");
  } catch (e) {
    return bad_(e.message || String(e));
  } finally {
    lock.releaseLock();
  }
}

function saveEvaluationV2(d) {
  ensureSetup();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const activityKey = code_(d.activityKey, "Activity Key"), classId = code_(d.classCode || d.classId, "Class Code");
    const evaluatorId = sid_(d.evaluatorStudentId || d.studentId), targetId = sid_(d.targetStudentId), passkey = validateStudentPasskey_(d.passkey);
    const comments = text_(d.comments, 1200), scores = d.scores || {};
    if (comments.length < 10) return bad_("Comment is required and must be at least 10 characters.");
    const profile = verifyStudentProfile_(evaluatorId, passkey);
    if (!profile.success) return profile;
    const me = findStudentActivityRegistration_(activityKey, classId, evaluatorId);
    if (!me) return bad_("Evaluator record not found. Please login to this activity first.");
    const target = regs_().find((x) => x.activityKey === activityKey && x.classId === classId && x.groupKey === me.groupKey && x.studentId === targetId && x.status === "Active");
    if (!target) return bad_("Target member is not registered in your group.");
    if (evaluatorId === targetId) return bad_("You cannot evaluate yourself.");
    const found = requireFacultyActivity_(activityKey, classId, me.groupKey);
    if (!found.success) return found;
    const w = windowCheck_(found.activity, "eval");
    if (!w.success) return w;
    if (evals_().some((x) => x.activityKey === activityKey && x.classId === classId && x.groupKey === me.groupKey && x.evaluatorStudentId === evaluatorId && x.targetStudentId === targetId)) return bad_("You already evaluated this member.");
    let raw = 0, maxRaw = 0, weighted = 0, totalWeight = 0;
    const detail = [];
    for (const c of rubricForActivityClass_(activityKey, classId)) {
      const v = Number(scores[c.criterionId]);
      if (c.required && (!Number.isInteger(v) || v < 1 || v > c.maxScore)) return bad_("Please rate " + c.label + " from 1 to " + c.maxScore + ".");
      if (Number.isInteger(v)) {
        raw += v; maxRaw += c.maxScore; weighted += (v / c.maxScore) * c.weight; totalWeight += c.weight;
        detail.push({ id: c.criterionId, label: c.label, score: v, maxScore: c.maxScore, weight: c.weight });
      }
    }
    const percent = totalWeight ? round2_((weighted / totalWeight) * 100) : 0;
    add_(CONFIG.S.E, H.E, { Timestamp: new Date(), "Activity Key": activityKey, "Class ID": classId, "Group Key": me.groupKey, "Evaluator Student ID": evaluatorId, "Evaluator Name": me.name, "Target Student ID": targetId, "Target Name": target.name, "Raw Score": raw, "Max Raw Score": maxRaw, "Weighted Score": round2_(weighted), "Weighted Percent": percent, "Scores JSON": JSON.stringify(detail), Comments: comments });
    audit_("Student", evaluatorId, "SUBMIT_EVALUATION", { activityKey, classId, target: targetId, percent });
    return loadStudentEvaluationSession_(me, passkey, "Evaluation for " + target.name + " saved.");
  } catch (e) {
    return bad_("System Error: " + (e.message || String(e)));
  } finally {
    lock.releaseLock();
  }
}

function saveActivitySettings(d, pin) {
  if (String(pin || "") !== pin_()) return bad_("Incorrect faculty code.");
  try {
    const x = { activityKey: code_(d.activityKey, "Activity"), classId: code_(d.classId, "Class ID"), groupKey: code_(d.groupKey, "Group Key"), activityName: text_(d.activityName || d.activityKey, 160), status: status_(d.status), allowRegistration: truth_(d.allowRegistration), regOpen: dateInput_(d.regOpen), regClose: dateInput_(d.regClose), evalOpen: dateInput_(d.evalOpen), evalClose: dateInput_(d.evalClose) };
    upsertActivity_(x);
    ensureRubric_(x.activityKey, x.classId);
    audit_("Faculty", "FACULTY", "SAVE_ACTIVITY", x);
    return good_("Activity settings saved.", facultyState_());
  } catch (e) {
    return bad_(e.message);
  }
}

function saveActivitySettingsOptional(d, p) { d = d || {}; d.groupKey = optionalCode_(d.groupKey) || CONFIG.ALL; return saveActivitySettings(d, p); }
function saveActivitySettingsAllClass(d, p) { d = d || {}; d.classId = optionalCode_(d.classId || d.classCode) || CONFIG.ALL; d.groupKey = optionalCode_(d.groupKey) || CONFIG.ALL; return saveActivitySettings(d, p); }

function saveRubric(d, pin) {
  if (String(pin || "") !== pin_()) return bad_("Incorrect faculty code.");
  try {
    const ak = code_(d.activityKey, "Activity"), ci = code_(d.classId, "Class ID"), items = (d.criteria || []).filter((x) => text_(x.label, 80));
    if (!items.length) return bad_("Add at least one criterion.");
    const s = sheet_(CONFIG.S.R), data = table_(s);
    for (let i = data.rows.length - 1; i >= 0; i--) {
      const o = data.rows[i].object;
      if (safeCode_(o["Activity Key"]) === ak && safeCode_(o["Class ID"]) === ci) s.deleteRow(data.rows[i].rowNumber);
    }
    items.forEach((c, i) => add_(CONFIG.S.R, H.R, { Timestamp: new Date(), "Activity Key": ak, "Class ID": ci, "Criterion ID": crit_(c.id || c.label || "c" + i), "Criterion Label": text_(c.label, 100), "Max Score": num_(c.maxScore) || 10, Weight: num_(c.weight) || 0, Required: truth_(c.required), Status: status_(c.status), "Last Updated": new Date() }));
    audit_("Faculty", "FACULTY", "SAVE_RUBRIC", { ak, ci, count: items.length });
    return good_("Rubric saved.", facultyState_());
  } catch (e) {
    return bad_(e.message);
  }
}

function saveRubricAllClass(d, pin) { d = d || {}; d.classId = optionalCode_(d.classId || d.classCode) || CONFIG.ALL; return saveRubric(d, pin); }

function moveStudentToGroup(d, pin) { return editReg_(d, pin, () => ({ "Group Key": code_(d.newGroupKey, "New Group Key"), "Last Updated": new Date() }), "Student moved to group."); }
function updateStudentContribution(d, pin) { return editReg_(d, pin, () => ({ Contribution: text_(d.contribution, 1600), "Last Updated": new Date() }), "Contribution updated."); }
function setRegistrationStatus(d, pin) { return editReg_(d, pin, () => ({ Status: status_(d.status), "Last Updated": new Date() }), "Registration status updated."); }

function cleanupDuplicateRegistrations(pin) {
  ensureSetup();
  if (String(pin || "") !== pin_()) return bad_("Incorrect faculty code.");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_(CONFIG.S.G), data = table_(s), seen = {}, extras = [];
    data.rows.forEach((row) => {
      const o = row.object, key = [safeCode_(o["Activity Key"]), safeCode_(o["Class ID"]), sid_(o["Student ID"])].join("||");
      if (!key.replace(/\|/g, "") || status_(o.Status) !== "Active") return;
      if (seen[key]) extras.push(row.rowNumber); else seen[key] = row.rowNumber;
    });
    extras.forEach((rowNumber) => setRow_(s, rowNumber, { Status: "Inactive", "Last Updated": new Date() }));
    audit_("Faculty", "FACULTY", "CLEAN_DUPLICATE_REGISTRATIONS", { count: extras.length });
    return good_(extras.length ? "Duplicate registrations deactivated." : "No active duplicate registrations found.", facultyState_());
  } catch (e) {
    return bad_(e.message);
  } finally {
    lock.releaseLock();
  }
}

function cleanupDuplicateEvaluations(pin) {
  ensureSetup();
  if (String(pin || "") !== pin_()) return bad_("Incorrect faculty code.");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sheet_(CONFIG.S.E), data = table_(s), seen = {}, extras = [];
    data.rows.forEach((row) => {
      const o = row.object, key = [safeCode_(o["Activity Key"]), safeCode_(o["Class ID"]), safeCode_(o["Group Key"]), sid_(o["Evaluator Student ID"]), sid_(o["Target Student ID"])].join("||");
      if (!key.replace(/\|/g, "")) return;
      if (seen[key]) extras.push(row.rowNumber); else seen[key] = row.rowNumber;
    });
    extras.sort((a, b) => b - a).forEach((rowNumber) => s.deleteRow(rowNumber));
    audit_("Faculty", "FACULTY", "CLEAN_DUPLICATE_EVALUATIONS", { count: extras.length });
    return good_(extras.length ? "Duplicate evaluations removed." : "No duplicate evaluations found.", facultyState_());
  } catch (e) {
    return bad_(e.message);
  } finally {
    lock.releaseLock();
  }
}

function editReg_(d, pin, changes, msg) {
  if (String(pin || "") !== pin_()) return bad_("Incorrect faculty code.");
  try {
    const ak = code_(d.activityKey, "Activity"), ci = code_(d.classId, "Class ID"), id = sid_(d.studentId), s = sheet_(CONFIG.S.G), data = table_(s);
    const row = data.rows.find((r) => safeCode_(r.object["Activity Key"]) === ak && safeCode_(r.object["Class ID"]) === ci && sid_(r.object["Student ID"]) === id);
    if (!row) return bad_("Student registration not found.");
    const c = changes(row.object);
    if (c["Group Key"]) ensureActivityExists_(ak, ci, c["Group Key"], ak);
    setRow_(s, row.rowNumber, c);
    audit_("Faculty", "FACULTY", msg, { ak, ci, id, c });
    return good_(msg, facultyState_());
  } catch (e) {
    return bad_(e.message);
  }
}

function generateReportData(f, pin) {
  ensureSetup();
  if (String(pin || "") !== pin_()) return bad_("Incorrect faculty code.");
  try {
    const ci = code_(f.classId, "Class ID"), af = filter_(f.activityKey, "Activity"), gf = filter_(f.groupKey, "Group Key");
    const rs = regs_().filter((x) => x.classId === ci && (af.all || x.activityKey === af.value) && (gf.all || x.groupKey === gf.value) && x.status === "Active");
    if (!rs.length) return bad_("No registered students found for the selected filters.");
    const keys = new Set(rs.map((x) => memberKey_(x)));
    const es = evals_().filter((x) => x.classId === ci && (af.all || x.activityKey === af.value) && (gf.all || x.groupKey === gf.value) && keys.has(memberKey_(x, x.evaluatorStudentId)) && keys.has(memberKey_(x, x.targetStudentId)));
    const cohorts = buildCohorts_(rs, es), summary = buildSummary_(rs, es), completion = buildCompletion_(rs, es);
    return good_("Report generated.", { generatedAt: fmt_(new Date()), filters: { classId: ci, activityKey: af.value, groupKey: gf.value, activityIsAll: af.all, groupIsAll: gf.all }, totals: { registeredStudents: rs.length, cohorts: cohorts.length, submittedEvaluations: es.length, expectedEvaluations: cohorts.reduce((t, c) => t + c.expectedEvaluations, 0), averagePercent: es.length ? round2_(sum_(es.map((x) => x.weightedPercent)) / es.length) : null }, cohorts, summary, completion, details: es.sort(sortEval_).map((x) => Object.assign({}, x, { timestamp: fmt_(x.timestamp) })), roster: rs.sort(sortRoster_) });
  } catch (e) {
    return bad_("System Error: " + e.message);
  }
}

function facultyState_() {
  const a = activities_(), r = regs_(), e = evals_();
  return { options: { activities: uniq_(a.map((x) => x.activityKey).concat(r.map((x) => x.activityKey))), classes: uniq_(a.map((x) => x.classId).concat(r.map((x) => x.classId))), groups: uniq_(a.map((x) => x.groupKey).concat(r.map((x) => x.groupKey))), students: uniqStudents_(r), cohorts: uniqCohorts_(a.concat(r)) }, dashboard: dashboard_(a, r, e), activities: a, rubrics: rubrics_(), roster: rosterOptions_(r), duplicates: duplicateSummary_(r, e) };
}

function dashboard_(a, r, e) {
  const active = r.filter((x) => x.status === "Active"), cohorts = uniqCohorts_(active), single = cohorts.filter((c) => active.filter((x) => cohortKey_(x) === c.key).length < 2), empty = cohorts.filter((c) => !e.some((x) => cohortKey_(x) === c.key)), incomplete = buildCompletion_(active, e).filter((x) => x.submittedCount < x.expectedCount);
  return { totalActivities: a.length, activeActivities: a.filter((x) => x.status === "Active").length, classes: uniq_(r.map((x) => x.classId)).length, cohorts: cohorts.length, registeredStudents: active.length, submittedEvaluations: e.length, incompleteEvaluators: incomplete.length, singleMemberGroups: single.length, cohortsWithNoSubmissions: empty.length, incomplete: incomplete.slice(0, 50), singleMemberCohorts: single, emptyCohorts: empty };
}

function requireFacultyActivity_(activityKey, classId, requestedGroupKey) {
  const rows = activities_().filter((x) => x.activityKey === activityKey && (x.classId === classId || x.classId === CONFIG.ALL));
  if (!rows.length) return bad_("Activity Key was not found for this Class Code. Ask the faculty to create it first.");
  const groupKey = requestedGroupKey || CONFIG.ALL;
  const exact = rows.find((x) => x.classId === classId && x.groupKey === groupKey) || rows.find((x) => x.classId === classId && x.groupKey === CONFIG.ALL);
  const allClass = rows.find((x) => x.classId === CONFIG.ALL && x.groupKey === groupKey) || rows.find((x) => x.classId === CONFIG.ALL && x.groupKey === CONFIG.ALL);
  const activity = exact || allClass || (rows.length === 1 ? rows[0] : null);
  if (!activity) return bad_("This Activity Key has multiple group settings. Ask the faculty to create an ALL group setting or move you to the correct group.");
  if (activity.status !== "Active") return bad_("This activity is currently inactive.");
  return good_("Activity found.", { activity });
}

function upsertStudentActivityRegistration_(x) {
  const s = sheet_(CONFIG.S.G), data = table_(s);
  const row = data.rows.find((r) => safeCode_(r.object["Activity Key"]) === x.activityKey && safeCode_(r.object["Class ID"]) === x.classId && sid_(r.object["Student ID"]) === x.studentId);
  const rec = { "Activity Key": x.activityKey, "Class ID": x.classId, "Group Key": x.groupKey || CONFIG.ALL, "Student ID": x.studentId, Name: x.name, Role: x.role, Contribution: x.contribution, Status: "Active", "Last Updated": new Date() };
  if (row) setRow_(s, row.rowNumber, rec); else add_(CONFIG.S.G, H.G, Object.assign({ Timestamp: new Date() }, rec));
  return { activityKey: x.activityKey, classId: x.classId, groupKey: rec["Group Key"], studentId: x.studentId, name: x.name, role: x.role, contribution: x.contribution, status: "Active" };
}

function findStudentActivityRegistration_(activityKey, classId, studentId) { return regs_().find((x) => x.activityKey === activityKey && x.classId === classId && x.studentId === studentId && x.status === "Active"); }
function loadStudentEvaluationSession_(me, passkey, message) {
  const found = requireFacultyActivity_(me.activityKey, me.classId, me.groupKey);
  if (!found.success) return found;
  const w = windowCheck_(found.activity, "eval");
  if (!w.success) return w;
  const done = new Set(evals_().filter((x) => x.activityKey === me.activityKey && x.classId === me.classId && x.groupKey === me.groupKey && x.evaluatorStudentId === me.studentId).map((x) => x.targetStudentId));
  const targets = regs_().filter((x) => x.activityKey === me.activityKey && x.classId === me.classId && x.groupKey === me.groupKey && x.status === "Active" && x.studentId !== me.studentId).sort((a, b) => a.name.localeCompare(b.name)).map((x) => ({ studentId: x.studentId, name: x.name, role: x.role, contribution: x.contribution, alreadyEvaluated: done.has(x.studentId) }));
  const remainingTargets = targets.filter((x) => !x.alreadyEvaluated);
  return good_(message || "Group loaded.", { evaluator: me, activity: found.activity, targets, remainingTargets, rubric: rubricForActivityClass_(me.activityKey, me.classId), completedCount: targets.length - remainingTargets.length, totalTargets: targets.length, passkey });
}
function rubricForActivityClass_(activityKey, classId) { const exact = rubric_(activityKey, classId); return exact.length ? exact : rubric_(activityKey, CONFIG.ALL); }
function verifyStudentProfile_(studentId, passkey) { if (!studentId) return bad_("ID number is required."); const p = rows_(CONFIG.S.P).map((r) => ({ studentId: sid_(r["Student ID"]), name: text_(r.Name, 120), hash: text_(r["Passkey Hash"], 300), status: status_(r.Status) })).find((x) => x.studentId === studentId && x.status === "Active"); if (!p) return bad_("Student profile not found. Register your ID number first."); if (p.hash !== hashStudentPasskey_(studentId, passkey)) return bad_("Incorrect passkey."); return good_("Profile verified.", { studentId: p.studentId, name: p.name }); }
function validateStudentPasskey_(v) { const p = text_(v, 120); if (!p) throw Error("Passkey is required."); if (p.length < 4) throw Error("Passkey must be at least 4 characters."); return p; }
function hashStudentPasskey_(studentId, passkey) { const salt = PropertiesService.getScriptProperties().getProperty("PASSKEY_SALT") || "peer-evaluation-v2"; const raw = sid_(studentId) + ":" + String(passkey || "") + ":" + salt; const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8); return bytes.map((b) => (b + 256).toString(16).slice(-2)).join(""); }

function ensureActivityExists_(ak, ci, gk, name) { const rows = activities_(); let a = rows.find((x) => x.activityKey === ak && x.classId === ci && x.groupKey === gk) || rows.find((x) => x.activityKey === ak && x.classId === ci && x.groupKey === CONFIG.ALL); if (a) return a; return upsertActivity_({ activityKey: ak, classId: ci, groupKey: gk, activityName: name || ak, status: "Active", allowRegistration: true, regOpen: "", regClose: "", evalOpen: "", evalClose: "" }); }
function upsertActivity_(x) { const s = sheet_(CONFIG.S.A), data = table_(s), found = data.rows.find((r) => safeCode_(r.object["Activity Key"]) === x.activityKey && safeCode_(r.object["Class ID"]) === x.classId && safeCode_(r.object["Group Key"]) === x.groupKey); const rec = { "Activity Key": x.activityKey, "Activity Name": x.activityName || x.activityKey, "Class ID": x.classId, "Group Key": x.groupKey, Status: x.status || "Active", "Allow Student Registration": x.allowRegistration !== false, "Registration Opens At": x.regOpen || "", "Registration Closes At": x.regClose || "", "Evaluation Opens At": x.evalOpen || "", "Evaluation Closes At": x.evalClose || "", "Last Updated": new Date() }; if (found) setRow_(s, found.rowNumber, rec); else add_(CONFIG.S.A, H.A, Object.assign({ Timestamp: new Date() }, rec)); return { activityKey: x.activityKey, activityName: rec["Activity Name"], classId: x.classId, groupKey: x.groupKey, status: rec.Status, allowRegistration: rec["Allow Student Registration"], regOpen: rec["Registration Opens At"], regClose: rec["Registration Closes At"], evalOpen: rec["Evaluation Opens At"], evalClose: rec["Evaluation Closes At"] }; }
function ensureRubric_(ak, ci) { if (rubric_(ak, ci).length) return; DEFAULT_RUBRIC.forEach((c) => add_(CONFIG.S.R, H.R, { Timestamp: new Date(), "Activity Key": ak, "Class ID": ci, "Criterion ID": c[0], "Criterion Label": c[1], "Max Score": c[2], Weight: c[3], Required: c[4], Status: "Active", "Last Updated": new Date() })); }
function rubric_(ak, ci) { return rubrics_().filter((x) => x.activityKey === ak && x.classId === ci && x.status === "Active"); }

function activities_() { return rows_(CONFIG.S.A).map((r) => ({ activityKey: safeCode_(r["Activity Key"]), activityName: text_(r["Activity Name"], 160), classId: safeCode_(r["Class ID"]), groupKey: safeCode_(r["Group Key"]), status: status_(r.Status), allowRegistration: truth_(r["Allow Student Registration"]), regOpen: r["Registration Opens At"], regClose: r["Registration Closes At"], evalOpen: r["Evaluation Opens At"], evalClose: r["Evaluation Closes At"] })).filter((x) => x.activityKey && x.classId && x.groupKey); }
function rubrics_() { return rows_(CONFIG.S.R).map((r) => ({ activityKey: safeCode_(r["Activity Key"]), classId: safeCode_(r["Class ID"]), criterionId: text_(r["Criterion ID"], 80), label: text_(r["Criterion Label"], 100), maxScore: num_(r["Max Score"]) || 10, weight: num_(r.Weight) || 0, required: truth_(r.Required), status: status_(r.Status) })).filter((x) => x.activityKey && x.classId && x.criterionId); }
function regs_() { return rows_(CONFIG.S.G).map((r) => ({ activityKey: safeCode_(r["Activity Key"]), classId: safeCode_(r["Class ID"]), groupKey: safeCode_(r["Group Key"]), studentId: sid_(r["Student ID"]), name: text_(r.Name, 120), role: text_(r.Role, 120), contribution: text_(r.Contribution, 1600), status: status_(r.Status) })).filter((x) => x.activityKey && x.classId && x.groupKey && x.studentId && x.name); }
function evals_() { return rows_(CONFIG.S.E).map((r) => ({ timestamp: r.Timestamp, activityKey: safeCode_(r["Activity Key"]), classId: safeCode_(r["Class ID"]), groupKey: safeCode_(r["Group Key"]), evaluatorStudentId: sid_(r["Evaluator Student ID"]), evaluatorName: text_(r["Evaluator Name"], 120), targetStudentId: sid_(r["Target Student ID"]), targetName: text_(r["Target Name"], 120), rawScore: num_(r["Raw Score"]), maxRawScore: num_(r["Max Raw Score"]), weightedScore: num_(r["Weighted Score"]), weightedPercent: num_(r["Weighted Percent"]), scoresJson: text_(r["Scores JSON"], 3000), comments: text_(r.Comments, 1200) })).filter((x) => x.activityKey && x.classId && x.groupKey && x.evaluatorStudentId && x.targetStudentId); }

function windowCheck_(a, type) { if (a.status !== "Active") return bad_("This activity is currently inactive."); if (type === "reg" && !a.allowRegistration) return bad_("Student registration is disabled for this activity."); const o = type === "reg" ? a.regOpen : a.evalOpen, c = type === "reg" ? a.regClose : a.evalClose, label = type === "reg" ? "registration" : "evaluation", now = new Date(); if (isFuture_(o, now)) return bad_("The " + label + " window has not opened yet."); if (isPast_(c, now)) return bad_("The " + label + " window is already closed."); return good_("Open"); }
function buildCohorts_(rs, es) { return uniqCohorts_(rs).map((c) => { const m = rs.filter((r) => cohortKey_(r) === c.key), ee = es.filter((e) => cohortKey_(e) === c.key), exp = m.length * Math.max(m.length - 1, 0); return Object.assign(c, { memberCount: m.length, submittedEvaluations: ee.length, expectedEvaluations: exp, completionPercent: exp ? Math.round((ee.length / exp) * 100) : 100, averagePercent: ee.length ? round2_(sum_(ee.map((x) => x.weightedPercent)) / ee.length) : null }); }); }
function buildSummary_(rs, es) { return rs.sort(sortRoster_).map((s) => { const members = rs.filter((r) => cohortKey_(r) === cohortKey_(s)), got = es.filter((e) => cohortKey_(e) === cohortKey_(s) && e.targetStudentId === s.studentId), ids = new Set(got.map((e) => e.evaluatorStudentId)), exp = Math.max(members.length - 1, 0); return Object.assign({}, s, { receivedCount: got.length, expectedCount: exp, completionPercent: exp ? Math.round((got.length / exp) * 100) : 100, averagePercent: got.length ? round2_(sum_(got.map((e) => e.weightedPercent)) / got.length) : null, missingEvaluators: members.filter((m) => m.studentId !== s.studentId && !ids.has(m.studentId)).map((m) => m.name) }); }); }
function buildCompletion_(rs, es) { return rs.sort(sortRoster_).map((s) => { const members = rs.filter((r) => cohortKey_(r) === cohortKey_(s)), sub = es.filter((e) => cohortKey_(e) === cohortKey_(s) && e.evaluatorStudentId === s.studentId), ids = new Set(sub.map((e) => e.targetStudentId)), exp = Math.max(members.length - 1, 0); return Object.assign({}, s, { submittedCount: sub.length, expectedCount: exp, completionPercent: exp ? Math.round((sub.length / exp) * 100) : 100, missingTargets: members.filter((m) => m.studentId !== s.studentId && !ids.has(m.studentId)).map((m) => m.name) }); }); }

function mkSheet_(ss, name, heads, color) { let s = ss.getSheetByName(name); if (!s) s = ss.insertSheet(name); s.getRange(1, 1, 1, heads.length).setValues([heads]).setFontWeight("bold").setBackground(color); s.setFrozenRows(1); return s; }
function sheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function add_(name, heads, obj) { sheet_(name).appendRow(heads.map((h) => (obj[h] === undefined ? "" : obj[h]))); }
function rows_(name) { const s = sheet_(name), lr = s.getLastRow(), lc = s.getLastColumn(); if (lr < 2) return []; const h = s.getRange(1, 1, 1, lc).getValues()[0].map((x) => text_(x, 120)); return s.getRange(2, 1, lr - 1, lc).getValues().map((row) => { const o = {}; h.forEach((k, i) => { if (k) o[k] = row[i]; }); return o; }); }
function table_(s) { return { rows: rowsFrom_(s).map((object, i) => ({ object, rowNumber: i + 2 })) }; }
function rowsFrom_(s) { const lr = s.getLastRow(), lc = s.getLastColumn(); if (lr < 2) return []; const h = s.getRange(1, 1, 1, lc).getValues()[0].map((x) => text_(x, 120)); return s.getRange(2, 1, lr - 1, lc).getValues().map((row) => { const o = {}; h.forEach((k, i) => { if (k) o[k] = row[i]; }); return o; }); }
function setRow_(s, row, obj) { const h = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map((x) => text_(x, 120)); Object.keys(obj).forEach((k) => { const i = h.indexOf(k); if (i > -1) s.getRange(row, i + 1).setValue(obj[k]); }); }
function audit_(type, id, action, detail) { try { add_(CONFIG.S.L, H.L, { Timestamp: new Date(), "Actor Type": type, "Actor ID": id, Action: action, Details: JSON.stringify(detail || {}) }); } catch (e) {} }
function filter_(v, label) { const x = text_(v || CONFIG.ALL, 100).toUpperCase(); return !x || x === CONFIG.ALL || x === "*" || x.indexOf("ALL ") === 0 ? { all: true, value: CONFIG.ALL } : { all: false, value: code_(x, label) }; }
function code_(v, label) { const x = text_(v, 100).toUpperCase(); if (!x) throw Error(label + " is required."); if (!/^[A-Z0-9][A-Z0-9 ._-]{0,99}$/.test(x)) throw Error(label + " may only contain letters, numbers, spaces, dashes, underscores, and periods."); return x; }
function optionalCode_(v) { return text_(v, 100).toUpperCase(); }
function safeCode_(v) { return text_(v, 100).toUpperCase(); }
function sid_(v) { return text_(v, 100).toUpperCase(); }
function status_(v) { const s = text_(v, 40).toLowerCase(); return s === "inactive" || s === "closed" ? "Inactive" : "Active"; }
function truth_(v) { return v === true || String(v).toLowerCase() === "true" || String(v).toLowerCase() === "yes" || String(v) === "1" || String(v).toLowerCase() === "active"; }
function text_(v, n) { return String(v == null ? "" : v).replace(/\s+/g, " ").trim().slice(0, n || 500); }
function num_(v) { const n = Number(v); return Number.isFinite(n) ? n : null; }
function dateInput_(v) { if (!v) return ""; const d = v instanceof Date ? v : new Date(String(v)); if (isNaN(d.getTime())) throw Error("Invalid date/time: " + v); return d; }
function isFuture_(v, now) { const d = v instanceof Date ? v : v ? new Date(v) : null; return d && !isNaN(d.getTime()) && d.getTime() > now.getTime(); }
function isPast_(v, now) { const d = v instanceof Date ? v : v ? new Date(v) : null; return d && !isNaN(d.getTime()) && d.getTime() < now.getTime(); }
function crit_(v) { return text_(v, 80).toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "criterion"; }
function sum_(a) { return a.reduce((t, n) => t + (Number(n) || 0), 0); }
function round2_(n) { return Math.round(n * 100) / 100; }
function fmt_(v) { if (!v) return ""; const d = v instanceof Date ? v : new Date(v); return isNaN(d.getTime()) ? String(v) : Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"); }
function uniq_(a) { return Array.from(new Set(a.filter(Boolean))).sort(); }
function uniqStudents_(rows) { const seen = {}; return rows.reduce((out, r) => { if (!r.studentId || seen[r.studentId]) return out; seen[r.studentId] = 1; out.push({ value: r.studentId, label: r.name ? r.studentId + " - " + r.name : r.studentId }); return out; }, []).sort((a, b) => a.value.localeCompare(b.value)); }
function rosterOptions_(rows) { return rows.map((r) => ({ activityKey: r.activityKey, classId: r.classId, groupKey: r.groupKey, studentId: r.studentId, name: r.name, status: r.status })).sort(sortRoster_); }
function duplicateSummary_(rs, es) { return { registrations: duplicateGroups_(rs.filter((r) => r.status === "Active"), (r) => [r.activityKey, r.classId, r.studentId].join("||"), (rows) => ({ activityKey: rows[0].activityKey, classId: rows[0].classId, studentId: rows[0].studentId, name: rows[0].name, count: rows.length, extraCount: rows.length - 1, groups: uniq_(rows.map((r) => r.groupKey)) })), evaluations: duplicateGroups_(es, (e) => [e.activityKey, e.classId, e.groupKey, e.evaluatorStudentId, e.targetStudentId].join("||"), (rows) => ({ activityKey: rows[0].activityKey, classId: rows[0].classId, groupKey: rows[0].groupKey, evaluatorStudentId: rows[0].evaluatorStudentId, evaluatorName: rows[0].evaluatorName, targetStudentId: rows[0].targetStudentId, targetName: rows[0].targetName, count: rows.length, extraCount: rows.length - 1 })) }; }
function duplicateGroups_(rows, keyFn, mapFn) { const buckets = {}; rows.forEach((row) => { const key = keyFn(row); if (!buckets[key]) buckets[key] = []; buckets[key].push(row); }); return Object.keys(buckets).filter((key) => buckets[key].length > 1).map((key) => mapFn(buckets[key])).sort((a, b) => a.classId.localeCompare(b.classId) || a.activityKey.localeCompare(b.activityKey) || String(a.groupKey || "").localeCompare(String(b.groupKey || "")) || String(a.studentId || a.evaluatorStudentId || "").localeCompare(String(b.studentId || b.evaluatorStudentId || ""))); }
function cohortKey_(r) { return [r.activityKey, r.classId, r.groupKey].join("||"); }
function memberKey_(r, id) { return cohortKey_(r) + "||" + String(id || r.studentId || "").toUpperCase(); }
function uniqCohorts_(rows) { const seen = {}; return rows.reduce((out, r) => { const k = cohortKey_(r); if (!seen[k]) { seen[k] = 1; out.push({ key: k, activityKey: r.activityKey, classId: r.classId, groupKey: r.groupKey }); } return out; }, []).sort((a, b) => a.classId.localeCompare(b.classId) || a.activityKey.localeCompare(b.activityKey) || a.groupKey.localeCompare(b.groupKey)); }
function sortRoster_(a, b) { return a.classId.localeCompare(b.classId) || a.activityKey.localeCompare(b.activityKey) || a.groupKey.localeCompare(b.groupKey) || a.name.localeCompare(b.name); }
function sortEval_(a, b) { return a.classId.localeCompare(b.classId) || a.activityKey.localeCompare(b.activityKey) || a.groupKey.localeCompare(b.groupKey) || a.targetName.localeCompare(b.targetName) || a.evaluatorName.localeCompare(b.evaluatorName); }
function pin_() { return PropertiesService.getScriptProperties().getProperty("FACULTY_PIN") || CONFIG.PIN; }
function good_(m, d) { return Object.assign({ success: true, message: m }, d || {}); }
function bad_(m) { return { success: false, message: m }; }
