/**
 * Student profile + faculty-controlled activity flow.
 * This file keeps the existing reporting/faculty backend intact while replacing
 * the student portal behavior with: register profile -> login to existing key -> evaluate.
 */
const AUTH_FLOW = {
  SHEET: "Student Profiles",
  HEADERS: ["Timestamp", "Student ID", "Name", "Passkey Hash", "Status", "Last Updated"]
};

function ensureStudentProfileSetup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  mkSheet_(ss, AUTH_FLOW.SHEET, AUTH_FLOW.HEADERS, "#dcfce7");
}

function getStudentHintsV2() {
  ensureSetup();
  ensureStudentProfileSetup_();
  const a = activities_().filter((x) => x.status === "Active");
  return good_("Hints loaded.", {
    activities: uniq_(a.map((x) => x.activityKey)),
    classes: uniq_(a.map((x) => x.classId).filter((x) => x !== CONFIG.ALL)),
    cohorts: uniqCohorts_(a)
  });
}

function saveActivitySettingsAllClass(d, pin) {
  d = d || {};
  d.classId = optionalCode_(d.classId || d.classCode) || CONFIG.ALL;
  d.groupKey = optionalCode_(d.groupKey) || CONFIG.ALL;
  return saveActivitySettings(d, pin);
}

function saveRubricAllClass(d, pin) {
  d = d || {};
  d.classId = optionalCode_(d.classId || d.classCode) || CONFIG.ALL;
  return saveRubric(d, pin);
}

function saveStudentProfile(f) {
  ensureSetup();
  ensureStudentProfileSetup_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const studentId = sid_(f.studentId);
    const name = text_(f.name, 120);
    const passkey = validateStudentPasskey_(f.passkey);
    if (!studentId || !name) return bad_("ID number and name are required.");

    const s = sheet_(AUTH_FLOW.SHEET);
    const data = table_(s);
    const row = data.rows.find((r) => sid_(r.object["Student ID"]) === studentId);
    const hash = hashStudentPasskey_(studentId, passkey);

    if (row) {
      if (status_(row.object.Status) !== "Active") return bad_("This student profile is inactive. Ask your faculty to reactivate it.");
      const stored = text_(row.object["Passkey Hash"], 300);
      if (stored && stored !== hash) return bad_("This ID number already has a profile. Enter the existing passkey.");
      setRow_(s, row.rowNumber, { Name: name, "Passkey Hash": hash, "Last Updated": new Date() });
      audit_("Student", studentId, "VERIFY_PROFILE", { studentId, name });
      return good_("Profile verified. Continue to activity login.", { studentId, name });
    }

    add_(AUTH_FLOW.SHEET, AUTH_FLOW.HEADERS, {
      Timestamp: new Date(),
      "Student ID": studentId,
      Name: name,
      "Passkey Hash": hash,
      Status: "Active",
      "Last Updated": new Date()
    });
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
  ensureStudentProfileSetup_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const studentId = sid_(d.studentId);
    const passkey = validateStudentPasskey_(d.passkey);
    const activityKey = code_(d.activityKey, "Activity Key");
    const classId = code_(d.classCode || d.classId, "Class Code");
    const role = text_(d.groupRole || d.role, 120);
    const contribution = text_(d.contribution, 1600);
    if (!role || !contribution) return bad_("Activity Key, Class Code, Group Role, and Contribution are required.");

    const profile = verifyStudentProfile_(studentId, passkey);
    if (!profile.success) return profile;

    const found = requireFacultyActivity_(activityKey, classId);
    if (!found.success) return found;
    const activity = found.activity;
    const regWindow = windowCheck_(activity, "reg");
    if (!regWindow.success) return regWindow;

    ensureRubric_(activityKey, activity.classId || classId);
    const me = upsertStudentActivityRegistration_({
      activityKey,
      classId,
      groupKey: activity.groupKey,
      studentId,
      name: profile.name,
      role,
      contribution
    });

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
  ensureStudentProfileSetup_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const activityKey = code_(d.activityKey, "Activity Key");
    const classId = code_(d.classCode || d.classId, "Class Code");
    const evaluatorId = sid_(d.evaluatorStudentId || d.studentId);
    const targetId = sid_(d.targetStudentId);
    const passkey = validateStudentPasskey_(d.passkey);
    const comments = text_(d.comments, 1200);
    const scores = d.scores || {};

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
    const evalWindow = windowCheck_(found.activity, "eval");
    if (!evalWindow.success) return evalWindow;

    if (evals_().some((x) => x.activityKey === activityKey && x.classId === classId && x.groupKey === me.groupKey && x.evaluatorStudentId === evaluatorId && x.targetStudentId === targetId)) {
      return bad_("You already evaluated this member.");
    }

    let raw = 0;
    let maxRaw = 0;
    let weighted = 0;
    let totalWeight = 0;
    const detail = [];
    const activeRubric = rubricForActivityClass_(activityKey, classId);
    for (const c of activeRubric) {
      const v = Number(scores[c.criterionId]);
      if (c.required && (!Number.isInteger(v) || v < 1 || v > c.maxScore)) return bad_("Please rate " + c.label + " from 1 to " + c.maxScore + ".");
      if (Number.isInteger(v)) {
        raw += v;
        maxRaw += c.maxScore;
        weighted += (v / c.maxScore) * c.weight;
        totalWeight += c.weight;
        detail.push({ id: c.criterionId, label: c.label, score: v, maxScore: c.maxScore, weight: c.weight });
      }
    }

    const percent = totalWeight ? round2_((weighted / totalWeight) * 100) : 0;
    add_(CONFIG.S.E, H.E, {
      Timestamp: new Date(),
      "Activity Key": activityKey,
      "Class ID": classId,
      "Group Key": me.groupKey,
      "Evaluator Student ID": evaluatorId,
      "Evaluator Name": me.name,
      "Target Student ID": targetId,
      "Target Name": target.name,
      "Raw Score": raw,
      "Max Raw Score": maxRaw,
      "Weighted Score": round2_(weighted),
      "Weighted Percent": percent,
      "Scores JSON": JSON.stringify(detail),
      Comments: comments
    });
    audit_("Student", evaluatorId, "SUBMIT_EVALUATION", { activityKey, classId, target: targetId, percent });
    return loadStudentEvaluationSession_(me, passkey, "Evaluation for " + target.name + " saved.");
  } catch (e) {
    return bad_("System Error: " + (e.message || String(e)));
  } finally {
    lock.releaseLock();
  }
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
  const s = sheet_(CONFIG.S.G);
  const data = table_(s);
  const row = data.rows.find((r) => safeCode_(r.object["Activity Key"]) === x.activityKey && safeCode_(r.object["Class ID"]) === x.classId && sid_(r.object["Student ID"]) === x.studentId);
  const rec = {
    "Activity Key": x.activityKey,
    "Class ID": x.classId,
    "Group Key": x.groupKey || CONFIG.ALL,
    "Student ID": x.studentId,
    Name: x.name,
    Role: x.role,
    Contribution: x.contribution,
    Status: "Active",
    "Last Updated": new Date()
  };
  if (row) setRow_(s, row.rowNumber, rec);
  else add_(CONFIG.S.G, H.G, Object.assign({ Timestamp: new Date() }, rec));
  return { activityKey: x.activityKey, classId: x.classId, groupKey: rec["Group Key"], studentId: x.studentId, name: x.name, role: x.role, contribution: x.contribution, status: "Active" };
}

function findStudentActivityRegistration_(activityKey, classId, studentId) {
  return regs_().find((x) => x.activityKey === activityKey && x.classId === classId && x.studentId === studentId && x.status === "Active");
}

function loadStudentEvaluationSession_(me, passkey, message) {
  const found = requireFacultyActivity_(me.activityKey, me.classId, me.groupKey);
  if (!found.success) return found;
  const evalWindow = windowCheck_(found.activity, "eval");
  if (!evalWindow.success) return evalWindow;
  const done = new Set(evals_().filter((x) => x.activityKey === me.activityKey && x.classId === me.classId && x.groupKey === me.groupKey && x.evaluatorStudentId === me.studentId).map((x) => x.targetStudentId));
  const targets = regs_()
    .filter((x) => x.activityKey === me.activityKey && x.classId === me.classId && x.groupKey === me.groupKey && x.status === "Active" && x.studentId !== me.studentId)
    .sort((a, b) => a.name.localeCompare(b.name))
    .map((x) => ({ studentId: x.studentId, name: x.name, role: x.role, contribution: x.contribution, alreadyEvaluated: done.has(x.studentId) }));
  const remainingTargets = targets.filter((x) => !x.alreadyEvaluated);
  return good_(message || "Group loaded.", {
    evaluator: me,
    activity: found.activity,
    targets,
    remainingTargets,
    rubric: rubricForActivityClass_(me.activityKey, me.classId),
    completedCount: targets.length - remainingTargets.length,
    totalTargets: targets.length,
    passkey
  });
}

function rubricForActivityClass_(activityKey, classId) {
  const exact = rubric_(activityKey, classId);
  return exact.length ? exact : rubric_(activityKey, CONFIG.ALL);
}

function verifyStudentProfile_(studentId, passkey) {
  if (!studentId) return bad_("ID number is required.");
  const p = rows_(AUTH_FLOW.SHEET).map((r) => ({ studentId: sid_(r["Student ID"]), name: text_(r.Name, 120), hash: text_(r["Passkey Hash"], 300), status: status_(r.Status) })).find((x) => x.studentId === studentId && x.status === "Active");
  if (!p) return bad_("Student profile not found. Register your ID number first.");
  if (p.hash !== hashStudentPasskey_(studentId, passkey)) return bad_("Incorrect passkey.");
  return good_("Profile verified.", { studentId: p.studentId, name: p.name });
}

function validateStudentPasskey_(v) {
  const p = text_(v, 120);
  if (!p) throw Error("Passkey is required.");
  if (p.length < 4) throw Error("Passkey must be at least 4 characters.");
  return p;
}

function hashStudentPasskey_(studentId, passkey) {
  const salt = PropertiesService.getScriptProperties().getProperty("PASSKEY_SALT") || "peer-evaluation-v2";
  const raw = sid_(studentId) + ":" + String(passkey || "") + ":" + salt;
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytes.map((b) => (b + 256).toString(16).slice(-2)).join("");
}
