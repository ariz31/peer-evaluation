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
    L: "Audit Log"
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
  A: [
    "Timestamp",
    "Activity Key",
    "Activity Name",
    "Class ID",
    "Group Key",
    "Status",
    "Allow Student Registration",
    "Registration Opens At",
    "Registration Closes At",
    "Evaluation Opens At",
    "Evaluation Closes At",
    "Last Updated"
  ],
  R: [
    "Timestamp",
    "Activity Key",
    "Class ID",
    "Criterion ID",
    "Criterion Label",
    "Max Score",
    "Weight",
    "Required",
    "Status",
    "Last Updated"
  ],
  G: [
    "Timestamp",
    "Activity Key",
    "Class ID",
    "Group Key",
    "Student ID",
    "Name",
    "Role",
    "Contribution",
    "Status",
    "Last Updated"
  ],
  E: [
    "Timestamp",
    "Activity Key",
    "Class ID",
    "Group Key",
    "Evaluator Student ID",
    "Evaluator Name",
    "Target Student ID",
    "Target Name",
    "Raw Score",
    "Max Raw Score",
    "Weighted Score",
    "Weighted Percent",
    "Scores JSON",
    "Comments"
  ],
  L: ["Timestamp", "Actor Type", "Actor ID", "Action", "Details"]
};
function doGet(e) {
  ensureSetup();
  const v = String((e && e.parameter && e.parameter.view) || "").toLowerCase();
  return HtmlService.createHtmlOutputFromFile(
    v === "faculty" ? "Faculty" : "Index"
  )
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
    [CONFIG.S.L, H.L, "#fef3c7"]
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
function saveRegistration(f) {
  ensureSetup();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const r = {
      activityKey: code_(f.activityKey, "Activity"),
      classId: code_(f.classId, "Class ID"),
      groupKey: code_(f.groupKey, "Group Key"),
      studentId: sid_(f.studentId),
      name: text_(f.name, 120),
      role: text_(f.role, 120),
      contribution: text_(f.contribution, 1600)
    };
    for (const k of [
      "activityKey",
      "classId",
      "groupKey",
      "studentId",
      "name",
      "role",
      "contribution"
    ])
      if (!r[k]) return bad_("Please complete all registration fields.");
    const a = getOrCreateActivity_(
      r.activityKey,
      r.classId,
      r.groupKey,
      r.activityKey
    );
    let w = windowCheck_(a, "reg");
    if (!w.success) return w;
    ensureRubric_(r.activityKey, r.classId);
    if (
      regs_().some(
        (x) =>
          x.activityKey === r.activityKey &&
          x.classId === r.classId &&
          x.studentId === r.studentId &&
          x.status === "Active"
      )
    )
      return bad_(
        "This Student ID is already registered for this Activity and Class."
      );
    add_(CONFIG.S.G, H.G, {
      Timestamp: new Date(),
      "Activity Key": r.activityKey,
      "Class ID": r.classId,
      "Group Key": r.groupKey,
      "Student ID": r.studentId,
      Name: r.name,
      Role: r.role,
      Contribution: r.contribution,
      Status: "Active",
      "Last Updated": new Date()
    });
    audit_("Student", r.studentId, "REGISTER", r);
    return good_(
      "Registration successful. You may now proceed to peer evaluation.",
      r
    );
  } catch (e) {
    return bad_("System Error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}
function verifyAndLoadGroup(studentId, activityKey, classId) {
  ensureSetup();
  try {
    const ak = code_(activityKey, "Activity"),
      ci = code_(classId, "Class ID"),
      id = sid_(studentId),
      rs = regs_(),
      es = evals_();
    const me = rs.find(
      (x) =>
        x.activityKey === ak &&
        x.classId === ci &&
        x.studentId === id &&
        x.status === "Active"
    );
    if (!me)
      return bad_(
        "Student ID not found for this Activity and Class. Please register first."
      );
    const a = getOrCreateActivity_(ak, ci, me.groupKey, ak),
      w = windowCheck_(a, "eval");
    if (!w.success) return w;
    const done = new Set(
      es
        .filter(
          (x) =>
            x.activityKey === ak &&
            x.classId === ci &&
            x.groupKey === me.groupKey &&
            x.evaluatorStudentId === id
        )
        .map((x) => x.targetStudentId)
    );
    const targets = rs
      .filter(
        (x) =>
          x.activityKey === ak &&
          x.classId === ci &&
          x.groupKey === me.groupKey &&
          x.status === "Active" &&
          x.studentId !== id
      )
      .sort((a, b) => a.name.localeCompare(b.name))
      .map((x) => ({
        studentId: x.studentId,
        name: x.name,
        role: x.role,
        contribution: x.contribution,
        alreadyEvaluated: done.has(x.studentId)
      }));
    const remainingTargets = targets.filter((x) => !x.alreadyEvaluated);
    return good_("Group loaded.", {
      evaluator: me,
      targets,
      remainingTargets,
      rubric: rubric_(ak, ci),
      completedCount: targets.length - remainingTargets.length,
      totalTargets: targets.length
    });
  } catch (e) {
    return bad_("System Error: " + e.message);
  }
}
function saveEvaluation(d) {
  ensureSetup();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ak = code_(d.activityKey, "Activity"),
      ci = code_(d.classId, "Class ID"),
      eid = sid_(d.evaluatorStudentId),
      tid = sid_(d.targetStudentId),
      scores = d.scores || {},
      rs = regs_(),
      es = evals_();
    const me = rs.find(
      (x) =>
        x.activityKey === ak &&
        x.classId === ci &&
        x.studentId === eid &&
        x.status === "Active"
    );
    if (!me) return bad_("Evaluator record not found.");
    const target = rs.find(
      (x) =>
        x.activityKey === ak &&
        x.classId === ci &&
        x.groupKey === me.groupKey &&
        x.studentId === tid &&
        x.status === "Active"
    );
    if (!target) return bad_("Target member is not registered in your group.");
    if (eid === tid) return bad_("You cannot evaluate yourself.");
    const w = windowCheck_(
      getOrCreateActivity_(ak, ci, me.groupKey, ak),
      "eval"
    );
    if (!w.success) return w;
    if (
      es.some(
        (x) =>
          x.activityKey === ak &&
          x.classId === ci &&
          x.groupKey === me.groupKey &&
          x.evaluatorStudentId === eid &&
          x.targetStudentId === tid
      )
    )
      return bad_("You already evaluated this member.");
    let raw = 0,
      maxRaw = 0,
      weighted = 0,
      totalWeight = 0,
      detail = [];
    for (const c of rubric_(ak, ci)) {
      const v = Number(scores[c.criterionId]);
      if (c.required && (!Number.isInteger(v) || v < 1 || v > c.maxScore))
        return bad_(
          "Please rate " + c.label + " from 1 to " + c.maxScore + "."
        );
      if (Number.isInteger(v)) {
        raw += v;
        maxRaw += c.maxScore;
        weighted += (v / c.maxScore) * c.weight;
        totalWeight += c.weight;
        detail.push({
          id: c.criterionId,
          label: c.label,
          score: v,
          maxScore: c.maxScore,
          weight: c.weight
        });
      }
    }
    const percent = totalWeight ? round2_((weighted / totalWeight) * 100) : 0;
    add_(CONFIG.S.E, H.E, {
      Timestamp: new Date(),
      "Activity Key": ak,
      "Class ID": ci,
      "Group Key": me.groupKey,
      "Evaluator Student ID": eid,
      "Evaluator Name": me.name,
      "Target Student ID": tid,
      "Target Name": target.name,
      "Raw Score": raw,
      "Max Raw Score": maxRaw,
      "Weighted Score": round2_(weighted),
      "Weighted Percent": percent,
      "Scores JSON": JSON.stringify(detail),
      Comments: text_(d.comments, 1200)
    });
    audit_("Student", eid, "SUBMIT_EVALUATION", {
      ak,
      ci,
      target: tid,
      percent
    });
    const snap = verifyAndLoadGroup(eid, ak, ci);
    return good_("Evaluation for " + target.name + " saved.", {
      targetName: target.name,
      weightedPercent: percent,
      remainingTargets: snap.remainingTargets || [],
      completedCount: snap.completedCount,
      totalTargets: snap.totalTargets
    });
  } catch (e) {
    return bad_("System Error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}
function unlockFaculty(pin) {
  ensureSetup();
  if (String(pin || "") !== pin_()) return bad_("Incorrect faculty code.");
  return good_("Unlocked.", facultyState_());
}
function saveActivitySettings(d, pin) {
  if (String(pin || "") !== pin_()) return bad_("Incorrect faculty code.");
  try {
    const x = {
      activityKey: code_(d.activityKey, "Activity"),
      classId: code_(d.classId, "Class ID"),
      groupKey: code_(d.groupKey, "Group Key"),
      activityName: text_(d.activityName || d.activityKey, 160),
      status: status_(d.status),
      allowRegistration: truth_(d.allowRegistration),
      regOpen: dateInput_(d.regOpen),
      regClose: dateInput_(d.regClose),
      evalOpen: dateInput_(d.evalOpen),
      evalClose: dateInput_(d.evalClose)
    };
    upsertActivity_(x);
    ensureRubric_(x.activityKey, x.classId);
    audit_("Faculty", "FACULTY", "SAVE_ACTIVITY", x);
    return good_("Activity settings saved.", facultyState_());
  } catch (e) {
    return bad_(e.message);
  }
}
function saveActivitySettingsOptional(d, p) {
  d = d || {};
  d.groupKey = optionalCode_(d.groupKey) || CONFIG.ALL;
  return saveActivitySettings(d, p);
}
function saveRubric(d, pin) {
  if (String(pin || "") !== pin_()) return bad_("Incorrect faculty code.");
  try {
    const ak = code_(d.activityKey, "Activity"),
      ci = code_(d.classId, "Class ID"),
      items = (d.criteria || []).filter((x) => text_(x.label, 80));
    if (!items.length) return bad_("Add at least one criterion.");
    const s = sheet_(CONFIG.S.R),
      data = table_(s);
    for (let i = data.rows.length - 1; i >= 0; i--) {
      const o = data.rows[i].object;
      if (
        safeCode_(o["Activity Key"]) === ak &&
        safeCode_(o["Class ID"]) === ci
      )
        s.deleteRow(data.rows[i].rowNumber);
    }
    items.forEach((c, i) =>
      add_(CONFIG.S.R, H.R, {
        Timestamp: new Date(),
        "Activity Key": ak,
        "Class ID": ci,
        "Criterion ID": crit_(c.id || c.label || "c" + i),
        "Criterion Label": text_(c.label, 100),
        "Max Score": num_(c.maxScore) || 10,
        Weight: num_(c.weight) || 0,
        Required: truth_(c.required),
        Status: status_(c.status),
        "Last Updated": new Date()
      })
    );
    audit_("Faculty", "FACULTY", "SAVE_RUBRIC", {
      ak,
      ci,
      count: items.length
    });
    return good_("Rubric saved.", facultyState_());
  } catch (e) {
    return bad_(e.message);
  }
}
function moveStudentToGroup(d, pin) {
  return editReg_(
    d,
    pin,
    () => ({
      "Group Key": code_(d.newGroupKey, "New Group Key"),
      "Last Updated": new Date()
    }),
    "Student moved to group."
  );
}
function updateStudentContribution(d, pin) {
  return editReg_(
    d,
    pin,
    () => ({
      Contribution: text_(d.contribution, 1600),
      "Last Updated": new Date()
    }),
    "Contribution updated."
  );
}
function setRegistrationStatus(d, pin) {
  return editReg_(
    d,
    pin,
    () => ({ Status: status_(d.status), "Last Updated": new Date() }),
    "Registration status updated."
  );
}
function editReg_(d, pin, changes, msg) {
  if (String(pin || "") !== pin_()) return bad_("Incorrect faculty code.");
  try {
    const ak = code_(d.activityKey, "Activity"),
      ci = code_(d.classId, "Class ID"),
      id = sid_(d.studentId),
      s = sheet_(CONFIG.S.G),
      data = table_(s);
    const row = data.rows.find(
      (r) =>
        safeCode_(r.object["Activity Key"]) === ak &&
        safeCode_(r.object["Class ID"]) === ci &&
        sid_(r.object["Student ID"]) === id
    );
    if (!row) return bad_("Student registration not found.");
    const c = changes(row.object);
    if (c["Group Key"]) getOrCreateActivity_(ak, ci, c["Group Key"], ak);
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
    const ci = code_(f.classId, "Class ID"),
      af = filter_(f.activityKey, "Activity"),
      gf = filter_(f.groupKey, "Group Key");
    const rs = regs_().filter(
      (x) =>
        x.classId === ci &&
        (af.all || x.activityKey === af.value) &&
        (gf.all || x.groupKey === gf.value) &&
        x.status === "Active"
    );
    if (!rs.length)
      return bad_("No registered students found for the selected filters.");
    const keys = new Set(rs.map((x) => memberKey_(x)));
    const es = evals_().filter(
      (x) =>
        x.classId === ci &&
        (af.all || x.activityKey === af.value) &&
        (gf.all || x.groupKey === gf.value) &&
        keys.has(memberKey_(x, x.evaluatorStudentId)) &&
        keys.has(memberKey_(x, x.targetStudentId))
    );
    const cohorts = buildCohorts_(rs, es),
      summary = buildSummary_(rs, es),
      completion = buildCompletion_(rs, es);
    return good_("Report generated.", {
      generatedAt: fmt_(new Date()),
      filters: {
        classId: ci,
        activityKey: af.value,
        groupKey: gf.value,
        activityIsAll: af.all,
        groupIsAll: gf.all
      },
      totals: {
        registeredStudents: rs.length,
        cohorts: cohorts.length,
        submittedEvaluations: es.length,
        expectedEvaluations: cohorts.reduce(
          (t, c) => t + c.expectedEvaluations,
          0
        ),
        averagePercent: es.length
          ? round2_(sum_(es.map((x) => x.weightedPercent)) / es.length)
          : null
      },
      cohorts,
      summary,
      completion,
      details: es
        .sort(sortEval_)
        .map((x) => Object.assign({}, x, { timestamp: fmt_(x.timestamp) })),
      roster: rs.sort(sortRoster_)
    });
  } catch (e) {
    return bad_("System Error: " + e.message);
  }
}
function facultyState_() {
  const a = activities_(),
    r = regs_(),
    e = evals_();
  return {
    options: {
      activities: uniq_(
        a.map((x) => x.activityKey).concat(r.map((x) => x.activityKey))
      ),
      classes: uniq_(a.map((x) => x.classId).concat(r.map((x) => x.classId))),
      groups: uniq_(a.map((x) => x.groupKey).concat(r.map((x) => x.groupKey))),
      students: uniqStudents_(r),
      cohorts: uniqCohorts_(a.concat(r))
    },
    dashboard: dashboard_(a, r, e),
    activities: a,
    rubrics: rubrics_()
  };
}
function dashboard_(a, r, e) {
  const active = r.filter((x) => x.status === "Active"),
    cohorts = uniqCohorts_(active),
    single = cohorts.filter(
      (c) => active.filter((x) => cohortKey_(x) === c.key).length < 2
    ),
    empty = cohorts.filter((c) => !e.some((x) => cohortKey_(x) === c.key)),
    incomplete = buildCompletion_(active, e).filter(
      (x) => x.submittedCount < x.expectedCount
    );
  return {
    totalActivities: a.length,
    activeActivities: a.filter((x) => x.status === "Active").length,
    classes: uniq_(r.map((x) => x.classId)).length,
    cohorts: cohorts.length,
    registeredStudents: active.length,
    submittedEvaluations: e.length,
    incompleteEvaluators: incomplete.length,
    singleMemberGroups: single.length,
    cohortsWithNoSubmissions: empty.length,
    incomplete: incomplete.slice(0, 50),
    singleMemberCohorts: single,
    emptyCohorts: empty
  };
}
function getOrCreateActivity_(ak, ci, gk, name) {
  const rows = activities_();
  let a = rows.find(
    (x) => x.activityKey === ak && x.classId === ci && x.groupKey === gk
  );
  if (a) return a;
  a = rows.find(
    (x) => x.activityKey === ak && x.classId === ci && x.groupKey === CONFIG.ALL
  );
  if (a) return a;
  return upsertActivity_({
    activityKey: ak,
    classId: ci,
    groupKey: gk,
    activityName: name || ak,
    status: "Active",
    allowRegistration: true,
    regOpen: "",
    regClose: "",
    evalOpen: "",
    evalClose: ""
  });
}
function upsertActivity_(x) {
  const s = sheet_(CONFIG.S.A),
    data = table_(s),
    found = data.rows.find(
      (r) =>
        safeCode_(r.object["Activity Key"]) === x.activityKey &&
        safeCode_(r.object["Class ID"]) === x.classId &&
        safeCode_(r.object["Group Key"]) === x.groupKey
    );
  const rec = {
    "Activity Key": x.activityKey,
    "Activity Name": x.activityName || x.activityKey,
    "Class ID": x.classId,
    "Group Key": x.groupKey,
    Status: x.status || "Active",
    "Allow Student Registration": x.allowRegistration !== false,
    "Registration Opens At": x.regOpen || "",
    "Registration Closes At": x.regClose || "",
    "Evaluation Opens At": x.evalOpen || "",
    "Evaluation Closes At": x.evalClose || "",
    "Last Updated": new Date()
  };
  if (found) setRow_(s, found.rowNumber, rec);
  else add_(CONFIG.S.A, H.A, Object.assign({ Timestamp: new Date() }, rec));
  return {
    activityKey: x.activityKey,
    activityName: rec["Activity Name"],
    classId: x.classId,
    groupKey: x.groupKey,
    status: rec.Status,
    allowRegistration: rec["Allow Student Registration"],
    regOpen: rec["Registration Opens At"],
    regClose: rec["Registration Closes At"],
    evalOpen: rec["Evaluation Opens At"],
    evalClose: rec["Evaluation Closes At"]
  };
}
function ensureRubric_(ak, ci) {
  if (rubric_(ak, ci).length) return;
  DEFAULT_RUBRIC.forEach((c) =>
    add_(CONFIG.S.R, H.R, {
      Timestamp: new Date(),
      "Activity Key": ak,
      "Class ID": ci,
      "Criterion ID": c[0],
      "Criterion Label": c[1],
      "Max Score": c[2],
      Weight: c[3],
      Required: c[4],
      Status: "Active",
      "Last Updated": new Date()
    })
  );
}
function rubric_(ak, ci) {
  return rubrics_().filter(
    (x) => x.activityKey === ak && x.classId === ci && x.status === "Active"
  );
}
function activities_() {
  return rows_(CONFIG.S.A)
    .map((r) => ({
      activityKey: safeCode_(r["Activity Key"]),
      activityName: text_(r["Activity Name"], 160),
      classId: safeCode_(r["Class ID"]),
      groupKey: safeCode_(r["Group Key"]),
      status: status_(r.Status),
      allowRegistration: truth_(r["Allow Student Registration"]),
      regOpen: r["Registration Opens At"],
      regClose: r["Registration Closes At"],
      evalOpen: r["Evaluation Opens At"],
      evalClose: r["Evaluation Closes At"]
    }))
    .filter((x) => x.activityKey && x.classId && x.groupKey);
}
function rubrics_() {
  return rows_(CONFIG.S.R)
    .map((r) => ({
      activityKey: safeCode_(r["Activity Key"]),
      classId: safeCode_(r["Class ID"]),
      criterionId: text_(r["Criterion ID"], 80),
      label: text_(r["Criterion Label"], 100),
      maxScore: num_(r["Max Score"]) || 10,
      weight: num_(r.Weight) || 0,
      required: truth_(r.Required),
      status: status_(r.Status)
    }))
    .filter((x) => x.activityKey && x.classId && x.criterionId);
}
function regs_() {
  return rows_(CONFIG.S.G)
    .map((r) => ({
      activityKey: safeCode_(r["Activity Key"]),
      classId: safeCode_(r["Class ID"]),
      groupKey: safeCode_(r["Group Key"]),
      studentId: sid_(r["Student ID"]),
      name: text_(r.Name, 120),
      role: text_(r.Role, 120),
      contribution: text_(r.Contribution, 1600),
      status: status_(r.Status)
    }))
    .filter(
      (x) => x.activityKey && x.classId && x.groupKey && x.studentId && x.name
    );
}
function evals_() {
  return rows_(CONFIG.S.E)
    .map((r) => ({
      timestamp: r.Timestamp,
      activityKey: safeCode_(r["Activity Key"]),
      classId: safeCode_(r["Class ID"]),
      groupKey: safeCode_(r["Group Key"]),
      evaluatorStudentId: sid_(r["Evaluator Student ID"]),
      evaluatorName: text_(r["Evaluator Name"], 120),
      targetStudentId: sid_(r["Target Student ID"]),
      targetName: text_(r["Target Name"], 120),
      rawScore: num_(r["Raw Score"]),
      maxRawScore: num_(r["Max Raw Score"]),
      weightedScore: num_(r["Weighted Score"]),
      weightedPercent: num_(r["Weighted Percent"]),
      scoresJson: text_(r["Scores JSON"], 3000),
      comments: text_(r.Comments, 1200)
    }))
    .filter(
      (x) =>
        x.activityKey &&
        x.classId &&
        x.groupKey &&
        x.evaluatorStudentId &&
        x.targetStudentId
    );
}
function windowCheck_(a, type) {
  if (a.status !== "Active")
    return bad_("This activity is currently inactive.");
  if (type === "reg" && !a.allowRegistration)
    return bad_("Student registration is disabled for this activity.");
  const o = type === "reg" ? a.regOpen : a.evalOpen,
    c = type === "reg" ? a.regClose : a.evalClose,
    label = type === "reg" ? "registration" : "evaluation",
    now = new Date();
  if (isFuture_(o, now))
    return bad_("The " + label + " window has not opened yet.");
  if (isPast_(c, now))
    return bad_("The " + label + " window is already closed.");
  return good_("Open");
}
function buildCohorts_(rs, es) {
  return uniqCohorts_(rs).map((c) => {
    const m = rs.filter((r) => cohortKey_(r) === c.key),
      ee = es.filter((e) => cohortKey_(e) === c.key),
      exp = m.length * Math.max(m.length - 1, 0);
    return Object.assign(c, {
      memberCount: m.length,
      submittedEvaluations: ee.length,
      expectedEvaluations: exp,
      completionPercent: exp ? Math.round((ee.length / exp) * 100) : 100,
      averagePercent: ee.length
        ? round2_(sum_(ee.map((x) => x.weightedPercent)) / ee.length)
        : null
    });
  });
}
function buildSummary_(rs, es) {
  return rs.sort(sortRoster_).map((s) => {
    const members = rs.filter((r) => cohortKey_(r) === cohortKey_(s)),
      got = es.filter(
        (e) =>
          cohortKey_(e) === cohortKey_(s) && e.targetStudentId === s.studentId
      ),
      ids = new Set(got.map((e) => e.evaluatorStudentId)),
      exp = Math.max(members.length - 1, 0);
    return Object.assign({}, s, {
      receivedCount: got.length,
      expectedCount: exp,
      completionPercent: exp ? Math.round((got.length / exp) * 100) : 100,
      averagePercent: got.length
        ? round2_(sum_(got.map((e) => e.weightedPercent)) / got.length)
        : null,
      missingEvaluators: members
        .filter((m) => m.studentId !== s.studentId && !ids.has(m.studentId))
        .map((m) => m.name)
    });
  });
}
function buildCompletion_(rs, es) {
  return rs.sort(sortRoster_).map((s) => {
    const members = rs.filter((r) => cohortKey_(r) === cohortKey_(s)),
      sub = es.filter(
        (e) =>
          cohortKey_(e) === cohortKey_(s) &&
          e.evaluatorStudentId === s.studentId
      ),
      ids = new Set(sub.map((e) => e.targetStudentId)),
      exp = Math.max(members.length - 1, 0);
    return Object.assign({}, s, {
      submittedCount: sub.length,
      expectedCount: exp,
      completionPercent: exp ? Math.round((sub.length / exp) * 100) : 100,
      missingTargets: members
        .filter((m) => m.studentId !== s.studentId && !ids.has(m.studentId))
        .map((m) => m.name)
    });
  });
}
function mkSheet_(ss, name, heads, color) {
  let s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  s.getRange(1, 1, 1, heads.length)
    .setValues([heads])
    .setFontWeight("bold")
    .setBackground(color);
  s.setFrozenRows(1);
  return s;
}
function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}
function add_(name, heads, obj) {
  sheet_(name).appendRow(
    heads.map((h) => (obj[h] === undefined ? "" : obj[h]))
  );
}
function rows_(name) {
  const s = sheet_(name),
    lr = s.getLastRow(),
    lc = s.getLastColumn();
  if (lr < 2) return [];
  const h = s
    .getRange(1, 1, 1, lc)
    .getValues()[0]
    .map((x) => text_(x, 120));
  return s
    .getRange(2, 1, lr - 1, lc)
    .getValues()
    .map((row) => {
      const o = {};
      h.forEach((k, i) => {
        if (k) o[k] = row[i];
      });
      return o;
    });
}
function table_(s) {
  return {
    rows: rowsFrom_(s).map((object, i) => ({ object, rowNumber: i + 2 }))
  };
}
function rowsFrom_(s) {
  const lr = s.getLastRow(),
    lc = s.getLastColumn();
  if (lr < 2) return [];
  const h = s
    .getRange(1, 1, 1, lc)
    .getValues()[0]
    .map((x) => text_(x, 120));
  return s
    .getRange(2, 1, lr - 1, lc)
    .getValues()
    .map((row) => {
      const o = {};
      h.forEach((k, i) => {
        if (k) o[k] = row[i];
      });
      return o;
    });
}
function setRow_(s, row, obj) {
  const h = s
    .getRange(1, 1, 1, s.getLastColumn())
    .getValues()[0]
    .map((x) => text_(x, 120));
  Object.keys(obj).forEach((k) => {
    const i = h.indexOf(k);
    if (i > -1) s.getRange(row, i + 1).setValue(obj[k]);
  });
}
function audit_(type, id, action, detail) {
  try {
    add_(CONFIG.S.L, H.L, {
      Timestamp: new Date(),
      "Actor Type": type,
      "Actor ID": id,
      Action: action,
      Details: JSON.stringify(detail || {})
    });
  } catch (e) {}
}
function filter_(v, label) {
  const x = text_(v || CONFIG.ALL, 100).toUpperCase();
  return !x || x === CONFIG.ALL || x === "*" || x.indexOf("ALL ") === 0
    ? { all: true, value: CONFIG.ALL }
    : { all: false, value: code_(x, label) };
}
function code_(v, label) {
  const x = text_(v, 100).toUpperCase();
  if (!x) throw Error(label + " is required.");
  if (!/^[A-Z0-9][A-Z0-9 ._-]{0,99}$/.test(x))
    throw Error(
      label +
        " may only contain letters, numbers, spaces, dashes, underscores, and periods."
    );
  return x;
}
function optionalCode_(v) {
  return text_(v, 100).toUpperCase();
}
function safeCode_(v) {
  return text_(v, 100).toUpperCase();
}
function sid_(v) {
  return text_(v, 100).toUpperCase();
}
function status_(v) {
  const s = text_(v, 40).toLowerCase();
  return s === "inactive" || s === "closed" ? "Inactive" : "Active";
}
function truth_(v) {
  return (
    v === true ||
    String(v).toLowerCase() === "true" ||
    String(v).toLowerCase() === "yes" ||
    String(v) === "1" ||
    String(v).toLowerCase() === "active"
  );
}
function text_(v, n) {
  return String(v == null ? "" : v)
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, n || 500);
}
function num_(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
function dateInput_(v) {
  if (!v) return "";
  const d = v instanceof Date ? v : new Date(String(v));
  if (isNaN(d.getTime())) throw Error("Invalid date/time: " + v);
  return d;
}
function isFuture_(v, now) {
  const d = v instanceof Date ? v : v ? new Date(v) : null;
  return d && !isNaN(d.getTime()) && d.getTime() > now.getTime();
}
function isPast_(v, now) {
  const d = v instanceof Date ? v : v ? new Date(v) : null;
  return d && !isNaN(d.getTime()) && d.getTime() < now.getTime();
}
function crit_(v) {
  return (
    text_(v, 80)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "") || "criterion"
  );
}
function sum_(a) {
  return a.reduce((t, n) => t + (Number(n) || 0), 0);
}
function round2_(n) {
  return Math.round(n * 100) / 100;
}
function fmt_(v) {
  if (!v) return "";
  const d = v instanceof Date ? v : new Date(v);
  return isNaN(d.getTime())
    ? String(v)
    : Utilities.formatDate(
        d,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH:mm:ss"
      );
}
function uniq_(a) {
  return Array.from(new Set(a.filter(Boolean))).sort();
}
function uniqStudents_(rows) {
  const seen = {};
  return rows
    .reduce((out, r) => {
      if (!r.studentId || seen[r.studentId]) return out;
      seen[r.studentId] = 1;
      out.push({
        value: r.studentId,
        label: r.name ? r.studentId + " - " + r.name : r.studentId
      });
      return out;
    }, [])
    .sort((a, b) => a.value.localeCompare(b.value));
}
function cohortKey_(r) {
  return [r.activityKey, r.classId, r.groupKey].join("||");
}
function memberKey_(r, id) {
  return cohortKey_(r) + "||" + String(id || r.studentId || "").toUpperCase();
}
function uniqCohorts_(rows) {
  const seen = {};
  return rows
    .reduce((out, r) => {
      const k = cohortKey_(r);
      if (!seen[k]) {
        seen[k] = 1;
        out.push({
          key: k,
          activityKey: r.activityKey,
          classId: r.classId,
          groupKey: r.groupKey
        });
      }
      return out;
    }, [])
    .sort(
      (a, b) =>
        a.classId.localeCompare(b.classId) ||
        a.activityKey.localeCompare(b.activityKey) ||
        a.groupKey.localeCompare(b.groupKey)
    );
}
function sortRoster_(a, b) {
  return (
    a.classId.localeCompare(b.classId) ||
    a.activityKey.localeCompare(b.activityKey) ||
    a.groupKey.localeCompare(b.groupKey) ||
    a.name.localeCompare(b.name)
  );
}
function sortEval_(a, b) {
  return (
    a.classId.localeCompare(b.classId) ||
    a.activityKey.localeCompare(b.activityKey) ||
    a.groupKey.localeCompare(b.groupKey) ||
    a.targetName.localeCompare(b.targetName) ||
    a.evaluatorName.localeCompare(b.evaluatorName)
  );
}
function pin_() {
  return (
    PropertiesService.getScriptProperties().getProperty("FACULTY_PIN") ||
    CONFIG.PIN
  );
}
function good_(m, d) {
  return Object.assign({ success: true, message: m }, d || {});
}
function bad_(m) {
  return { success: false, message: m };
}
