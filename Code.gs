/**
 * Peer Evaluation Web App - Fresh Google Apps Script backend
 *
 * Files required:
 * - Code.gs
 * - Index.html
 * - Faculty.html
 *
 * Faculty portal:
 * <web-app-url>?view=faculty
 */

const CONFIG = Object.freeze({
  DEFAULT_FACULTY_PIN: 'Admin2026',
  MIN_SCORE: 1,
  MAX_SCORE: 10,
  ALL_VALUE: 'ALL',
  SHEETS: {
    REGISTRATIONS: 'Registrations',
    EVALUATIONS: 'Evaluations',
    AUDIT: 'Audit Log'
  }
});

const REGISTRATION_HEADERS = [
  'Timestamp',
  'Activity Key',
  'Class ID',
  'Group Key',
  'Student ID',
  'Name',
  'Role',
  'Contribution',
  'Status'
];

const EVALUATION_HEADERS = [
  'Timestamp',
  'Activity Key',
  'Class ID',
  'Group Key',
  'Evaluator Student ID',
  'Evaluator Name',
  'Target Student ID',
  'Target Name',
  'Cooperation (10)',
  'Contribution (10)',
  'Communication (10)',
  'Problem Solving (10)',
  'Quality of Work (10)',
  'Total Score',
  'Comments'
];

const AUDIT_HEADERS = [
  'Timestamp',
  'Actor Type',
  'Actor ID',
  'Action',
  'Details'
];

function doGet(e) {
  ensureSetup();
  const view = e && e.parameter && String(e.parameter.view || '').toLowerCase();
  const fileName = view === 'faculty' ? 'Faculty' : 'Index';
  const title = view === 'faculty' ? 'Faculty Report Portal' : 'Student Peer Evaluation';

  return HtmlService.createHtmlOutputFromFile(fileName)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ensureSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, CONFIG.SHEETS.REGISTRATIONS, REGISTRATION_HEADERS, '#e2e8f0');
  ensureSheet_(ss, CONFIG.SHEETS.EVALUATIONS, EVALUATION_HEADERS, '#e2e8f0');
  ensureSheet_(ss, CONFIG.SHEETS.AUDIT, AUDIT_HEADERS, '#fef3c7');

  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1 && sheet1.getLastRow() === 0) {
    ss.deleteSheet(sheet1);
  }
}

function changeFacultyPin(currentPin, newPin) {
  if (String(currentPin || '') !== getFacultyPin_()) {
    throw new Error('Current PIN is incorrect.');
  }
  const pin = cleanText_(newPin, 100);
  if (pin.length < 6) throw new Error('Faculty PIN must be at least 6 characters.');
  PropertiesService.getScriptProperties().setProperty('FACULTY_PIN', pin);
  return 'Faculty PIN updated.';
}

function saveRegistration(formObject) {
  ensureSetup();

  let registration;
  try {
    registration = {
      activityKey: normalizeCode_(formObject.activityKey, 'Activity'),
      classId: normalizeCode_(formObject.classId, 'Class ID'),
      groupKey: normalizeCode_(formObject.groupKey, 'Group Key'),
      studentId: normalizeStudentId_(formObject.studentId),
      name: cleanText_(formObject.name, 120),
      role: cleanText_(formObject.role, 120),
      contribution: cleanText_(formObject.contribution, 1500)
    };
  } catch (error) {
    return fail_(error.message);
  }

  const validationError = validateRegistration_(registration);
  if (validationError) return fail_(validationError);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const existing = getActiveRegistrations_();
    const duplicate = existing.find(row =>
      row.activityKey === registration.activityKey &&
      row.classId === registration.classId &&
      row.studentId === registration.studentId
    );

    if (duplicate) {
      return fail_('This Student ID is already registered for this Activity and Class. A student may register again only for a different activity or a different class.');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.REGISTRATIONS);
    appendRecord_(sheet, REGISTRATION_HEADERS, {
      'Timestamp': new Date(),
      'Activity Key': registration.activityKey,
      'Class ID': registration.classId,
      'Group Key': registration.groupKey,
      'Student ID': registration.studentId,
      'Name': registration.name,
      'Role': registration.role,
      'Contribution': registration.contribution,
      'Status': 'Active'
    });

    logAudit_('Student', registration.studentId, 'REGISTER', registration);
    return ok_('Registration successful. You may now proceed to peer evaluation.', registration);
  } catch (error) {
    return fail_('System Error: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

function verifyAndLoadGroup(studentId, activityKey, classId) {
  ensureSetup();

  let cleanActivityKey;
  let cleanClassId;
  try {
    cleanActivityKey = normalizeCode_(activityKey, 'Activity');
    cleanClassId = normalizeCode_(classId, 'Class ID');
  } catch (error) {
    return fail_(error.message);
  }

  const evaluatorId = normalizeStudentId_(studentId);
  if (!evaluatorId) return fail_('Please enter your Student ID.');

  try {
    const registrations = getActiveRegistrations_();
    const evaluator = registrations.find(row =>
      row.activityKey === cleanActivityKey &&
      row.classId === cleanClassId &&
      row.studentId === evaluatorId
    );

    if (!evaluator) {
      return fail_('Student ID not found for this Activity and Class. Please register first, or check the Activity and Class ID.');
    }

    const groupMembers = registrations
      .filter(row =>
        row.activityKey === evaluator.activityKey &&
        row.classId === evaluator.classId &&
        row.groupKey === evaluator.groupKey
      )
      .sort((a, b) => a.name.localeCompare(b.name));

    const evaluatedTargetIds = new Set(
      getEvaluations_()
        .filter(ev =>
          ev.activityKey === evaluator.activityKey &&
          ev.classId === evaluator.classId &&
          ev.groupKey === evaluator.groupKey &&
          ev.evaluatorStudentId === evaluator.studentId
        )
        .map(ev => ev.targetStudentId)
    );

    const targets = groupMembers
      .filter(member => member.studentId !== evaluator.studentId)
      .map(member => ({
        studentId: member.studentId,
        name: member.name,
        role: member.role,
        contribution: member.contribution,
        alreadyEvaluated: evaluatedTargetIds.has(member.studentId)
      }));

    const remainingTargets = targets.filter(member => !member.alreadyEvaluated);

    return ok_('Group loaded.', {
      evaluator,
      groupSize: groupMembers.length,
      targets,
      remainingTargets,
      completedCount: targets.length - remainingTargets.length,
      totalTargets: targets.length
    });
  } catch (error) {
    return fail_('System Error: ' + error.message);
  }
}

function saveEvaluation(evalData) {
  ensureSetup();

  let activityKey;
  let classId;
  try {
    activityKey = normalizeCode_(evalData.activityKey, 'Activity');
    classId = normalizeCode_(evalData.classId, 'Class ID');
  } catch (error) {
    return fail_(error.message);
  }

  const evaluatorStudentId = normalizeStudentId_(evalData.evaluatorStudentId);
  const targetStudentId = normalizeStudentId_(evalData.targetStudentId);
  const comments = cleanText_(evalData.comments || '', 1200);
  const scores = parseScores_(evalData.scores || evalData);

  if (!evaluatorStudentId) return fail_('Missing evaluator Student ID. Please verify again.');
  if (!targetStudentId) return fail_('Please select a team member to evaluate.');
  if (evaluatorStudentId === targetStudentId) return fail_('You cannot evaluate yourself.');
  if (!scores.success) return fail_(scores.message);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const registrations = getActiveRegistrations_();
    const evaluator = registrations.find(row =>
      row.activityKey === activityKey &&
      row.classId === classId &&
      row.studentId === evaluatorStudentId
    );

    if (!evaluator) return fail_('Evaluator record not found. Please register first.');

    const target = registrations.find(row =>
      row.activityKey === evaluator.activityKey &&
      row.classId === evaluator.classId &&
      row.groupKey === evaluator.groupKey &&
      row.studentId === targetStudentId
    );

    if (!target) return fail_('Target member is not registered in your Activity, Class, and Group.');

    const duplicate = getEvaluations_().some(ev =>
      ev.activityKey === evaluator.activityKey &&
      ev.classId === evaluator.classId &&
      ev.groupKey === evaluator.groupKey &&
      ev.evaluatorStudentId === evaluator.studentId &&
      ev.targetStudentId === target.studentId
    );

    if (duplicate) return fail_('You already evaluated this member for this activity. Duplicate evaluations are blocked.');

    const totalScore = scores.values.cooperation +
      scores.values.contribution +
      scores.values.communication +
      scores.values.problemSolving +
      scores.values.qualityOfWork;

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EVALUATIONS);
    appendRecord_(sheet, EVALUATION_HEADERS, {
      'Timestamp': new Date(),
      'Activity Key': evaluator.activityKey,
      'Class ID': evaluator.classId,
      'Group Key': evaluator.groupKey,
      'Evaluator Student ID': evaluator.studentId,
      'Evaluator Name': evaluator.name,
      'Target Student ID': target.studentId,
      'Target Name': target.name,
      'Cooperation (10)': scores.values.cooperation,
      'Contribution (10)': scores.values.contribution,
      'Communication (10)': scores.values.communication,
      'Problem Solving (10)': scores.values.problemSolving,
      'Quality of Work (10)': scores.values.qualityOfWork,
      'Total Score': totalScore,
      'Comments': comments
    });

    logAudit_('Student', evaluator.studentId, 'SUBMIT_EVALUATION', {
      activityKey: evaluator.activityKey,
      classId: evaluator.classId,
      groupKey: evaluator.groupKey,
      targetStudentId: target.studentId,
      totalScore
    });

    const snapshot = verifyAndLoadGroup(evaluator.studentId, evaluator.activityKey, evaluator.classId);
    return ok_('Evaluation for ' + target.name + ' saved.', {
      totalScore,
      targetName: target.name,
      remainingTargets: snapshot.success ? snapshot.remainingTargets : [],
      completedCount: snapshot.success ? snapshot.completedCount : null,
      totalTargets: snapshot.success ? snapshot.totalTargets : null
    });
  } catch (error) {
    return fail_('System Error: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

function unlockFaculty(pin) {
  ensureSetup();
  if (String(pin || '') !== getFacultyPin_()) return fail_('Access denied. Incorrect Faculty PIN.');
  return ok_('Faculty portal unlocked.', { options: getFacultyOptions_() });
}

function generateReportData(filters, providedPin) {
  ensureSetup();
  if (String(providedPin || '') !== getFacultyPin_()) return fail_('Access denied. Incorrect Faculty PIN.');

  let classId;
  let activityFilter;
  let groupFilter;
  try {
    classId = normalizeCode_(filters.classId, 'Class ID');
    activityFilter = normalizeOptionalFilter_(filters.activityKey, 'Activity');
    groupFilter = normalizeOptionalFilter_(filters.groupKey, 'Group Key');
  } catch (error) {
    return fail_(error.message);
  }

  try {
    const registrations = getActiveRegistrations_().filter(row =>
      row.classId === classId &&
      (activityFilter.includeAll || row.activityKey === activityFilter.value) &&
      (groupFilter.includeAll || row.groupKey === groupFilter.value)
    );

    if (registrations.length === 0) return fail_('No registered students found for the selected filters.');

    const memberKeys = new Set(registrations.map(row => cohortMemberKey_(row)));
    const evaluations = getEvaluations_().filter(ev =>
      ev.classId === classId &&
      (activityFilter.includeAll || ev.activityKey === activityFilter.value) &&
      (groupFilter.includeAll || ev.groupKey === groupFilter.value) &&
      memberKeys.has(cohortMemberKey_(ev, ev.evaluatorStudentId)) &&
      memberKeys.has(cohortMemberKey_(ev, ev.targetStudentId))
    );

    const cohorts = buildCohorts_(registrations, evaluations);
    const summary = buildStudentSummary_(registrations, evaluations);
    const completion = buildCompletion_(registrations, evaluations);
    const details = evaluations
      .sort((a, b) =>
        a.activityKey.localeCompare(b.activityKey) ||
        a.groupKey.localeCompare(b.groupKey) ||
        a.targetName.localeCompare(b.targetName) ||
        a.evaluatorName.localeCompare(b.evaluatorName)
      )
      .map(ev => ({
        timestamp: formatDateForClient_(ev.timestamp),
        activityKey: ev.activityKey,
        classId: ev.classId,
        groupKey: ev.groupKey,
        evaluatorStudentId: ev.evaluatorStudentId,
        evaluatorName: ev.evaluatorName,
        targetStudentId: ev.targetStudentId,
        targetName: ev.targetName,
        cooperation: ev.cooperation,
        contribution: ev.contribution,
        communication: ev.communication,
        problemSolving: ev.problemSolving,
        qualityOfWork: ev.qualityOfWork,
        totalScore: ev.totalScore,
        comments: ev.comments
      }));

    const totals = evaluations.map(ev => ev.totalScore).filter(Number.isFinite);

    logAudit_('Faculty', 'FACULTY', 'GENERATE_REPORT', {
      classId,
      activityKey: activityFilter.value,
      groupKey: groupFilter.value,
      submitted: evaluations.length
    });

    return ok_('Report generated.', {
      generatedAt: formatDateForClient_(new Date()),
      filters: {
        classId,
        activityKey: activityFilter.value,
        groupKey: groupFilter.value,
        activityIsAll: activityFilter.includeAll,
        groupIsAll: groupFilter.includeAll
      },
      totals: {
        registeredStudents: registrations.length,
        cohorts: cohorts.length,
        submittedEvaluations: evaluations.length,
        expectedEvaluations: cohorts.reduce((total, cohort) => total + cohort.expectedEvaluations, 0),
        averageScore: totals.length ? round1_(sum_(totals) / totals.length) : null
      },
      cohorts,
      summary,
      completion,
      details,
      roster: registrations.sort(sortRoster_)
    });
  } catch (error) {
    return fail_('System Error: ' + error.message);
  }
}

function getFacultyOptions_() {
  const registrations = getActiveRegistrations_();
  const activities = uniqueSorted_(registrations.map(row => row.activityKey));
  const classes = uniqueSorted_(registrations.map(row => row.classId));
  const groups = uniqueSorted_(registrations.map(row => row.groupKey));
  const cohorts = registrations
    .reduce((list, row) => {
      const key = cohortKey_(row);
      if (!list.some(item => item.key === key)) {
        list.push({
          key,
          activityKey: row.activityKey,
          classId: row.classId,
          groupKey: row.groupKey
        });
      }
      return list;
    }, [])
    .sort((a, b) => a.classId.localeCompare(b.classId) || a.activityKey.localeCompare(b.activityKey) || a.groupKey.localeCompare(b.groupKey));

  return { activities, classes, groups, cohorts };
}

function buildCohorts_(registrations, evaluations) {
  const byCohort = groupBy_(registrations, cohortKey_);
  return Object.keys(byCohort).sort().map(key => {
    const members = byCohort[key];
    const first = members[0];
    const cohortEvals = evaluations.filter(ev => cohortKey_(ev) === key);
    const expected = members.length * Math.max(members.length - 1, 0);
    const totals = cohortEvals.map(ev => ev.totalScore).filter(Number.isFinite);
    return {
      activityKey: first.activityKey,
      classId: first.classId,
      groupKey: first.groupKey,
      memberCount: members.length,
      submittedEvaluations: cohortEvals.length,
      expectedEvaluations: expected,
      completionPercent: expected ? Math.round((cohortEvals.length / expected) * 100) : 100,
      averageScore: totals.length ? round1_(sum_(totals) / totals.length) : null
    };
  });
}

function buildStudentSummary_(registrations, evaluations) {
  return registrations.sort(sortRoster_).map(student => {
    const received = evaluations.filter(ev =>
      cohortKey_(ev) === cohortKey_(student) && ev.targetStudentId === student.studentId
    );
    const cohortMembers = registrations.filter(row => cohortKey_(row) === cohortKey_(student));
    const expected = Math.max(cohortMembers.length - 1, 0);
    const receivedEvaluatorIds = new Set(received.map(ev => ev.evaluatorStudentId));
    const missingEvaluators = cohortMembers
      .filter(member => member.studentId !== student.studentId && !receivedEvaluatorIds.has(member.studentId))
      .map(member => member.name);
    const totals = received.map(ev => ev.totalScore).filter(Number.isFinite);

    return {
      activityKey: student.activityKey,
      classId: student.classId,
      groupKey: student.groupKey,
      studentId: student.studentId,
      name: student.name,
      role: student.role,
      contribution: student.contribution,
      receivedCount: received.length,
      expectedCount: expected,
      completionPercent: expected ? Math.round((received.length / expected) * 100) : 100,
      average: totals.length ? round1_(sum_(totals) / totals.length) : null,
      criteria: averageCriteria_(received),
      missingEvaluators
    };
  });
}

function buildCompletion_(registrations, evaluations) {
  return registrations.sort(sortRoster_).map(evaluator => {
    const submitted = evaluations.filter(ev =>
      cohortKey_(ev) === cohortKey_(evaluator) && ev.evaluatorStudentId === evaluator.studentId
    );
    const cohortMembers = registrations.filter(row => cohortKey_(row) === cohortKey_(evaluator));
    const expected = Math.max(cohortMembers.length - 1, 0);
    const submittedTargetIds = new Set(submitted.map(ev => ev.targetStudentId));
    const missingTargets = cohortMembers
      .filter(member => member.studentId !== evaluator.studentId && !submittedTargetIds.has(member.studentId))
      .map(member => member.name);

    return {
      activityKey: evaluator.activityKey,
      classId: evaluator.classId,
      groupKey: evaluator.groupKey,
      studentId: evaluator.studentId,
      name: evaluator.name,
      submittedCount: submitted.length,
      expectedCount: expected,
      completionPercent: expected ? Math.round((submitted.length / expected) * 100) : 100,
      missingTargets
    };
  });
}

function getActiveRegistrations_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.REGISTRATIONS);
  return getRowsAsObjects_(sheet)
    .filter(row => cleanText_(row.Status || 'Active', 40).toLowerCase() !== 'inactive')
    .map(row => ({
      timestamp: row.Timestamp,
      activityKey: cleanText_(row['Activity Key'], 100).toUpperCase(),
      classId: cleanText_(row['Class ID'], 100).toUpperCase(),
      groupKey: cleanText_(row['Group Key'], 100).toUpperCase(),
      studentId: normalizeStudentId_(row['Student ID']),
      name: cleanText_(row.Name, 120),
      role: cleanText_(row.Role, 120),
      contribution: cleanText_(row.Contribution, 1500),
      status: cleanText_(row.Status || 'Active', 40)
    }))
    .filter(row => row.activityKey && row.classId && row.groupKey && row.studentId && row.name);
}

function getEvaluations_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EVALUATIONS);
  return getRowsAsObjects_(sheet)
    .map(row => ({
      timestamp: row.Timestamp,
      activityKey: cleanText_(row['Activity Key'], 100).toUpperCase(),
      classId: cleanText_(row['Class ID'], 100).toUpperCase(),
      groupKey: cleanText_(row['Group Key'], 100).toUpperCase(),
      evaluatorStudentId: normalizeStudentId_(row['Evaluator Student ID']),
      evaluatorName: cleanText_(row['Evaluator Name'], 120),
      targetStudentId: normalizeStudentId_(row['Target Student ID']),
      targetName: cleanText_(row['Target Name'], 120),
      cooperation: toNumber_(row['Cooperation (10)']),
      contribution: toNumber_(row['Contribution (10)']),
      communication: toNumber_(row['Communication (10)']),
      problemSolving: toNumber_(row['Problem Solving (10)']),
      qualityOfWork: toNumber_(row['Quality of Work (10)']),
      totalScore: toNumber_(row['Total Score']),
      comments: cleanText_(row.Comments, 1200)
    }))
    .filter(row => row.activityKey && row.classId && row.groupKey && row.evaluatorStudentId && row.targetStudentId);
}

function parseScores_(rawScores) {
  const scoreMap = {
    cooperation: rawScores.cooperation || rawScores.q1,
    contribution: rawScores.contribution || rawScores.q2,
    communication: rawScores.communication || rawScores.q3,
    problemSolving: rawScores.problemSolving || rawScores.q4,
    qualityOfWork: rawScores.qualityOfWork || rawScores.q5
  };

  const labels = {
    cooperation: 'Cooperation',
    contribution: 'Contribution',
    communication: 'Communication',
    problemSolving: 'Problem Solving',
    qualityOfWork: 'Quality of Work'
  };

  const values = {};
  Object.keys(scoreMap).forEach(key => {
    const value = Number(scoreMap[key]);
    if (!Number.isInteger(value) || value < CONFIG.MIN_SCORE || value > CONFIG.MAX_SCORE) {
      throw new Error(labels[key] + ' must be a whole number from 1 to 10.');
    }
    values[key] = value;
  });

  return ok_('Scores valid.', { values });
}

function validateRegistration_(registration) {
  if (!registration.activityKey) return 'Activity is required.';
  if (!registration.classId) return 'Class ID is required.';
  if (!registration.groupKey) return 'Group Key is required.';
  if (!registration.studentId) return 'Student ID is required.';
  if (!registration.name) return 'Full Name is required.';
  if (!registration.role) return 'Project Role is required.';
  if (!registration.contribution) return 'Contribution Details are required.';
  return '';
}

function ensureSheet_(ss, sheetName, headers, headerColor) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground(headerColor);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function appendRecord_(sheet, headers, record) {
  sheet.appendRow(headers.map(header => record[header] === undefined ? '' : record[header]));
}

function getRowsAsObjects_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(value => cleanText_(value, 120));
  return sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues().map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      if (header) obj[header] = row[index];
    });
    return obj;
  });
}

function logAudit_(actorType, actorId, action, details) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.AUDIT);
    appendRecord_(sheet, AUDIT_HEADERS, {
      'Timestamp': new Date(),
      'Actor Type': actorType,
      'Actor ID': actorId,
      'Action': action,
      'Details': JSON.stringify(details || {})
    });
  } catch (error) {
    // Audit failure should not block the app.
  }
}

function averageCriteria_(evaluations) {
  const fields = ['cooperation', 'contribution', 'communication', 'problemSolving', 'qualityOfWork'];
  const output = {};
  fields.forEach(field => {
    const values = evaluations.map(ev => ev[field]).filter(Number.isFinite);
    output[field] = values.length ? round1_(sum_(values) / values.length) : null;
  });
  return output;
}

function normalizeOptionalFilter_(value, label) {
  const raw = cleanText_(value || CONFIG.ALL_VALUE, 100).toUpperCase();
  if (!raw || raw === CONFIG.ALL_VALUE || raw === '*' || raw.indexOf('ALL ') === 0) {
    return { includeAll: true, value: CONFIG.ALL_VALUE };
  }
  return { includeAll: false, value: normalizeCode_(raw, label) };
}

function normalizeCode_(value, label) {
  const normalized = cleanText_(value, 100).toUpperCase();
  if (!normalized) throw new Error(label + ' is required.');
  if (!/^[A-Z0-9][A-Z0-9 ._-]{0,99}$/.test(normalized)) {
    throw new Error(label + ' may only contain letters, numbers, spaces, dashes, underscores, and periods.');
  }
  return normalized;
}

function normalizeStudentId_(value) {
  return cleanText_(value, 100).toUpperCase();
}

function cleanText_(value, maxLength) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/[\u0000-\u001F\u007F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, maxLength || 500);
}

function toNumber_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function sum_(numbers) {
  return numbers.reduce((total, number) => total + number, 0);
}

function round1_(value) {
  return Math.round(value * 10) / 10;
}

function uniqueSorted_(values) {
  return Array.from(new Set(values.filter(Boolean))).sort((a, b) => String(a).localeCompare(String(b)));
}

function groupBy_(rows, keyFn) {
  return rows.reduce((acc, row) => {
    const key = keyFn(row);
    if (!acc[key]) acc[key] = [];
    acc[key].push(row);
    return acc;
  }, {});
}

function cohortKey_(row) {
  return [row.activityKey, row.classId, row.groupKey].map(value => String(value || '').toUpperCase()).join('||');
}

function cohortMemberKey_(row, overrideStudentId) {
  return cohortKey_(row) + '||' + String(overrideStudentId || row.studentId || '').toUpperCase();
}

function sortRoster_(a, b) {
  return a.classId.localeCompare(b.classId) ||
    a.activityKey.localeCompare(b.activityKey) ||
    a.groupKey.localeCompare(b.groupKey) ||
    a.name.localeCompare(b.name);
}

function formatDateForClient_(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function getFacultyPin_() {
  return PropertiesService.getScriptProperties().getProperty('FACULTY_PIN') || CONFIG.DEFAULT_FACULTY_PIN;
}

function ok_(message, data) {
  return Object.assign({ success: true, message }, data || {});
}

function fail_(message) {
  return { success: false, message };
}
