// TIEcon Match — Admin Google Sheet Sync v6
// Startup (was Expert), Challenge (was Problem), Match-a-Thon
//
// v6 changes vs the deployed v5: pullExperts() and pullIndustries() now also
// pull the new profile fields added by the July 2026 review-doc update
// (LinkedIn/Website, WhatsApp, Stage, Pitch Deck Link, Special Category on
// STARTUPS; LinkedIn/Website, WhatsApp, Timeline, Budget Range, Secondary
// Industry Type on INDUSTRIES). "Special Category" combines the three
// checkbox flags (Student-Led / Women-Led / Service Provider) into one
// comma-joined column, same pattern as the existing Industries/Problem
// Domains columns, instead of a separate Yes/No column per flag.
// Nothing else changed — pushExpertApprovals/pushIndustryApprovals/
// pushProblemApprovals, matching, notify, and all other functions are
// identical to v5.
//
// Before deploying: add the new column headers (exact text, case-sensitive)
// to the live Google Sheet tabs:
//   STARTUPS:   LinkedIn/Website, WhatsApp, Stage, Pitch Deck Link,
//               Special Category
//   INDUSTRIES: LinkedIn/Website, WhatsApp, Timeline, Budget Range,
//               Secondary Industry Type

const SUPABASE_URL = 'https://wrzpgultvahxbrgooibn.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6IndyenBndWx0dmFoeGJyZ29vaWJuIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzc1Mzc2MTcsImV4cCI6MjA5MzExMzYxN30.sZM4GQiwom34d7TjeK_Jc_n8KYoC8VhI48jfq71dwb8';
const EDGE_FUNCTION_URL = 'https://wrzpgultvahxbrgooibn.supabase.co/functions/v1/generate-embedding';
const SHEET_ID = '1abSLhdEOEfdvMZ7Yi--kVGKckL1dRcZlQpIXQOTTj6w';
const EXPERTS_TAB = 'STARTUPS';
const INDUSTRIES_TAB = 'INDUSTRIES';
const PROBLEMS_TAB = 'CHALLENGES';
const MATCHES_TAB = 'MATCHES';
const ADMIN_EMAIL = 'unigoods2026@gmail.com';

// =============================================
// HEADER HELPER
// =============================================

function getHeaders(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[String(h).trim()] = i; });
  return map;
}

function getCol(map, name) {
  const idx = map[name];
  if (idx === undefined) throw new Error('Column not found: ' + name);
  return idx;
}

// =============================================
// MAIN FUNCTIONS
// =============================================

function pullAllFromSupabase() {
  pullExperts();
  pullIndustries();
  pullProblems();
}

function pushApprovalsToSupabase() {
  pushExpertApprovals();
  pushIndustryApprovals();
  pushProblemApprovals();
}

// =============================================
// PULL STARTUPS → SHEET
// =============================================

function pullExperts() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EXPERTS_TAB);
  const headers = getHeaders(sheet);
  const idCol = getCol(headers, 'Supabase ID');
  const existingIds = getExistingIds(sheet, idCol + 1);

  const experts = supabaseGet('tie_experts?select=*&status=eq.pending&order=created_at.desc');
  if (!experts || !experts.length) { Logger.log('No new startups'); return; }

  const allIndustries = supabaseGet('tie_industry_types?select=id,name');
  const allDomains = supabaseGet('tie_problem_domains?select=id,name');

  const industryMap = {};
  const domainMap = {};
  if (allIndustries) allIndustries.forEach(i => { industryMap[i.id] = i.name; });
  if (allDomains) allDomains.forEach(d => { domainMap[d.id] = d.name; });

  experts.forEach(expert => {
    if (existingIds.includes(expert.id)) return;

    const tags = supabaseGet(`tie_expert_tags?expert_id=eq.${expert.id}&select=industry_type_id,problem_domain_id`);
    let industryNames = '';
    let domainNames = '';

    if (tags && tags.length) {
      const indIds = [...new Set(tags.map(t => t.industry_type_id).filter(Boolean))];
      const domIds = [...new Set(tags.map(t => t.problem_domain_id).filter(Boolean))];
      industryNames = indIds.map(id => industryMap[id] || '').filter(Boolean).join(', ');
      domainNames = domIds.map(id => domainMap[id] || '').filter(Boolean).join(', ');
    }

    const specialCategories = [];
    if (expert.is_student_led) specialCategories.push('Student-Led');
    if (expert.is_women_led) specialCategories.push('Women-Led');
    if (expert.is_service_provider) specialCategories.push('Service Provider');

    const rowData = buildRow(headers, {
      'Registered On': formatDate(expert.created_at),
      'Name': expert.name || '',
      'Organisation': expert.organisation || '',
      'Designation': expert.designation || '',
      'Email': expert.email || '',
      'Phone': expert.phone || '',
      'City': expert.city || '',
      'State': expert.state || '',
      'Industries': industryNames,
      'Problem Domains': domainNames,
      'Description': expert.description || '',
      'Suggested Domain': expert.suggested_domain || '',
      'Status': 'PENDING',
      'Admin Notes': '',
      'Supabase ID': expert.id,
      'LinkedIn/Website': expert.linkedin_url || '',
      'WhatsApp': expert.whatsapp || '',
      'Stage': expert.stage || '',
      'Pitch Deck Link': expert.pitch_deck_url || '',
      'Special Category': specialCategories.join(', ')
    });

    sheet.appendRow(rowData);
  });

  setupDropdownAndColor(sheet, headers, ['PENDING','APPROVED','ON_HOLD','REJECT']);
  SpreadsheetApp.flush();
  Logger.log('Startups pulled successfully');
}

// =============================================
// PULL INDUSTRIES → SHEET
// =============================================

function pullIndustries() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(INDUSTRIES_TAB);
  const headers = getHeaders(sheet);
  const idCol = getCol(headers, 'Supabase ID');
  const existingIds = getExistingIds(sheet, idCol + 1);

  const allIndustries = supabaseGet('tie_industry_types?select=id,name');
  const industryMap = {};
  if (allIndustries) allIndustries.forEach(i => { industryMap[i.id] = i.name; });

  const industries = supabaseGet('tie_industries?select=*&status=eq.pending&order=created_at.desc');
  if (!industries || !industries.length) { Logger.log('No new industries'); return; }

  industries.forEach(ind => {
    if (existingIds.includes(ind.id)) return;
    const industryTypeName = industryMap[ind.industry_type_id] || '';
    const secondaryIndustryTypeName = ind.secondary_industry_type_id ? (industryMap[ind.secondary_industry_type_id] || '') : '';

    const rowData = buildRow(headers, {
      'Registered On': formatDate(ind.created_at),
      'Name': ind.name || '',
      'Organisation': ind.organisation || '',
      'Designation': ind.designation || '',
      'Email': ind.email || '',
      'Phone': ind.phone || '',
      'City': ind.city || '',
      'State': ind.state || '',
      'Industry Type': industryTypeName,
      'Status': 'PENDING',
      'Admin Notes': '',
      'Supabase ID': ind.id,
      'LinkedIn/Website': ind.linkedin_url || '',
      'WhatsApp': ind.whatsapp || '',
      'Timeline': ind.timeline || '',
      'Budget Range': ind.budget_range || '',
      'Secondary Industry Type': secondaryIndustryTypeName
    });

    sheet.appendRow(rowData);
  });

  setupDropdownAndColor(sheet, headers, ['PENDING','APPROVED','ON_HOLD','REJECT']);
  SpreadsheetApp.flush();
  Logger.log('Industries pulled successfully');
}

// =============================================
// PULL CHALLENGES → SHEET
// =============================================

function pullProblems() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(PROBLEMS_TAB);

  if (!sheet) {
    sheet = ss.insertSheet(PROBLEMS_TAB);
    sheet.appendRow([
      'Submitted On','Industry Name','Organisation','Phone','Email',
      'Challenge Description','Expected Outcome','Status','Admin Notes','Supabase ID'
    ]);
    sheet.getRange(1, 1, 1, 10).setBackground('#0a2540').setFontColor('white').setFontWeight('bold');
  }

  const headers = getHeaders(sheet);
  const idCol = getCol(headers, 'Supabase ID');
  const existingIds = getExistingIds(sheet, idCol + 1);

  const problems = supabaseGet('tie_problems?select=*&status=eq.pending&order=created_at.desc');
  if (!problems || !problems.length) { Logger.log('No new challenges'); return; }

  const industries = supabaseGet('tie_industries?select=id,name,organisation,phone,email');
  const indMap = {};
  if (industries) industries.forEach(i => { indMap[i.id] = i; });

  problems.forEach(prob => {
    if (existingIds.includes(prob.id)) return;
    const ind = indMap[prob.industry_id] || {};

    const rowData = buildRow(headers, {
      'Submitted On': formatDate(prob.created_at),
      'Industry Name': ind.name || '',
      'Organisation': ind.organisation || '',
      'Phone': ind.phone || '',
      'Email': ind.email || '',
      'Challenge Description': prob.description || '',
      'Expected Outcome': prob.expected_outcome || '',
      'Status': 'PENDING',
      'Admin Notes': '',
      'Supabase ID': prob.id
    });

    sheet.appendRow(rowData);
  });

  setupDropdownAndColor(sheet, headers, ['PENDING','APPROVED','REJECT']);
  SpreadsheetApp.flush();
  Logger.log('Challenges pulled successfully');
}

// =============================================
// PUSH STARTUP APPROVALS → SUPABASE
// =============================================

function pushExpertApprovals() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EXPERTS_TAB);
  const headers = getHeaders(sheet);
  const statusCol = getCol(headers, 'Status');
  const idCol = getCol(headers, 'Supabase ID');
  const emailCol = getCol(headers, 'Email');
  const nameCol = getCol(headers, 'Name');
  const notesCol = getCol(headers, 'Admin Notes');

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[statusCol]).trim().toUpperCase();
    const supabaseId = String(row[idCol]).trim();
    const email = String(row[emailCol]).trim();
    const name = String(row[nameCol]).trim();
    const notes = String(row[notesCol]).trim();

    if (!supabaseId || status === 'PENDING' || status === 'APPROVED' ||
        status === 'REJECTED' || status === 'ON HOLD') continue;

    if (status === 'APPROVED') {
      supabasePatch('tie_experts', supabaseId, { status: 'active' });
      sendEmail(email, name, 'expert', 'approved', notes);
      setStatusCell(sheet, i + 1, statusCol + 1, 'APPROVED', '#00C853', '#FFFFFF');
      try {
        generateEmbedding({ expert_id: supabaseId });
        Logger.log('Embedding saved for startup: ' + name);
      } catch(e) { Logger.log('Embedding error: ' + e); }

    } else if (status === 'REJECT') {
      supabasePatch('tie_experts', supabaseId, { status: 'rejected' });
      setStatusCell(sheet, i + 1, statusCol + 1, 'REJECTED', '#D32F2F', '#FFFFFF');

    } else if (status === 'ON_HOLD') {
      supabasePatch('tie_experts', supabaseId, { status: 'on_hold' });
      sendEmail(email, name, 'expert', 'on_hold', notes);
      setStatusCell(sheet, i + 1, statusCol + 1, 'ON HOLD', '#F9A825', '#000000');
    }
  }

  SpreadsheetApp.flush();
  Logger.log('Startup approvals pushed');
}

// =============================================
// PUSH INDUSTRY APPROVALS → SUPABASE
// =============================================

function pushIndustryApprovals() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(INDUSTRIES_TAB);
  const headers = getHeaders(sheet);
  const statusCol = getCol(headers, 'Status');
  const idCol = getCol(headers, 'Supabase ID');
  const emailCol = getCol(headers, 'Email');
  const nameCol = getCol(headers, 'Name');
  const notesCol = getCol(headers, 'Admin Notes');

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[statusCol]).trim().toUpperCase();
    const supabaseId = String(row[idCol]).trim();
    const email = String(row[emailCol]).trim();
    const name = String(row[nameCol]).trim();
    const notes = String(row[notesCol]).trim();

    if (!supabaseId || status === 'PENDING' || status === 'APPROVED' ||
        status === 'REJECTED' || status === 'ON HOLD') continue;

    if (status === 'APPROVED') {
      supabasePatch('tie_industries', supabaseId, { status: 'active' });
      sendEmail(email, name, 'industry', 'approved', notes);
      setStatusCell(sheet, i + 1, statusCol + 1, 'APPROVED', '#00C853', '#FFFFFF');

    } else if (status === 'REJECT') {
      supabasePatch('tie_industries', supabaseId, { status: 'rejected' });
      setStatusCell(sheet, i + 1, statusCol + 1, 'REJECTED', '#D32F2F', '#FFFFFF');

    } else if (status === 'ON_HOLD') {
      supabasePatch('tie_industries', supabaseId, { status: 'on_hold' });
      sendEmail(email, name, 'industry', 'on_hold', notes);
      setStatusCell(sheet, i + 1, statusCol + 1, 'ON HOLD', '#F9A825', '#000000');
    }
  }

  SpreadsheetApp.flush();
  Logger.log('Industry approvals pushed');
}

// =============================================
// PUSH CHALLENGE APPROVALS → SUPABASE + MATCHING
// =============================================

function pushProblemApprovals() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(PROBLEMS_TAB);
  if (!sheet) return;

  const headers = getHeaders(sheet);
  const statusCol = getCol(headers, 'Status');
  const idCol = getCol(headers, 'Supabase ID');

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[statusCol]).trim().toUpperCase();
    const supabaseId = String(row[idCol]).trim();

    if (!supabaseId || status === 'PENDING' || status === 'APPROVED' || status === 'REJECTED') continue;

    if (status === 'APPROVED') {
      supabasePatch('tie_problems', supabaseId, { status: 'active' });
      setStatusCell(sheet, i + 1, statusCol + 1, 'APPROVED', '#00C853', '#FFFFFF');

      try {
        generateEmbedding({ problem_id: supabaseId });
        Logger.log('Challenge embedding saved: ' + supabaseId);
      } catch(e) { Logger.log('Challenge embedding error: ' + e); }

      try {
        const matches = runMatching(supabaseId);
        Logger.log('Matches found: ' + matches.length);
        if (matches && matches.length > 0) {
          saveMatchesToSheet(ss, supabaseId, matches, row, headers);
        }
      } catch(e) { Logger.log('Matching error: ' + e); }

    } else if (status === 'REJECT') {
      supabasePatch('tie_problems', supabaseId, { status: 'rejected' });
      setStatusCell(sheet, i + 1, statusCol + 1, 'REJECTED', '#D32F2F', '#FFFFFF');
    }
  }

  SpreadsheetApp.flush();
  Logger.log('Challenge approvals pushed');
}

// =============================================
// AI MATCHING
// =============================================

function runMatching(problemId) {
  const problems = supabaseGet(`tie_problems?id=eq.${problemId}&select=embedding`);
  if (!problems || !problems.length || !problems[0].embedding) {
    throw new Error('Challenge embedding not found');
  }

  let emb = problems[0].embedding;
  if (typeof emb === 'string') { emb = JSON.parse(emb); }

  const response = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/rpc/match_experts`, {
    method: 'POST',
    headers: {
      'apikey': SUPABASE_KEY,
      'Authorization': 'Bearer ' + SUPABASE_KEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      query_embedding: emb,
      match_threshold: 0.2,
      match_count: 3
    }),
    muteHttpExceptions: true
  });

  const matches = JSON.parse(response.getContentText());
  Logger.log('Match results: ' + JSON.stringify(matches));
  return matches || [];
}

// =============================================
// SAVE MATCHES → SUPABASE + SHEET
// =============================================

function saveMatchesToSheet(ss, problemId, matches, problemRow, problemHeaders) {
  let matchSheet = ss.getSheetByName(MATCHES_TAB);

  if (!matchSheet) {
    matchSheet = ss.insertSheet(MATCHES_TAB);
    matchSheet.appendRow([
      'Match Date','Score','Challenge Description',
      'Industry Name','Industry Org','Industry Phone','Industry Email',
      '','Startup Name','Startup Org','Startup Phone','Startup Email',
      'Status','Notify','Supabase Match ID'
    ]);
    matchSheet.getRange(1, 1, 1, 7).setBackground('#f4a11d').setFontColor('#0a2540').setFontWeight('bold');
    matchSheet.getRange(1, 8, 1, 1).setBackground('#e5e7eb');
    matchSheet.setColumnWidth(8, 18);
    matchSheet.getRange(1, 9, 1, 7).setBackground('#0a2540').setFontColor('white').setFontWeight('bold');
  }

  const descCol = problemHeaders['Challenge Description'] !== undefined
    ? problemHeaders['Challenge Description']
    : problemHeaders['Problem Description'];
  const indNameCol = problemHeaders['Industry Name'];
  const indOrgCol = problemHeaders['Organisation'];
  const indPhoneCol = problemHeaders['Phone'];
  const indEmailCol = problemHeaders['Email'];

  const challengeDesc = problemRow[descCol] || '';
  const indName = problemRow[indNameCol] || '';
  const indOrg = problemRow[indOrgCol] || '';
  const indPhone = problemRow[indPhoneCol] || '';
  const indEmail = problemRow[indEmailCol] || '';

  matches.forEach(match => {
    const matchRes = supabasePost('tie_matches', {
      problem_id: problemId,
      expert_id: match.id,
      match_score: match.similarity,
      status: 'pending'
    });

    const matchId = matchRes && matchRes.length > 0 ? matchRes[0].id : '';

    matchSheet.appendRow([
      formatDate(new Date().toISOString()),
      (match.similarity * 100).toFixed(1) + '%',
      challengeDesc,
      indName, indOrg, indPhone, indEmail,
      '',
      match.name || '',
      match.organisation || '',
      match.phone || '',
      match.email || '',
      'PENDING',
      'NOTIFY',
      matchId
    ]);
  });

  const lastRow = matchSheet.getLastRow();
  const firstDataRow = lastRow - matches.length + 1;
  if (firstDataRow >= 2) {
    matchSheet.getRange(firstDataRow, 14, matches.length, 1)
      .setBackground('#1a3a5c').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  }

  SpreadsheetApp.flush();
  Logger.log('Matches saved to sheet');
}

// =============================================
// ON EDIT — color + NOTIFY click handler
// =============================================

function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();

    if (sheetName === MATCHES_TAB) {
      if (e.range.getColumn() === 14 && e.range.getRow() >= 2) {
        const val = String(e.value || '').trim().toUpperCase();
        if (val === 'NOTIFY') {
          handleNotify(sheet, e.range.getRow());
        }
      }
      return;
    }

    if (sheetName !== EXPERTS_TAB && sheetName !== INDUSTRIES_TAB && sheetName !== PROBLEMS_TAB) return;

    const headers = getHeaders(sheet);
    const statusCol = getCol(headers, 'Status') + 1;
    if (e.range.getColumn() !== statusCol) return;
    if (e.range.getRow() < 2) return;

    const value = String(e.value || '').trim().toUpperCase();
    const cell = e.range;

    if (value === 'APPROVED') {
      cell.setBackground('#00C853'); cell.setFontColor('#FFFFFF'); cell.setFontWeight('bold');
    } else if (value === 'REJECT') {
      cell.setBackground('#FFEBEE'); cell.setFontColor('#C62828'); cell.setFontWeight('bold');
    } else if (value === 'ON_HOLD') {
      cell.setBackground('#FFFDE7'); cell.setFontColor('#F57F17'); cell.setFontWeight('bold');
    } else if (value === 'PENDING') {
      cell.setBackground('#E3F2FD'); cell.setFontColor('#1565C0'); cell.setFontWeight('bold');
    }
  } catch(err) {
    Logger.log('onEdit error: ' + err);
  }
}

// =============================================
// NOTIFY BOTH PARTIES
// =============================================

function handleNotify(sheet, row) {
  const data = sheet.getRange(row, 1, 1, 15).getValues()[0];

  const challengeDesc = data[2];
  const indName  = data[3];
  const indOrg   = data[4];
  const indPhone = data[5];
  const indEmail = data[6];
  const startupName  = data[8];
  const startupOrg   = data[9];
  const startupPhone = data[10];
  const startupEmail = data[11];
  const matchId      = data[14];
  const scoreRaw     = data[1];
  const score = (parseFloat(scoreRaw) * 100).toFixed(1) + '%';

  const platformName = 'TIEcon Match-a-Thon';

  const indSubject = `TIEcon Match-a-Thon — Startup Match Found for Your Challenge`;
  const indBody = `Dear ${indName},\n\nGreat news! We have found a startup match for your submitted challenge.\n\nMatch Score: ${score}\n\nStartup Details:\nName: ${startupName}\nOrganisation: ${startupOrg}\nPhone: ${startupPhone}\nEmail: ${startupEmail}\n\nPlease reach out to connect and explore how they can help.\n\nRegards,\nTIEcon Match-a-Thon Team\n${ADMIN_EMAIL}`;

  const expSubject = `TIEcon Match-a-Thon — New Challenge Match for You`;
  const expBody = `Dear ${startupName},\n\nYou have been matched with an industry partner who needs your expertise.\n\nChallenge Summary:\n${challengeDesc}\n\nIndustry Partner Details:\nName: ${indName}\nOrganisation: ${indOrg}\nPhone: ${indPhone}\nEmail: ${indEmail}\n\nPlease reach out to connect and explore how you can help.\n\nRegards,\nTIEcon Match-a-Thon Team\n${ADMIN_EMAIL}`;

  try {
    GmailApp.sendEmail(indEmail, indSubject, indBody, { name: platformName });
    GmailApp.sendEmail(startupEmail, expSubject, expBody, { name: platformName });

    sheet.getRange(row, 13).setValue('NOTIFIED').setBackground('#00C853').setFontColor('white').setFontWeight('bold');
    sheet.getRange(row, 14).setValue('✓ SENT').setBackground('#00C853').setFontColor('white').setFontWeight('bold');

    if (matchId) supabasePatch('tie_matches', matchId, { status: 'notified' });

    Logger.log('Notifications sent for match: ' + matchId);
  } catch(e) {
    Logger.log('Notify error: ' + e);
  }
}

// =============================================
// GENERATE EMBEDDING via Edge Function
// =============================================

function generateEmbedding(payload) {
  const response = UrlFetchApp.fetch(EDGE_FUNCTION_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + SUPABASE_KEY
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  if (!result.success) throw new Error('Edge Function error: ' + result.error);
  Logger.log('Embedding saved. Dims: ' + result.dims);
}

// =============================================
// EMAIL
// =============================================

function sendEmail(email, name, type, decision, notes) {
  const platformName = 'TIEcon Match-a-Thon';
  const typeLabel = type === 'expert' ? 'Startup' : 'Industry Partner';
  let subject, body;

  if (decision === 'approved') {
    subject = `✅ Your TIEcon Match-a-Thon registration is approved!`;
    body = `Dear ${name},\n\nGreat news! Your registration as a ${typeLabel} on ${platformName} has been approved.\n\n`
      + (type === 'industry'
          ? `You can now submit your challenge at:\nhttps://vkv-coder.github.io/tiecon-match/submit-challenge.html\n\n`
          : '')
      + `Our team will be in touch with you shortly.\n\n`
      + (notes ? `Note from our team: ${notes}\n\n` : '')
      + `Regards,\nTIEcon Match-a-Thon Team\n${ADMIN_EMAIL}`;

  } else if (decision === 'on_hold') {
    subject = `TIEcon Match-a-Thon — Additional Information Required`;
    body = `Dear ${name},\n\nThank you for registering on ${platformName}.\n\nWe need a little more information before processing your registration.\n\n`
      + (notes ? `Our team notes: ${notes}\n\n` : '')
      + `Please reply to this email or contact us at ${ADMIN_EMAIL}.\n\nRegards,\nTIEcon Match-a-Thon Team`;
  }

  GmailApp.sendEmail(email, subject, body, { name: platformName });
}

// =============================================
// STATUS CELL COLOR
// =============================================

function setStatusCell(sheet, row, col, text, bgColor, fontColor) {
  const cell = sheet.getRange(row, col);
  cell.setValue(text);
  cell.setBackground(bgColor);
  cell.setFontColor(fontColor);
  cell.setFontWeight('bold');
}

// =============================================
// DROPDOWN + PENDING COLOR
// =============================================

function setupDropdownAndColor(sheet, headers, options) {
  const statusCol = getCol(headers, 'Status') + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, statusCol, lastRow - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .build();
  range.setDataValidation(rule);

  const values = range.getValues();
  values.forEach((val, i) => {
    const cell = sheet.getRange(i + 2, statusCol);
    const v = String(val[0]).trim().toUpperCase();
    if (v === 'PENDING') {
      cell.setBackground('#E3F2FD');
      cell.setFontColor('#1565C0');
      cell.setFontWeight('bold');
    }
  });
}

// =============================================
// HELPERS
// =============================================

function buildRow(headers, dataMap) {
  const totalCols = Object.keys(headers).length;
  const row = new Array(totalCols).fill('');
  Object.entries(dataMap).forEach(([key, value]) => {
    if (headers[key] !== undefined) row[headers[key]] = value;
  });
  return row;
}

function formatDate(isoString) {
  if (!isoString) return '';
  const d = new Date(isoString);
  return `${d.getDate().toString().padStart(2,'0')}/${(d.getMonth()+1).toString().padStart(2,'0')}/${d.getFullYear()}`;
}

function supabaseGet(endpoint) {
  try {
    const response = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/${endpoint}`, {
      method: 'GET',
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    return JSON.parse(response.getContentText());
  } catch(e) { Logger.log('GET error: ' + e); return null; }
}

function supabasePatch(table, id, data) {
  try {
    UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/${table}?id=eq.${id}`, {
      method: 'PATCH',
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Content-Type': 'application/json',
        'Prefer': 'return=minimal'
      },
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    });
  } catch(e) { Logger.log('PATCH error: ' + e); }
}

function supabasePost(table, data) {
  try {
    const response = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/${table}`, {
      method: 'POST',
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Content-Type': 'application/json',
        'Prefer': 'return=representation'
      },
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    });
    return JSON.parse(response.getContentText());
  } catch(e) { Logger.log('POST error: ' + e); return null; }
}

function getExistingIds(sheet, idColNum) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, idColNum, last - 1, 1)
    .getValues()
    .map(r => String(r[0]).trim())
    .filter(Boolean);
}

// =============================================
// TRIGGERS
// =============================================

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pullAllFromSupabase').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('pushApprovalsToSupabase').timeBased().everyMinutes(10).create();
  Logger.log('Triggers set up successfully');
}

// =============================================
// MATCHES SHEET FORMATTER (run once)
// =============================================

function formatMatchesSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(MATCHES_TAB);
  if (!sheet) { Logger.log('MATCHES sheet not found'); return; }

  sheet.getRange(1, 1, 1, 7).setBackground('#f4a11d').setFontColor('#0a2540').setFontWeight('bold');
  sheet.getRange(1, 8, 1, 1).setBackground('#e5e7eb');
  sheet.setColumnWidth(8, 18);
  sheet.getRange(1, 9, 1, 7).setBackground('#0a2540').setFontColor('white').setFontWeight('bold');

  SpreadsheetApp.flush();
  Logger.log('MATCHES sheet formatted');
}

// =============================================
// WEB APP — Acknowledgement Email
// =============================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const name = data.name || '';
    const email = data.email || '';
    const type = data.type || 'expert';
    const platformName = 'TIEcon Match-a-Thon';

    const subject = `Thank you for registering on ${platformName}`;

    let body;
    if (type === 'problem') {
      body = `Dear ${name},\n\nThank you for submitting your challenge on ${platformName}.\n\nOur team will review it and our AI will match you with the best startups.\nYou will be notified by email once matches are found.\n\nRegards,\nTIEcon Match-a-Thon Team`;
    } else {
      body = `Dear ${name},\n\nThank you for registering on ${platformName} as a ${type === 'expert' ? 'Startup' : 'Industry Partner'}.\n\nWe have received your registration and our team will review your profile within 2-3 working days.\n\nYou will receive a confirmation once your profile is activated.\n\nFor any queries, please contact us at ${ADMIN_EMAIL}.\n\nRegards,\nTIEcon Match-a-Thon Team`;
    }

    GmailApp.sendEmail(email, subject, body, { name: platformName });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =============================================
// TEST FUNCTIONS
// =============================================

function testMatchNow() {
  const problemId = 'cbd00808-db17-4b81-96fc-c32879766952';
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const problems = supabaseGet(`tie_problems?id=eq.${problemId}&select=embedding`);
  Logger.log('Embedding type: ' + typeof problems[0].embedding);

  let embedding = problems[0].embedding;
  if (typeof embedding === 'string') { embedding = JSON.parse(embedding); }
  Logger.log('Embedding length: ' + embedding.length);

  const response = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/rpc/match_experts`, {
    method: 'POST',
    headers: {
      'apikey': SUPABASE_KEY,
      'Authorization': 'Bearer ' + SUPABASE_KEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      query_embedding: embedding,
      match_threshold: 0.1,
      match_count: 3
    }),
    muteHttpExceptions: true
  });

  const matches = JSON.parse(response.getContentText());
  Logger.log('Matches: ' + JSON.stringify(matches));

  if (matches && matches.length > 0) {
    const sheet = ss.getSheetByName(PROBLEMS_TAB);
    const headers = getHeaders(sheet);
    const data = sheet.getDataRange().getValues();
    saveMatchesToSheet(ss, problemId, matches, data[1], headers);
    Logger.log('MATCHES tab updated!');
  }
}

function testNotifyNow() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(MATCHES_TAB);
  handleNotify(sheet, 2);
}
