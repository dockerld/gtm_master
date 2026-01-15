/**
 * Dropdown action in "Create Task" column:
 * Values: "Docker" or "Camden"
 *
 * - Responds ONLY on sheet: Sauron
 * - On dropdown selection:
 *    * Reads Task Type column to decide the message template
 *    * If Task Type is empty/unknown, uses a generic fallback message
 *    * Creates a Notion page with:
 *         Task (title)    = built title  [title property named "Task"]
 *         Owner (people)  = Docker or Camden (mapped via NOTION_PERSON_MAP)
 *    * No dedupe or throttle: you can create tasks repeatedly from the same row
 *    * Resets the dropdown cell back to blank
 */

const NOTION_MIN = {
  ENABLED_SHEETS: ['Sauron'],
  HEADER_ROW_BY_SHEET: {
    'Sauron': 3
  },
  HEADERS: {
    ACTION: 'Create Task',
    NAME: 'Name',
    ORG: 'Org Name',
    EMAIL: 'Email',
    TASK_TYPE: 'Task Type',
    NOTE: 'Note',
  },
  // Templates keyed by task type (lowercased)
  TASK_TEMPLATES: {
    'onboarding': (name, org, email) =>
      `Onboarding task for ${name || '{{Name}}'} from ${org || '{{Org Name}}'} – help them get fully set up Email: ${email || '{{Email}}'}`,

    'reachout': (name, org, email) =>
      `Reach out to ${name || '{{Name}}'} from ${org || '{{Org Name}}'} – check in and see how things are going Email: ${email || '{{Email}}'}`,

    'send video': (name, org, email) =>
      `Send demo / explainer video to ${name || '{{Name}}'} from ${org || '{{Org Name}}'} Email:${email || '{{Email}}'}`,

    'get clients': (name, org, email) =>
      `Help ${name || '{{Name}}'} from ${org || '{{Org Name}}'} get more clients – schedule strategy conversation Email: ${email || '{{Email}}'}`,

    'white glove': (name, org, email) =>
      `WHITE GLOVE!!!  ${name || '{{Name}}'} from ${org || '{{Org Name}}'} They need some attention Email: ${email || '{{Email}}'}`,

    'trial ending': (name, org, email) =>
      `TRIAL ENDING, reach out to  ${name || '{{Name}}'} from ${org || '{{Org Name}}'} Email: ${email || '{{Email}}'}`,  

    'trial expired': (name, org, email) =>
      `TRIAL expired, reach out to  ${name || '{{Name}}'} from ${org || '{{Org Name}}'} Email: ${email || '{{Email}}'}`, 
  },
  DEFAULT_TEMPLATE: (name, org, email) =>
    `Follow up with ${name || '{{Name}}'} from ${org || '{{Org Name}}'} Email: ${email || '{{Email}}'}`,
  NOTION_VERSION: '2022-06-28',
};

// Simple trigger placeholder (don’t use this for UrlFetch)
function onEdit(e) { return; }

// INSTALLABLE TRIGGER: From spreadsheet → On edit → run onEditInstallable
function onEditInstallable(e) {
  processDropdownToNotion_(e);
}

function processDropdownToNotion_(e) {
  if (!e || !e.range || !e.source) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  if (!NOTION_MIN.ENABLED_SHEETS.includes(sheetName)) return;

  const HEADER_ROW = NOTION_MIN.HEADER_ROW_BY_SHEET[sheetName] || 1;

  // Header lookups
  const headers = sheet
    .getRange(HEADER_ROW, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(h => String(h).trim());

  const col = (name) =>
    headers.findIndex(h => h.toLowerCase() === name.toLowerCase()) + 1;

  const cAction    = col(NOTION_MIN.HEADERS.ACTION);
  const cName      = col(NOTION_MIN.HEADERS.NAME);
  const cOrg       = col(NOTION_MIN.HEADERS.ORG);
  const cEmail     = col(NOTION_MIN.HEADERS.EMAIL);
  const cTaskType  = col(NOTION_MIN.HEADERS.TASK_TYPE);
  const cNote      = col(NOTION_MIN.HEADERS.NOTE);   // may be -1 if not present

  if ([cAction, cName, cOrg, cEmail, cTaskType].some(x => x <= 0)) return;

  const row = e.range.getRow();
  const column = e.range.getColumn();

  // Ignore edits on header row or above
  if (row <= HEADER_ROW || column !== cAction) return;

  const assignee = String(e.range.getValue() || '').trim();
  const validAssignee = assignee === 'Docker' || assignee === 'Camden';
  if (!validAssignee) {
    sheet.getRange(row, cAction).setValue('');
    return;
  }

  try {
    const name  = String(sheet.getRange(row, cName).getValue()  || '').trim();
    const org   = String(sheet.getRange(row, cOrg).getValue()   || '').trim();
    const email = String(sheet.getRange(row, cEmail).getValue() || '').trim();
    const taskTypeRaw = String(sheet.getRange(row, cTaskType).getValue() || '').trim();
    const taskTypeKey = taskTypeRaw.toLowerCase();

    const note =
      cNote > 0
        ? String(sheet.getRange(row, cNote).getValue() || '').trim()
        : '';

    if (!name && !org && !email) {
      sheet.getRange(row, cAction).setValue('');
      return;
    }

    let title = '';

    if (taskTypeKey) {
      const typeTemplate = NOTION_MIN.TASK_TEMPLATES[taskTypeKey];
      const tpl = typeTemplate || NOTION_MIN.DEFAULT_TEMPLATE;
      title = tpl(name, org, email);
      if (note) title += `  Note: ${note}`;
    } else {
      if (note) {
        title = `Note regarding ${name || org || email}: ${note}`;
        if (email) title += ` (Email: ${email})`;
      } else {
        const fallback = NOTION_MIN.DEFAULT_TEMPLATE;
        title = fallback(name, org, email);
      }
    }

    createNotionPageWithPerson_({ title, personName: assignee, email });

  } catch (err) {
    Logger.log('Error creating Notion task: ' + (err && err.message ? err.message : err));
  } finally {
    sheet.getRange(row, cAction).setValue('');
    sheet.getRange(row, cTaskType).setValue('');
    if (cNote > 0) sheet.getRange(row, cNote).setValue('');
  }
}

// ----- Notion helpers -----

function createNotionPageWithPerson_({ title, personName, email }) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('NOTION_TOKEN');
  const dbId  = props.getProperty('NOTION_TASKS_DB_ID');
  if (!token || !dbId) throw new Error('Missing NOTION_TOKEN or NOTION_TASKS_DB_ID');

  const mapJson = props.getProperty('NOTION_PERSON_MAP') || '{}';
  const idMap = JSON.parse(mapJson);
  const notionUserId = idMap[personName];
  if (!notionUserId) throw new Error(`No Notion userId mapped for "${personName}". Set NOTION_PERSON_MAP script property.`);

  const headers = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json',
    'Notion-Version': NOTION_MIN.NOTION_VERSION
  };

  const titleRichText = buildNotionTitleWithEmailLink_(title, email);

  const url = 'https://api.notion.com/v1/pages';
  const body = {
    parent: { database_id: normalizeDbId_(dbId) },
    properties: {
      'Task':   { title: titleRichText },
      'Owner':  { people: [{ id: notionUserId }] },
      'Priority Level': { select: { name: 'High' } }
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code >= 300) throw new Error(`Notion API error ${code}: ${res.getContentText()}`);
  return JSON.parse(res.getContentText());
}

function buildNotionTitleWithEmailLink_(title, email) {
  if (!email) {
    return [{ type: 'text', text: { content: title } }];
  }

  const lower = title.toLowerCase();
  const labelIndex = lower.lastIndexOf('email:');

  let searchStart = 0;
  if (labelIndex >= 0) searchStart = labelIndex;

  const emailIndex = title.indexOf(email, searchStart);

  if (emailIndex >= 0) {
    const beforeEmail = title.substring(0, emailIndex);
    const afterEmail = title.substring(emailIndex + email.length);

    const rich = [];
    if (beforeEmail) rich.push({ type: 'text', text: { content: beforeEmail } });

    rich.push({
      type: 'text',
      text: { content: email, link: { url: `mailto:${email}` } }
    });

    if (afterEmail) rich.push({ type: 'text', text: { content: afterEmail } });

    return rich;
  }

  return [
    { type: 'text', text: { content: title + ' ' } },
    { type: 'text', text: { content: email, link: { url: `mailto:${email}` } } }
  ];
}

function normalizeDbId_(rawId) {
  if (rawId.includes('-')) return rawId;
  return rawId.replace(/^(.{8})(.{4})(.{4})(.{4})(.{12})$/, '$1-$2-$3-$4-$5');
}

/**
 * Manual test: one-off page with Docker as Person
 */
function testSendToNotionDocker() {
  const title = '✅ Test task (Docker)';
  const page = createNotionPageWithPerson_({ title, personName: 'Docker', email: '' });
  Logger.log('Created test page: ' + (page?.url || '(no url)'));
}

function getNotionUsers() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('NOTION_TOKEN');
  if (!token) throw new Error('Missing NOTION_TOKEN script property');

  const res = UrlFetchApp.fetch('https://api.notion.com/v1/users', {
    method: 'get',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Notion-Version': '2022-06-28'
    }
  });

  const users = JSON.parse(res.getContentText());
  Logger.log(JSON.stringify(users, null, 2));
}