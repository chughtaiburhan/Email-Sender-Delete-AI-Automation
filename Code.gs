const CONFIG = {
  appName: 'Email Automation System',
  recipientsSheetName: 'Recipients',
  profileSheetName: 'Profile',
  defaultBatchSize: 50,
  defaultDeleteDays: 30,
  templateFolderId: '', // Optional: set a Drive folder ID for templates.
};

function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  template.appName = CONFIG.appName;
  return template
    .evaluate()
    .setTitle(CONFIG.appName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function createUserProfile(profile) {
  validateProfile_(profile);
  const props = PropertiesService.getUserProperties();
  const spreadsheetId = getOrCreateSpreadsheet_(profile);
  props.setProperty('profile', JSON.stringify(profile));
  props.setProperty('spreadsheetId', spreadsheetId);
  return { spreadsheetId };
}

function listDocTemplates() {
  const docs = [];
  const folderId = CONFIG.templateFolderId;
  const iterator = folderId
    ? DriveApp.getFolderById(folderId).getFilesByType(MimeType.GOOGLE_DOCS)
    : DriveApp.getFilesByType(MimeType.GOOGLE_DOCS);

  while (iterator.hasNext() && docs.length < 50) {
    const file = iterator.next();
    docs.push({ id: file.getId(), name: file.getName() });
  }
  return docs;
}

function saveSelectedTemplate(templateId) {
  if (!templateId) {
    throw new Error('Template ID is required.');
  }
  PropertiesService.getUserProperties().setProperty('templateId', templateId);
  return { templateId };
}

function saveDraftMessage(subject, body) {
  if (!subject || !body) {
    throw new Error('Subject and body are required.');
  }
  const props = PropertiesService.getUserProperties();
  props.setProperty('draftSubject', subject);
  props.setProperty('draftBody', body);
  return { ok: true };
}

function sendPersonalizedEmails() {
  const props = PropertiesService.getUserProperties();
  const spreadsheetId = props.getProperty('spreadsheetId');
  const templateId = props.getProperty('templateId');
  const draftSubject = props.getProperty('draftSubject') || '';
  const draftBody = props.getProperty('draftBody') || '';

  if (!spreadsheetId) {
    throw new Error('Spreadsheet not found. Please create a profile first.');
  }
  if (!templateId && !draftBody) {
    throw new Error('Please select a template or save a draft message.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(CONFIG.recipientsSheetName);
  if (!sheet) {
    throw new Error(`Sheet "${CONFIG.recipientsSheetName}" not found.`);
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { sent: 0 };
  }

  const headers = normalizeHeaders_(data[0]);
  const rows = data.slice(1).filter(row => row[0]);
  const templateBody = templateId ? extractDocBody_(templateId) : draftBody;
  const templateSubject = templateId ? extractDocTitle_(templateId) : draftSubject;

  const firstHalf = rows.slice(0, Math.ceil(rows.length / 2));
  const secondHalf = rows.slice(Math.ceil(rows.length / 2));

  let sentCount = 0;
  sentCount += processRowsInBatches_(firstHalf, headers, templateSubject, templateBody);
  sentCount += processRowsInBatches_(secondHalf, headers, templateSubject, templateBody);

  return { sent: sentCount };
}

function deleteOldEmails(days) {
  const deleteDays = Number(days) || CONFIG.defaultDeleteDays;
  const query = `older_than:${deleteDays}d`;
  const threads = GmailApp.search(query);
  const batchSize = CONFIG.defaultBatchSize;

  let deleted = 0;
  for (let i = 0; i < threads.length; i += batchSize) {
    const batch = threads.slice(i, i + batchSize);
    GmailApp.moveThreadsToTrash(batch);
    deleted += batch.length;
    Utilities.sleep(500);
  }

  return { deleted };
}

function getSpreadsheetUrl() {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty('spreadsheetId');
  if (!spreadsheetId) {
    return null;
  }
  return SpreadsheetApp.openById(spreadsheetId).getUrl();
}

function validateProfile_(profile) {
  if (!profile || !profile.name || !profile.email || !profile.phone) {
    throw new Error('Name, email, and phone are required.');
  }
}

function getOrCreateSpreadsheet_(profile) {
  const props = PropertiesService.getUserProperties();
  const existing = props.getProperty('spreadsheetId');
  if (existing) {
    return existing;
  }

  const spreadsheet = SpreadsheetApp.create(`${CONFIG.appName} - ${profile.name}`);
  const profileSheet = spreadsheet.getSheetByName(CONFIG.profileSheetName) || spreadsheet.insertSheet(CONFIG.profileSheetName);
  profileSheet.getRange('A1:B4').setValues([
    ['Name', profile.name],
    ['Email', profile.email],
    ['Phone', profile.phone],
    ['Created', new Date()],
  ]);

  const recipientsSheet = spreadsheet.getSheetByName(CONFIG.recipientsSheetName) || spreadsheet.insertSheet(CONFIG.recipientsSheetName);
  recipientsSheet.getRange('A1:E1').setValues([[
    'email',
    'name',
    'about',
    'company',
    'custom',
  ]]);

  return spreadsheet.getId();
}

function extractDocBody_(docId) {
  const doc = DocumentApp.openById(docId);
  return doc.getBody().getText();
}

function extractDocTitle_(docId) {
  const doc = DocumentApp.openById(docId);
  return doc.getName();
}

function processRowsInBatches_(rows, headers, subjectTemplate, bodyTemplate) {
  const batchSize = CONFIG.defaultBatchSize;
  let sent = 0;

  for (let i = 0; i < rows.length; i += batchSize) {
    const batch = rows.slice(i, i + batchSize);
    sent += sendBatch_(batch, headers, subjectTemplate, bodyTemplate);
    Utilities.sleep(300);
  }

  return sent;
}

function sendBatch_(rows, headers, subjectTemplate, bodyTemplate) {
  rows.forEach(row => {
    const data = mapRowToObject_(headers, row);
    const email = data.email;
    if (!email) {
      return;
    }
    const subject = applyTemplate_(subjectTemplate, data);
    const body = applyTemplate_(bodyTemplate, data);

    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      name: CONFIG.appName,
    });
  });

  return rows.length;
}

function normalizeHeaders_(headers) {
  return headers.map(header => String(header || '').trim().toLowerCase());
}

function mapRowToObject_(headers, row) {
  return headers.reduce((acc, header, index) => {
    if (header) {
      acc[header] = row[index];
    }
    return acc;
  }, {});
}

function applyTemplate_(template, data) {
  return template.replace(/\$\{([^}]+)\}/g, (match, key) => {
    const value = data[String(key).trim().toLowerCase()];
    return value !== undefined ? value : match;
  });
}
