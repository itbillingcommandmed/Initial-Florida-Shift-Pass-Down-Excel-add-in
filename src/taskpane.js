import { createNestablePublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';
import { config } from './config';

let msalInstance;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('submit').addEventListener('click', submitPassdown);
    setStatus('Ready.');
  }
});

async function initMsal() {
  if (!msalInstance) {
    msalInstance = await createNestablePublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`
      },
      cache: { cacheLocation: 'localStorage' }
    });
  }
}

async function acquireGraphToken() {
  await initMsal();
  const request = { scopes: ['Lists.SelectedOperations.Selected'] };
  try {
    return (await msalInstance.acquireTokenSilent(request)).accessToken;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) throw error;
    return (await msalInstance.acquireTokenPopup(request)).accessToken;
  }
}

async function readSubmissionFromWorkbook() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(config.worksheetName);
    const shiftRange = sheet.getRange(config.shiftCell);
    shiftRange.load('text');
    await context.sync();

    const shift = (shiftRange.text?.[0]?.[0] || '').trim();
    if (!shift) throw new Error(`Shift is blank at ${config.worksheetName}!${config.shiftCell}.`);

    // The queue currently has only the writable Title field.
    // The existing Power Automate flow can use this item as its trigger and then read the workbook.
    return {
      Title: `Florida Shift Pass Down - ${shift} - ${new Date().toISOString()}`
    };
  });
}

async function createQueueItem(accessToken, fields) {
  const endpoint = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(config.siteId)}/lists/${config.listId}/items`;
  const response = await fetch(endpoint, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ fields })
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Graph returned ${response.status}: ${body}`);
  }
  return response.json();
}

async function submitPassdown() {
  const button = document.getElementById('submit');
  button.disabled = true;
  setStatus('Reading the workbook...');
  try {
    const fields = await readSubmissionFromWorkbook();
    setStatus('Signing in and creating the queue item...');
    const accessToken = await acquireGraphToken();
    const item = await createQueueItem(accessToken, fields);
    setStatus(`Submitted successfully. Queue item ID: ${item.id}`, 'ok');
  } catch (error) {
    console.error(error);
    setStatus(error?.message || String(error), 'error');
  } finally {
    button.disabled = false;
  }
}

function setStatus(message, cssClass = '') {
  const status = document.getElementById('status');
  status.textContent = message;
  status.className = cssClass;
}
