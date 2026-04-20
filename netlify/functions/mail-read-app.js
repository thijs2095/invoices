const DEFAULT_ALLOWED_HEADERS = 'Content-Type, Authorization';
const DEFAULT_ALLOWED_METHODS = 'POST, OPTIONS';

class HttpError extends Error {
  constructor(statusCode, message) {
    super(message);
    this.name = 'HttpError';
    this.statusCode = statusCode;
  }
}

function getCorsHeaders(event) {
  const reqOrigin = (event.headers && (event.headers.origin || event.headers.Origin)) || '';
  const configured = (process.env.ALLOWED_ORIGINS || '').split(',').map(s => s.trim()).filter(Boolean);
  const allowOrigin = configured.length
    ? (configured.includes(reqOrigin) ? reqOrigin : configured[0])
    : '*';

  return {
    'Access-Control-Allow-Origin': allowOrigin,
    'Access-Control-Allow-Headers': DEFAULT_ALLOWED_HEADERS,
    'Access-Control-Allow-Methods': DEFAULT_ALLOWED_METHODS,
    'Content-Type': 'application/json'
  };
}

function requiredEnv(name) {
  const value = process.env[name];
  if (!value) throw new HttpError(500, `Missing environment variable: ${name}`);
  return value;
}

function getBearerToken(event) {
  const auth = (event.headers && (event.headers.authorization || event.headers.Authorization)) || '';
  if (!auth.startsWith('Bearer ')) return '';
  return auth.slice(7).trim();
}

async function verifyFrontendUserToken(event) {
  const token = getBearerToken(event);
  if (!token) return null;

  const resp = await fetch('https://graph.microsoft.com/v1.0/me?$select=id,displayName,userPrincipalName,mail', {
    headers: { Authorization: `Bearer ${token}` }
  });

  if (!resp.ok) return null;
  return resp.json();
}

async function getAzureAppToken() {
  const tenantId = requiredEnv('AZURE_TENANT_ID');
  const clientId = requiredEnv('AZURE_CLIENT_ID');
  const clientSecret = requiredEnv('AZURE_CLIENT_SECRET');

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default'
  });

  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });

  if (!resp.ok) {
    const errText = await resp.text();
    throw new Error(`Azure token error (${resp.status}): ${errText}`);
  }

  const data = await resp.json();
  return data.access_token;
}

async function graphGet(accessToken, pathOrUrl, extraHeaders = {}) {
  const url = pathOrUrl.startsWith('http')
    ? pathOrUrl
    : `https://graph.microsoft.com/v1.0${pathOrUrl}`;

  const resp = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...extraHeaders
    }
  });

  if (!resp.ok) {
    const text = await resp.text();
    throw new HttpError(resp.status, `Graph error (${resp.status}): ${text}`);
  }

  return resp.json();
}

async function listAllMailboxes(accessToken) {
  const mailboxes = [];
  let nextUrl = '/users?$select=id,displayName,mail,userPrincipalName&$top=999';

  while (nextUrl) {
    const data = await graphGet(accessToken, nextUrl);
    const users = data.value || [];
    for (const u of users) {
      const mail = u.mail || u.userPrincipalName;
      if (!mail) continue;
      mailboxes.push({
        id: u.id,
        displayName: u.displayName || mail,
        mail
      });
    }
    nextUrl = data['@odata.nextLink'] || null;
  }

  mailboxes.sort((a, b) => a.mail.localeCompare(b.mail));
  return mailboxes;
}

function parseMailboxListFromEnv() {
  const raw = process.env.MAILBOX_LIST || process.env.DHL_WAREHOUSE_EMAIL || '';
  const values = raw
    .split(',')
    .map(v => v.trim())
    .filter(Boolean);

  const uniq = Array.from(new Set(values));
  return uniq.map(mail => ({
    id: mail,
    displayName: mail,
    mail
  }));
}

async function searchMessages(
  accessToken,
  terms,
  mailboxes,
  top = 5,
  requestedScope = 'all',
  requestedExcludeLouvenberg = false,
  requestedYear = 'all',
  requestedOnlyWithAttachments = false
) {
  const unique = new Map();
  const safeTop = Math.min(Math.max(Number(top) || 5, 1), 25);
  const rawScope = String(requestedScope || 'all').trim().toLowerCase();
  const searchScope = ['all', 'subject', 'from', 'body'].includes(rawScope) ? rawScope : 'all';
  const excludeLouvenberg = !!requestedExcludeLouvenberg;
  const onlyWithAttachments = !!requestedOnlyWithAttachments;
  const parsedYear = Number(requestedYear);
  const yearFilter = Number.isInteger(parsedYear) && parsedYear >= 2000 && parsedYear <= 2100
    ? parsedYear
    : null;

  function buildAqs(term) {
    const escaped = String(term).replace(/"/g, '\\"');
    if (searchScope === 'subject') return `subject:"${escaped}"`;
    if (searchScope === 'from') return `from:"${escaped}"`;
    if (searchScope === 'body') return `body:"${escaped}"`;
    return `"${escaped}"`;
  }

  for (const rawTerm of terms || []) {
    const term = String(rawTerm || '').trim();
    if (!term) continue;
    const q = encodeURIComponent(buildAqs(term));

    for (const rawMailbox of mailboxes || []) {
      const mailbox = String(rawMailbox || '').trim();
      if (!mailbox) continue;

      const path = `/users/${encodeURIComponent(mailbox)}/messages`
        + `?$search=${q}`
        + `&$select=id,subject,from,receivedDateTime,hasAttachments,bodyPreview,webLink`
        + `&$top=${safeTop}`;

      try {
        const data = await graphGet(accessToken, path, { ConsistencyLevel: 'eventual' });
        for (const m of data.value || []) {
          if (onlyWithAttachments && !m.hasAttachments) continue;

          if (yearFilter) {
            const y = new Date(m.receivedDateTime).getFullYear();
            if (y !== yearFilter) continue;
          }

          const fromAddress = String(m?.from?.emailAddress?.address || '').toLowerCase().trim();
          if (excludeLouvenberg && fromAddress.endsWith('@louvenbergadvies.nl')) continue;

          if (unique.has(m.id)) continue;
          unique.set(m.id, {
            id: m.id,
            subject: m.subject,
            from: m.from,
            receivedDateTime: m.receivedDateTime,
            hasAttachments: !!m.hasAttachments,
            bodyPreview: m.bodyPreview,
            webLink: m.webLink,
            _mailbox: mailbox,
            _matchedTerm: term,
            _matchedScope: searchScope
          });
        }
      } catch (e) {
        // Ignore mailbox-level failures so one inaccessible mailbox does not fail all searches.
      }
    }
  }

  return Array.from(unique.values())
    .sort((a, b) => (new Date(b.receivedDateTime) - new Date(a.receivedDateTime)));
}

async function getAttachments(accessToken, mailbox, messageId) {
  const safeMailbox = String(mailbox || '').trim();
  const safeMessageId = String(messageId || '').trim();
  if (!safeMailbox || !safeMessageId) {
    throw new Error('mailbox and messageId are required');
  }

  const basePath = `/users/${encodeURIComponent(safeMailbox)}/messages/${encodeURIComponent(safeMessageId)}/attachments`;
  const listPath = `${basePath}?$select=id,name,size,contentType`;
  const data = await graphGet(accessToken, listPath);

  const result = [];
  for (const a of (data.value || [])) {
    // Load attachment details first; type info is reliable there.
    try {
      const detailPath = `${basePath}/${encodeURIComponent(a.id)}`;
      let detail = await graphGet(accessToken, detailPath);

      if (String(detail['@odata.type'] || '') !== '#microsoft.graph.fileAttachment') continue;

      // Fallback: some tenants/versions omit contentBytes in the generic detail payload.
      if (!detail.contentBytes) {
        try {
          const castPath = `${basePath}/${encodeURIComponent(a.id)}/microsoft.graph.fileAttachment?$select=id,name,size,contentType,contentBytes`;
          detail = await graphGet(accessToken, castPath);
        } catch (e) {
          // Keep original detail and let the contentBytes check below decide.
        }
      }

      if (!detail.contentBytes) continue;

      result.push({
        id: detail.id,
        name: detail.name || a.name,
        size: detail.size || a.size,
        contentType: detail.contentType || a.contentType,
        contentBytes: detail.contentBytes
      });
    } catch (e) {
      // Skip attachments we cannot fetch (permissions/type-specific constraints).
      continue;
    }
  }

  return result;
}

async function getMessage(accessToken, mailbox, messageId) {
  const safeMailbox = String(mailbox || '').trim();
  const safeMessageId = String(messageId || '').trim();
  if (!safeMailbox || !safeMessageId) {
    throw new Error('mailbox and messageId are required');
  }

  const path = `/users/${encodeURIComponent(safeMailbox)}/messages/${encodeURIComponent(safeMessageId)}`
    + '?$select=id,subject,from,receivedDateTime,body,bodyPreview,hasAttachments,webLink';

  return graphGet(accessToken, path);
}

exports.handler = async (event) => {
  const headers = getCorsHeaders(event);

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  try {
    const payload = event.body ? JSON.parse(event.body) : {};
    const action = payload.action;

    if (!action) {
      return { statusCode: 400, headers, body: JSON.stringify({ error: 'Missing action' }) };
    }

    if (action === 'health') {
      return { statusCode: 200, headers, body: JSON.stringify({ ok: true }) };
    }

    const frontendUser = await verifyFrontendUserToken(event);
    if (!frontendUser) {
      return { statusCode: 401, headers, body: JSON.stringify({ error: 'Unauthorized: log in with Azure AD first' }) };
    }

    const accessToken = await getAzureAppToken();

    if (action === 'listMailboxes') {
      let mailboxes = [];
      let source = 'graph';
      try {
        mailboxes = await listAllMailboxes(accessToken);
      } catch (e) {
        // If app permission for /users is missing, allow a configured static list.
        if (e instanceof HttpError && (e.statusCode === 401 || e.statusCode === 403)) {
          mailboxes = parseMailboxListFromEnv();
          source = 'MAILBOX_LIST';
          if (mailboxes.length === 0) {
            throw new HttpError(
              403,
              'Geen toegang tot /users. Voeg Application permission User.Read.All toe met admin consent, of zet MAILBOX_LIST in Netlify env.'
            );
          }
        } else {
          throw e;
        }
      }
      return { statusCode: 200, headers, body: JSON.stringify({ mailboxes, source }) };
    }

    if (action === 'searchMessages') {
      const messages = await searchMessages(
        accessToken,
        payload.terms || [],
        payload.mailboxes || [],
        payload.top || 5,
        payload.searchScope || 'all',
        payload.excludeLouvenberg === true,
        payload.searchYear || 'all',
        payload.onlyWithAttachments === true
      );
      return { statusCode: 200, headers, body: JSON.stringify({ messages }) };
    }

    if (action === 'getAttachments') {
      const attachments = await getAttachments(accessToken, payload.mailbox, payload.messageId);
      return { statusCode: 200, headers, body: JSON.stringify({ attachments }) };
    }

    if (action === 'getMessage') {
      const message = await getMessage(accessToken, payload.mailbox, payload.messageId);
      return { statusCode: 200, headers, body: JSON.stringify({ message }) };
    }

    return { statusCode: 400, headers, body: JSON.stringify({ error: `Unknown action: ${action}` }) };
  } catch (error) {
    const msg = error && error.message ? error.message : 'Unknown backend error';
    const statusCode = (error && error.statusCode) ? error.statusCode : 500;
    return { statusCode, headers, body: JSON.stringify({ error: msg }) };
  }
};
