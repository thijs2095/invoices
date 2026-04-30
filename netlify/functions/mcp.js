// Remote MCP server for Claude (Custom Connectors) over Streamable HTTP.
// Exposes Microsoft Graph mailbox search/read tools using the same Azure AD
// app credentials as mail-read-app.js.
//
// Required env vars (already used by mail-read-app.js):
//   AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
// Recommended env var for this endpoint:
//   MCP_BEARER_TOKEN  -> long random string. If unset, the endpoint is OPEN.
// Optional:
//   MAILBOX_LIST      -> comma-separated fallback if /users is not permitted.

const PROTOCOL_VERSION = '2025-06-18';
const SERVER_INFO = { name: 'impossible-mail-mcp', version: '1.0.0' };

class HttpError extends Error {
  constructor(statusCode, message) {
    super(message);
    this.name = 'HttpError';
    this.statusCode = statusCode;
  }
}

function requiredEnv(name) {
  const v = process.env[name];
  if (!v) throw new HttpError(500, `Missing environment variable: ${name}`);
  return v;
}

// ---------- Microsoft Graph helpers (same logic as mail-read-app.js) ----------

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
    throw new HttpError(500, `Azure token error (${resp.status}): ${errText}`);
  }
  const data = await resp.json();
  return data.access_token;
}

async function graphGet(accessToken, pathOrUrl, extraHeaders = {}) {
  const url = pathOrUrl.startsWith('http')
    ? pathOrUrl
    : `https://graph.microsoft.com/v1.0${pathOrUrl}`;

  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}`, ...extraHeaders }
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
    for (const u of data.value || []) {
      const mail = u.mail || u.userPrincipalName;
      if (!mail) continue;
      mailboxes.push({ id: u.id, displayName: u.displayName || mail, mail });
    }
    nextUrl = data['@odata.nextLink'] || null;
  }
  mailboxes.sort((a, b) => a.mail.localeCompare(b.mail));
  return mailboxes;
}

function parseMailboxListFromEnv() {
  const raw = process.env.MAILBOX_LIST || process.env.DHL_WAREHOUSE_EMAIL || '';
  const values = raw.split(',').map(v => v.trim()).filter(Boolean);
  const uniq = Array.from(new Set(values));
  return uniq.map(mail => ({ id: mail, displayName: mail, mail }));
}

async function resolveAllMailboxes(accessToken) {
  try {
    return await listAllMailboxes(accessToken);
  } catch (e) {
    if (e instanceof HttpError && (e.statusCode === 401 || e.statusCode === 403)) {
      const fallback = parseMailboxListFromEnv();
      if (fallback.length === 0) {
        throw new HttpError(403, 'Geen toegang tot /users en geen MAILBOX_LIST geconfigureerd.');
      }
      return fallback;
    }
    throw e;
  }
}

async function searchMessages(
  accessToken,
  terms,
  mailboxes,
  top = 10,
  requestedScope = 'all',
  requestedExcludeLouvenberg = false,
  requestedYear = 'all',
  requestedOnlyWithAttachments = false
) {
  const unique = new Map();
  const safeTop = Math.min(Math.max(Number(top) || 10, 1), 25);
  const rawScope = String(requestedScope || 'all').trim().toLowerCase();
  const searchScope = ['all', 'subject', 'from', 'body'].includes(rawScope) ? rawScope : 'all';
  const excludeLouvenberg = !!requestedExcludeLouvenberg;
  const onlyWithAttachments = !!requestedOnlyWithAttachments;
  const parsedYear = Number(requestedYear);
  const yearFilter = Number.isInteger(parsedYear) && parsedYear >= 2000 && parsedYear <= 2100
    ? parsedYear : null;

  function buildAqs(term) {
    const escaped = String(term).replace(/"/g, '\\"');
    if (searchScope === 'subject') return `subject:"${escaped}"`;
    if (searchScope === 'from') return `"${escaped}"`;
    if (searchScope === 'body') return `body:"${escaped}"`;
    return `"${escaped}"`;
  }

  for (const rawTerm of terms || []) {
    const term = String(rawTerm || '').trim();
    if (!term) continue;
    const normalizedTerm = term.toLowerCase();
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
          const fromAddress = String(m?.from?.emailAddress?.address || '').toLowerCase().trim();
          const fromName = String(m?.from?.emailAddress?.name || '').toLowerCase().trim();

          if (searchScope === 'from') {
            const senderMatches = fromAddress.includes(normalizedTerm) || fromName.includes(normalizedTerm);
            if (!senderMatches) continue;
          }
          if (onlyWithAttachments && !m.hasAttachments) continue;
          if (yearFilter) {
            const y = new Date(m.receivedDateTime).getFullYear();
            if (y !== yearFilter) continue;
          }
          if (excludeLouvenberg && fromAddress.endsWith('@louvenbergadvies.nl')) continue;
          if (unique.has(m.id)) continue;

          unique.set(m.id, {
            id: m.id,
            subject: m.subject,
            from: m.from?.emailAddress || null,
            receivedDateTime: m.receivedDateTime,
            hasAttachments: !!m.hasAttachments,
            bodyPreview: m.bodyPreview,
            webLink: m.webLink,
            mailbox,
            matchedTerm: term,
            matchedScope: searchScope
          });
        }
      } catch (e) {
        // ignore single-mailbox failures
      }
    }
  }

  return Array.from(unique.values())
    .sort((a, b) => (new Date(b.receivedDateTime) - new Date(a.receivedDateTime)));
}

async function getMessage(accessToken, mailbox, messageId) {
  const safeMailbox = String(mailbox || '').trim();
  const safeMessageId = String(messageId || '').trim();
  if (!safeMailbox || !safeMessageId) throw new HttpError(400, 'mailbox and messageId are required');

  const path = `/users/${encodeURIComponent(safeMailbox)}/messages/${encodeURIComponent(safeMessageId)}`
    + '?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,hasAttachments,webLink';
  return graphGet(accessToken, path);
}

async function listAttachmentsMeta(accessToken, mailbox, messageId) {
  const safeMailbox = String(mailbox || '').trim();
  const safeMessageId = String(messageId || '').trim();
  if (!safeMailbox || !safeMessageId) throw new HttpError(400, 'mailbox and messageId are required');

  const path = `/users/${encodeURIComponent(safeMailbox)}/messages/${encodeURIComponent(safeMessageId)}/attachments`
    + `?$select=id,name,size,contentType`;
  const data = await graphGet(accessToken, path);
  return (data.value || []).map(a => ({
    id: a.id, name: a.name, size: a.size, contentType: a.contentType
  }));
}

async function getAttachment(accessToken, mailbox, messageId, attachmentId) {
  const safeMailbox = String(mailbox || '').trim();
  const safeMessageId = String(messageId || '').trim();
  const safeAttId = String(attachmentId || '').trim();
  if (!safeMailbox || !safeMessageId || !safeAttId) {
    throw new HttpError(400, 'mailbox, messageId and attachmentId are required');
  }

  const basePath = `/users/${encodeURIComponent(safeMailbox)}/messages/${encodeURIComponent(safeMessageId)}/attachments/${encodeURIComponent(safeAttId)}`;
  let detail = await graphGet(accessToken, basePath);
  if (String(detail['@odata.type'] || '') !== '#microsoft.graph.fileAttachment') {
    throw new HttpError(415, 'Attachment is not a fileAttachment');
  }
  if (!detail.contentBytes) {
    const castPath = `${basePath}/microsoft.graph.fileAttachment?$select=id,name,size,contentType,contentBytes`;
    detail = await graphGet(accessToken, castPath);
  }
  return {
    id: detail.id,
    name: detail.name,
    size: detail.size,
    contentType: detail.contentType,
    contentBytes: detail.contentBytes
  };
}

// ---------- MCP tool catalog ----------

const TOOLS = [
  {
    name: 'list_mailboxes',
    description: 'Geef alle beschikbare mailboxen (e-mailadressen) in de tenant terug.',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false }
  },
  {
    name: 'search_messages',
    description:
      'Zoek e-mails in één of meerdere mailboxen via Microsoft Graph $search. ' +
      'Laat "mailboxes" leeg om in alle mailboxen te zoeken. ' +
      'Resultaat bevat id, mailbox, subject, from, receivedDateTime, bodyPreview en webLink.',
    inputSchema: {
      type: 'object',
      properties: {
        terms: {
          type: 'array',
          items: { type: 'string' },
          description: 'Zoektermen (resultaten worden samengevoegd, dedupe op message id).'
        },
        mailboxes: {
          type: 'array',
          items: { type: 'string' },
          description: 'Mailbox e-mailadressen. Laat leeg voor alle mailboxen.'
        },
        top: { type: 'integer', minimum: 1, maximum: 25, default: 10 },
        searchScope: {
          type: 'string',
          enum: ['all', 'subject', 'from', 'body'],
          default: 'all'
        },
        searchYear: {
          type: ['integer', 'string'],
          description: 'Filter op jaar (bv. 2025) of "all".'
        },
        onlyWithAttachments: { type: 'boolean', default: false },
        excludeLouvenberg: { type: 'boolean', default: false }
      },
      required: ['terms'],
      additionalProperties: false
    }
  },
  {
    name: 'get_message',
    description: 'Haal de volledige inhoud (body) en metadata van één bericht op.',
    inputSchema: {
      type: 'object',
      properties: {
        mailbox: { type: 'string', description: 'Mailbox e-mailadres (UPN).' },
        messageId: { type: 'string', description: 'Graph message id.' }
      },
      required: ['mailbox', 'messageId'],
      additionalProperties: false
    }
  },
  {
    name: 'list_attachments',
    description: 'Geef metadata (id, naam, type, grootte) van alle bijlagen van een bericht.',
    inputSchema: {
      type: 'object',
      properties: {
        mailbox: { type: 'string' },
        messageId: { type: 'string' }
      },
      required: ['mailbox', 'messageId'],
      additionalProperties: false
    }
  },
  {
    name: 'get_attachment',
    description:
      'Download één bijlage als base64. Let op: grote bijlagen kunnen het token-budget van Claude snel opmaken; gebruik list_attachments eerst.',
    inputSchema: {
      type: 'object',
      properties: {
        mailbox: { type: 'string' },
        messageId: { type: 'string' },
        attachmentId: { type: 'string' }
      },
      required: ['mailbox', 'messageId', 'attachmentId'],
      additionalProperties: false
    }
  }
];

async function executeTool(name, args) {
  const accessToken = await getAzureAppToken();

  if (name === 'list_mailboxes') {
    return await resolveAllMailboxes(accessToken);
  }
  if (name === 'search_messages') {
    let mailboxes = Array.isArray(args.mailboxes) ? args.mailboxes : [];
    if (mailboxes.length === 0) {
      const all = await resolveAllMailboxes(accessToken);
      mailboxes = all.map(m => m.mail);
    }
    return await searchMessages(
      accessToken,
      args.terms || [],
      mailboxes,
      args.top || 10,
      args.searchScope || 'all',
      args.excludeLouvenberg === true,
      args.searchYear || 'all',
      args.onlyWithAttachments === true
    );
  }
  if (name === 'get_message') {
    return await getMessage(accessToken, args.mailbox, args.messageId);
  }
  if (name === 'list_attachments') {
    return await listAttachmentsMeta(accessToken, args.mailbox, args.messageId);
  }
  if (name === 'get_attachment') {
    return await getAttachment(accessToken, args.mailbox, args.messageId, args.attachmentId);
  }
  throw new HttpError(404, `Unknown tool: ${name}`);
}

// ---------- JSON-RPC / MCP transport ----------

function rpcResult(id, result) { return { jsonrpc: '2.0', id, result }; }
function rpcError(id, code, message) { return { jsonrpc: '2.0', id, error: { code, message } }; }

async function handleRpc(msg) {
  const id = (msg && (msg.id !== undefined)) ? msg.id : null;
  const method = msg && msg.method;
  const params = (msg && msg.params) || {};
  const isNotification = msg && msg.id === undefined;

  try {
    if (method === 'initialize') {
      return rpcResult(id, {
        protocolVersion: PROTOCOL_VERSION,
        capabilities: { tools: { listChanged: false } },
        serverInfo: SERVER_INFO,
        instructions:
          'Microsoft 365 mailbox-zoekserver voor Impossible Drinks. Gebruik search_messages om e-mails te vinden, daarna get_message voor de volledige inhoud.'
      });
    }
    if (method && method.startsWith('notifications/')) {
      return null; // notifications get no response
    }
    if (method === 'ping') return rpcResult(id, {});
    if (method === 'tools/list') return rpcResult(id, { tools: TOOLS });
    if (method === 'tools/call') {
      const name = params.name;
      const args = params.arguments || {};
      try {
        const data = await executeTool(name, args);
        return rpcResult(id, {
          content: [{ type: 'text', text: JSON.stringify(data, null, 2) }],
          isError: false
        });
      } catch (e) {
        return rpcResult(id, {
          content: [{ type: 'text', text: `Tool error: ${e.message || e}` }],
          isError: true
        });
      }
    }

    if (isNotification) return null;
    return rpcError(id, -32601, `Method not found: ${method}`);
  } catch (e) {
    if (isNotification) return null;
    return rpcError(id, -32000, e.message || 'Server error');
  }
}

function extractToken(event) {
  // 1. Authorization: Bearer <token>
  const auth = (event.headers && (event.headers.authorization || event.headers.Authorization)) || '';
  if (auth.startsWith('Bearer ')) return auth.slice(7).trim();

  // 2. Query string ?token=...
  const qs = (event.queryStringParameters && event.queryStringParameters.token) || '';
  if (qs) return String(qs).trim();

  // 3. URL path: /mcp/<token>  ->  function gets path /.netlify/functions/mcp/<token>
  const rawPath = event.path || '';
  const m = rawPath.match(/\/mcp\/([^/?#]+)/i)
        || rawPath.match(/\/functions\/mcp\/([^/?#]+)/i);
  if (m) return decodeURIComponent(m[1]).trim();

  return '';
}

function checkAuth(event) {
  const expected = process.env.MCP_BEARER_TOKEN;
  if (!expected) return; // auth disabled (NOT recommended in production)
  const token = extractToken(event);
  if (!token || token !== expected) {
    throw new HttpError(401, 'Unauthorized: invalid or missing token');
  }
}

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization, Mcp-Session-Id, MCP-Protocol-Version',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS, DELETE',
  'Access-Control-Expose-Headers': 'Mcp-Session-Id'
};

exports.handler = async (event) => {
  const headers = { 'Content-Type': 'application/json', ...CORS };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers, body: '' };
  }

  // GET on the MCP endpoint is used by some clients to open an SSE stream
  // for server-initiated messages. We don't need it (stateless tool server),
  // so respond with 405 per the Streamable HTTP spec.
  if (event.httpMethod === 'GET' || event.httpMethod === 'DELETE') {
    return {
      statusCode: 405,
      headers: { ...headers, Allow: 'POST, OPTIONS' },
      body: JSON.stringify({ error: 'Method not allowed; use POST with a JSON-RPC body.' })
    };
  }

  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  try {
    checkAuth(event);

    let payload;
    try {
      payload = event.body ? JSON.parse(event.body) : null;
    } catch (e) {
      return { statusCode: 400, headers, body: JSON.stringify(rpcError(null, -32700, 'Parse error')) };
    }
    if (!payload) {
      return { statusCode: 400, headers, body: JSON.stringify(rpcError(null, -32600, 'Empty body')) };
    }

    if (Array.isArray(payload)) {
      const results = [];
      for (const m of payload) {
        const r = await handleRpc(m);
        if (r) results.push(r);
      }
      if (results.length === 0) return { statusCode: 202, headers, body: '' };
      return { statusCode: 200, headers, body: JSON.stringify(results) };
    }

    const result = await handleRpc(payload);
    if (!result) return { statusCode: 202, headers, body: '' };
    return { statusCode: 200, headers, body: JSON.stringify(result) };
  } catch (e) {
    const status = e.statusCode || 500;
    return {
      statusCode: status,
      headers,
      body: JSON.stringify(rpcError(null, status === 401 ? -32001 : -32000, e.message || 'Server error'))
    };
  }
};
