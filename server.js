require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const http = require('http');
const WebSocket = require('ws');
const { BotFrameworkAdapter } = require('botbuilder');

console.log('\n=== TEAMS CTI SERVER INITIALIZING ===\n');

const app = express();
const server = http.createServer(app);

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(express.static(path.join(__dirname, '.')));

// =============================================================================
// CONFIG
// =============================================================================
const PORT = process.env.PORT || 3000;
const MS_APP_ID = process.env.MS_APP_ID || '';
const MS_APP_PASSWORD = process.env.MS_APP_PASSWORD || '';
const MS_APP_TENANT_ID = process.env.MS_APP_TENANT_ID || '';
const FUSION_BASE_URL = process.env.FUSION_BASE_URL || '';
const FUSION_USER = process.env.FUSION_USER || '';
const FUSION_PASS = process.env.FUSION_PASS || '';

console.log('[CONFIG] MS_APP_ID        :', MS_APP_ID ? 'SET' : 'MISSING');
console.log('[CONFIG] MS_APP_PASSWORD  :', MS_APP_PASSWORD ? 'SET' : 'MISSING');
console.log('[CONFIG] MS_APP_TENANT_ID :', MS_APP_TENANT_ID ? 'SET' : 'MISSING');
console.log('[CONFIG] FUSION_BASE_URL  :', FUSION_BASE_URL || '(not set)');
console.log('[CONFIG] PORT             :', PORT);

function fusionAuthHeader() {
    return 'Basic ' + Buffer.from(FUSION_USER + ':' + FUSION_PASS).toString('base64');
}

// =============================================================================
// BOT FRAMEWORK ADAPTER
// =============================================================================
const adapter = new BotFrameworkAdapter({
    appId: MS_APP_ID,
    appPassword: MS_APP_PASSWORD
});

adapter.onTurnError = async (context, error) => {
    console.error('[BOT] onTurnError:', error.message);
};

// =============================================================================
// WEBSOCKET SERVER
// =============================================================================
const wss = new WebSocket.Server({ server });
const agentSessions = new Map();

wss.on('connection', (ws, req) => {
    console.log('[WS] New connection from:', req.socket.remoteAddress);
    let registeredAgentId = null;

    ws.on('message', (raw) => {
        try {
            const msg = JSON.parse(raw);
            if (msg.type === 'register' && msg.agentId) {
                registeredAgentId = msg.agentId;
                agentSessions.set(registeredAgentId, ws);
                console.log('[WS] Agent registered:', registeredAgentId);
                ws.send(JSON.stringify({ type: 'registered', agentId: registeredAgentId }));
            }
        } catch (e) {
            console.error('[WS] Message parse error:', e.message);
        }
    });

    ws.on('close', () => {
        if (registeredAgentId) {
            agentSessions.delete(registeredAgentId);
            console.log('[WS] Agent disconnected:', registeredAgentId);
        }
    });

    ws.on('error', (err) => {
        console.error('[WS] Error:', err.message);
    });

    ws.isAlive = true;
    ws.on('pong', () => { ws.isAlive = true; });
});

setInterval(() => {
    wss.clients.forEach((ws) => {
        if (!ws.isAlive) { return ws.terminate(); }
        ws.isAlive = false;
        ws.ping();
    });
}, 30000);

// =============================================================================
// PUSH EVENT TO AGENT
// =============================================================================
function pushToAgent(agentId, payload) {
    const ws = agentSessions.get(agentId);
    if (!ws || ws.readyState !== WebSocket.OPEN) {
        console.warn('[WS] No active session for agent:', agentId, '— event dropped');
        return false;
    }
    ws.send(JSON.stringify(payload));
    console.log('[WS] Pushed to agent:', agentId, '| eventType:', payload.eventType);
    return true;
}

// =============================================================================
// AGENT MAP
// =============================================================================
const AGENT_MAP = {
    '+12146180369': 'agent-default',
};

function resolveAgentId(calledParty) {
    return AGENT_MAP[calledParty] || 'agent-default';
}

// =============================================================================
// ROUTES — all routes BEFORE the 404 handler
// =============================================================================

// -----------------------------------------------------------------------------
// GET /api/calling — Graph subscription validation handshake
// Microsoft sends a GET with ?validationToken= before activating subscription.
// Must echo the token back as plain text with 200.
// -----------------------------------------------------------------------------
app.get('/api/calling', (req, res) => {
    const validationToken = req.query.validationToken;
    if (validationToken) {
        console.log('[CALLING] Graph validation handshake received');
        res.set('Content-Type', 'text/plain');
        return res.status(200).send(validationToken);
    }
    res.json({ status: 'Teams CTI calling endpoint active. POST only.' });
});

// -----------------------------------------------------------------------------
// POST /api/calling — Teams Graph webhook (real call notifications)
// -----------------------------------------------------------------------------
app.post('/api/calling', express.raw({ type: '*/*' }), async (req, res) => {
    console.log('[CALLING] POST /api/calling received');

    // Microsoft validation POST — body is plain text token
    const validationToken = req.query.validationToken;
    if (validationToken) {
        console.log('[CALLING] Validation POST received');
        res.set('Content-Type', 'text/plain');
        return res.status(200).send(validationToken);
    }

    // Parse JSON body manually since we used raw middleware
    let body;
    try {
        body = JSON.parse(req.body.toString());
    } catch (e) {
        console.warn('[CALLING] Body parse error:', e.message);
        return res.status(400).json({ error: 'Invalid JSON body' });
    }

    // Legacy Skype format
    if (!body.value && body['@type']) {
        console.log('[CALLING] Legacy Skype format — sending 204');
        return res.status(204).send();
    }

    if (!body.value || !Array.isArray(body.value)) {
        console.warn('[CALLING] Unexpected payload:', JSON.stringify(body).slice(0, 200));
        return res.status(400).json({ error: 'Unexpected payload' });
    }

    // Respond 202 immediately
    res.status(202).send();

    for (const notification of body.value) {
        await handleGraphNotification(notification);
    }
});

// -----------------------------------------------------------------------------
// GET /api/messages — Bot Framework health check
// -----------------------------------------------------------------------------
app.get('/api/messages', (req, res) => {
    res.json({ status: 'Teams CTI messages endpoint active. POST only.' });
});

// -----------------------------------------------------------------------------
// POST /api/messages — Bot Framework adapter (Teams channel handshake)
// -----------------------------------------------------------------------------
app.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        console.log('[BOT] Activity type:', context.activity.type);
        const text = (context.activity.text || '').trim().toLowerCase();

        // Optional: trigger inbound via Teams bot message for testing
        // Send "incoming +12146180369" to the bot in Teams
        if (text.includes('incoming')) {
            const match = text.match(/(\+?[0-9]{10,15})/);
            const ani = match ? match[1] : '+12146180369';
            const callId = 'bot-msg-' + Date.now();
            console.log('[BOT] Simulating inbound from message:', ani);
            await handleIncomingCall('agent-default', ani, callId);
            await context.sendActivity('CTI triggered for ANI: ' + ani);
        }
    });
});

// -----------------------------------------------------------------------------
// POST /api/call-event — manual test endpoint (Postman simulation)
// -----------------------------------------------------------------------------
app.post('/api/call-event', async (req, res) => {
    const { eventType, callId, callerANI, agentId, direction, extraData } = req.body;

    if (!eventType || !agentId) {
        return res.status(400).json({ error: 'eventType and agentId are required' });
    }

    console.log('[CALL-EVENT] Manual event:', { eventType, callId, callerANI, agentId });

    if (eventType === 'newCommEvent' && callerANI) {
        await handleIncomingCall(agentId, callerANI, callId || ('manual-' + Date.now()));
        return res.json({ success: true, note: 'Full inbound pipeline triggered' });
    }

    const pushed = pushToAgent(agentId, req.body);
    res.json({ success: pushed, agentId, eventType });
});

// -----------------------------------------------------------------------------
// GET /health
// -----------------------------------------------------------------------------
app.get('/health', (req, res) => {
    res.json({
        status: 'ok',
        timestamp: new Date().toISOString(),
        connectedAgents: Array.from(agentSessions.keys())
    });
});

// -----------------------------------------------------------------------------
// GET / — serve iFrame
// -----------------------------------------------------------------------------
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

// -----------------------------------------------------------------------------
// 404 — MUST be last
// -----------------------------------------------------------------------------
app.use((req, res) => {
    console.log('[404]', req.method, req.path);
    res.status(404).json({ error: 'Route not found' });
});

app.use((err, req, res, next) => {
    console.error('[ERROR]', err);
    res.status(500).json({ error: err.message });
});

// =============================================================================
// GRAPH NOTIFICATION HANDLER
// =============================================================================
async function handleGraphNotification(notification) {
    try {
        const resource = notification.resource || '';
        const resourceData = notification.resourceData || {};
        const callId = resourceData.id || resource.split('/').pop();
        const callState = resourceData.state || '';

        console.log('[GRAPH] callId:', callId, '| state:', callState);

        const callerANI = resourceData?.source?.identity?.phone?.id
            || resourceData?.source?.identity?.user?.displayName
            || 'Unknown';

        const calledParty = resourceData?.targets?.[0]?.identity?.phone?.id
            || resourceData?.targets?.[0]?.identity?.user?.id
            || null;

        const agentId = resolveAgentId(calledParty);

        console.log('[GRAPH] ANI:', callerANI, '| calledParty:', calledParty, '| agentId:', agentId);

        switch (callState) {
            case 'incoming':
                await handleIncomingCall(agentId, callerANI, callId);
                break;
            case 'established':
                pushToAgent(agentId, { eventType: 'startCommEvent', callId, callerANI, agentId });
                break;
            case 'terminated':
            case 'disconnected':
                pushToAgent(agentId, { eventType: 'closeCommEvent', callId, reason: 'WRAPUP', agentId });
                break;
            default:
                console.log('[GRAPH] Unhandled call state:', callState);
        }
    } catch (e) {
        console.error('[GRAPH] handleGraphNotification error:', e.message);
    }
}

// =============================================================================
// INCOMING CALL HANDLER
// =============================================================================
async function handleIncomingCall(agentId, callerANI, callId) {
    console.log('[INBOUND] agentId:', agentId, '| ANI:', callerANI, '| callId:', callId);

    // Phase 1: immediate newCommEvent with just ANI
    pushToAgent(agentId, {
        eventType: 'newCommEvent',
        callId,
        callerANI,
        direction: 'ORA_SVC_INBOUND',
        extraData: null,
        agentId
    });

    if (!FUSION_BASE_URL) {
        console.warn('[INBOUND] FUSION_BASE_URL not set — skipping contact lookup');
        return;
    }

    // Phase 2: enrich with contact + SR lookup
    try {
        const contactData = await lookupContact(callerANI);
        if (!contactData.found) {
            console.log('[INBOUND] No contact found for ANI:', callerANI);
            return;
        }

        const { contactId } = contactData;
        const srData = await lookupSR(contactId);

        if (srData.found && srData.openSRs.length === 1) {
            const sr = srData.openSRs[0];
            pushToAgent(agentId, {
                eventType: 'newCommEvent',
                callId,
                callerANI,
                direction: 'ORA_SVC_INBOUND',
                extraData: {
                    SVCMCA_CONTACT_ID: String(contactId),
                    SVCMCA_SR_ID: String(sr.srId),
                    SVCMCA_SR_NUM: String(sr.srNumber)
                },
                agentId
            });
        } else {
            pushToAgent(agentId, {
                eventType: 'newCommEvent',
                callId,
                callerANI,
                direction: 'ORA_SVC_INBOUND',
                extraData: { SVCMCA_CONTACT_ID: String(contactId) },
                agentId
            });
        }
    } catch (e) {
        console.error('[INBOUND] Lookup error:', e.message);
    }
}

// =============================================================================
// CONTACT LOOKUP
// =============================================================================
async function lookupContact(phone) {
    const formats = buildPhoneFormats(phone);
    for (const fmt of formats) {
        try {
            const url = FUSION_BASE_URL +
                '/crmRestApi/resources/latest/contacts' +
                '?q=MobileNumber=' + encodeURIComponent(fmt);
            const response = await fetch(url, {
                headers: { 'Authorization': fusionAuthHeader(), 'Content-Type': 'application/json' }
            });
            if (!response.ok) continue;
            const data = await response.json();
            if (data.items && data.items.length > 0) {
                const contact = data.items[0];
                return {
                    found: true,
                    contactId: contact.PartyId,
                    contactName: contact.ContactName || contact.FormattedName || 'Unknown',
                    phone: contact.MobileNumber || fmt
                };
            }
        } catch (e) {
            console.error('[LOOKUP-CONTACT] Error for format', fmt, ':', e.message);
        }
    }
    return { found: false, phone };
}

// =============================================================================
// SR LOOKUP
// =============================================================================
async function lookupSR(contactId) {
    try {
        const query = 'PrimaryContactPartyId=' + contactId + " AND StatusCd='ORA_SVC_OPEN'";
        const url = FUSION_BASE_URL +
            '/crmRestApi/resources/latest/serviceRequests' +
            '?q=' + encodeURIComponent(query) +
            '&fields=SrId,SrNumber,Title,StatusCd,PrimaryContactPartyId';
        const response = await fetch(url, {
            headers: { 'Authorization': fusionAuthHeader(), 'Content-Type': 'application/json' }
        });
        if (!response.ok) {
            return { found: false, openSRs: [], error: 'HTTP ' + response.status };
        }
        const data = await response.json();
        const OPEN_STATUSES = ['ORA_SVC_OPEN', 'ORA_SVC_NEW', 'ORA_SVC_IN_PROGRESS'];
        const openSRs = (data.items || [])
            .filter(sr => OPEN_STATUSES.includes(sr.StatusCd))
            .map(sr => ({ srId: sr.SrId, srNumber: sr.SrNumber, title: sr.Title || '', status: sr.StatusCd }));
        return { found: openSRs.length > 0, openSRs, count: openSRs.length };
    } catch (e) {
        return { found: false, openSRs: [], error: e.message };
    }
}

// =============================================================================
// HELPER
// =============================================================================
function buildPhoneFormats(raw) {
    const digits = raw.replace(/\D/g, '');
    const last10 = digits.slice(-10);
    const withPlus = raw.startsWith('+') ? raw : '+' + digits;
    let dashed = '';
    if (last10.length === 10) {
        dashed = last10.slice(0, 3) + '-' + last10.slice(3, 6) + '-' + last10.slice(6);
    }
    return [...new Set([raw, withPlus, last10, digits, dashed].filter(Boolean))];
}

// =============================================================================
// START
// =============================================================================
server.listen(PORT, () => {
    console.log('\n=== TEAMS CTI SERVER STARTED ===');
    console.log('Port   : ' + PORT);
    console.log('URL    : http://localhost:' + PORT);
    console.log('Health : http://localhost:' + PORT + '/health');
    console.log('WS     : ws://localhost:' + PORT);
    console.log('================================\n');
});