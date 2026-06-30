// =============================================================================
// STATE
// =============================================================================
let ws = null;
let wsReconnectTimer = null;
let onHold = false;
let muted = false;
let isAvailable = false;
let uef = null;
let phoneContext = null;
let currentEventId = null;
let currentCallerANI = null;

const AGENT_ID = 'agent-default'; // must match the key in server.js AGENT_MAP
const WS_URL = (location.protocol === 'https:' ? 'wss://' : 'ws://') + location.host;

console.log('\n========================================');
console.log('TEAMS CTI APPLICATION LOADING');
console.log('========================================\n');

// =============================================================================
// BOOT
// =============================================================================
async function boot() {
    try {
        console.log('[BOOT] Starting initialization...');
        await Promise.all([initWebSocket(), initUEF()]);
        console.log('[BOOT] Initialization complete\n');
    } catch (err) {
        console.error('[BOOT] Initialization failed:', err);
    }
}
boot();

// =============================================================================
// WEBSOCKET — connects to MCA Adapter Middleware
// Receives call events pushed by server.js: newCommEvent, startCommEvent, closeCommEvent
// =============================================================================
function initWebSocket() {
    return new Promise((resolve) => {
        console.log('[WS] Connecting to:', WS_URL);

        ws = new WebSocket(WS_URL);

        ws.addEventListener('open', () => {
            console.log('[WS] Connected');
            setBadge('ws', 'WS: connected', 'ok');
            log('in', 'WebSocket connected', 'ok');

            // Register this agent session with the server
            ws.send(JSON.stringify({ type: 'register', agentId: AGENT_ID }));

            // Clear any pending reconnect timer
            if (wsReconnectTimer) {
                clearTimeout(wsReconnectTimer);
                wsReconnectTimer = null;
            }

            resolve();
        });

        ws.addEventListener('message', (event) => {
            try {
                const payload = JSON.parse(event.data);
                console.log('[WS] Message received:', payload);
                handleServerEvent(payload);
            } catch (e) {
                console.error('[WS] Message parse error:', e.message);
            }
        });

        ws.addEventListener('close', () => {
            console.warn('[WS] Connection closed — reconnecting in 5s...');
            setBadge('ws', 'WS: reconnecting...', 'err');
            log('in', 'WebSocket closed — reconnecting...', 'err');
            wsReconnectTimer = setTimeout(initWebSocket, 5000);
        });

        ws.addEventListener('error', (err) => {
            console.error('[WS] Error:', err);
            setBadge('ws', 'WS: error', 'err');
        });
    });
}

// =============================================================================
// HANDLE EVENTS PUSHED FROM SERVER
// Server pushes: { eventType, callId, callerANI, direction, extraData, agentId }
// =============================================================================
async function handleServerEvent(payload) {
    const { eventType, callId, callerANI, direction, extraData, reason } = payload;

    switch (eventType) {

        case 'registered':
            log('in', 'Agent session registered: ' + payload.agentId, 'ok');
            break;

        case 'newCommEvent':
            currentEventId = callId;
            currentCallerANI = callerANI;
            setStatus('Incoming: ' + callerANI, 'ringing');
            showInboundControls(true);
            log('in', 'Incoming call: ' + callerANI, 'ok');
            await fireNewCommEvent(callerANI, direction || 'ORA_SVC_INBOUND', extraData, callId);
            break;

        case 'startCommEvent':
            await fireStartCommEvent();
            break;

        case 'closeCommEvent':
            await fireCloseCommEvent(reason || 'WRAPUP');
            resetCallState();
            break;

        default:
            console.warn('[WS] Unknown eventType:', eventType);
    }
}

// =============================================================================
// UEF INIT
// =============================================================================
async function initUEF() {
    try {
        console.log('[UEF] Initializing...');
        if (!window.CX_SVC_UI_EVENTS_FRAMEWORK) {
            throw new Error('UEF framework not loaded — is this running inside Fusion?');
        }
        console.log('[UEF] Framework available');

        uef = await CX_SVC_UI_EVENTS_FRAMEWORK.uiEventsFramework.initialize('MCA_APP', 'v1');
        console.log('[UEF] Framework initialized');

        const mcaContext = await uef.getMultiChannelAdaptorContext();
        phoneContext = await mcaContext.getCommunicationChannelContext('PHONE');
        console.log('[UEF] Phone context retrieved');

        setBadge('uef', 'UEF: connected', 'ok');
        log('in', 'UEF connected', 'ok');
        document.getElementById('availability-toggle').disabled = false;

        subscribeToToolbarInteraction();
        subscribeToOutgoingEvent();

        console.log('[UEF] Initialization complete\n');

    } catch (e) {
        console.error('[UEF] INIT FAILED:', e.message);
        setBadge('uef', 'UEF: failed', 'err');
        log('in', 'UEF init failed: ' + e.message, 'err');
    }
}

// =============================================================================
// UEF SUBSCRIPTIONS (Fusion toolbar -> iFrame)
// =============================================================================
function subscribeToToolbarInteraction() {
    try {
        const req = uef.requestHelper.createSubscriptionRequest('onToolbarInteractionCommand');
        phoneContext.subscribe(req, async (eventResponse) => {
            const command = eventResponse.getResponseData().getCommand();
            console.log('[UEF-EVENT] Toolbar command:', command);
            log('in', 'Fusion command: ' + command, 'cmd');

            switch (command) {
                case 'accept':
                    // Agent accepted in Fusion toolbar
                    // Teams handles the actual audio — we just advance the UEF state
                    showInboundControls(false);
                    showActiveControls(true);
                    setStatus('On Call', 'oncall');
                    await fireStartCommEvent();
                    break;

                case 'reject':
                    showInboundControls(false);
                    setStatus('Call rejected', 'error');
                    await fireCloseCommEvent('REJECT');
                    resetCallState();
                    break;

                case 'disconnect':
                    showActiveControls(false);
                    await fireCloseCommEvent('WRAPUP');
                    resetCallState();
                    break;

                case 'hold':
                    if (!onHold) toggleHold();
                    break;

                case 'unhold':
                    if (onHold) toggleHold();
                    break;

                case 'mute':
                    if (!muted) toggleMute();
                    break;

                case 'unmute':
                    if (muted) toggleMute();
                    break;

                default:
                    console.warn('[UEF-CMD] Unknown command:', command);
            }
        });
        log('in', 'Subscribed to toolbar commands', 'ok');
        console.log('[UEF-SUB] onToolbarInteractionCommand subscribed');
    } catch (e) {
        console.error('[UEF-SUB] Subscribe to toolbar failed:', e.message);
    }
}

function subscribeToOutgoingEvent() {
    try {
        const req = uef.requestHelper.createSubscriptionRequest('onOutgoingEvent');
        req.setAppClassification('ORA_SERVICE');
        phoneContext.subscribe(req, async (eventResponse) => {
            const outData = eventResponse.getResponseData().getOutData();
            const phoneNumber = outData.SVCMCA_ANI || outData.phoneNumber || '';
            const callId = outData.SVCMCA_CALL_ID || ('out-' + Date.now());
            console.log('[UEF-EVENT] onOutgoingEvent:', { phoneNumber, callId });
            log('in', 'Click-to-Dial: ' + phoneNumber, 'cmd');

            // In Teams integration, outbound dial intent is logged but Teams handles the actual call.
            // Optionally POST to server to log or trigger a Graph call initiation.
            currentEventId = callId;
            currentCallerANI = phoneNumber;
            showActiveControls(true);
            setStatus('Dialing (Teams): ' + phoneNumber, 'ringing');

            await fireNewCommEvent(phoneNumber, 'ORA_SVC_OUTBOUND', outData, callId);
        });
        log('in', 'Subscribed to outgoing events', 'ok');
        console.log('[UEF-SUB] onOutgoingEvent subscribed');
    } catch (e) {
        console.error('[UEF-SUB] Subscribe to outgoing failed:', e.message);
    }
}

// =============================================================================
// UEF ACTIONS (iFrame -> Fusion)
// =============================================================================
async function setAgentAvailability(available) {
    if (!phoneContext) return;
    try {
        const req = uef.requestHelper.createPublishRequest('agentStateEventOperation');
        req.setEventId('1');
        req.setIsAvailable(available);
        req.setIsLoggedIn(available);
        req.setState(available ? 'AVAILABLE' : 'UNAVAILABLE');
        req.setStateDisplayString('Idle');
        req.setReason('');
        req.setReasonDisplayString('Idle');
        req.setInData({ phoneLineId: '1' });
        await phoneContext.publish(req);
        log('out', 'agentState: ' + (available ? 'AVAILABLE' : 'UNAVAILABLE'), 'ok');
        console.log('[UEF-ACTION] agentState published:', available ? 'AVAILABLE' : 'UNAVAILABLE');
    } catch (e) {
        console.error('[UEF-ACTION] setAgentAvailability failed:', e.message);
        log('out', 'agentStateEvent failed: ' + e.message, 'err');
    }
}

async function fireNewCommEvent(ani, direction, extraData, eventId) {
    if (!phoneContext) {
        console.warn('[UEF-ACTION] fireNewCommEvent skipped — phoneContext not ready');
        return;
    }
    try {
        console.log('[UEF-ACTION] Publishing newCommEvent:', { ani, direction, extraData, eventId });
        const req = uef.requestHelper.createPublishRequest('newCommEvent');
        req.setEventId(eventId || currentEventId || 'call-001');
        req.getInData().setInDataValueByAttribute('SVCMCA_ANI', ani);
        req.getInData().setInDataValueByAttribute('SVCMCA_COMMUNICATION_DIRECTION', direction);
        if (extraData) {
            for (const key in extraData) {
                req.getInData().setInDataValueByAttribute(key, extraData[key]);
            }
        }
        req.setAppClassification('ORA_SERVICE');
        const res = await phoneContext.publish(req);
        try {
            const contactName = res.getResponseData().getData()['SVCMCA_CONTACT_NAME'] || 'Unknown';
            console.log('[UEF-ACTION] newCommEvent matched:', contactName);
            log('out', 'newCommEvent - matched: ' + contactName, 'ok');
            setStatus('Popup shown - ' + contactName, 'ringing');
        } catch (_) {
            log('out', 'newCommEvent fired', 'ok');
        }
        return res;
    } catch (e) {
        console.error('[UEF-ACTION] fireNewCommEvent failed:', e.message);
        log('out', 'newCommEvent failed: ' + e.message, 'err');
    }
}

async function fireStartCommEvent() {
    if (!phoneContext) return;
    try {
        console.log('[UEF-ACTION] Publishing startCommEvent, eventId:', currentEventId);
        const req = uef.requestHelper.createPublishRequest('startCommEvent');
        req.setAppClassification('ORA_SERVICE');
        req.setEventId(currentEventId || 'call-001');
        const res = await phoneContext.publish(req);
        const contactName = res.getResponseData().getData()['SVCMCA_CONTACT_NAME'] || 'Unknown';
        console.log('[UEF-ACTION] startCommEvent - screen pop:', contactName);
        log('out', 'startCommEvent - screen pop: ' + contactName, 'ok');
        setStatus('On Call - ' + contactName, 'oncall');
        showInboundControls(false);
        showActiveControls(true);
        return res;
    } catch (e) {
        console.error('[UEF-ACTION] fireStartCommEvent failed:', e.message);
        log('out', 'startCommEvent failed: ' + e.message, 'err');
    }
}

async function fireCloseCommEvent(reason) {
    if (!phoneContext) return;
    try {
        console.log('[UEF-ACTION] Publishing closeCommEvent, reason:', reason, 'eventId:', currentEventId);
        const req = uef.requestHelper.createPublishRequest('closeCommEvent');
        req.setAppClassification('ORA_SERVICE');
        req.setReason(reason);
        req.setEventId(currentEventId || 'call-001');
        await phoneContext.publish(req);
        log('out', 'closeCommEvent: ' + reason, 'ok');
        console.log('[UEF-ACTION] closeCommEvent published:', reason);
    } catch (e) {
        console.error('[UEF-ACTION] fireCloseCommEvent failed:', e.message);
        log('out', 'closeCommEvent failed: ' + e.message, 'err');
    }
}

// =============================================================================
// CALL STATE HELPERS
// In Teams integration, Teams handles audio.
// Hold/mute are signalling-only — you can extend these to call Graph API
// (PATCH /communications/calls/{callId}) if you need server-side hold.
// =============================================================================
function toggleHold() {
    onHold = !onHold;
    document.getElementById('btn-hold').textContent = onHold ? 'Unhold' : 'Hold';
    log('out', 'hold: ' + onHold + ' (Teams handles audio)', 'ok');
    console.log('[CALL-ACTION] Hold toggled:', onHold);
    // TODO: PATCH /communications/calls/{callId} via server if server-side hold needed
}

function toggleMute() {
    muted = !muted;
    document.getElementById('btn-mute').textContent = muted ? 'Unmute' : 'Mute';
    log('out', 'mute: ' + muted + ' (Teams handles audio)', 'ok');
    console.log('[CALL-ACTION] Mute toggled:', muted);
    // TODO: PATCH /communications/calls/{callId} via server if server-side mute needed
}

function resetCallState() {
    currentEventId = null;
    currentCallerANI = null;
    onHold = false;
    muted = false;
    showInboundControls(false);
    showActiveControls(false);
    const btnHold = document.getElementById('btn-hold');
    const btnMute = document.getElementById('btn-mute');
    if (btnHold) btnHold.textContent = 'Hold';
    if (btnMute) btnMute.textContent = 'Mute';
    setStatus(isAvailable ? 'Ready' : 'Unavailable', isAvailable ? 'ready' : '');
}

// =============================================================================
// SIMULATE INBOUND — for testing without Teams
// Sends a fake newCommEvent payload directly to the server's /api/call-event endpoint.
// The server runs the full inbound pipeline (contact + SR lookup) and pushes back via WS.
// =============================================================================
async function simulateInbound() {
    const ani = document.getElementById('phone-input').value.trim() || '+12146180369';
    const callId = 'sim-' + Date.now();
    log('out', 'Simulating inbound: ' + ani, 'ok');
    console.log('[TEST] Simulating inbound:', ani);

    try {
        const resp = await fetch('/api/call-event', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                eventType: 'newCommEvent',
                callId,
                callerANI: ani,
                agentId: AGENT_ID,
                direction: 'ORA_SVC_INBOUND',
                extraData: null
            })
        });
        const data = await resp.json();
        console.log('[TEST] Simulate response:', data);
    } catch (e) {
        log('out', 'Simulate error: ' + e.message, 'err');
        console.error('[TEST] Simulate error:', e.message);
    }
}

// =============================================================================
// AGENT AVAILABILITY TOGGLE
// =============================================================================
async function toggleAvailability() {
    isAvailable = document.getElementById('availability-toggle').checked;
    console.log('[UI] Toggling availability:', isAvailable ? 'AVAILABLE' : 'UNAVAILABLE');
    await setAgentAvailability(isAvailable);

    const avatar = document.getElementById('agent-avatar');
    const state = document.getElementById('agent-state');
    const label = document.getElementById('toggle-label');
    const btnSim = document.getElementById('btn-simulate');

    if (isAvailable) {
        avatar.className = 'agent-avatar available';
        state.textContent = 'Available';
        state.className = 'agent-state available';
        label.textContent = 'Go Unavailable';
        if (btnSim) btnSim.disabled = false;
        setStatus('Ready - waiting for calls', 'ready');
    } else {
        avatar.className = 'agent-avatar';
        state.textContent = 'Unavailable';
        state.className = 'agent-state unavailable';
        label.textContent = 'Go Available';
        if (btnSim) btnSim.disabled = true;
        setStatus('Unavailable', '');
    }
}

// =============================================================================
// UI HELPERS
// =============================================================================
function setStatus(text, type) {
    document.getElementById('status-text').textContent = text;
    document.getElementById('status-dot').className = 'status-dot ' + (type || '');
}
function showInboundControls(show) {
    document.getElementById('inbound-controls').style.display = show ? 'flex' : 'none';
}
function showActiveControls(show) {
    document.getElementById('active-controls').style.display = show ? 'flex' : 'none';
}
function setBadge(type, text, state) {
    const el = document.getElementById('badge-' + type);
    const footer = document.getElementById('footer-' + type);
    if (el) { el.textContent = text; el.className = 'badge ' + (state || ''); }
    if (footer) { footer.textContent = text; footer.className = 'footer-badge ' + (state || ''); }
}
function log(dir, msg, type) {
    const logEl = document.getElementById('log');
    const row = document.createElement('div');
    row.className = 'log-row';
    const now = new Date();
    const time = now.toTimeString().slice(0, 8);
    row.innerHTML =
        '<span class="log-time">' + time + '</span>' +
        '<span class="log-dir ' + dir + '">' + (dir === 'in' ? 'IN' : 'OUT') + '</span>' +
        '<span class="log-msg ' + (type || '') + '">' + msg + '</span>';
    logEl.appendChild(row);
    logEl.scrollTop = logEl.scrollHeight;
}
function clearLog() {
    document.getElementById('log').innerHTML = '';
}