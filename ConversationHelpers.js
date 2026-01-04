/**
 * Conversation Helpers
 * Single canonical implementation for Conversation_History append + dedupe + last-message summary fields.
 *
 * Canonical truth = Conversation_History (append-only).
 * Derived truth is computed elsewhere after append (GeminiAI.analyzeConversationAndUpdateState).
 */

/**
 * Append a message/event to Conversation_History with dedupe + size guardrails.
 *
 * Expected message shape (minimum):
 * {
 *   id?: string,
 *   messageId?: string,
 *   timestamp?: string (ISO),
 *   senderEmail?: string,
 *   senderName?: string,
 *   type?: string, // boss_message / employee_reply / system_override / etc.
 *   content?: string,
 *   metadata?: Object
 * }
 */
function appendToConversationHistory(taskId, message) {
  try {
    const task = getTask(taskId);
    if (!task) return false;

    const normalized = _normalizeConversationEvent_(message);
    normalized.actor = _inferConversationActor_(task, normalized);

    // Parse existing conversation history
    let history = [];
    if (task.Conversation_History) {
      try {
        history = JSON.parse(task.Conversation_History);
        if (!Array.isArray(history)) history = [];
      } catch (e) {
        history = [];
      }
    }

    // Dedupe primarily by messageId (stable when available)
    const msgId = normalized.messageId;
    const isDuplicate = history.some(m => {
      if (!m) return false;
      if (msgId && (m.messageId === msgId || m.id === msgId)) return true;
      // Secondary heuristic when messageId is missing: same content within 1s
      if (normalized.content && m.content === normalized.content) {
        const t1 = new Date(m.timestamp || 0).getTime();
        const t2 = new Date(normalized.timestamp || 0).getTime();
        if (t1 && t2 && Math.abs(t1 - t2) < 1000) return true;
      }
      return false;
    });

    if (isDuplicate) {
      Logger.log(`Duplicate conversation event detected for task ${taskId}, skipping (messageId=${msgId || 'none'})`);
      return false;
    }

    history.push(normalized);

    // Keep a bounded number of events (also helps Sheets cell limits)
    const MAX_EVENTS = 50;
    if (history.length > MAX_EVENTS) {
      history = history.slice(-MAX_EVENTS);
    }

    // Enforce Sheets cell size limit (~50k chars); keep buffer
    let historyJson = JSON.stringify(history);
    const MAX_CELL_CHARS = 45000;
    if (historyJson.length > MAX_CELL_CHARS) {
      // Drop older events first
      history = history.slice(-20);
      historyJson = JSON.stringify(history);

      // As a last resort truncate individual content fields
      if (historyJson.length > MAX_CELL_CHARS) {
        history.forEach(ev => {
          if (ev && ev.content && ev.content.length > 500) {
            ev.content = ev.content.substring(0, 500) + '... [truncated]';
          }
        });
        historyJson = JSON.stringify(history);
      }
    }

    const lastSnippet = _makeSnippet_(normalized.content || '');
    const lastSenderEmail = normalized.senderEmail || '';
    const lastSender = lastSenderEmail || normalized.senderName || normalized.type || 'unknown';
    const lastActor = normalized.actor || 'system';

    const updateData = {
      Conversation_History: historyJson,
      Message_Count: history.length,
      // Last-message summary for fast Task Card rendering
      Last_Message_Timestamp: normalized.timestamp,
      Last_Message_Sender: lastSender,
      Last_Message_Snippet: lastSnippet,
      // Canonical last-message identity fields (avoid name heuristics)
      Last_Message_Sender_Email: lastSenderEmail,
      Last_Message_Actor: lastActor,
      Last_Updated: new Date()
    };

    updateTask(taskId, updateData);
    return true;
  } catch (error) {
    Logger.log(`Error appending to conversation history for task ${taskId}: ${error.toString()}`);
    return false;
  }
}

/**
 * Append a system event to the canonical Conversation_History.
 * Used to replace legacy Interaction_Log for internal/audit events.
 */
function appendSystemEvent(taskId, type, content, metadata) {
  const nowIso = new Date().toISOString();

  // Normalize type to always start with "system_" so actor inference + UI filtering are consistent.
  const rawType = String(type || 'event');
  const normalizedType = rawType.indexOf('system_') === 0 ? rawType : `system_${rawType}`;

  // Stable messageId is critical: prevents duplicate system/audit rows when triggers retry/double-fire.
  // Prefer correlating to a Gmail message id when provided by the caller.
  const md = metadata || {};
  const correlationId = md.messageId || md.gmailMessageId || md.sentMessageId || null;
  const stableMessageId = correlationId
    ? `${String(correlationId)}:${normalizedType}`
    : `system_${Date.now()}_${Math.floor(Math.random() * 1e6)}`;

  return appendToConversationHistory(taskId, {
    id: stableMessageId,
    messageId: stableMessageId,
    timestamp: nowIso,
    senderEmail: '',
    senderName: 'System',
    type: normalizedType,
    content: content || '',
    metadata: md
  });
}

function _normalizeConversationEvent_(message) {
  const nowIso = new Date().toISOString();
  const msg = message || {};

  const messageId = msg.messageId || msg.id || `local_${Date.now()}`;
  const timestamp = msg.timestamp || nowIso;

  const normalized = {
    id: msg.id || messageId,
    messageId: messageId,
    timestamp: timestamp,
    senderEmail: msg.senderEmail || '',
    senderName: msg.senderName || '',
    type: msg.type || msg.role || 'system',
    content: msg.content || '',
    metadata: msg.metadata || {}
  };
  
  // Store raw content separately if provided (for debugging truncation issues)
  // This allows us to compare what was received vs what was cleaned
  if (msg.rawContent && msg.rawContent !== msg.content) {
    normalized.rawContent = msg.rawContent;
    // Add flag indicating content was cleaned
    normalized.metadata = normalized.metadata || {};
    normalized.metadata.wasContentCleaned = true;
    normalized.metadata.originalLength = msg.rawContent.length;
    normalized.metadata.cleanedLength = (msg.content || '').length;
  }
  
  return normalized;
}

/**
 * Infer canonical actor (boss|employee|system) for a conversation event.
 * Root-cause fix: never rely on display names to decide boss vs employee.
 */
function _inferConversationActor_(task, normalized) {
  const type = (normalized.type || '').toString().toLowerCase();
  const senderEmail = (normalized.senderEmail || '').toString().trim().toLowerCase();
  const assigneeEmail = (task.Assignee_Email || '').toString().trim().toLowerCase();
  const bossEmail = (CONFIG.BOSS_EMAIL ? CONFIG.BOSS_EMAIL() : '') || '';
  const bossEmailNorm = bossEmail.toString().trim().toLowerCase();

  // System events (explicit)
  if (type.indexOf('system') !== -1 || type.indexOf('ai_') === 0) return 'system';

  // Explicit message types (strongest signal)
  if (type.indexOf('boss') !== -1) return 'boss';
  if (type.indexOf('employee') !== -1) return 'employee';

  // Email identity (canonical)
  if (assigneeEmail && senderEmail && senderEmail === assigneeEmail) return 'employee';
  if (bossEmailNorm && senderEmail && senderEmail === bossEmailNorm) return 'boss';

  // If we can't tell, fall back safely based on presence of assignee email:
  // - If the task has an assignee, any non-assignee sender is treated as boss.
  // - Otherwise default to system.
  if (assigneeEmail && senderEmail && senderEmail !== assigneeEmail) return 'boss';
  return 'system';
}

function _makeSnippet_(text) {
  if (!text) return '';
  const clean = String(text).replace(/\s+/g, ' ').trim();
  return clean.length > 160 ? clean.substring(0, 157) + '...' : clean;
}

// NOTE: New-system assumption: we do NOT backfill/migrate old tasks.
// Canonical last-message identity fields are always written on new messages via appendToConversationHistory().


