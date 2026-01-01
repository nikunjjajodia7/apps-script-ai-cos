/**
 * Email Negotiation Engine
 * Handles task assignment emails and reply processing
 */

/**
 * Canonical inbound message ingestion pipeline.
 * Correlates message -> task via Primary_Thread_ID (preferred) or strict "Task ID:" match.
 * Ensures Processed_Message_IDs is used for idempotency.
 */
function ingestInboundMessage(message) {
  try {
    if (!message) return { success: false, error: 'No message provided' };

    const messageId = message.getId();
    const threadId = message.getThread ? message.getThread().getId() : null;
    const fromEmail = extractEmailFromString(message.getFrom());
    const bossEmail = CONFIG.BOSS_EMAIL();

    // Try correlation by Primary_Thread_ID first
    let taskId = null;
    if (threadId) {
      const matches = findRowsByCondition(SHEETS.TASKS_DB, t => (t.Primary_Thread_ID || '') === threadId);
      if (matches && matches.length > 0) {
        taskId = matches[0].Task_ID;
      }
    }

    // Fallback: strict Task ID match in body
    if (!taskId) {
      const body = message.getPlainBody ? message.getPlainBody() : '';
      const match = body.match(/Task ID:\s*([A-Za-z0-9_-]+)/i);
      if (match && match[1]) {
        taskId = match[1].trim();
      }
    }

    if (!taskId) {
      Logger.log(`ingestInboundMessage: Could not correlate message ${messageId} (threadId=${threadId}) to a task`);
      return { success: false, error: 'Could not correlate message to task' };
    }

    // Idempotency via Processed_Message_IDs
    const processedIds = getProcessedMessageIds(taskId);
    if (processedIds.has(messageId)) {
      Logger.log(`ingestInboundMessage: Message ${messageId} already processed for ${taskId}, skipping`);
      return { success: true, skipped: true, taskId: taskId };
    }

    const task = getTask(taskId);
    if (!task) return { success: false, error: `Task ${taskId} not found` };
    const assigneeEmail = extractEmailFromString(task.Assignee_Email);

    const isBoss = fromEmail && bossEmail && fromEmail.toLowerCase() === bossEmail.toLowerCase();
    const isEmployee = fromEmail && assigneeEmail && fromEmail.toLowerCase() === assigneeEmail.toLowerCase();
    if (!isBoss && !isEmployee) {
      Logger.log(`ingestInboundMessage: Unknown sender ${fromEmail} for task ${taskId}, skipping`);
      return { success: false, error: 'Unknown sender' };
    }

    // Process
    if (isBoss) {
      processBossMessage(taskId, message);
    } else {
      processReplyEmail(taskId, message);
    }

    // Mark processed and run deterministic post-step analysis
    markMessageAsProcessed(taskId, messageId);
    try {
      analyzeConversationAndUpdateState(taskId);
    } catch (e) {
      Logger.log(`ingestInboundMessage: post-step analysis failed for ${taskId}: ${e.toString()}`);
    }

    return { success: true, taskId: taskId, messageId: messageId, threadId: threadId };
  } catch (error) {
    Logger.log(`ingestInboundMessage error: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Process a message from the boss
 * Analyzes for date changes, instructions, and automatically updates task
 */
function processBossMessage(taskId, message) {
  try {
    Logger.log(`=== Processing boss message for task: ${taskId} ===`);
    
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`ERROR: Task ${taskId} not found`);
      return;
    }
    
    // Get email content
    const rawEmailContent = message.getPlainBody();
    const emailContent = cleanEmailContent(rawEmailContent);
    const messageId = message.getId();
    
    Logger.log(`Boss message preview: ${emailContent.substring(0, 200)}...`);
    
    // Check if we've already processed this message
    const log = task.Interaction_Log || '';
    if (log.includes(`Message ID: ${messageId}`)) {
      Logger.log(`Message ${messageId} already processed, skipping`);
      return;
    }
    
    // Analyze boss message for date changes and instructions
    const analysis = analyzeBossMessageForDateChange(taskId, emailContent);
    
    Logger.log(`Boss message analysis:`);
    Logger.log(`  Contains date change: ${analysis.hasDateChange}`);
    Logger.log(`  New date: ${analysis.newDate || 'None'}`);
    Logger.log(`  Requires employee approval: ${analysis.requiresApproval}`);
    Logger.log(`  Reasoning: ${analysis.reasoning}`);
    
    // Store boss message in conversation history
    appendToConversationHistory(taskId, {
      id: messageId,
      timestamp: new Date().toISOString(),
      senderEmail: CONFIG.BOSS_EMAIL(),
      senderName: 'Boss',
      type: 'boss_message',
      content: emailContent,
      messageId: messageId,
      metadata: {
        hasDateChange: analysis.hasDateChange,
        newDate: analysis.newDate,
        messageType: analysis.messageType
      }
    });
    
    // Update Last_Boss_Message timestamp
    updateTask(taskId, {
      Last_Boss_Message: new Date().toISOString(),
      Last_Updated: new Date()
    });
    
    // If boss specified a new date, update task and request employee approval
    if (analysis.hasDateChange && analysis.newDate) {
      handleBossDateChange(taskId, analysis.newDate, emailContent, messageId);
    } else {
      // Just log the boss message
      logInteraction(taskId, `Boss message received: ${emailContent.substring(0, 100)}... (Message ID: ${messageId})`);

      // Ensure Conversation_State/Pending_Changes stay in sync for dashboard task cards.
      // Boss messages can change scope/plan without a date change; those should still reflect on the card.
      try {
        Logger.log(`Analyzing conversation to update state after boss message for task ${taskId}...`);
        const analysisResult = analyzeConversationAndUpdateState(taskId);
        if (analysisResult && analysisResult.success) {
          Logger.log(`Conversation state updated to: ${analysisResult.data.conversationState}`);
        } else {
          Logger.log(`Warning: Failed to analyze conversation after boss message: ${(analysisResult && analysisResult.error) || 'unknown error'}`);
        }
      } catch (analysisError) {
        Logger.log(`Error analyzing conversation after boss message: ${analysisError.toString()}`);
        // Don't fail boss message processing if analysis fails
      }
    }
    
  } catch (error) {
    Logger.log(`ERROR processing boss message: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
    logError(ERROR_TYPE.API_ERROR, 'processBossMessage', error.toString(), taskId, error.stack);
  }
}

/**
 * Analyze boss message to detect date changes
 * Uses AI to understand context and extract dates
 */
function analyzeBossMessageForDateChange(taskId, messageContent) {
  try {
    const task = getTask(taskId);
    const bossEmail = CONFIG.BOSS_EMAIL();
    
    // Build context about the task
    const currentDueDate = task.Due_Date ? 
      Utilities.formatDate(new Date(task.Due_Date), Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy') : 
      'Not set';
    const proposedDate = task.Proposed_Date ? 
      Utilities.formatDate(new Date(task.Proposed_Date), Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy') : 
      'None';
    
    // Get conversation history to understand context
    let conversationContext = '';
    if (task.Conversation_History) {
      try {
        const history = JSON.parse(task.Conversation_History);
        const recentMessages = history.slice(-3); // Last 3 messages for context
        conversationContext = recentMessages.map(m => 
          `${m.senderName || m.senderEmail}: ${m.content.substring(0, 200)}`
        ).join('\n\n');
      } catch (e) {
        Logger.log('Could not parse conversation history');
      }
    }
    
    const prompt = `You are analyzing a message from a boss to an employee about a task deadline. The boss may be:
1. Rejecting an employee's date change request
2. Proposing a new date (possibly one the employee suggested earlier)
3. Requesting an earlier deadline
4. Just providing general feedback

Task Details:
- Task Name: "${task.Task_Name}"
- Current Due Date: ${currentDueDate}
- Employee Proposed Date: ${proposedDate}
- Task Status: ${task.Status}

${conversationContext ? `Recent Conversation Context:\n${conversationContext}\n\n` : ''}

Boss's Message:
"${messageContent}"

Analyze this message and determine:
1. Does the boss mention a specific new date/deadline?
2. Is the boss requesting the employee to approve/confirm a date?
3. Is this a date change instruction that should be applied to the task?

IMPORTANT:
- Extract the NEW date the boss wants (not the current date)
- Look for phrases like: "I need it by [date]", "deadline is [date]", "by [date]", "earlier by [date]", "use [date]"
- If boss mentions a date that was previously suggested by employee, that's the date to extract
- Convert all date formats to YYYY-MM-DD

Return ONLY a valid JSON object:
{
  "hasDateChange": true/false,
  "newDate": "YYYY-MM-DD or null",
  "requiresApproval": true/false,
  "reasoning": "explanation of what the boss is requesting",
  "messageType": "DATE_CHANGE|REJECTION|APPROVAL|GENERAL"
}`;

    const response = callGeminiPro(prompt, { temperature: 0.3 });
    let jsonText = response.trim();
    
    // Clean JSON response
    if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```json\n?/, '').replace(/```$/, '');
      jsonText = jsonText.replace(/^```\n?/, '').replace(/```$/, '');
    }
    
    const jsonMatch = jsonText.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      jsonText = jsonMatch[0];
    }
    
    const analysis = JSON.parse(jsonText);
    
    // Fallback: Try manual date extraction if AI didn't find one
    if (analysis.hasDateChange && !analysis.newDate) {
      const extractedDate = extractDateFromText(messageContent);
      if (extractedDate) {
        Logger.log(`Manual date extraction found: ${extractedDate}`);
        analysis.newDate = extractedDate;
      }
    }
    
    return analysis;
    
  } catch (error) {
    Logger.log(`Error analyzing boss message: ${error.toString()}`);
    // Fallback: Try simple date extraction
    const extractedDate = extractDateFromText(messageContent);
    return {
      hasDateChange: !!extractedDate,
      newDate: extractedDate,
      requiresApproval: true,
      reasoning: 'Date detected in message',
      messageType: 'DATE_CHANGE'
    };
  }
}

/**
 * Handle boss's date change instruction
 * Updates task due date and sends approval request to employee
 */
function handleBossDateChange(taskId, newDate, bossMessage, messageId) {
  try {
    Logger.log(`=== Handling boss date change for task ${taskId} ===`);
    Logger.log(`New date: ${newDate}`);
    
    const task = getTask(taskId);
    if (!task || !task.Assignee_Email) {
      Logger.log(`ERROR: Task or assignee not found`);
      return;
    }
    
    // Format dates for display
    const tz = Session.getScriptTimeZone();
    const oldDateFormatted = task.Due_Date ? 
      Utilities.formatDate(new Date(task.Due_Date), tz, 'EEEE, MMMM d, yyyy') : 
      'Not set';
    const newDateFormatted = Utilities.formatDate(new Date(newDate), tz, 'EEEE, MMMM d, yyyy');
    
    // Boss-proposed date: NOT final until employee confirms.
    // Record a structured Pending_Decision so employee "ACCEPTANCE" can deterministically confirm/apply.
    const pendingDecision = {
      type: 'date_change',
      parameter: 'dueDate',
      currentValue: task.Due_Date || null,
      proposedValue: newDate,
      requestedBy: 'boss',
      awaitingFrom: 'employee',
      messageId: messageId,
      createdAt: new Date().toISOString()
    };
    
    // Update task with new proposed date (not final until employee confirms)
    updateTask(taskId, {
      Proposed_Date: newDate,
      // Keep lifecycle status as-is (unless this is first response)
      Status: task.Status === TASK_STATUS.NOT_ACTIVE ? TASK_STATUS.ON_TIME : task.Status,
      // Await explicit assignee confirmation; do NOT treat as resolved
      Conversation_State: CONVERSATION_STATE.AWAITING_EMPLOYEE,
      Pending_Decision: JSON.stringify(pendingDecision),
      Pending_Changes: JSON.stringify([{
        id: 'date_change_' + Date.now(),
        parameter: 'dueDate',
        currentValue: task.Due_Date,
        proposedValue: newDate,
        requestedBy: 'boss',
        awaitingFrom: 'employee',
        reasoning: bossMessage || 'Boss proposed new date'
      }]),
      Last_Updated: new Date()
    });
    
    Logger.log(`Task ${taskId} proposed date updated to ${newDate}`);
    logInteraction(taskId, `Boss changed proposed date from ${oldDateFormatted} to ${newDateFormatted} (Message ID: ${messageId})`);
    
    // Send approval request email to employee
    sendEmployeeApprovalRequest(taskId, newDateFormatted, bossMessage);
    
  } catch (error) {
    Logger.log(`ERROR in handleBossDateChange: ${error.toString()}`);
    logError(ERROR_TYPE.API_ERROR, 'handleBossDateChange', error.toString(), taskId, error.stack);
  }
}

/**
 * Append message to conversation history
 */
function appendToConversationHistory(taskId, message) {
  try {
    const task = getTask(taskId);
    let history = [];
    
    // Parse existing conversation history
    if (task.Conversation_History) {
      try {
        history = JSON.parse(task.Conversation_History);
      } catch (e) {
        Logger.log('Error parsing conversation history: ' + e.toString());
      }
    }
    
    // Check for duplicates
    const isDuplicate = history.some(m => 
      m.messageId === message.messageId || 
      (m.content === message.content && Math.abs(new Date(m.timestamp) - new Date(message.timestamp)) < 1000)
    );
    
    if (!isDuplicate) {
      history.push(message);
      
      // Keep only last 30 messages to avoid storage bloat AND cell size limits
      if (history.length > 30) {
        history = history.slice(-30);
        Logger.log(`Conversation history truncated to last 30 messages`);
      }
      
      // Check size before updating (Google Sheets limit is 50,000 characters)
      let historyJson = JSON.stringify(history);
      if (historyJson.length > 45000) { // Leave buffer
        // Truncate further - keep only last 20 messages
        history = history.slice(-20);
        Logger.log(`Conversation history truncated to last 20 messages due to size limit`);
        
        historyJson = JSON.stringify(history);
        
        // Also truncate individual message content if still too large
        if (historyJson.length > 45000) {
          history.forEach(msg => {
            if (msg.content && msg.content.length > 500) {
              msg.content = msg.content.substring(0, 500) + '... [truncated]';
            }
          });
          historyJson = JSON.stringify(history);
        }
      }
      
      // Update message count and timestamps
      const updateData = {
        Conversation_History: historyJson,
        Message_Count: history.length,
        Last_Updated: new Date()
      };
      
      // Update appropriate timestamp
      if (message.type === 'boss_message') {
        updateData.Last_Boss_Message = message.timestamp;
      } else if (message.type === 'email_reply' || message.type === 'employee_reply') {
        updateData.Last_Employee_Message = message.timestamp;
      }
      
      updateTask(taskId, updateData);
      
      Logger.log(`Message appended to conversation history for task ${taskId}`);
    } else {
      Logger.log(`Duplicate message detected, skipping: ${message.messageId}`);
    }
    
  } catch (error) {
    Logger.log(`Error appending to conversation history: ${error.toString()}`);
  }
}

// ============================================
// ⭐⭐ REPROCESS TASK REPLY - START HERE ⭐⭐
// ============================================
// 
// TO USE THIS FUNCTION:
// 
// 1. Find your Task ID:
//    - Look in your Tasks_DB Google Sheet, column "Task_ID"
//    - Or check the task card in Lovable dashboard
//    - Format: TASK-20251222235746 (example)
//
// 2. Scroll down to find the function "reprocessMyTask" below
//    (It's around line 910, or search for "reprocessMyTask")
//
// 3. You'll see this line:
//    const taskId = 'PUT-TASK-ID-HERE';
//
// 4. Replace 'PUT-TASK-ID-HERE' with your actual task ID
//    Example: const taskId = 'TASK-20251222235746';
//
// 5. Save the file (Ctrl+S or Cmd+S)
//
// 6. In Apps Script:
//    - Click the function dropdown (top left, shows function names)
//    - Select "reprocessMyTask"
//    - Click the Run button (▶️)
//
// 7. Check the execution log (View → Execution log) for results
//
// ============================================

/**
 * Trigger when task status changes to "Assigned"
 * Note: This should be called from updateTask or via onChange trigger
 */
function onTaskAssigned(taskId) {
  try {
    const task = getTask(taskId);
    // Use new status system - NOT_ACTIVE is the assigned state
    if (!task || task.Status !== TASK_STATUS.NOT_ACTIVE) {
      return;
    }
    
    // Send assignment email
    sendTaskAssignmentEmail(taskId);
    
    // Update task with email sent timestamp
    logInteraction(taskId, 'Task assigned and email sent');
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'onTaskAssigned', error.toString(), taskId, error.stack);
  }
}

/**
 * Gmail trigger for new emails (replies)
 * Note: This requires setting up a Gmail trigger in the Apps Script editor
 */
function onGmailReply(e) {
  try {
    const message = e.message;
    // Canonical ingestion (threadId/TaskId correlation + Processed_Message_IDs idempotency)
    const result = ingestInboundMessage(message);
    if (!result.success) {
      Logger.log(`onGmailReply: ingest failed: ${result.error || 'unknown error'}`);
    }
    
  } catch (error) {
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'onGmailReply', error.toString(), null, error.stack);
  }
}

/**
 * Process a reply email for a task
 */
function processReplyEmail(taskId, message) {
  try {
    Logger.log(`=== Processing reply for task: ${taskId} ===`);
    
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`ERROR: Task ${taskId} not found`);
      return;
    }
    
    // Get email content
    const rawEmailContent = message.getPlainBody();
    const senderEmail = extractEmailFromString(message.getFrom());
    const messageId = message.getId();
    
    Logger.log(`Sender: ${senderEmail}`);
    Logger.log(`Message ID: ${messageId}`);
    Logger.log(`Raw email preview: ${rawEmailContent.substring(0, 200)}...`);
    
    // Clean email content - remove quoted replies and signatures
    const emailContent = cleanEmailContent(rawEmailContent);
    Logger.log(`Cleaned email content: ${emailContent.substring(0, 200)}...`);
    
    // Verify sender is the assignee (allow for email variations)
    const assigneeEmail = extractEmailFromString(task.Assignee_Email);
    if (senderEmail.toLowerCase() !== assigneeEmail.toLowerCase()) {
      Logger.log(`WARNING: Reply from ${senderEmail} doesn't match assignee ${assigneeEmail}`);
      Logger.log(`Skipping this reply`);
      return;
    }
    
    // Check if we've already processed this message
    const log = task.Interaction_Log || '';
    if (log.includes(`Message ID: ${messageId}`)) {
      Logger.log(`Message ${messageId} already processed, skipping`);
      return;
    }
    
    // Classify reply type using AI (pass original due date for context)
    Logger.log('Classifying reply type using AI...');
    const originalDueDate = task.Due_Date ? Utilities.formatDate(new Date(task.Due_Date), Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
    Logger.log(`Original task due date: ${originalDueDate || 'Not set'}`);
    const classification = classifyReplyType(emailContent, originalDueDate);
    Logger.log(`Reply classified as: ${classification.type} (confidence: ${classification.confidence})`);
    Logger.log(`Reasoning: ${classification.reasoning || 'N/A'}`);
    Logger.log(`Extracted date: ${classification.extracted_date || 'None'}`);
    
    // Additional check: If classified as OTHER but contains date-related keywords, reclassify
    if (classification.type === 'OTHER' && classification.confidence < 0.7) {
      const dateKeywords = ['not feasible', 'deadline', 'by', 'before', 'need more time', 'can\'t make', 'feasible', 'propose', 'suggest', '10th', 'jan', 'january'];
      const hasDateKeyword = dateKeywords.some(keyword => emailContent.toLowerCase().includes(keyword));
      const hasDatePattern = /(\d{1,2}(?:st|nd|rd|th)?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2})/i.test(emailContent);
      
      if (hasDateKeyword && hasDatePattern) {
        Logger.log('Reclassifying OTHER to DATE_CHANGE based on date keywords and patterns');
        classification.type = 'DATE_CHANGE';
        classification.confidence = 0.7;
        classification.reasoning = 'Reclassified based on date keywords and patterns';
        
        // Try to extract date if not already extracted
        if (!classification.extracted_date) {
          const extractedDate = extractDateFromText(emailContent);
          if (extractedDate) {
            classification.extracted_date = extractedDate;
            Logger.log(`Extracted date from text: ${extractedDate}`);
          }
        }
      }
    }
    
    // Additional fallback: Manual date extraction if still OTHER
    if (classification.type === 'OTHER') {
      const manuallyExtractedDate = extractDateFromText(emailContent);
      if (manuallyExtractedDate) {
        Logger.log(`Manual date extraction found: ${manuallyExtractedDate}`);
        Logger.log('Reclassifying OTHER to DATE_CHANGE based on manual date detection');
        classification.type = 'DATE_CHANGE';
        classification.extracted_date = manuallyExtractedDate;
        classification.confidence = 0.7;
        classification.reasoning = 'Reclassified based on manual date extraction';
      }
    }
    
    // Process based on classification
    if (classification.type === 'ACCEPTANCE') {
      handleAcceptanceReply(taskId, emailContent, messageId);
    } else if (classification.type === 'DATE_CHANGE') {
      handleDateChangeReply(taskId, classification.extracted_date, emailContent, messageId);
    } else if (classification.type === 'SCOPE_QUESTION') {
      handleScopeQuestionReply(taskId, emailContent, messageId);
    } else if (classification.type === 'ROLE_REJECTION') {
      handleRoleRejectionReply(taskId, emailContent, messageId);
    } else {
      handleOtherReply(taskId, emailContent, messageId);
    }
    
    // Append to conversation history
    appendToConversationHistory(taskId, {
      id: messageId,
      timestamp: new Date().toISOString(),
      senderEmail: senderEmail,
      senderName: message.getFrom().match(/^(.+?)\s*</)?.[1] || senderEmail,
      type: 'email_reply',
      content: emailContent,
      messageId: messageId,
      metadata: {
        classification: classification.type,
        extractedDate: classification.extracted_date,
        confidence: classification.confidence
      }
    });
    
    // Log the interaction with message ID
    logInteraction(taskId, `Reply received from ${senderEmail}: ${classification.type} (Message ID: ${messageId})`);
    
    // IMPORTANT: Analyze conversation and update state (new unified model)
    // This derives the conversation state from the full conversation history
    try {
      Logger.log(`Analyzing conversation to update state for task ${taskId}...`);
      const analysisResult = analyzeConversationAndUpdateState(taskId);
      if (analysisResult.success) {
        Logger.log(`Conversation state updated to: ${analysisResult.data.conversationState}`);
        Logger.log(`AI Summary: ${analysisResult.data.summary}`);
      } else {
        Logger.log(`Warning: Failed to analyze conversation: ${analysisResult.error}`);
      }
    } catch (analysisError) {
      Logger.log(`Error analyzing conversation: ${analysisError.toString()}`);
      // Don't fail the reply processing if analysis fails
    }
    
    // Execute workflows for email_reply trigger
    try {
      const updatedTask = getTask(taskId);
      executeWorkflow('email_reply', {
        taskId: taskId,
        task: updatedTask,
        replyType: classification.type,
        senderEmail: senderEmail,
        messageId: messageId,
        conversationState: updatedTask.Conversation_State
      });
    } catch (error) {
      Logger.log(`Error executing workflow for email_reply: ${error.toString()}`);
      // Don't fail reply processing if workflow fails
    }
    
    Logger.log(`✓ Reply processed successfully`);
    
  } catch (error) {
    Logger.log(`ERROR processing reply: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
    logError(ERROR_TYPE.API_ERROR, 'processReplyEmail', error.toString(), taskId, error.stack);
  }
}

/**
 * Extract email address from string (handles "Name <email@domain.com>" format)
 */
function extractEmailFromString(emailString) {
  if (!emailString) return '';
  
  // Check if it's in "Name <email@domain.com>" format
  const match = emailString.match(/<(.+?)>/);
  if (match) {
    return match[1].trim();
  }
  
  // Otherwise return as-is
  return emailString.trim();
}

/**
 * Clean email content by removing quoted replies and signatures
 * This ensures we only process the actual reply, not quoted text
 */
function cleanEmailContent(emailContent) {
  if (!emailContent) return '';
  
  let cleaned = emailContent;
  
  // Gmail-style quote patterns (can span multiple lines)
  // Pattern: "On [date] at [time], [name] <email@domain.com>\nwrote:" or similar
  // These patterns match the START of a quoted section
  const gmailQuotePatterns = [
    // Gmail: "On Sun, 28 Dec 2025 at 12:17 AM, Name <email@domain.com>\nwrote:"
    /On\s+(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)[,\s]+\d{1,2}\s+\w+\s+\d{4}\s+at\s+\d{1,2}:\d{2}\s*(?:AM|PM)?[,\s]+[^<]*<[^>]+>[\s\S]*?wrote:/gi,
    // Gmail variant: "On Dec 28, 2025, at 12:17 AM, Name wrote:"
    /On\s+\w+\s+\d{1,2},?\s+\d{4},?\s+at\s+\d{1,2}:\d{2}\s*(?:AM|PM)?[,\s]+[^w]*wrote:/gi,
    // Generic: "On [any date format] ... wrote:"
    /On\s+[^]*?wrote:/gi,
  ];
  
  // Try each Gmail pattern and cut off at the first match
  for (const pattern of gmailQuotePatterns) {
    const match = cleaned.match(pattern);
    if (match) {
      const index = cleaned.indexOf(match[0]);
      if (index > 0) {
        cleaned = cleaned.substring(0, index).trim();
        break; // Stop after first successful cut
      }
    }
  }
  
  // Additional single-line quote markers
  const singleLineMarkers = [
    /^From:\s*.*$/m,                    // "From: ..."
    /^-----Original Message-----.*$/m,  // Outlook style
    /^________________________________.*$/m, // Outlook separator
    /^_{10,}.*$/m,                      // Multiple underscores
    /^={10,}.*$/m,                      // Multiple equals signs
    /^-{10,}.*$/m,                      // Multiple dashes
    /^Forwarded message.*$/mi,          // Forwarded message header
    /^Begin forwarded message.*$/mi,    // Apple Mail forward
    /^\*From:\*.*$/m,                   // Bold From (some clients)
  ];
  
  for (const marker of singleLineMarkers) {
    const match = cleaned.match(marker);
    if (match) {
      const index = cleaned.indexOf(match[0]);
      if (index > 0) {
        cleaned = cleaned.substring(0, index).trim();
      }
    }
  }
  
  // Remove lines starting with ">" (standard email quote)
  cleaned = cleaned.split('\n')
    .filter(line => !line.trim().startsWith('>'))
    .join('\n');
  
  // Remove email signatures (common patterns)
  // These patterns match from the signature start to the end of the content
  const signaturePatterns = [
    /\n\n--\s*\n[\s\S]*$/,              // Standard signature delimiter "--"
    /\nBest regards,[\s\S]*$/i,
    /\nKind regards,[\s\S]*$/i,
    /\nRegards,[\s\S]*$/i,
    /\nThanks,[\s\S]*$/i,
    /\nThank you,[\s\S]*$/i,
    /\nSincerely,[\s\S]*$/i,
    /\nCheers,[\s\S]*$/i,
    /\nSent from my [\s\S]*$/i,         // Mobile signatures
    /\nGet Outlook[\s\S]*$/i,           // Outlook mobile
    /\nSent from Mail[\s\S]*$/i,        // Apple Mail
  ];
  
  for (const pattern of signaturePatterns) {
    cleaned = cleaned.replace(pattern, '');
  }
  
  // Clean up excessive whitespace
  cleaned = cleaned
    .replace(/\n{3,}/g, '\n\n')  // Replace 3+ newlines with 2
    .trim();
  
  return cleaned;
}

/**
 * Handle acceptance reply
 */
function handleAcceptanceReply(taskId, emailContent, messageId) {
  Logger.log(`Handling ACCEPTANCE reply for task ${taskId}`);
  
  const task = getTask(taskId);
  
  // If we are awaiting explicit employee confirmation (boss proposed or boss approved),
  // treat an ACCEPTANCE reply as confirmation and APPLY the change.
  let pendingDecision = null;
  try {
    pendingDecision = task.Pending_Decision ? JSON.parse(task.Pending_Decision) : null;
  } catch (e) {
    pendingDecision = null;
  }
  
  if (pendingDecision && pendingDecision.type === 'date_change' && pendingDecision.awaitingFrom === 'employee') {
    Logger.log(`Employee confirming pending date change`);
    
    // Apply the date change now
    updateTask(taskId, {
      Due_Date: pendingDecision.proposedValue,
      Proposed_Date: '',
      Status: task.Status === TASK_STATUS.NOT_ACTIVE ? TASK_STATUS.ON_TIME : task.Status,
      Conversation_State: CONVERSATION_STATE.RESOLVED,
      Pending_Decision: '',
      Pending_Changes: '',
      Employee_Reply: '',
      Last_Updated: new Date()
    });
    
    logInteraction(taskId, `Employee confirmed date change. Due date updated to ${pendingDecision.proposedValue} (Message ID: ${messageId})`);
    
    // Send confirmation to boss
    try {
      const bossEmail = CONFIG.BOSS_EMAIL();
      const tz = Session.getScriptTimeZone();
      const confirmedDateFormatted = Utilities.formatDate(
        new Date(pendingDecision.proposedValue), 
        tz, 
        'EEEE, MMMM d, yyyy'
      );
      
      const subject = `Date Change Confirmed: ${task.Task_Name}`;
      const body = `The employee has confirmed the new due date:\n\n` +
        `Task: ${task.Task_Name}\n` +
        `Confirmed Due Date: ${confirmedDateFormatted}\n\n` +
        `The task is now active with the updated deadline.`;
      
      GmailApp.sendEmail(
        bossEmail,
        subject,
        body,
        {
          name: 'Chief of Staff AI',
        }
      );
      
      Logger.log(`Confirmation sent to boss`);
    } catch (error) {
      Logger.log(`Failed to send boss confirmation: ${error.toString()}`);
    }
    
    return;
  }
  
  // Regular acceptance - just update status
  updateTask(taskId, {
    Status: TASK_STATUS.ON_TIME,
    Conversation_State: CONVERSATION_STATE.ACTIVE,
  });
  
  logInteraction(taskId, `Assignee accepted task (Message ID: ${messageId})`);
  
  // Optionally send confirmation to assignee
  try {
    if (task && task.Assignee_Email) {
      const subject = `Re: Task Assignment: ${task.Task_Name}`;
      const body = `Thank you for confirming. I've updated the task status to Active. Please let me know if you need any support.\n\nBest regards,\nChief of Staff AI`;
      
      GmailApp.sendEmail(
        task.Assignee_Email,
        subject,
        body,
        {
          name: 'Chief of Staff AI',
          replyTo: CONFIG.BOSS_EMAIL(),
        }
      );
      
      Logger.log(`Confirmation email sent to ${task.Assignee_Email}`);
    }
  } catch (error) {
    Logger.log(`Failed to send confirmation email: ${error.toString()}`);
  }
}

/**
 * Handle date change request
 * Note: Status stays at ON_TIME (lifecycle = active), Conversation_State drives UI
 */
function handleDateChangeReply(taskId, proposedDate, emailContent, messageId) {
  Logger.log(`Handling DATE_CHANGE reply for task ${taskId}`);
  Logger.log(`Proposed date from AI: ${proposedDate || 'Not extracted'}`);
  
  // Try to extract date from email content if AI didn't extract it
  let finalProposedDate = proposedDate;
  if (!finalProposedDate) {
    Logger.log('AI did not extract date, trying manual extraction...');
    finalProposedDate = extractDateFromText(emailContent);
    Logger.log(`Manual extraction result: ${finalProposedDate || 'No date found'}`);
  }
  
  const task = getTask(taskId);

  // If we were awaiting employee confirmation on a boss-proposed date,
  // and the employee replied with a (new) date change request, treat this as
  // a counter-proposal: clear the pending decision and surface as boss action needed.
  let pendingDecision = null;
  try {
    pendingDecision = task.Pending_Decision ? JSON.parse(task.Pending_Decision) : null;
  } catch (e) {
    pendingDecision = null;
  }
  
  const isCounterProposal =
    pendingDecision &&
    pendingDecision.type === 'date_change' &&
    pendingDecision.awaitingFrom === 'employee' &&
    pendingDecision.requestedBy === 'boss';
  
  // Keep task active (ON_TIME) - Conversation_State will be set by analyzeConversationAndUpdateState()
  // Legacy: Also set REVIEW_DATE for backward compatibility with old frontend
  updateTask(taskId, {
    Proposed_Date: finalProposedDate || '',
    Status: task.Status === TASK_STATUS.NOT_ACTIVE ? TASK_STATUS.ON_TIME : task.Status, // Move to active if first response
    Employee_Reply: emailContent,
    Last_Employee_Message: new Date().toISOString(),
    ...(isCounterProposal ? {
      Pending_Decision: '',
      Conversation_State: CONVERSATION_STATE.CHANGE_REQUESTED,
      Pending_Changes: JSON.stringify([{
        id: 'date_change_' + Date.now(),
        parameter: 'dueDate',
        currentValue: task.Due_Date,
        proposedValue: finalProposedDate || '',
        requestedBy: 'employee',
        reasoning: emailContent.substring(0, 500)
      }])
    } : {})
  });
  
  Logger.log(`Task ${taskId} employee reply stored. Proposed date: ${finalProposedDate || 'none'}`);
  logInteraction(taskId, `Assignee requested date change to ${finalProposedDate || 'unspecified date'} (Message ID: ${messageId})`);
  
  // Notify boss about date change request
  notifyBossOfReviewRequest(taskId, 'DATE_CHANGE', emailContent, finalProposedDate);
}

/**
 * Handle scope question reply
 * Note: Status stays at ON_TIME (lifecycle = active), Conversation_State drives UI
 */
function handleScopeQuestionReply(taskId, emailContent, messageId) {
  Logger.log(`Handling SCOPE_QUESTION reply for task ${taskId}`);
  
  const task = getTask(taskId);
  const currentLog = task.Interaction_Log || '';
  const newLog = currentLog + `\n\nScope question from assignee: ${emailContent.substring(0, 500)}`;
  
  // Keep task active (ON_TIME) - Conversation_State will be set by analyzeConversationAndUpdateState()
  updateTask(taskId, {
    Status: task.Status === TASK_STATUS.NOT_ACTIVE ? TASK_STATUS.ON_TIME : task.Status,
    Interaction_Log: newLog,
    Employee_Reply: emailContent,
    Last_Employee_Message: new Date().toISOString(),
  });
  
  logInteraction(taskId, `Assignee has scope questions (Message ID: ${messageId})`);
  
  // Notify boss about scope question
  notifyBossOfReviewRequest(taskId, 'SCOPE_QUESTION', emailContent);
}

/**
 * Handle role rejection reply
 * Note: Status stays at ON_TIME (lifecycle = active), Conversation_State drives UI
 */
function handleRoleRejectionReply(taskId, emailContent, messageId) {
  Logger.log(`Handling ROLE_REJECTION reply for task ${taskId}`);
  
  const task = getTask(taskId);
  const currentLog = task.Interaction_Log || '';
  const newLog = currentLog + `\n\nRole rejection from assignee: ${emailContent.substring(0, 500)}`;
  
  // Keep task active (ON_TIME) - Conversation_State will be set by analyzeConversationAndUpdateState()
  updateTask(taskId, {
    Status: task.Status === TASK_STATUS.NOT_ACTIVE ? TASK_STATUS.ON_TIME : task.Status,
    Interaction_Log: newLog,
    Employee_Reply: emailContent,
    Last_Employee_Message: new Date().toISOString(),
  });
  
  logInteraction(taskId, `Assignee claims task is not their responsibility (Message ID: ${messageId})`);
  
  // Notify boss about role rejection
  notifyBossOfReviewRequest(taskId, 'ROLE_REJECTION', emailContent);
}

/**
 * Handle other types of replies
 */
function handleOtherReply(taskId, emailContent, messageId) {
  Logger.log(`Handling OTHER reply for task ${taskId}`);
  
  const task = getTask(taskId);
  const currentLog = task.Interaction_Log || '';
  const newLog = currentLog + `\n\nReply from assignee: ${emailContent.substring(0, 500)}`;
  
  updateTask(taskId, {
    Interaction_Log: newLog,
  });
  
  logInteraction(taskId, `Received reply (unclassified type) (Message ID: ${messageId})`);
  
  // For unclassified replies, still notify boss if task is in NOT_ACTIVE status
  if (task.Status === TASK_STATUS.NOT_ACTIVE) {
    notifyBossOfReviewRequest(taskId, 'OTHER', emailContent);
  }
}

/**
 * Notify boss about review request
 * Enhanced with more context, AI summary, and quick action guidance
 */
function notifyBossOfReviewRequest(taskId, reviewType, emailContent, proposedDate = null) {
  try {
    const task = getTask(taskId);
    if (!task) return;
    
    const bossEmail = CONFIG.BOSS_EMAIL();
    if (!bossEmail) return;
    
    // Get staff name for personalization
    const staff = getStaff(task.Assignee_Email);
    const assigneeName = staff ? staff.Name : task.Assignee_Email;
    
    // Format dates nicely
    const tz = Session.getScriptTimeZone();
    const currentDueDateFormatted = task.Due_Date 
      ? Utilities.formatDate(new Date(task.Due_Date), tz, 'EEEE, MMMM d, yyyy')
      : 'Not specified';
    const proposedDateFormatted = proposedDate 
      ? Utilities.formatDate(new Date(proposedDate), tz, 'EEEE, MMMM d, yyyy')
      : 'Not specified in email';
    
    // Generate AI summary if possible (with fallback)
    let aiSummary = '';
    try {
      aiSummary = summarizeReviewRequest(
        reviewType, 
        emailContent, 
        task.Task_Name, 
        currentDueDateFormatted, 
        proposedDateFormatted
      );
    } catch (e) {
      Logger.log(`Could not generate AI summary: ${e.toString()}`);
    }
    
    let subject = '';
    let body = '';
    let htmlBody = '';
    
    // Header with emoji for quick visual identification
    const emoji = {
      'DATE_CHANGE': '📅',
      'SCOPE_QUESTION': '❓',
      'ROLE_REJECTION': '🔄',
      'OTHER': '📧'
    }[reviewType] || '📧';
    
    switch (reviewType) {
      case 'DATE_CHANGE':
        subject = `${emoji} Date Change Request: ${task.Task_Name}`;
        body = `═══════════════════════════════════════════════════════════════\n`;
        body += `DATE CHANGE REQUEST\n`;
        body += `═══════════════════════════════════════════════════════════════\n\n`;
        body += `${assigneeName} has requested a different due date.\n\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `TASK DETAILS\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `• Task: ${task.Task_Name}\n`;
        body += `• Assignee: ${assigneeName} (${task.Assignee_Email})\n`;
        body += `• Project: ${task.Project_Tag || 'No project'}\n`;
        body += `• Current Due Date: ${currentDueDateFormatted}\n`;
        body += `• PROPOSED DATE: ${proposedDateFormatted}\n`;
        if (task.Priority) body += `• Priority: ${task.Priority}\n`;
        body += `\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `EMPLOYEE'S MESSAGE\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `"${emailContent.substring(0, 800)}"\n\n`;
        if (aiSummary) {
          body += `───────────────────────────────────────────────────────────────\n`;
          body += `AI SUMMARY\n`;
          body += `───────────────────────────────────────────────────────────────\n`;
          body += `${aiSummary}\n\n`;
        }
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `RECOMMENDED ACTIONS\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `1. APPROVE the new date (${proposedDateFormatted})\n`;
        body += `   → Update the task due date in your dashboard\n\n`;
        body += `2. NEGOTIATE a different date\n`;
        body += `   → Reply to the assignee with your preferred date\n\n`;
        body += `3. REJECT and keep the original date\n`;
        body += `   → Send a message explaining why the deadline must be kept\n`;
        break;
        
      case 'SCOPE_QUESTION':
        subject = `${emoji} Scope Question: ${task.Task_Name}`;
        body = `═══════════════════════════════════════════════════════════════\n`;
        body += `SCOPE CLARIFICATION NEEDED\n`;
        body += `═══════════════════════════════════════════════════════════════\n\n`;
        body += `${assigneeName} has questions about this task.\n\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `TASK DETAILS\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `• Task: ${task.Task_Name}\n`;
        body += `• Assignee: ${assigneeName} (${task.Assignee_Email})\n`;
        body += `• Project: ${task.Project_Tag || 'No project'}\n`;
        body += `• Due Date: ${currentDueDateFormatted}\n`;
        if (task.Context_Hidden) body += `• Original Context: ${task.Context_Hidden.substring(0, 200)}...\n`;
        body += `\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `EMPLOYEE'S QUESTION\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `"${emailContent.substring(0, 800)}"\n\n`;
        if (aiSummary) {
          body += `───────────────────────────────────────────────────────────────\n`;
          body += `AI SUMMARY\n`;
          body += `───────────────────────────────────────────────────────────────\n`;
          body += `${aiSummary}\n\n`;
        }
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `RECOMMENDED ACTIONS\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `1. Reply directly to clarify the task scope\n`;
        body += `2. Update the task description with more details\n`;
        body += `3. Schedule a quick call if the scope is complex\n`;
        break;
        
      case 'ROLE_REJECTION':
        subject = `${emoji} Role Concern: ${task.Task_Name}`;
        body = `═══════════════════════════════════════════════════════════════\n`;
        body += `TASK REASSIGNMENT MAY BE NEEDED\n`;
        body += `═══════════════════════════════════════════════════════════════\n\n`;
        body += `${assigneeName} believes this task should be assigned to someone else.\n\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `TASK DETAILS\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `• Task: ${task.Task_Name}\n`;
        body += `• Current Assignee: ${assigneeName} (${task.Assignee_Email})\n`;
        body += `• Project: ${task.Project_Tag || 'No project'}\n`;
        body += `• Due Date: ${currentDueDateFormatted}\n`;
        body += `\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `EMPLOYEE'S CONCERN\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `"${emailContent.substring(0, 800)}"\n\n`;
        if (aiSummary) {
          body += `───────────────────────────────────────────────────────────────\n`;
          body += `AI SUMMARY\n`;
          body += `───────────────────────────────────────────────────────────────\n`;
          body += `${aiSummary}\n\n`;
        }
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `RECOMMENDED ACTIONS\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `1. REASSIGN to the correct person\n`;
        body += `   → Update the assignee in your dashboard\n\n`;
        body += `2. CLARIFY why this person is the right fit\n`;
        body += `   → Reply explaining the assignment rationale\n\n`;
        body += `3. DISCUSS with the team to determine ownership\n`;
        break;
        
      default:
        subject = `${emoji} Reply Received: ${task.Task_Name}`;
        body = `═══════════════════════════════════════════════════════════════\n`;
        body += `NEW REPLY RECEIVED\n`;
        body += `═══════════════════════════════════════════════════════════════\n\n`;
        body += `${assigneeName} has sent a reply about this task.\n\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `TASK DETAILS\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `• Task: ${task.Task_Name}\n`;
        body += `• Assignee: ${assigneeName} (${task.Assignee_Email})\n`;
        body += `• Project: ${task.Project_Tag || 'No project'}\n`;
        body += `• Due Date: ${currentDueDateFormatted}\n`;
        body += `• Current Status: ${task.Status}\n`;
        body += `\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `EMPLOYEE'S MESSAGE\n`;
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `"${emailContent.substring(0, 800)}"\n\n`;
        if (aiSummary) {
          body += `───────────────────────────────────────────────────────────────\n`;
          body += `AI SUMMARY\n`;
          body += `───────────────────────────────────────────────────────────────\n`;
          body += `${aiSummary}\n\n`;
        }
        body += `───────────────────────────────────────────────────────────────\n`;
        body += `Please review this reply and take appropriate action.\n`;
    }
    
    // Add footer with task metadata
    body += `\n═══════════════════════════════════════════════════════════════\n`;
    body += `QUICK REFERENCE\n`;
    body += `═══════════════════════════════════════════════════════════════\n`;
    body += `Task ID: ${taskId}\n`;
    body += `Task Status: ${task.Status}\n`;
    body += `Reply Type: ${reviewType}\n`;
    body += `Processed: ${Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss z')}\n`;
    body += `\n`;
    body += `To take action, open your TaskFlow Hub dashboard or reply to the\n`;
    body += `original task assignment email thread.\n`;
    body += `═══════════════════════════════════════════════════════════════\n`;
    body += `\nThis notification was sent by Chief of Staff AI\n`;
    
    GmailApp.sendEmail(
      bossEmail,
      subject,
      body,
      {
        name: 'Chief of Staff AI',
      }
    );
    
    Logger.log(`Boss notification sent for ${reviewType} review request`);
    
  } catch (error) {
    Logger.log(`Failed to notify boss: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

/**
 * Extract date from text (helper function)
 * Handles various date formats including relative dates
 */
function extractDateFromText(text) {
  Logger.log(`Extracting date from text: ${text.substring(0, 200)}...`);
  
  // Clean the text first to remove quoted sections
  const cleanedText = cleanEmailContent(text);
  Logger.log(`Cleaned text for date extraction: ${cleanedText.substring(0, 200)}...`);
  
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth();
  const currentDay = now.getDate();
  const tz = Session.getScriptTimeZone();
  
  // Helper function to format date
  function formatDate(date) {
    return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  }
  
  // Helper function to get next occurrence of a weekday
  function getNextWeekday(dayIndex) {
    const result = new Date(now);
    const diff = (dayIndex - now.getDay() + 7) % 7;
    result.setDate(now.getDate() + (diff === 0 ? 7 : diff)); // If today, get next week
    return result;
  }
  
  const lowerText = cleanedText.toLowerCase();
  
  // ========================================
  // FIRST: Check for relative date expressions
  // ========================================
  
  // Tomorrow
  if (/\btomorrow\b/i.test(lowerText) || /\bkal\b/i.test(lowerText)) {
    const tomorrow = new Date(now);
    tomorrow.setDate(tomorrow.getDate() + 1);
    Logger.log('Found relative date: tomorrow');
    return formatDate(tomorrow);
  }
  
  // Day after tomorrow
  if (/\bday after tomorrow\b/i.test(lowerText) || /\bparson\b/i.test(lowerText) || /\bparso\b/i.test(lowerText)) {
    const dayAfter = new Date(now);
    dayAfter.setDate(dayAfter.getDate() + 2);
    Logger.log('Found relative date: day after tomorrow');
    return formatDate(dayAfter);
  }
  
  // Next week / next Monday
  if (/\bnext\s+week\b/i.test(lowerText) || /\bnext\s+monday\b/i.test(lowerText) || /\bagle\s+hafte\b/i.test(lowerText)) {
    const daysUntilMonday = (8 - now.getDay()) % 7 || 7;
    const nextMonday = new Date(now);
    nextMonday.setDate(now.getDate() + daysUntilMonday);
    Logger.log('Found relative date: next week/next monday');
    return formatDate(nextMonday);
  }
  
  // This week / end of week / by Friday
  if (/\bthis\s+week\b/i.test(lowerText) || /\bend\s+of\s+week\b/i.test(lowerText) || /\bby\s+friday\b/i.test(lowerText) || /\bis\s+hafte\b/i.test(lowerText)) {
    const daysUntilFriday = (5 - now.getDay() + 7) % 7;
    const friday = new Date(now);
    friday.setDate(now.getDate() + (daysUntilFriday === 0 ? 0 : daysUntilFriday));
    Logger.log('Found relative date: end of week/friday');
    return formatDate(friday);
  }
  
  // Next Friday specifically
  if (/\bnext\s+friday\b/i.test(lowerText)) {
    const daysUntilNextFriday = ((5 - now.getDay() + 7) % 7) + 7;
    const nextFriday = new Date(now);
    nextFriday.setDate(now.getDate() + daysUntilNextFriday);
    Logger.log('Found relative date: next friday');
    return formatDate(nextFriday);
  }
  
  // In X days
  const inDaysMatch = lowerText.match(/\bin\s+(\d+)\s+days?\b/i);
  if (inDaysMatch) {
    const days = parseInt(inDaysMatch[1], 10);
    const futureDate = new Date(now);
    futureDate.setDate(now.getDate() + days);
    Logger.log(`Found relative date: in ${days} days`);
    return formatDate(futureDate);
  }
  
  // In a week / one week
  if (/\bin\s+(?:a|one|1)\s+week\b/i.test(lowerText) || /\bek\s+hafte\s+mein\b/i.test(lowerText)) {
    const oneWeek = new Date(now);
    oneWeek.setDate(now.getDate() + 7);
    Logger.log('Found relative date: in a week');
    return formatDate(oneWeek);
  }
  
  // In two weeks
  if (/\bin\s+(?:two|2)\s+weeks\b/i.test(lowerText) || /\bdo\s+hafte\s+mein\b/i.test(lowerText)) {
    const twoWeeks = new Date(now);
    twoWeeks.setDate(now.getDate() + 14);
    Logger.log('Found relative date: in two weeks');
    return formatDate(twoWeeks);
  }
  
  // End of month / month end
  if (/\bend\s+of\s+month\b/i.test(lowerText) || /\bmonth\s+end\b/i.test(lowerText) || /\bmahine\s+ke\s+end\b/i.test(lowerText)) {
    const endOfMonth = new Date(currentYear, currentMonth + 1, 0);
    Logger.log('Found relative date: end of month');
    return formatDate(endOfMonth);
  }
  
  // Weekday names (next occurrence)
  const weekdayMatch = lowerText.match(/\b(?:this\s+)?(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i);
  if (weekdayMatch && !lowerText.includes('next')) {
    const weekdays = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    const targetDay = weekdays.indexOf(weekdayMatch[1].toLowerCase());
    if (targetDay >= 0) {
      const result = getNextWeekday(targetDay);
      Logger.log(`Found weekday: ${weekdayMatch[1]}`);
      return formatDate(result);
    }
  }
  
  // ========================================
  // SECOND: Try absolute date patterns
  // ========================================
  
  const datePatterns = [
    // "by 5th January 2026" or "by the 5th of January 2026" - prioritize "by" phrases
    {
      pattern: /by\s+(?:the\s+)?(\d{1,2})(?:st|nd|rd|th)?\s+(?:of\s+)?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})/i,
      parser: function(match) {
        const day = parseInt(match[1]);
        const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        const month = monthNames.indexOf(match[2].toLowerCase().substring(0, 3));
        const year = parseInt(match[3]);
        if (month >= 0 && day >= 1 && day <= 31) {
          return new Date(year, month, day);
        }
        return null;
      }
    },
    // "by 5th January" without year
    {
      pattern: /by\s+(?:the\s+)?(\d{1,2})(?:st|nd|rd|th)?\s+(?:of\s+)?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*/i,
      parser: function(match) {
        const day = parseInt(match[1]);
        const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        const month = monthNames.indexOf(match[2].toLowerCase().substring(0, 3));
        if (month >= 0 && day >= 1 && day <= 31) {
          const testDate = new Date(currentYear, month, day);
          const year = (testDate < now) ? currentYear + 1 : currentYear;
          return new Date(year, month, day);
        }
        return null;
      }
    },
    // "7th of Jan 2026" or "7th of January 2026" format
    {
      pattern: /(\d{1,2})(?:st|nd|rd|th)?\s+of\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})/i,
      parser: function(match) {
        const day = parseInt(match[1]);
        const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        const month = monthNames.indexOf(match[2].toLowerCase().substring(0, 3));
        const year = parseInt(match[3]);
        if (month >= 0 && day >= 1 && day <= 31) {
          return new Date(year, month, day);
        }
        return null;
      }
    },
    // "7th of Jan" or "7th of January" format without year
    {
      pattern: /(\d{1,2})(?:st|nd|rd|th)?\s+of\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b/i,
      parser: function(match) {
        const day = parseInt(match[1]);
        const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        const month = monthNames.indexOf(match[2].toLowerCase().substring(0, 3));
        if (month >= 0 && day >= 1 && day <= 31) {
          const testDate = new Date(currentYear, month, day);
          const year = (testDate < now) ? currentYear + 1 : currentYear;
          return new Date(year, month, day);
        }
        return null;
      }
    },
    // "10th Jan 2025" or "10 Jan 2025" format with year
    {
      pattern: /(\d{1,2})(?:st|nd|rd|th)?\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})/i,
      parser: function(match) {
        const day = parseInt(match[1]);
        const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        const month = monthNames.indexOf(match[2].toLowerCase().substring(0, 3));
        const year = parseInt(match[3]);
        if (month >= 0 && day >= 1 && day <= 31) {
          return new Date(year, month, day);
        }
        return null;
      }
    },
    // "Jan 10th, 2025" or "January 10, 2025" format with year
    {
      pattern: /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})/i,
      parser: function(match) {
        const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        const month = monthNames.indexOf(match[1].toLowerCase().substring(0, 3));
        const day = parseInt(match[2]);
        const year = parseInt(match[3]);
        if (month >= 0 && day >= 1 && day <= 31) {
          return new Date(year, month, day);
        }
        return null;
      }
    },
    // "10th Jan" or "10 Jan" format without year
    {
      pattern: /(\d{1,2})(?:st|nd|rd|th)?\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b/i,
      parser: function(match) {
        const day = parseInt(match[1]);
        const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        const month = monthNames.indexOf(match[2].toLowerCase().substring(0, 3));
        if (month >= 0 && day >= 1 && day <= 31) {
          const testDate = new Date(currentYear, month, day);
          const year = (testDate < now) ? currentYear + 1 : currentYear;
          return new Date(year, month, day);
        }
        return null;
      }
    },
    // "Jan 10th" or "January 10" format without year
    {
      pattern: /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2})(?:st|nd|rd|th)?\b/i,
      parser: function(match) {
        const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        const month = monthNames.indexOf(match[1].toLowerCase().substring(0, 3));
        const day = parseInt(match[2]);
        if (month >= 0 && day >= 1 && day <= 31) {
          const testDate = new Date(currentYear, month, day);
          const year = (testDate < now) ? currentYear + 1 : currentYear;
          return new Date(year, month, day);
        }
        return null;
      }
    },
    // YYYY-MM-DD
    {
      pattern: /(\d{4})-(\d{2})-(\d{2})/,
      parser: function(match) {
        return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
      }
    },
    // DD/MM/YYYY (more common internationally)
    {
      pattern: /(\d{1,2})\/(\d{1,2})\/(\d{4})/,
      parser: function(match) {
        const day = parseInt(match[1]);
        const month = parseInt(match[2]) - 1;
        const year = parseInt(match[3]);
        // Validate: if month > 12, it's probably MM/DD/YYYY format
        if (month > 11) {
          return new Date(year, day - 1, month);
        }
        return new Date(year, month, day);
      }
    },
    // DD-MM-YYYY (Indian format)
    {
      pattern: /(\d{1,2})-(\d{1,2})-(\d{4})/,
      parser: function(match) {
        const day = parseInt(match[1]);
        const month = parseInt(match[2]) - 1;
        const year = parseInt(match[3]);
        if (month <= 11 && day >= 1 && day <= 31) {
          return new Date(year, month, day);
        }
        return null;
      }
    },
    // Full month names with year
    {
      pattern: /(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})/i,
      parser: function(match) {
        const monthNames = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
        const month = monthNames.indexOf(match[1].toLowerCase());
        const day = parseInt(match[2]);
        const year = parseInt(match[3]);
        if (month >= 0 && day >= 1 && day <= 31) {
          return new Date(year, month, day);
        }
        return null;
      }
    },
    // Full month names without year - "January 5th" or "5th January"
    {
      pattern: /(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2})(?:st|nd|rd|th)?/i,
      parser: function(match) {
        const monthNames = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
        const month = monthNames.indexOf(match[1].toLowerCase());
        const day = parseInt(match[2]);
        if (month >= 0 && day >= 1 && day <= 31) {
          const testDate = new Date(currentYear, month, day);
          const year = (testDate < now) ? currentYear + 1 : currentYear;
          return new Date(year, month, day);
        }
        return null;
      }
    },
    // "5th January" or "5 January" without year
    {
      pattern: /(\d{1,2})(?:st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December)/i,
      parser: function(match) {
        const day = parseInt(match[1]);
        const monthNames = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
        const month = monthNames.indexOf(match[2].toLowerCase());
        if (month >= 0 && day >= 1 && day <= 31) {
          const testDate = new Date(currentYear, month, day);
          const year = (testDate < now) ? currentYear + 1 : currentYear;
          return new Date(year, month, day);
        }
        return null;
      }
    },
  ];
  
  // Use cleaned text for extraction
  for (const datePattern of datePatterns) {
    const match = cleanedText.match(datePattern.pattern);
    if (match) {
      try {
        Logger.log(`Found date pattern match: ${match[0]}`);
        const date = datePattern.parser(match);
        if (date && !isNaN(date.getTime())) {
          const formatted = formatDate(date);
          Logger.log(`Parsed date: ${formatted}`);
          return formatted;
        }
      } catch (e) {
        Logger.log(`Error parsing date ${match[0]}: ${e.toString()}`);
        // Continue to next pattern
      }
    }
  }
  
  Logger.log('No date pattern matched in cleaned text');
  return null;
}

/**
 * Handle silence escalation - check for tasks with no reply
 */
function handleSilenceEscalation() {
  try {
    const followupHours = CONFIG.ESCALATION_FOLLOWUP_HOURS();
    const bossAlertHours = CONFIG.ESCALATION_BOSS_ALERT_HOURS();
    
    const now = new Date();
    const followupThreshold = new Date(now.getTime() - followupHours * 60 * 60 * 1000);
    const bossAlertThreshold = new Date(now.getTime() - bossAlertHours * 60 * 60 * 1000);
    
    // Get all not_active tasks (awaiting response)
    const assignedTasks = getTasksByStatus(TASK_STATUS.NOT_ACTIVE);
    
    assignedTasks.forEach(task => {
      // Check last interaction time
      const lastUpdated = new Date(task.Last_Updated);
      
      // Check if we need to send follow-up
      if (lastUpdated < followupThreshold && lastUpdated >= bossAlertThreshold) {
        // Check if we already sent a follow-up (check Interaction_Log)
        const log = task.Interaction_Log || '';
        if (!log.includes('Follow-up email sent')) {
          sendFollowUpEmail(task.Task_ID);
          logInteraction(task.Task_ID, 'Follow-up email sent (no response in 24h)');
        }
      }
      
      // Check if we need to alert Boss
      if (lastUpdated < bossAlertThreshold) {
        // Check if we already escalated
        const log = task.Interaction_Log || '';
        if (!log.includes('Escalated to Boss')) {
          // Send escalation email
          sendEscalationEmail(task.Task_ID);
          
          // Update status to pending_action (new status for stagnation)
          updateTask(task.Task_ID, {
            Status: TASK_STATUS.PENDING_ACTION,
          });
          
          logInteraction(task.Task_ID, 'Escalated to Boss (no response in 48h)');
        }
      }
    });
    
  } catch (error) {
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'handleSilenceEscalation', error.toString(), null, error.stack);
  }
}

/**
 * Manual function to check and process replies
 * Can be run periodically or triggered
 */
function checkForReplies() {
  try {
    Logger.log('=== Checking for email replies ===');
    
    // Get all tasks that need reply checking - new status system
    const notActiveTasks = getTasksByStatus(TASK_STATUS.NOT_ACTIVE);
    const onTimeTasks = getTasksByStatus(TASK_STATUS.ON_TIME);
    const slowTasks = getTasksByStatus(TASK_STATUS.SLOW_PROGRESS);
    const allTasks = [...notActiveTasks, ...onTimeTasks, ...slowTasks];
    
    Logger.log(`Found ${allTasks.length} tasks to check`);
    
    allTasks.forEach(task => {
      try {
        // Refresh task data
        const freshTask = getTask(task.Task_ID);
        if (!freshTask) {
          Logger.log(`Task ${task.Task_ID} not found`);
          return;
        }
        
        // Find email thread
        const thread = findTaskEmailThread(freshTask.Task_ID);
        if (!thread) {
          Logger.log(`No email thread for task ${freshTask.Task_ID}`);
          return;
        }
        
        // Get ALL messages in thread
        const allMessages = thread.getMessages();
        Logger.log(`Thread has ${allMessages.length} total message(s)`);
        
        // Get unprocessed replies
        const replies = getTaskReplies(freshTask.Task_ID);
        Logger.log(`Found ${replies.length} unprocessed reply(ies) for task ${freshTask.Task_ID}`);
        
        if (replies.length === 0) {
          return;
        }
        
        // Process each reply via canonical ingestion
        replies.forEach((reply, index) => {
          try {
            Logger.log(`\n--- Reply ${index + 1}/${replies.length} ---`);
            const ingestResult = ingestInboundMessage(reply);
            if (ingestResult.success) {
              Logger.log(`  ✓ Ingested message ${ingestResult.messageId} for task ${ingestResult.taskId}`);
            } else {
              Logger.log(`  ❌ Ingest failed: ${ingestResult.error || 'unknown error'}`);
            }
            
          } catch (error) {
            Logger.log(`  ❌ ERROR processing reply ${index + 1}: ${error.toString()}`);
          }
        });
        
        
      } catch (error) {
        Logger.log(`Error processing task ${task.Task_ID}: ${error.toString()}`);
        logError(ERROR_TYPE.API_ERROR, 'checkForReplies', error.toString(), task.Task_ID, error.stack);
      }
    });
    
    Logger.log(`\n=== Check complete ===`);
    
  } catch (error) {
    Logger.log(`ERROR in checkForReplies: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack}`);
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'checkForReplies', error.toString(), null, error.stack);
  }
}

// ============================================
// TEST FUNCTIONS
// ============================================

/**
 * Test function: Check for replies manually
 * Run this to test the email reply processing
 */
function testCheckForReplies() {
  Logger.log('=== Testing Email Reply Processing ===');
  checkForReplies();
  Logger.log('=== Test Complete ===');
}

/**
 * Force reprocess an email reply for a task (bypasses "already processed" check)
 * This is useful when you want to reprocess with improved classification code
 * Usage: forceReprocessReply('TASK-20251222235746')
 */
function forceReprocessReply(taskId) {
  try {
    Logger.log(`=== Force Reprocessing Reply for Task: ${taskId} ===`);
    
    if (!taskId) {
      Logger.log('ERROR: No task ID provided');
      Logger.log('Usage: forceReprocessReply("TASK-20251222235746")');
      return;
    }
    
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`ERROR: Task ${taskId} not found`);
      return;
    }
    
    Logger.log(`Task: ${task.Task_Name}`);
    Logger.log(`Assignee: ${task.Assignee_Email}`);
    Logger.log(`Current Status: ${task.Status}`);
    
    // Find email thread
    const thread = findTaskEmailThread(taskId);
    if (!thread) {
      Logger.log('ERROR: No email thread found for this task');
      Logger.log('Make sure an assignment email was sent first');
      return;
    }
    
    Logger.log(`Found email thread: ${thread.getFirstMessageSubject()}`);
    
    // Get replies
    const replies = getTaskReplies(taskId);
    Logger.log(`Found ${replies.length} reply(ies)`);
    
    if (replies.length === 0) {
      Logger.log('No replies found. Send a test reply to the assignment email first.');
      return;
    }
    
    // Process most recent reply
    const latestReply = replies[replies.length - 1];
    const messageId = latestReply.getId();
    Logger.log(`Processing latest reply from: ${latestReply.getFrom()}`);
    Logger.log(`Message ID: ${messageId}`);
    
    // Get email content
    const rawEmailContent = latestReply.getPlainBody();
    const senderEmail = extractEmailFromString(latestReply.getFrom());
    
    Logger.log(`Raw email preview: ${rawEmailContent.substring(0, 300)}...`);
    
    // Clean email content - remove quoted replies and signatures
    const emailContent = cleanEmailContent(rawEmailContent);
    Logger.log(`Cleaned email content: ${emailContent.substring(0, 300)}...`);
    
    // Verify sender is the assignee
    const assigneeEmail = extractEmailFromString(task.Assignee_Email);
    if (senderEmail.toLowerCase() !== assigneeEmail.toLowerCase()) {
      Logger.log(`WARNING: Reply from ${senderEmail} doesn't match assignee ${assigneeEmail}`);
      Logger.log(`Proceeding anyway (force reprocess mode)`);
    }
    
    // Remove the Message ID from Interaction_Log to allow reprocessing
    const currentLog = task.Interaction_Log || '';
    const logLines = currentLog.split('\n');
    const filteredLog = logLines.filter(line => !line.includes(`Message ID: ${messageId}`)).join('\n');
    
    if (currentLog !== filteredLog) {
      Logger.log(`Removed Message ID ${messageId} from Interaction_Log to allow reprocessing`);
      updateTask(taskId, { Interaction_Log: filteredLog });
    }
    
    // Classify reply type using AI (with improved classification, pass original due date for context)
    Logger.log('Classifying reply type using AI (with improved classification)...');
    const originalDueDate = task.Due_Date ? Utilities.formatDate(new Date(task.Due_Date), Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
    Logger.log(`Original task due date: ${originalDueDate || 'Not set'}`);
    const classification = classifyReplyType(emailContent, originalDueDate);
    Logger.log(`Reply classified as: ${classification.type} (confidence: ${classification.confidence})`);
    Logger.log(`Reasoning: ${classification.reasoning || 'N/A'}`);
    Logger.log(`Extracted date: ${classification.extracted_date || 'None'}`);
    
    // Enhanced fallback: If classified as OTHER but contains date-related keywords, reclassify
    if (classification.type === 'OTHER' && classification.confidence < 0.7) {
      const dateKeywords = ['not feasible', 'deadline', 'by', 'before', 'need more time', 'can\'t make', 'feasible', 'propose', 'suggest', '10th', 'jan', 'january'];
      const hasDateKeyword = dateKeywords.some(keyword => emailContent.toLowerCase().includes(keyword));
      const hasDatePattern = /(\d{1,2}(?:st|nd|rd|th)?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2})/i.test(emailContent);
      
      if (hasDateKeyword && hasDatePattern) {
        Logger.log('Reclassifying OTHER to DATE_CHANGE based on date keywords and patterns');
        classification.type = 'DATE_CHANGE';
        classification.confidence = 0.7;
        classification.reasoning = 'Reclassified based on date keywords and patterns';
        
        // Try to extract date if not already extracted
        if (!classification.extracted_date) {
          const extractedDate = extractDateFromText(emailContent);
          if (extractedDate) {
            classification.extracted_date = extractedDate;
            Logger.log(`Extracted date from text: ${extractedDate}`);
          }
        }
      }
    }
    
    // Additional fallback: Manual date extraction if still OTHER
    if (classification.type === 'OTHER') {
      const manuallyExtractedDate = extractDateFromText(emailContent);
      if (manuallyExtractedDate) {
        Logger.log(`Manual date extraction found: ${manuallyExtractedDate}`);
        Logger.log('Reclassifying OTHER to DATE_CHANGE based on manual date detection');
        classification.type = 'DATE_CHANGE';
        classification.extracted_date = manuallyExtractedDate;
        classification.confidence = 0.7;
        classification.reasoning = 'Reclassified based on manual date extraction';
      }
    }
    
    // Process based on classification
    if (classification.type === 'ACCEPTANCE') {
      handleAcceptanceReply(taskId, emailContent, messageId);
    } else if (classification.type === 'DATE_CHANGE') {
      handleDateChangeReply(taskId, classification.extracted_date, emailContent, messageId);
      Logger.log(`✓ Task status updated to REVIEW_DATE`);
      Logger.log(`✓ Task should now appear in Lovable dashboard`);
    } else if (classification.type === 'SCOPE_QUESTION') {
      handleScopeQuestionReply(taskId, emailContent, messageId);
    } else if (classification.type === 'ROLE_REJECTION') {
      handleRoleRejectionReply(taskId, emailContent, messageId);
    } else {
      handleOtherReply(taskId, emailContent, messageId);
    }
    
    Logger.log('✓ Reply reprocessed successfully!');
    Logger.log('Check the task status in your spreadsheet and Lovable dashboard.');
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
    logError(ERROR_TYPE.API_ERROR, 'forceReprocessReply', error.toString(), taskId, error.stack);
  }
}

/**
 * Test function: Process a specific task's replies
 * Usage: testProcessTaskReplies('TASK-20251221031304')
 */
function testProcessTaskReplies(taskId) {
  try {
    Logger.log(`=== Testing Reply Processing for Task: ${taskId} ===`);
    
    if (!taskId) {
      Logger.log('ERROR: No task ID provided');
      Logger.log('Usage: testProcessTaskReplies("TASK-20251221031304")');
      return;
    }
    
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`ERROR: Task ${taskId} not found`);
      return;
    }
    
    Logger.log(`Task: ${task.Task_Name}`);
    Logger.log(`Assignee: ${task.Assignee_Email}`);
    Logger.log(`Status: ${task.Status}`);
    
    // Find email thread
    const thread = findTaskEmailThread(taskId);
    if (!thread) {
      Logger.log('ERROR: No email thread found for this task');
      Logger.log('Make sure an assignment email was sent first');
      return;
    }
    
    Logger.log(`Found email thread: ${thread.getFirstMessageSubject()}`);
    
    // Get replies
    const replies = getTaskReplies(taskId);
    Logger.log(`Found ${replies.length} reply(ies)`);
    
    if (replies.length === 0) {
      Logger.log('No replies found. Send a test reply to the assignment email first.');
      return;
    }
    
    // Process most recent reply
    const latestReply = replies[replies.length - 1];
    Logger.log(`Processing latest reply from: ${latestReply.getFrom()}`);
    
    processReplyEmail(taskId, latestReply);
    
    Logger.log('✓ Reply processed!');
    Logger.log('Check the task status and Interaction_Log in your spreadsheet.');
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

/**
 * Test function: List all tasks with email threads
 */
function testListTasksWithEmails() {
  try {
    Logger.log('=== Tasks with Email Threads ===');
    
    const notActiveTasks = getTasksByStatus(TASK_STATUS.NOT_ACTIVE);
    const onTimeTasks = getTasksByStatus(TASK_STATUS.ON_TIME);
    const slowTasks = getTasksByStatus(TASK_STATUS.SLOW_PROGRESS);
    const allTasks = [...notActiveTasks, ...onTimeTasks, ...slowTasks];
    
    Logger.log(`Found ${allTasks.length} tasks to check`);
    
    allTasks.forEach((task, index) => {
      const thread = findTaskEmailThread(task.Task_ID);
      if (thread) {
        const replies = getTaskReplies(task.Task_ID);
        Logger.log(`\n${index + 1}. ${task.Task_ID} - ${task.Task_Name}`);
        Logger.log(`   Assignee: ${task.Assignee_Email}`);
        Logger.log(`   Status: ${task.Status}`);
        Logger.log(`   Thread: ${thread.getFirstMessageSubject()}`);
        Logger.log(`   Replies: ${replies.length}`);
      }
    });
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
  }
}

// ============================================
// ⭐⭐ REPROCESS TASK REPLY - PUT TASK ID HERE ⭐⭐
// ============================================
// 
// INSTRUCTIONS:
// 
// 1. Find your Task ID:
//    - Open your Tasks_DB Google Sheet
//    - Look in the "Task_ID" column
//    - Or check the task card in Lovable dashboard
//    - Example: TASK-20251222235746
//
// 2. Look at the line below that says:
//    const taskId = 'PUT-TASK-ID-HERE';
//
// 3. Replace 'PUT-TASK-ID-HERE' with your actual task ID
//    Example: const taskId = 'TASK-20251222235746';
//
// 4. Save the file (Ctrl+S or Cmd+S)
//
// 5. In Apps Script editor:
//    - Click the function dropdown (top left, shows function names)
//    - Select "reprocessMyTask"
//    - Click the Run button (▶️)
//
// 6. Check the execution log (View → Execution log) for results
//
// ============================================

function reprocessMyTask() {
  
  // ═══════════════════════════════════════════════════════════════
  // ⬇️⬇️⬇️ CHANGE THIS LINE - PUT YOUR TASK ID HERE ⬇️⬇️⬇️
  // ═══════════════════════════════════════════════════════════════
  
  const taskId = 'PUT-TASK-ID-HERE';
  
  // ═══════════════════════════════════════════════════════════════
  // ⬆️⬆️⬆️ CHANGE THE LINE ABOVE - PUT YOUR TASK ID THERE ⬆️⬆️⬆️
  // ═══════════════════════════════════════════════════════════════
  
  // Don't modify anything below this line
  if (taskId === 'PUT-TASK-ID-HERE') {
    Logger.log('');
    Logger.log('═══════════════════════════════════════════════════════════════');
    Logger.log('❌ ERROR: You need to replace PUT-TASK-ID-HERE with your task ID!');
    Logger.log('═══════════════════════════════════════════════════════════════');
    Logger.log('');
    Logger.log('HOW TO FIX:');
    Logger.log('1. Find this line in the code: const taskId = "PUT-TASK-ID-HERE";');
    Logger.log('2. Replace "PUT-TASK-ID-HERE" with your actual task ID');
    Logger.log('3. Example: const taskId = "TASK-20251222235746";');
    Logger.log('4. Save the file (Ctrl+S or Cmd+S)');
    Logger.log('5. Run the function again');
    Logger.log('');
    Logger.log('WHERE TO FIND YOUR TASK ID:');
    Logger.log('- Open your Tasks_DB Google Sheet');
    Logger.log('- Look in the "Task_ID" column');
    Logger.log('- Or check the task card in Lovable dashboard');
    Logger.log('');
    return;
  }
  
  Logger.log('✅ Processing task: ' + taskId);
  forceReprocessReply(taskId);
}

/**
 * Test function: Simulate processing a reply with sample text
 * Usage: testClassifyReply('I accept this task and will complete it by the deadline.')
 */
function testClassifyReply(replyText) {
  try {
    Logger.log('=== Testing Reply Classification ===');
    
    if (!replyText) {
      replyText = 'I accept this task and will complete it by the deadline.';
      Logger.log(`Using sample text: "${replyText}"`);
    }
    
    Logger.log(`Classifying: "${replyText}"`);
    
    // Test function - no original due date available, pass null
    const classification = classifyReplyType(replyText, null);
    
    Logger.log(`\nClassification Result:`);
    Logger.log(`  Type: ${classification.type}`);
    Logger.log(`  Confidence: ${classification.confidence}`);
    Logger.log(`  Extracted Date: ${classification.extracted_date || 'N/A'}`);
    Logger.log(`  Reasoning: ${classification.reasoning || 'N/A'}`);
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

