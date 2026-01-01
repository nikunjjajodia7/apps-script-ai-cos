/**
 * Email Helper Utilities
 * Functions for sending and processing emails via Gmail
 */

/**
 * Get email template from Config sheet
 */
function getEmailTemplate(templateType, variables = {}) {
  const config = getConfig();
  const templateKey = `EMAIL_TEMPLATE_${templateType.toUpperCase()}`;
  let template = config[templateKey];
  
  if (!template) {
    // Fallback templates
    if (templateType === 'assignment') {
      template = `Hello {{ASSIGNEE_NAME}},

On behalf of {{BOSS_NAME}}, I'm assigning you the following task:

• Task: {{TASK_NAME}}
• Context: {{CONTEXT}}
• Due Date: {{DUE_DATE}}

Please confirm that you can meet this deadline. If you have any questions, need clarification, or foresee any issues, let me know by replying to this email. If the date is not feasible, please propose an alternate date.

Thank you,
{{SIGNATURE}}`;
    } else if (templateType === 'followup') {
      template = `Hello {{ASSIGNEE_NAME}},

I wanted to follow up on the task assignment below. Please confirm receipt and your availability to complete this by {{DUE_DATE}}.

• Task: {{TASK_NAME}}
• Due Date: {{DUE_DATE}}

If you have any concerns, please reply to this email.

Thank you,
{{SIGNATURE}}`;
    } else if (templateType === 'escalation') {
      template = `Hello {{ASSIGNEE_NAME}},

This is a follow-up regarding the task assignment below. We haven't received a response yet, and this requires your attention.

• Task: {{TASK_NAME}}
• Due Date: {{DUE_DATE}}

Please respond as soon as possible to confirm you can complete this task.

Thank you,
{{SIGNATURE}}`;
    }
  }
  
  // Replace variables
  const bossEmail = CONFIG.BOSS_EMAIL();
  const bossName = bossEmail.split('@')[0];
  
  const replacements = {
    '{{ASSIGNEE_NAME}}': variables.assigneeName || variables.assignee_email || 'Team Member',
    '{{BOSS_NAME}}': bossName,
    '{{TASK_NAME}}': variables.taskName || variables.Task_Name || 'Task',
    '{{CONTEXT}}': variables.context || variables.Context_Hidden || 'No additional context',
    '{{DUE_DATE}}': variables.dueDate || variables.Due_Date || 'Not specified',
    '{{SIGNATURE}}': CONFIG.EMAIL_SIGNATURE(),
  };
  
  let emailBody = template;
  Object.keys(replacements).forEach(key => {
    emailBody = emailBody.replace(new RegExp(key, 'g'), replacements[key]);
  });
  
  return emailBody;
}

/**
 * Send task assignment email
 */
function sendTaskAssignmentEmail(taskId) {
  try {
    const task = getTask(taskId);
    if (!task) {
      throw new Error(`Task ${taskId} not found`);
    }
    
    if (!task.Assignee_Email) {
      throw new Error(`Task ${taskId} has no assignee`);
    }
    
    const staff = getStaff(task.Assignee_Email);
    const assigneeName = staff ? staff.Name : task.Assignee_Email.split('@')[0];
    
    // Generate email using AI or template
    let emailBody;
    try {
      emailBody = generateAssignmentEmail(task);
    } catch (error) {
      Logger.log('AI email generation failed, using template: ' + error);
      emailBody = getEmailTemplate('assignment', {
        assigneeName: assigneeName,
        taskName: task.Task_Name,
        context: task.Context_Hidden || '',
        dueDate: task.Due_Date ? Utilities.formatDate(new Date(task.Due_Date), Session.getScriptTimeZone(), 'MMMM d, yyyy') : 'Not specified',
      });
    }
    
    const subject = `Task Assignment: ${task.Task_Name}`;
    
    // Add Task ID to email body for correlation (in case subject changes)
    emailBody = `${emailBody}\n\n---\nTask ID: ${taskId}`;
    
    const emailOptions = {
      htmlBody: emailBody.replace(/\n/g, '<br>'),
      name: 'Chief of Staff AI',
      replyTo: CONFIG.BOSS_EMAIL(),
    };
    
    // Send email
    GmailApp.sendEmail(
      task.Assignee_Email,
      subject,
      emailBody,
      emailOptions
    );
    
    // Store email thread ID for correlation
    // Wait a moment for Gmail to index the email
    Utilities.sleep(1000);
    
    const emailThread = GmailApp.search(`to:${task.Assignee_Email} subject:"${subject}"`, 0, 1)[0];
    if (emailThread) {
      const threadId = emailThread.getId();
      // Store thread ID in dedicated field and Interaction_Log
      updateTask(taskId, {
        Primary_Thread_ID: threadId,
        Last_Reply_Check: new Date().toISOString()
      });
      logInteraction(taskId, `Assignment email sent to ${task.Assignee_Email}. Thread ID: ${threadId}`);
      Logger.log(`Stored Primary Thread ID: ${threadId} for task ${taskId}`);
    } else {
      logInteraction(taskId, `Assignment email sent to ${task.Assignee_Email} (Thread ID lookup failed)`);
      Logger.log(`Warning: Could not find thread for task ${taskId} immediately after sending`);
    }
    
    return true;
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'sendTaskAssignmentEmail', error.toString(), taskId, error.stack);
    throw error;
  }
}

/**
 * Send follow-up email
 */
function sendFollowUpEmail(taskId) {
  try {
    const task = getTask(taskId);
    if (!task || !task.Assignee_Email) {
      return false;
    }
    
    const staff = getStaff(task.Assignee_Email);
    const assigneeName = staff ? staff.Name : task.Assignee_Email.split('@')[0];
    
    const emailBody = getEmailTemplate('followup', {
      assigneeName: assigneeName,
      taskName: task.Task_Name,
      dueDate: task.Due_Date ? Utilities.formatDate(new Date(task.Due_Date), Session.getScriptTimeZone(), 'MMMM d, yyyy') : 'Not specified',
    });
    
    const subject = `Follow-up: Task Assignment: ${task.Task_Name}`;
    
    const result = sendEmailToAssignee(
      taskId,
      task.Assignee_Email,
      subject,
      emailBody,
      {
        htmlBody: emailBody.replace(/\n/g, '<br>')
      }
    );
    
    logInteraction(taskId, `Follow-up email sent to ${task.Assignee_Email}${result.threadId ? ` (Thread ID: ${result.threadId})` : ''}`);
    return result.success;
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'sendFollowUpEmail', error.toString(), taskId);
    return false;
  }
}

/**
 * Send escalation email
 */
function sendEscalationEmail(taskId) {
  try {
    const task = getTask(taskId);
    if (!task || !task.Assignee_Email) {
      return false;
    }
    
    const staff = getStaff(task.Assignee_Email);
    const assigneeName = staff ? staff.Name : task.Assignee_Email.split('@')[0];
    
    const emailBody = getEmailTemplate('escalation', {
      assigneeName: assigneeName,
      taskName: task.Task_Name,
      dueDate: task.Due_Date ? Utilities.formatDate(new Date(task.Due_Date), Session.getScriptTimeZone(), 'MMMM d, yyyy') : 'Not specified',
    });
    
    const subject = `Urgent: Task Assignment: ${task.Task_Name}`;
    
    const result = sendEmailToAssignee(
      taskId,
      task.Assignee_Email,
      subject,
      emailBody,
      {
        htmlBody: emailBody.replace(/\n/g, '<br>')
      }
    );
    
    logInteraction(taskId, `Escalation email sent to ${task.Assignee_Email}${result.threadId ? ` (Thread ID: ${result.threadId})` : ''}`);
    return result.success;
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'sendEscalationEmail', error.toString(), taskId);
    return false;
  }
}

/**
 * Send confirmation email to Boss
 */
function sendBossConfirmation(taskId, summary) {
  try {
    const task = getTask(taskId);
    if (!task) return;
    
    const bossEmail = CONFIG.BOSS_EMAIL();
    const subject = `Task Created: ${task.Task_Name}`;
    
    const emailBody = `Hello,

I've created a new task based on your voice command:

${summary}

Task ID: ${taskId}
Status: ${task.Status}

You can review and manage this task in your dashboard.

Best regards,
Chief of Staff AI`;

    GmailApp.sendEmail(
      bossEmail,
      subject,
      emailBody,
      {
        name: 'Chief of Staff AI',
      }
    );
    
    logInteraction(taskId, `Confirmation email sent to Boss`);
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'sendBossConfirmation', error.toString(), taskId);
  }
}

/**
 * Send email to assignee in the same thread (if thread exists)
 * This ensures all system emails continue the conversation thread
 * 
 * @param {string} taskId - Task ID
 * @param {string} to - Recipient email (usually task.Assignee_Email)
 * @param {string} subject - Email subject
 * @param {string} body - Email body (plain text)
 * @param {Object} options - Additional options (htmlBody, etc.)
 * @returns {Object} { success: boolean, threadId: string|null, messageId: string|null }
 */
function sendEmailToAssignee(taskId, to, subject, body, options = {}) {
  // #region agent log
  Logger.log('[DEBUG] sendEmailToAssignee entry: ' + JSON.stringify({location:'EmailHelper.gs:282',message:'sendEmailToAssignee entry',data:{taskId:taskId,to:to,subject:subject,bodyLength:body.length},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'A'}));
  // #endregion
  
  try {
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`sendEmailToAssignee: Task ${taskId} not found`);
      return { success: false, threadId: null, messageId: null };
    }
    
    // #region agent log
    Logger.log('[DEBUG] sendEmailToAssignee task loaded: ' + JSON.stringify({location:'EmailHelper.gs:290',message:'sendEmailToAssignee task loaded',data:{taskId:taskId,primaryThreadId:task.Primary_Thread_ID,hasTaskIdInBody:body.includes('Task ID: ' + taskId)},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'A'}));
    // #endregion
    
    // Ensure Task ID is in body
    let emailBody = body;
    if (!emailBody.includes(`Task ID: ${taskId}`)) {
      emailBody = `${emailBody}\n\n---\nTask ID: ${taskId}`;
    }
    
    // #region agent log
    Logger.log('[DEBUG] sendEmailToAssignee after Task ID check: ' + JSON.stringify({location:'EmailHelper.gs:297',message:'sendEmailToAssignee after Task ID check',data:{taskId:taskId,hasTaskIdInBody:emailBody.includes('Task ID: ' + taskId),emailBodyLength:emailBody.length},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'B'}));
    // #endregion
    
    // Get the original thread
    const thread = findTaskEmailThread(taskId);
    let threadId = null;
    let originalSubject = subject;
    
    // #region agent log
    Logger.log('[DEBUG] sendEmailToAssignee thread lookup result: ' + JSON.stringify({location:'EmailHelper.gs:301',message:'sendEmailToAssignee thread lookup result',data:{taskId:taskId,threadFound:!!thread,threadId:thread ? thread.getId() : null},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'C'}));
    // #endregion
    
    if (thread) {
      threadId = thread.getId();
      originalSubject = thread.getFirstMessageSubject();
      Logger.log(`Replying to existing thread: ${threadId}`);
      
      // #region agent log
      Logger.log('[DEBUG] sendEmailToAssignee thread found, preparing reply: ' + JSON.stringify({location:'EmailHelper.gs:304',message:'sendEmailToAssignee thread found, preparing reply',data:{taskId:taskId,threadId:threadId,originalSubject:originalSubject,emailBodyHasTaskId:emailBody.includes('Task ID: ' + taskId)},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'D'}));
      // #endregion
      
      // Get the last message in thread for In-Reply-To header
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];
      const lastMessageId = lastMessage.getId();
      
      try {
        // Get raw message to extract Message-ID header
        const rawMessage = Gmail.Users.Messages.get('me', lastMessageId, { format: 'raw' });
        const messageIdHeader = rawMessage.payload.headers.find(
          h => h.name.toLowerCase() === 'message-id'
        );
        const inReplyTo = messageIdHeader ? messageIdHeader.value : null;
        
        // Build email with proper headers for threading
        const emailLines = [
          `To: ${to}`,
          `From: ${CONFIG.BOSS_EMAIL()}`,
          `Reply-To: ${CONFIG.BOSS_EMAIL()}`,
          `Subject: Re: ${originalSubject}`,
        ];
        
        if (inReplyTo) {
          emailLines.push(`In-Reply-To: ${inReplyTo}`);
          emailLines.push(`References: ${inReplyTo}`);
        }
        
        // Add thread ID to email body (as requested)
        emailBody = `${emailBody}\n\n---\nThread ID: ${threadId}`;
        
        // #region agent log
        Logger.log('[DEBUG] sendEmailToAssignee before Gmail API send: ' + JSON.stringify({location:'EmailHelper.gs:333',message:'sendEmailToAssignee before Gmail API send',data:{taskId:taskId,threadId:threadId,hasTaskIdInBody:emailBody.includes('Task ID: ' + taskId),hasThreadIdInBody:emailBody.includes('Thread ID: ' + threadId),inReplyTo:inReplyTo},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'D'}));
        // #endregion
        
        emailLines.push(''); // Empty line before body
        emailLines.push(emailBody);
        
        // Encode email in base64url format
        const rawEmail = emailLines.join('\n');
        const encodedEmail = Utilities.base64EncodeWebSafe(rawEmail);
        
        // Send using Gmail API with threadId
        const sentMessage = Gmail.Users.Messages.send({
          raw: encodedEmail,
          threadId: threadId
        }, 'me');
        
        const sentMessageId = sentMessage.id;
        Logger.log(`Email sent in thread ${threadId}, message ID: ${sentMessageId}`);
        
        // #region agent log
        Logger.log('[DEBUG] sendEmailToAssignee Gmail API send success: ' + JSON.stringify({location:'EmailHelper.gs:349',message:'sendEmailToAssignee Gmail API send success',data:{taskId:taskId,threadId:threadId,sentMessageId:sentMessageId},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'D'}));
        // #endregion
        
        // Update Last_Reply_Check timestamp
        updateTask(taskId, {
          Last_Reply_Check: new Date().toISOString()
        });
        
        logInteraction(taskId, `Email sent to ${to} in thread ${threadId} (Message ID: ${sentMessageId})`);
        
        return { 
          success: true, 
          threadId: threadId, 
          messageId: sentMessageId 
        };
        
      } catch (apiError) {
        Logger.log(`Error sending via Gmail API: ${apiError.toString()}`);
        Logger.log(`Falling back to GmailApp.sendEmail`);
        
        // #region agent log
        Logger.log('[DEBUG] sendEmailToAssignee Gmail API failed, falling back: ' + JSON.stringify({location:'EmailHelper.gs:365',message:'sendEmailToAssignee Gmail API failed, falling back',data:{taskId:taskId,threadId:threadId,error:apiError.toString()},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'E'}));
        // #endregion
        
        // Fall through to fallback
      }
    }
    
    // Fallback: Send new email if thread not found or API failed
    Logger.log(`Sending new email (no thread found or API failed)`);
    
    // #region agent log
    Logger.log('[DEBUG] sendEmailToAssignee fallback path: ' + JSON.stringify({location:'EmailHelper.gs:372',message:'sendEmailToAssignee fallback path',data:{taskId:taskId,threadId:threadId,hasTaskIdInBody:emailBody.includes('Task ID: ' + taskId),subject:subject},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'E'}));
    // #endregion
    
    // Add thread ID to body if we have one (even if sending new email)
    if (threadId) {
      emailBody = `${emailBody}\n\n---\nThread ID: ${threadId}`;
    } else {
      emailBody = `${emailBody}\n\n---\nThread ID: (New thread - will be created)`;
    }
    
    // #region agent log
    Logger.log('[DEBUG] sendEmailToAssignee before GmailApp.sendEmail: ' + JSON.stringify({location:'EmailHelper.gs:380',message:'sendEmailToAssignee before GmailApp.sendEmail',data:{taskId:taskId,hasTaskIdInBody:emailBody.includes('Task ID: ' + taskId),emailBodyLength:emailBody.length},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'B'}));
    // #endregion
    
    // Regenerate htmlBody from emailBody (which now includes Task ID) to ensure Task ID is in HTML version too
    const emailOptions = {
      name: 'Chief of Staff AI',
      replyTo: CONFIG.BOSS_EMAIL(),
      ...options,
      htmlBody: emailBody.replace(/\n/g, '<br>') // Override htmlBody to include Task ID
    };
    
    GmailApp.sendEmail(to, subject, emailBody, emailOptions);
    
    // Try to capture thread ID after sending (for new threads)
    if (!threadId) {
      Utilities.sleep(1000); // Wait for Gmail to index
      const searchQuery = `to:${to} subject:"${subject}"`;
      const threads = GmailApp.search(searchQuery, 0, 1);
      if (threads.length > 0) {
        const newThreadId = threads[0].getId();
        // Update Primary_Thread_ID if not set
        if (!task.Primary_Thread_ID) {
          updateTask(taskId, {
            Primary_Thread_ID: newThreadId
          });
          Logger.log(`Stored new Primary_Thread_ID: ${newThreadId}`);
        }
        threadId = newThreadId;
        // Update email body with actual thread ID
        emailBody = emailBody.replace('Thread ID: (New thread - will be created)', `Thread ID: ${threadId}`);
      }
    }
    
    logInteraction(taskId, `Email sent to ${to}${threadId ? ` (Thread ID: ${threadId})` : ' (New thread)'}`);
    
    return { 
      success: true, 
      threadId: threadId, 
      messageId: null 
    };
    
  } catch (error) {
    Logger.log(`ERROR in sendEmailToAssignee: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
    logError(ERROR_TYPE.API_ERROR, 'sendEmailToAssignee', error.toString(), taskId, error.stack);
    return { success: false, threadId: null, messageId: null };
  }
}

/**
 * Find email thread for a task
 * SIMPLIFIED: Uses Primary_Thread_ID and Task ID search only
 * Returns the PRIMARY thread (original assignment email thread)
 */
function findTaskEmailThread(taskId) {
  // #region agent log
  Logger.log('[DEBUG] findTaskEmailThread entry: ' + JSON.stringify({location:'EmailHelper.gs:431',message:'findTaskEmailThread entry',data:{taskId:taskId},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'C'}));
  // #endregion
  
  const task = getTask(taskId);
  if (!task || !task.Assignee_Email) {
    Logger.log(`findTaskEmailThread: Task ${taskId} not found or no assignee email`);
    
    // #region agent log
    Logger.log('[DEBUG] findTaskEmailThread task not found: ' + JSON.stringify({location:'EmailHelper.gs:436',message:'findTaskEmailThread task not found',data:{taskId:taskId,hasTask:!!task,hasAssigneeEmail:task && !!task.Assignee_Email},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'C'}));
    // #endregion
    
    return null;
  }
  
  Logger.log(`Searching for thread: Task="${taskId}", Assignee="${task.Assignee_Email}"`);
  
  // METHOD 1: Use Primary_Thread_ID (fastest - direct access)
  if (task.Primary_Thread_ID) {
    // #region agent log
    Logger.log('[DEBUG] findTaskEmailThread trying Primary_Thread_ID: ' + JSON.stringify({location:'EmailHelper.gs:441',message:'findTaskEmailThread trying Primary_Thread_ID',data:{taskId:taskId,primaryThreadId:task.Primary_Thread_ID},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'C'}));
    // #endregion
    
    try {
      const thread = GmailApp.getThreadById(task.Primary_Thread_ID);
      if (thread) {
        Logger.log(`✓ Found thread by Primary_Thread_ID: ${task.Primary_Thread_ID}`);
        Logger.log(`  Subject: "${thread.getFirstMessageSubject()}"`);
        Logger.log(`  Messages: ${thread.getMessages().length}`);
        
        // #region agent log
        Logger.log('[DEBUG] findTaskEmailThread found by Primary_Thread_ID: ' + JSON.stringify({location:'EmailHelper.gs:448',message:'findTaskEmailThread found by Primary_Thread_ID',data:{taskId:taskId,threadId:thread.getId(),messageCount:thread.getMessages().length},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'C'}));
        // #endregion
        
        return thread;
      }
    } catch (error) {
      Logger.log(`Thread ${task.Primary_Thread_ID} not found, trying fallback`);
      
      // #region agent log
      Logger.log('[DEBUG] findTaskEmailThread Primary_Thread_ID failed: ' + JSON.stringify({location:'EmailHelper.gs:451',message:'findTaskEmailThread Primary_Thread_ID failed',data:{taskId:taskId,primaryThreadId:task.Primary_Thread_ID,error:error.toString()},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'C'}));
      // #endregion
    }
  } else {
    // #region agent log
    Logger.log('[DEBUG] findTaskEmailThread no Primary_Thread_ID: ' + JSON.stringify({location:'EmailHelper.gs:454',message:'findTaskEmailThread no Primary_Thread_ID',data:{taskId:taskId},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'C'}));
    // #endregion
  }
  
  // METHOD 2: Search by Task ID (fallback)
  const searchQuery = `"Task ID: ${taskId}" in:inbox`;
  
  // #region agent log
  Logger.log('[DEBUG] findTaskEmailThread searching by Task ID: ' + JSON.stringify({location:'EmailHelper.gs:456',message:'findTaskEmailThread searching by Task ID',data:{taskId:taskId,searchQuery:searchQuery},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'C'}));
  // #endregion
  
  const threads = GmailApp.search(searchQuery, 0, 1);
  
  if (threads.length > 0) {
    const thread = threads[0];
    
    // Verify Task ID is in the first message
    const firstMessage = thread.getMessages()[0];
    if (firstMessage.getPlainBody().includes(`Task ID: ${taskId}`)) {
      // Store Primary_Thread_ID for future use
      updateTask(taskId, {
        Primary_Thread_ID: thread.getId()
      });
      
      Logger.log(`✓ Found thread by search, stored Primary_Thread_ID`);
      
      // #region agent log
      Logger.log('[DEBUG] findTaskEmailThread found by search: ' + JSON.stringify({location:'EmailHelper.gs:470',message:'findTaskEmailThread found by search',data:{taskId:taskId,threadId:thread.getId()},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'C'}));
      // #endregion
      
      return thread;
    }
  }
  
  Logger.log(`❌ No matching thread found for task ${taskId}`);
  
  // #region agent log
  Logger.log('[DEBUG] findTaskEmailThread no thread found: ' + JSON.stringify({location:'EmailHelper.gs:475',message:'findTaskEmailThread no thread found',data:{taskId:taskId,searchResultsCount:threads.length},timestamp:Date.now(),sessionId:'debug-session',runId:'run1',hypothesisId:'C'}));
  // #endregion
  
  return null;
}

/**
 * Get replies to a task assignment email
 * SIMPLIFIED: Uses Primary_Thread_ID and Task ID search
 * 1. Get thread by Primary_Thread_ID (fastest - direct access)
 * 2. Fallback search by Task ID (strict verification)
 * 3. Uses Processed_Message_IDs field instead of parsing log
 */
function getTaskReplies(taskId) {
  const task = getTask(taskId);
  if (!task || !task.Assignee_Email) {
    return [];
  }
  
  const assigneeEmail = task.Assignee_Email;
  const bossEmail = CONFIG.BOSS_EMAIL();
  const processedIds = getProcessedMessageIds(taskId);
  const allMessages = [];
  
  Logger.log(`Getting replies for task ${taskId}`);
  Logger.log(`Found ${processedIds.size} already processed message ID(s)`);
  
  // STEP 1: Get thread by Primary_Thread_ID (fastest - direct access)
  if (task.Primary_Thread_ID) {
    try {
      const thread = GmailApp.getThreadById(task.Primary_Thread_ID);
      if (thread) {
        const messages = thread.getMessages();
        Logger.log(`Found thread with ${messages.length} message(s)`);
        
        for (const message of messages) {
          const messageId = message.getId();
          const messageDate = message.getDate();
          
          // Skip if already processed
          if (processedIds.has(messageId)) {
            continue;
          }
          
          // Skip if before task creation
          if (task.Created_Date && messageDate < new Date(task.Created_Date)) {
            continue;
          }
          
          const senderEmail = extractEmailFromString(message.getFrom());
          const isBoss = senderEmail.toLowerCase() === bossEmail.toLowerCase();
          const isAssignee = senderEmail.toLowerCase() === assigneeEmail.toLowerCase();
          
          if (isBoss || isAssignee) {
            allMessages.push({
              message: message,
              messageId: messageId,
              date: messageDate,
              threadId: task.Primary_Thread_ID,
              source: 'primary_thread'
            });
          }
        }
      }
    } catch (error) {
      Logger.log(`Could not access thread ${task.Primary_Thread_ID}: ${error.toString()}`);
    }
  }
  
  Logger.log(`Extracted ${allMessages.length} message(s) from primary thread`);
  
  // STEP 2: Fallback search by Task ID (for emails not in primary thread)
  // Only search after Last_Reply_Check for efficiency
  const lastCheckDate = task.Last_Reply_Check ? new Date(task.Last_Reply_Check) : new Date(task.Created_Date);
  const searchDate = Utilities.formatDate(lastCheckDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const searchQuery = `"Task ID: ${taskId}" in:inbox after:${searchDate}`;
  
  Logger.log(`Searching emails after ${searchDate}`);
  
  try {
    const threads = GmailApp.search(searchQuery, 0, 5); // Limit to 5 threads
    Logger.log(`Found ${threads.length} thread(s) in search`);
    
    for (const thread of threads) {
      const threadId = thread.getId();
      
      // Skip if already processed via primary thread
      if (task.Primary_Thread_ID === threadId) {
        continue;
      }
      
      // STRICT VERIFICATION: Task ID must be in message body
      const messages = thread.getMessages();
      for (const message of messages) {
        const body = message.getPlainBody();
        if (!body.includes(`Task ID: ${taskId}`)) {
          continue; // Skip - doesn't match
        }
        
        const messageId = message.getId();
        const messageDate = message.getDate();
        
        // Skip if before task creation
        if (task.Created_Date && messageDate < new Date(task.Created_Date)) {
          continue;
        }
        
        // Skip if already processed
        if (processedIds.has(messageId)) {
          continue;
        }
        
        const senderEmail = extractEmailFromString(message.getFrom());
        const isBoss = senderEmail.toLowerCase() === bossEmail.toLowerCase();
        const isAssignee = senderEmail.toLowerCase() === assigneeEmail.toLowerCase();
        
        if (isBoss || isAssignee) {
          allMessages.push({
            message: message,
            messageId: messageId,
            date: messageDate,
            threadId: threadId,
            source: 'search_fallback'
          });
          
          // If this is a new thread, consider updating Primary_Thread_ID
          // (only if current Primary_Thread_ID is invalid)
          if (!task.Primary_Thread_ID) {
            updateTask(taskId, {
              Primary_Thread_ID: threadId
            });
            Logger.log(`Set Primary_Thread_ID to ${threadId} from search`);
          }
        }
      }
    }
  } catch (error) {
    Logger.log(`Error searching: ${error.toString()}`);
  }
  
  // Sort chronologically
  allMessages.sort((a, b) => a.date.getTime() - b.date.getTime());
  
  Logger.log(`Total unprocessed messages found: ${allMessages.length}`);
  
  // Update Last_Reply_Check timestamp
  updateTask(taskId, {
    Last_Reply_Check: new Date().toISOString(),
    Last_Updated: new Date()
  });
  
  // Return just the message objects
  return allMessages.map(msgData => msgData.message);
}

/**
 * Send email to employee requesting approval of new date
 * Includes the boss's message for context
 */
function sendEmployeeApprovalRequest(taskId, newDateFormatted, bossMessage) {
  try {
    const task = getTask(taskId);
    if (!task || !task.Assignee_Email) {
      return;
    }
    
    const staff = getStaff(task.Assignee_Email);
    const assigneeName = staff ? staff.Name : task.Assignee_Email.split('@')[0];
    const bossName = CONFIG.BOSS_NAME() || 'Boss';
    
    // Build email body
    let emailBody = `Hello ${assigneeName},\n\n`;
    emailBody += `We need to update the deadline for the following task:\n\n`;
    emailBody += `Task: ${task.Task_Name}\n`;
    emailBody += `New Due Date: ${newDateFormatted}\n\n`;
    
    // Include boss's message for context
    if (bossMessage) {
      emailBody += `Message from ${bossName}:\n`;
      emailBody += `"${bossMessage.substring(0, 500)}"\n\n`;
    }

    // Make the reply machine-decidable so "ACCEPTANCE" can be treated as confirmation.
    emailBody += `Please reply with ONE of the following:\n`;
    emailBody += `- CONFIRM → if you can commit to ${newDateFormatted}\n`;
    emailBody += `- REJECT → if you cannot commit (and optionally propose an alternate date)\n\n`;
    emailBody += `If you reply normally (e.g., "ok", "confirmed", "sure"), we will treat it as CONFIRM.\n\n`;
    emailBody += `Thank you,\n${CONFIG.EMAIL_SIGNATURE()}`;
    
    const subject = `Date Change Request: ${task.Task_Name}`;
    
    const result = sendEmailToAssignee(
      taskId,
      task.Assignee_Email,
      subject,
      emailBody,
      {
        htmlBody: emailBody.replace(/\n/g, '<br>')
      }
    );
    
    Logger.log(`Approval request email sent to ${task.Assignee_Email}${result.threadId ? ` (Thread ID: ${result.threadId})` : ''}`);
    logInteraction(taskId, `Employee approval request sent for new date: ${newDateFormatted}${result.threadId ? ` (Thread ID: ${result.threadId})` : ''}`);
    
  } catch (error) {
    Logger.log(`ERROR sending approval request: ${error.toString()}`);
    logError(ERROR_TYPE.API_ERROR, 'sendEmployeeApprovalRequest', error.toString(), taskId, error.stack);
  }
}

