/**
 * Email Negotiation Engine
 * Handles task assignment emails and reply processing
 */

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
    if (!task || task.Status !== TASK_STATUS.ASSIGNED) {
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
    const thread = message.getThread();
    const subject = message.getSubject();
    
    // Check if this is a reply to a task assignment
    if (!subject.includes('Task Assignment:')) {
      return;
    }
    
    // Extract task name from subject
    const taskNameMatch = subject.match(/Task Assignment: (.+)/);
    if (!taskNameMatch) {
      return;
    }
    
    const taskName = taskNameMatch[1];
    
    // Find task by name
    const tasks = findRowsByCondition(SHEETS.TASKS_DB, task => 
      task.Task_Name === taskName && 
      (task.Status === TASK_STATUS.ASSIGNED || task.Status === TASK_STATUS.ACTIVE)
    );
    
    if (tasks.length === 0) {
      Logger.log(`No task found for reply: ${taskName}`);
      return;
    }
    
    const task = tasks[0]; // Use first match
    const taskId = task.Task_ID;
    
    // Process the reply
    processReplyEmail(taskId, message);
    
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
    
    // Log the interaction with message ID
    logInteraction(taskId, `Reply received from ${senderEmail}: ${classification.type} (Message ID: ${messageId})`);
    
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
  
  // Remove quoted sections (common patterns):
  // - Lines starting with ">" (standard email quote)
  // - "On [date] [person] wrote:" followed by quoted text
  // - "From:" followed by quoted text
  // - "-----Original Message-----" and everything after
  // - "________________________________" (signature separator)
  
  // Remove everything after common quote markers
  const quoteMarkers = [
    /^On .+ wrote:.*$/m,  // "On Thu, Dec 25, 2025 at 10:05 PM ... wrote:"
    /^From:.*$/m,         // "From: ..."
    /^-----Original Message-----.*$/m,
    /^________________________________.*$/m,
    /^_{10,}.*$/m,        // Multiple underscores
    /^={10,}.*$/m,        // Multiple equals signs
  ];
  
  for (const marker of quoteMarkers) {
    const match = cleaned.match(marker);
    if (match) {
      const index = cleaned.indexOf(match[0]);
      if (index > 0) {
        cleaned = cleaned.substring(0, index).trim();
      }
    }
  }
  
  // Remove lines starting with ">" (quoted text)
  cleaned = cleaned.split('\n')
    .filter(line => !line.trim().startsWith('>'))
    .join('\n');
  
  // Remove email signatures (common patterns)
  const signaturePatterns = [
    /Best regards,.*$/is,
    /Regards,.*$/is,
    /Thanks,.*$/is,
    /Sincerely,.*$/is,
    /Sent from.*$/is,
    /Get Outlook.*$/is,
  ];
  
  for (const pattern of signaturePatterns) {
    cleaned = cleaned.replace(pattern, '');
  }
  
  return cleaned.trim();
}

/**
 * Handle acceptance reply
 */
function handleAcceptanceReply(taskId, emailContent, messageId) {
  Logger.log(`Handling ACCEPTANCE reply for task ${taskId}`);
  
  updateTask(taskId, {
    Status: TASK_STATUS.ACTIVE,
  });
  
  logInteraction(taskId, `Assignee accepted task (Message ID: ${messageId})`);
  
  // Optionally send confirmation to assignee
  try {
    const task = getTask(taskId);
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
  
  // Update task status to REVIEW_DATE so it shows up in Lovable
  updateTask(taskId, {
    Proposed_Date: finalProposedDate || '',
    Status: TASK_STATUS.REVIEW_DATE,
    Employee_Reply: emailContent, // Store full employee reply
  });
  
  Logger.log(`Task ${taskId} status updated to REVIEW_DATE`);
  logInteraction(taskId, `Assignee requested date change to ${finalProposedDate || 'unspecified date'} (Message ID: ${messageId})`);
  
  // Notify boss about date change request
  notifyBossOfReviewRequest(taskId, 'DATE_CHANGE', emailContent, finalProposedDate);
}

/**
 * Handle scope question reply
 */
function handleScopeQuestionReply(taskId, emailContent, messageId) {
  Logger.log(`Handling SCOPE_QUESTION reply for task ${taskId}`);
  
  const task = getTask(taskId);
  const currentLog = task.Interaction_Log || '';
  const newLog = currentLog + `\n\nScope question from assignee: ${emailContent.substring(0, 500)}`;
  
  updateTask(taskId, {
    Status: TASK_STATUS.REVIEW_SCOPE,
    Interaction_Log: newLog,
    Employee_Reply: emailContent, // Store full employee reply
  });
  
  logInteraction(taskId, `Assignee has scope questions (Message ID: ${messageId})`);
  
  // Notify boss about scope question
  notifyBossOfReviewRequest(taskId, 'SCOPE_QUESTION', emailContent);
}

/**
 * Handle role rejection reply
 */
function handleRoleRejectionReply(taskId, emailContent, messageId) {
  Logger.log(`Handling ROLE_REJECTION reply for task ${taskId}`);
  
  const task = getTask(taskId);
  const currentLog = task.Interaction_Log || '';
  const newLog = currentLog + `\n\nRole rejection from assignee: ${emailContent.substring(0, 500)}`;
  
  updateTask(taskId, {
    Status: TASK_STATUS.REVIEW_ROLE,
    Interaction_Log: newLog,
    Employee_Reply: emailContent, // Store full employee reply
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
  
  // For unclassified replies, still notify boss if task is in Assigned status
  if (task.Status === TASK_STATUS.ASSIGNED) {
    notifyBossOfReviewRequest(taskId, 'OTHER', emailContent);
  }
}

/**
 * Notify boss about review request
 */
function notifyBossOfReviewRequest(taskId, reviewType, emailContent, proposedDate = null) {
  try {
    const task = getTask(taskId);
    if (!task) return;
    
    const bossEmail = CONFIG.BOSS_EMAIL();
    if (!bossEmail) return;
    
    let subject = '';
    let body = '';
    
    switch (reviewType) {
      case 'DATE_CHANGE':
        subject = `Action Required: Date Change Request - ${task.Task_Name}`;
        body = `The assignee (${task.Assignee_Email}) has requested a different due date for this task.\n\n`;
        body += `Task: ${task.Task_Name}\n`;
        body += `Current Due Date: ${task.Due_Date || 'Not specified'}\n`;
        body += `Proposed Date: ${proposedDate || 'Not specified in email'}\n\n`;
        body += `Assignee's message:\n${emailContent.substring(0, 500)}\n\n`;
        body += `Please review and update the task in your dashboard.`;
        break;
        
      case 'SCOPE_QUESTION':
        subject = `Action Required: Scope Question - ${task.Task_Name}`;
        body = `The assignee (${task.Assignee_Email}) has questions about the scope of this task.\n\n`;
        body += `Task: ${task.Task_Name}\n\n`;
        body += `Assignee's message:\n${emailContent.substring(0, 500)}\n\n`;
        body += `Please review and provide clarification.`;
        break;
        
      case 'ROLE_REJECTION':
        subject = `Action Required: Role Rejection - ${task.Task_Name}`;
        body = `The assignee (${task.Assignee_Email}) claims this task is not their responsibility.\n\n`;
        body += `Task: ${task.Task_Name}\n\n`;
        body += `Assignee's message:\n${emailContent.substring(0, 500)}\n\n`;
        body += `Please review and reassign if needed.`;
        break;
        
      default:
        subject = `Action Required: Reply Received - ${task.Task_Name}`;
        body = `A reply was received from the assignee (${task.Assignee_Email}) that requires your attention.\n\n`;
        body += `Task: ${task.Task_Name}\n\n`;
        body += `Assignee's message:\n${emailContent.substring(0, 500)}\n\n`;
        body += `Please review the task in your dashboard.`;
    }
    
    body += `\n\nTask ID: ${taskId}\n`;
    body += `View in dashboard: [Link to your dashboard]`;
    
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
  }
}

/**
 * Extract date from text (helper function)
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
  
  // Try common date patterns (order matters - try more specific first)
  // Prioritize patterns that match "7th of Jan 2026" format
  const datePatterns = [
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
    // "10th Jan" or "10 Jan" format without year - assume current or next year
    {
      pattern: /(\d{1,2})(?:st|nd|rd|th)?\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b/i,
      parser: function(match) {
        const day = parseInt(match[1]);
        const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        const month = monthNames.indexOf(match[2].toLowerCase().substring(0, 3));
        if (month >= 0 && day >= 1 && day <= 31) {
          // If date has passed this year, assume next year
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
  ];
  
  // Use cleaned text for extraction
  for (const datePattern of datePatterns) {
    const match = cleanedText.match(datePattern.pattern);
    if (match) {
      try {
        Logger.log(`Found date pattern match: ${match[0]}`);
        const date = datePattern.parser(match);
        if (date && !isNaN(date.getTime())) {
          const formatted = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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
    
    // Get all assigned tasks
    const assignedTasks = getTasksByStatus(TASK_STATUS.ASSIGNED);
    
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
          
          // Update status to Review_Stagnation
          updateTask(task.Task_ID, {
            Status: TASK_STATUS.REVIEW_STAGNATION,
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
    
    // Get all assigned tasks and active tasks (in case they reply after accepting)
    const assignedTasks = getTasksByStatus(TASK_STATUS.ASSIGNED);
    const activeTasks = getTasksByStatus(TASK_STATUS.ACTIVE);
    const allTasks = [...assignedTasks, ...activeTasks];
    
    Logger.log(`Found ${allTasks.length} tasks to check (${assignedTasks.length} assigned, ${activeTasks.length} active)`);
    
    let processedCount = 0;
    
    allTasks.forEach(task => {
      try {
        // Find email thread for this task
        const thread = findTaskEmailThread(task.Task_ID);
        if (!thread) {
          Logger.log(`No email thread found for task ${task.Task_ID}`);
          return;
        }
        
        // Get replies
        const replies = getTaskReplies(task.Task_ID);
        
        if (replies.length === 0) {
          Logger.log(`No replies found for task ${task.Task_ID}`);
          return;
        }
        
        Logger.log(`Found ${replies.length} reply(ies) for task ${task.Task_ID}`);
        
        // Process most recent reply
        const latestReply = replies[replies.length - 1];
        
        // Check if we've already processed this message
        const messageId = latestReply.getId();
        const log = task.Interaction_Log || '';
        
        if (log.includes(`Message ID: ${messageId}`)) {
          Logger.log(`Message ${messageId} already processed for task ${task.Task_ID}`);
          return;
        }
        
        // Process the reply
        processReplyEmail(task.Task_ID, latestReply);
        processedCount++;
        
      } catch (error) {
        Logger.log(`Error processing task ${task.Task_ID}: ${error.toString()}`);
        logError(ERROR_TYPE.API_ERROR, 'checkForReplies', error.toString(), task.Task_ID, error.stack);
      }
    });
    
    Logger.log(`=== Check complete: Processed ${processedCount} new reply(ies) ===`);
    
  } catch (error) {
    Logger.log(`ERROR in checkForReplies: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
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
    
    const assignedTasks = getTasksByStatus(TASK_STATUS.ASSIGNED);
    const activeTasks = getTasksByStatus(TASK_STATUS.ACTIVE);
    const allTasks = [...assignedTasks, ...activeTasks];
    
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

