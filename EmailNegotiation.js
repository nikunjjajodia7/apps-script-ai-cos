/**
 * Email Negotiation Engine
 * Handles task assignment emails and reply processing
 */

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
    const emailContent = message.getPlainBody();
    const senderEmail = extractEmailFromString(message.getFrom());
    const messageId = message.getId();
    
    Logger.log(`Sender: ${senderEmail}`);
    Logger.log(`Message ID: ${messageId}`);
    Logger.log(`Email preview: ${emailContent.substring(0, 200)}...`);
    
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
    
    // Classify reply type using AI
    Logger.log('Classifying reply type using AI...');
    const classification = classifyReplyType(emailContent);
    Logger.log(`Reply classified as: ${classification.type} (confidence: ${classification.confidence})`);
    Logger.log(`Reasoning: ${classification.reasoning || 'N/A'}`);
    
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
  Logger.log(`Proposed date: ${proposedDate || 'Not extracted'}`);
  
  // Try to extract date from email content if AI didn't extract it
  let finalProposedDate = proposedDate;
  if (!finalProposedDate) {
    finalProposedDate = extractDateFromText(emailContent);
  }
  
  updateTask(taskId, {
    Proposed_Date: finalProposedDate || '',
    Status: TASK_STATUS.REVIEW_DATE,
  });
  
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
  // Try common date patterns
  const datePatterns = [
    /(\d{4}-\d{2}-\d{2})/,  // YYYY-MM-DD
    /(\d{1,2}\/\d{1,2}\/\d{4})/,  // MM/DD/YYYY
    /(\d{1,2}-\d{1,2}-\d{4})/,  // MM-DD-YYYY
    /(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}/i,
    /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},?\s+\d{4}/i,
  ];
  
  for (const pattern of datePatterns) {
    const match = text.match(pattern);
    if (match) {
      try {
        const date = new Date(match[1]);
        if (!isNaN(date.getTime())) {
          return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      } catch (e) {
        // Continue to next pattern
      }
    }
  }
  
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
    
    const classification = classifyReplyType(replyText);
    
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

