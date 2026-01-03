/**
 * Configuration Constants
 * Central configuration for the Chief of Staff AI System
 */

// Get configuration from Config sheet
function getConfig() {
  let spreadsheet;
  
  // Try to get active spreadsheet first (if script is bound to a sheet)
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      spreadsheet = active;
    }
  } catch (e) {
    // Not bound to spreadsheet, continue
  }
  
  // If null, it's a standalone script - use spreadsheet ID from Script Properties
  if (!spreadsheet) {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('SPREADSHEET_ID not found in Script Properties. Please run quickSetup() or setupStandaloneScript() first with your Spreadsheet ID.');
    }
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      throw new Error('Could not open spreadsheet with ID: ' + spreadsheetId + '. Error: ' + e.toString() + '\nPlease verify the Spreadsheet ID is correct.');
    }
  }
  
  if (!spreadsheet) {
    throw new Error('Could not access spreadsheet. Make sure you have run setupStandaloneScript() or the script is bound to a spreadsheet.');
  }
  
  const configSheet = spreadsheet.getSheetByName('Config');
  if (!configSheet) {
    throw new Error('Config sheet not found. Please run createAllSheets() first in your spreadsheet.');
  }
  
  const config = {};
  const data = configSheet.getDataRange().getValues();
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    if (key && value) {
      config[key] = value;
    }
  }
  
  return config;
}

// Get a specific config value
function getConfigValue(key, defaultValue = null) {
  try {
    const config = getConfig();
    return config[key] || defaultValue;
  } catch (error) {
    Logger.log('Error getting config value: ' + error);
    return defaultValue;
  }
}

// Set a config value in the Config sheet
function setConfigValue(key, value, description = '', category = 'System') {
  try {
    // Use the same logic as getConfig() to access the spreadsheet
    let spreadsheet;
    try {
      const active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) {
        spreadsheet = active;
      }
    } catch (e) {
      // Not bound to spreadsheet, continue
    }
    
    if (!spreadsheet) {
      const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      if (!spreadsheetId) {
        throw new Error('SPREADSHEET_ID not found in Script Properties. Please run quickSetup() first.');
      }
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    }
    
    if (!spreadsheet) {
      throw new Error('Could not access spreadsheet.');
    }
    
    const configSheet = spreadsheet.getSheetByName('Config');
    
    if (!configSheet) {
      throw new Error('Config sheet not found. Please run createAllSheets() first.');
    }
    
    const data = configSheet.getDataRange().getValues();
    let found = false;
    
    // Check if key already exists
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        // Update existing row
        configSheet.getRange(i + 1, 2).setValue(value);
        if (description) {
          configSheet.getRange(i + 1, 3).setValue(description);
        }
        if (category) {
          configSheet.getRange(i + 1, 4).setValue(category);
        }
        found = true;
        Logger.log(`Updated config value: ${key} = ${value}`);
        break;
      }
    }
    
    // If not found, add new row
    if (!found) {
      configSheet.appendRow([key, value, description, category]);
      Logger.log(`Added new config value: ${key} = ${value}`);
    }
    
    return true;
  } catch (error) {
    Logger.log('Error setting config value: ' + error.toString());
    throw error;
  }
}

// Configuration constants
const CONFIG = {
  // System
  BOSS_EMAIL: () => getConfigValue('BOSS_EMAIL', Session.getActiveUser().getEmail()),
  SPREADSHEET_ID: () => {
    // Try to get from active spreadsheet first (if bound)
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      return active.getId();
    }
    // Otherwise get from Config sheet
    return getConfigValue('SPREADSHEET_ID');
  },
  VOICE_INBOX_FOLDER_ID: () => getConfigValue('VOICE_INBOX_FOLDER_ID'),
  TASK_RECORDINGS_FOLDER_ID: () => getConfigValue('TASK_RECORDINGS_FOLDER_ID'),
  MEETING_LAKE_FOLDER_ID: () => getConfigValue('MEETING_LAKE_FOLDER_ID'),
  
  // Timing
  ESCALATION_FOLLOWUP_HOURS: () => parseInt(getConfigValue('ESCALATION_FOLLOWUP_HOURS', '24')),
  ESCALATION_BOSS_ALERT_HOURS: () => parseInt(getConfigValue('ESCALATION_BOSS_ALERT_HOURS', '48')),
  
  // Scheduling
  DEFAULT_MEETING_DURATION_MINUTES: () => parseInt(getConfigValue('DEFAULT_MEETING_DURATION_MINUTES', '30')),
  FOCUS_TIME_DURATION_MINUTES: () => parseInt(getConfigValue('FOCUS_TIME_DURATION_MINUTES', '60')),
  WORKING_HOURS_START: () => getConfigValue('WORKING_HOURS_START', '09:00'),
  WORKING_HOURS_END: () => getConfigValue('WORKING_HOURS_END', '17:00'),
  WEEKLY_MEETING_TITLE: () => getConfigValue('WEEKLY_MEETING_TITLE', 'Weekly Ops'),
  
  // AI
  AI_CONFIDENCE_THRESHOLD: () => parseFloat(getConfigValue('AI_CONFIDENCE_THRESHOLD', '0.6')),
  
  // Scoring
  RELIABILITY_UPDATE_INTERVAL_HOURS: () => parseInt(getConfigValue('RELIABILITY_UPDATE_INTERVAL_HOURS', '24')),
  
  // Email
  EMAIL_SIGNATURE: () => getConfigValue('EMAIL_SIGNATURE', '[Boss\'s Chief of Staff AI]'),
  BOSS_NAME: () => getConfigValue('BOSS_NAME', 'Boss'),
  
  // Notifications
  // If false, the system will still record DATE_CHANGE requests but will not email the boss about them.
  NOTIFY_BOSS_ON_DATE_CHANGE: () => getConfigValue('NOTIFY_BOSS_ON_DATE_CHANGE', 'false') === 'true',
  
  // Vertex AI
  VERTEX_AI_PROJECT_ID: () => getConfigValue('VERTEX_AI_PROJECT_ID'),
  VERTEX_AI_LOCATION: () => getConfigValue('VERTEX_AI_LOCATION', 'us-central1'),
  
  // Model names
  GEMINI_FLASH_MODEL: 'gemini-2.5-flash',
  GEMINI_PRO_MODEL: 'gemini-2.5-pro',
  GEMINI_PRO_FALLBACK_MODEL: 'gemini-1.5-pro',  // Fallback if 2.5 not available
  
  // Speech-to-Text API
  SPEECH_TO_TEXT_ENABLED: () => getConfigValue('SPEECH_TO_TEXT_ENABLED', 'true') === 'true',
  SPEECH_TO_TEXT_MODEL: () => getConfigValue('SPEECH_TO_TEXT_MODEL', 'latest_long'),
  SPEECH_TO_TEXT_LANGUAGE: () => getConfigValue('SPEECH_TO_TEXT_LANGUAGE', 'en-US'),
  SPEECH_TO_TEXT_ALTERNATIVE_LANGUAGES: () => {
    const altLangs = getConfigValue('SPEECH_TO_TEXT_ALTERNATIVE_LANGUAGES', '');
    return altLangs ? altLangs.split(',').map(lang => lang.trim()).filter(lang => lang) : [];
  },
};

// Sheet names
const SHEETS = {
  TASKS_DB: 'Tasks_DB',
  STAFF_DB: 'Staff_DB',
  PROJECTS_DB: 'Projects_DB',
  KNOWLEDGE_LAKE: 'Knowledge_Lake',
  CONFIG: 'Config',
  ERROR_LOG: 'Error_Log',
  VOICE_PROMPTS: 'VoicePrompts',
  EMAIL_PROMPTS: 'EmailPrompts',
  MOM_PROMPTS: 'MoMPrompts',
  WORKFLOWS: 'Workflows',
};

// Task statuses - LIFECYCLE-only status system
// Status tracks WHERE the task is in its lifecycle
// Conversation_State (separate field) tracks conversation/approval state
const TASK_STATUS = {
  // Pre-Active - Task not yet assigned or needs setup
  AI_ASSIST: 'ai_assist',             // Needs clarification before assignment
  NOT_ACTIVE: 'not_active',           // Assigned, awaiting first response
  PENDING_ACTION: 'pending_action',   // Legacy/compat: needs attention (mapped to slow_progress in new buckets)
  
  // Active - Task is in progress
  ON_TIME: 'on_time',                 // Active, on track
  SLOW_PROGRESS: 'slow_progress',     // Active, behind schedule
  
  // Done States
  COMPLETED: 'completed',             // Employee claims done, pending boss review
  CLOSED: 'closed',                   // Verified complete or cancelled
  
  // Paused States
  ON_HOLD: 'on_hold',                 // Temporarily paused
  SOMEDAY: 'someday',                 // Deferred to future

  // ------------------------------------------------------------------
  // Legacy review statuses (backward compatibility)
  // NOTE: New system uses lifecycle Status + Conversation_State.
  // These are kept so older code paths and existing sheet rows don't break.
  // ------------------------------------------------------------------
  REVIEW_DATE: 'review_date',
  REVIEW_DATE_BOSS_APPROVED: 'review_date_boss_approved',
  REVIEW_DATE_BOSS_REJECTED: 'review_date_boss_rejected',
  REVIEW_DATE_BOSS_PROPOSED: 'review_date_boss_proposed',
  REVIEW_SCOPE: 'review_scope',
  REVIEW_SCOPE_CLARIFIED: 'review_scope_clarified',
  REVIEW_ROLE: 'review_role',
};

// Conversation State - Derived from analyzing conversation history
// This determines what actions/UI to show, independent of lifecycle status
const CONVERSATION_STATE = {
  // Normal States
  ACTIVE: 'active',                           // Normal operation, no pending items
  UPDATE_RECEIVED: 'update_received',         // Employee sent update, FYI only
  
  // Approval Needed - Employee requested something
  CHANGE_REQUESTED: 'change_requested',       // Employee requested parameter change (date/scope/role)
  COMPLETION_PENDING: 'completion_pending',   // Employee claims done, needs verification
  BLOCKER_REPORTED: 'blocker_reported',       // Employee reported a blocker
  
  // Awaiting Response States
  AWAITING_EMPLOYEE: 'awaiting_employee',     // Boss sent message, waiting for employee
  AWAITING_CONFIRMATION: 'awaiting_confirmation', // Boss approved change, employee needs to confirm
  
  // Negotiation States
  BOSS_PROPOSED: 'boss_proposed',             // Boss proposed alternative (date, scope, etc.)
  NEGOTIATING: 'negotiating',                 // Active back-and-forth discussion
  
  // Resolution States
  RESOLVED: 'resolved',                       // Change was applied or issue resolved
  REJECTED: 'rejected',                       // Boss rejected request, conversation may continue
};

// Conversation states that require boss attention (employee requested something)
// Used by frontend "Needs Attention" bucket - cross-cutting filter
const ATTENTION_STATES = [
  CONVERSATION_STATE.CHANGE_REQUESTED,    // Employee requested date/scope/role change
  CONVERSATION_STATE.COMPLETION_PENDING,  // Employee claims task is done
  CONVERSATION_STATE.BLOCKER_REPORTED,    // Employee reported a blocker
  CONVERSATION_STATE.NEGOTIATING,         // Active back-and-forth discussion
];

// Helper to check if a conversation state needs boss attention
function taskNeedsAttention(conversationState) {
  return conversationState && ATTENTION_STATES.includes(conversationState);
}

// Legacy status mapping for backward compatibility with existing data
// Old review statuses now map to lifecycle statuses (conversation state handles the rest)
const LEGACY_STATUS_MAP = {
  'Draft': TASK_STATUS.AI_ASSIST,
  'New': TASK_STATUS.AI_ASSIST,
  'Assigned': TASK_STATUS.NOT_ACTIVE,
  'Active': TASK_STATUS.ON_TIME,
  'Done Pending Review': TASK_STATUS.COMPLETED,
  'Done': TASK_STATUS.CLOSED,
  'Review_AI_Assist': TASK_STATUS.AI_ASSIST,
  'Review_Date': TASK_STATUS.ON_TIME,             // Conversation state handles the approval flow
  'Review_Date_Boss_Approved': TASK_STATUS.ON_TIME,
  'Review_Date_Boss_Rejected': TASK_STATUS.ON_TIME,
  'Review_Date_Boss_Proposed': TASK_STATUS.ON_TIME,
  'Review_Scope': TASK_STATUS.ON_TIME,
  'Review_Scope_Clarified': TASK_STATUS.ON_TIME,
  'Review_Role': TASK_STATUS.ON_TIME,
  'Review_Stagnation': TASK_STATUS.SLOW_PROGRESS,
  'Review_Update': TASK_STATUS.ON_TIME,
  'Scheduled': TASK_STATUS.ON_TIME,
  'Scheduling_Conflict': TASK_STATUS.AI_ASSIST,
  'Cancelled': TASK_STATUS.CLOSED,
  'Reopened': TASK_STATUS.ON_TIME,
  // Lowercase variants
  'review_date': TASK_STATUS.ON_TIME,
  'review_date_boss_approved': TASK_STATUS.ON_TIME,
  'review_date_boss_rejected': TASK_STATUS.ON_TIME,
  'review_date_boss_proposed': TASK_STATUS.ON_TIME,
  'review_scope': TASK_STATUS.ON_TIME,
  'review_scope_clarified': TASK_STATUS.ON_TIME,
  'review_role': TASK_STATUS.ON_TIME,
  'pending_action': TASK_STATUS.SLOW_PROGRESS,
};

// Helper function to normalize legacy statuses to new system
function normalizeStatus(status) {
  if (!status) return TASK_STATUS.AI_ASSIST;
  // Check if it's already a new status
  if (Object.values(TASK_STATUS).includes(status)) {
    return status;
  }
  // Map legacy status to new
  return LEGACY_STATUS_MAP[status] || TASK_STATUS.AI_ASSIST;
}

// Meeting actions
const MEETING_ACTION = {
  ONE_ON_ONE: '1-on-1',
  WEEKLY: 'Weekly',
  SELF: 'Self',
};

// Source types for Knowledge_Lake
const SOURCE_TYPE = {
  MEETING_SUMMARY: 'Meeting_Summary',
  PROJECT_PLAN: 'Project_Plan',
  REFERENCE_DOC: 'Reference_Doc',
  EMAIL_THREAD: 'Email_Thread',
  DECISION_RECORD: 'Decision_Record',
  ACTION_PLAN: 'Action_Plan',
};

// Legacy APPROVAL_STATE - Maps to new CONVERSATION_STATE for backward compatibility
const APPROVAL_STATE = {
  NONE: CONVERSATION_STATE.ACTIVE,
  AWAITING_BOSS: CONVERSATION_STATE.CHANGE_REQUESTED,
  BOSS_APPROVED: CONVERSATION_STATE.AWAITING_CONFIRMATION,
  BOSS_REJECTED: CONVERSATION_STATE.REJECTED,
  BOSS_PROPOSED: CONVERSATION_STATE.BOSS_PROPOSED,
  EMPLOYEE_CONFIRMED: CONVERSATION_STATE.RESOLVED,
  NEGOTIATING: CONVERSATION_STATE.NEGOTIATING
};

// Error types
const ERROR_TYPE = {
  API_ERROR: 'API_ERROR',
  PERMISSION_ERROR: 'PERMISSION_ERROR',
  DATA_ERROR: 'DATA_ERROR',
  TIMEOUT_ERROR: 'TIMEOUT_ERROR',
  QUOTA_ERROR: 'QUOTA_ERROR',
  UNKNOWN_ERROR: 'UNKNOWN_ERROR',
};

