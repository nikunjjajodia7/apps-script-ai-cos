/**
 * Meeting Intelligence Pipeline
 * Processes meeting audio recordings and extracts action items
 */

/**
 * Drive trigger handler for new files in Meeting_Lake
 */
function onMeetingFileAdded(e) {
  try {
    const file = e.source;
    const fileId = file.getId();
    
    // Check if file is in Meeting_Lake folder
    const meetingLakeFolderId = CONFIG.MEETING_LAKE_FOLDER_ID();
    if (!meetingLakeFolderId) {
      Logger.log('MEETING_LAKE_FOLDER_ID not configured');
      return;
    }
    
    const folders = file.getParents();
    let isInMeetingLake = false;
    
    folders.forEach(folder => {
      if (folder.getId() === meetingLakeFolderId) {
        isInMeetingLake = true;
      }
    });
    
    if (isInMeetingLake) {
      // Process the meeting file
      processMeetingAudio(fileId);
    }
  } catch (error) {
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'onMeetingFileAdded', error.toString(), null, error.stack);
  }
}

/**
 * Main function to process meeting audio
 */
function processMeetingAudio(fileId) {
  try {
    Logger.log(`Processing meeting audio: ${fileId}`);
    
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    
    // Transcribe meeting audio
    let transcript;
    try {
      transcript = transcribeAudio(fileId);
      Logger.log(`Transcript length: ${transcript.length} characters`);
    } catch (error) {
      logError(ERROR_TYPE.API_ERROR, 'processMeetingAudio', `Transcription failed: ${error.toString()}`, null, error.stack);
      throw error;
    }
    
    // Analyze transcript
    const meetingData = analyzeMeetingTranscript(transcript);
    Logger.log(`Meeting data extracted: ${JSON.stringify(meetingData)}`);
    
    // Create meeting summary document (Pipeline A)
    const summaryDoc = createMeetingSummaryDoc(meetingData, fileName);
    
    // Add to Knowledge_Lake
    addToKnowledgeLake(summaryDoc.getUrl(), meetingData.executive_summary, fileName);
    
    // Process action items (Pipeline B)
    createTasksFromActionItems(meetingData.action_items, {
      meetingDate: new Date(),
      meetingName: fileName,
      meetingDocUrl: summaryDoc.getUrl(),
    });
    
    Logger.log(`Meeting processed successfully. Summary doc: ${summaryDoc.getUrl()}`);
    
  } catch (error) {
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'processMeetingAudio', error.toString(), null, error.stack);
  }
}

/**
 * Create meeting summary document
 */
function createMeetingSummaryDoc(meetingData, audioFileName) {
  try {
    const doc = DocumentApp.create(`Meeting Summary - ${audioFileName} - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}`);
    const body = doc.getBody();
    
    // Title
    body.appendParagraph('Meeting Summary').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`Date: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy')}`);
    body.appendParagraph(`Source: ${audioFileName}`);
    body.appendParagraph('');
    
    // Executive Summary
    body.appendParagraph('Executive Summary').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(meetingData.executive_summary || 'No summary available.');
    body.appendParagraph('');
    
    // Decisions Made
    if (meetingData.decisions_made && meetingData.decisions_made.length > 0) {
      body.appendParagraph('Decisions Made').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      const decisionsList = body.appendListItem(meetingData.decisions_made[0]);
      for (let i = 1; i < meetingData.decisions_made.length; i++) {
        decisionsList.appendListItem(meetingData.decisions_made[i]);
      }
      body.appendParagraph('');
    }
    
    // Action Items
    if (meetingData.action_items && meetingData.action_items.length > 0) {
      body.appendParagraph('Action Items').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      meetingData.action_items.forEach(item => {
        const itemText = `${item.description} - Owner: ${item.owner || 'TBD'} - Deadline: ${item.deadline || 'TBD'}`;
        body.appendListItem(itemText);
      });
      body.appendParagraph('');
    }
    
    // Risks & Sentiment
    if (meetingData.risks_sentiment) {
      body.appendParagraph('Risks & Sentiment').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(meetingData.risks_sentiment);
    }
    
    // Save and return
    doc.saveAndClose();
    return doc;
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'createMeetingSummaryDoc', error.toString(), null, error.stack);
    throw error;
  }
}

/**
 * Add entry to Knowledge_Lake
 */
function addToKnowledgeLake(docUrl, summary, sourceName) {
  try {
    const infoId = `INFO-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss')}`;
    
    const knowledgeData = {
      Info_ID: infoId,
      Link: docUrl,
      Source_Type: SOURCE_TYPE.MEETING_SUMMARY,
      Summary: summary || `Meeting summary: ${sourceName}`,
      Created_Date: new Date(),
      Meeting_Date: new Date(),
      Related_Tasks: '', // Will be updated when tasks are created
      Tags: '',
    };
    
    addRow(SHEETS.KNOWLEDGE_LAKE, knowledgeData);
    return infoId;
  } catch (error) {
    logError(ERROR_TYPE.DATA_ERROR, 'addToKnowledgeLake', error.toString(), null, error.stack);
    throw error;
  }
}

/**
 * Create tasks from action items
 */
function createTasksFromActionItems(actionItems, meetingContext) {
  try {
    if (!actionItems || actionItems.length === 0) {
      Logger.log('No action items to process');
      return [];
    }
    
    const createdTaskIds = [];
    
    actionItems.forEach((item, index) => {
      try {
        // Resolve owner email if name was provided
        let assigneeEmail = null;
        if (item.owner) {
          // Try to find staff by name
          const staff = getSheetData(SHEETS.STAFF_DB);
          const matchedStaff = staff.find(s => 
            s.Name && s.Name.toLowerCase().includes(item.owner.toLowerCase())
          );
          if (matchedStaff) {
            assigneeEmail = matchedStaff.Email;
          }
        }
        
        // Create task
        const taskData = {
          Task_Name: item.description || `Action Item ${index + 1}`,
          Status: assigneeEmail ? TASK_STATUS.DRAFT : TASK_STATUS.REVIEW_AI_ASSIST,
          Assignee_Email: assigneeEmail || '',
          Due_Date: item.deadline || '',
          Context_Hidden: `From meeting: ${meetingContext.meetingName} on ${Utilities.formatDate(meetingContext.meetingDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}`,
          Created_By: 'Meeting',
          Priority: 'Medium',
        };
        
        const taskId = createTask(taskData);
        createdTaskIds.push(taskId);
        
        // Link to meeting in Knowledge_Lake
        if (meetingContext.meetingDocUrl) {
          const knowledgeEntries = findRowsByCondition(SHEETS.KNOWLEDGE_LAKE, entry => 
            entry.Link === meetingContext.meetingDocUrl
          );
          if (knowledgeEntries.length > 0) {
            const knowledgeEntry = knowledgeEntries[0];
            const relatedTasks = knowledgeEntry.Related_Tasks || '';
            const updatedTasks = relatedTasks ? `${relatedTasks}, ${taskId}` : taskId;
            updateRowByValue(SHEETS.KNOWLEDGE_LAKE, 'Info_ID', knowledgeEntry.Info_ID, {
              Related_Tasks: updatedTasks,
            });
          }
        }
        
        Logger.log(`Created task from action item: ${taskId}`);
        
      } catch (error) {
        logError(ERROR_TYPE.DATA_ERROR, 'createTasksFromActionItems', `Failed to create task for action item ${index}: ${error.toString()}`, null, error.stack);
      }
    });
    
    return createdTaskIds;
    
  } catch (error) {
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'createTasksFromActionItems', error.toString(), null, error.stack);
    return [];
  }
}

/**
 * Match existing tasks mentioned in meeting
 */
function matchExistingTasks(actionItems) {
  // This would use AI to match action items to existing tasks
  // For MVP, we'll keep it simple and create new tasks
  // Future enhancement: use Gemini to match by name/description
  return [];
}

/**
 * Update existing tasks based on meeting discussion
 */
function updateTasksFromMeeting(taskIds, updates) {
  taskIds.forEach(taskId => {
    try {
      updateTask(taskId, updates);
      logInteraction(taskId, `Updated from meeting: ${JSON.stringify(updates)}`);
    } catch (error) {
      logError(ERROR_TYPE.DATA_ERROR, 'updateTasksFromMeeting', error.toString(), taskId);
    }
  });
}

// ============================================
// TEST FUNCTIONS
// ============================================

/**
 * Test function: Process the most recent meeting file
 */
function testProcessMeetingAudio() {
  try {
    Logger.log('=== Testing Meeting Audio Processing ===');
    
    const meetingLakeFolderId = CONFIG.MEETING_LAKE_FOLDER_ID();
    if (!meetingLakeFolderId) {
      Logger.log('ERROR: MEETING_LAKE_FOLDER_ID not configured');
      Logger.log('Please set MEETING_LAKE_FOLDER_ID in your Config sheet');
      return;
    }
    
    const folder = DriveApp.getFolderById(meetingLakeFolderId);
    const files = folder.getFiles();
    
    let latestFile = null;
    let latestDate = new Date(0);
    
    // Find the most recent file
    while (files.hasNext()) {
      const file = files.next();
      const lastModified = file.getLastUpdated();
      if (lastModified > latestDate) {
        latestDate = lastModified;
        latestFile = file;
      }
    }
    
    if (!latestFile) {
      Logger.log('No files found in Meeting_Lake folder');
      Logger.log(`Folder ID: ${meetingLakeFolderId}`);
      return;
    }
    
    Logger.log(`Processing file: ${latestFile.getName()}`);
    Logger.log(`Modified: ${latestDate}`);
    
    processMeetingAudio(latestFile.getId());
    
    Logger.log('✓ Meeting processed successfully!');
    Logger.log('Check your Google Drive for the meeting summary document');
    Logger.log('Check your Tasks_DB sheet for new tasks created from action items');
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

/**
 * Test function: Process a specific meeting file by name
 * Usage: testProcessMeetingFileByName('Meeting Recording.m4a')
 */
function testProcessMeetingFileByName(fileName) {
  try {
    Logger.log(`=== Processing Meeting File: ${fileName} ===`);
    
    if (!fileName) {
      Logger.log('ERROR: No file name provided');
      Logger.log('Usage: testProcessMeetingFileByName("Meeting Recording.m4a")');
      return;
    }
    
    const meetingLakeFolderId = CONFIG.MEETING_LAKE_FOLDER_ID();
    if (!meetingLakeFolderId) {
      Logger.log('ERROR: MEETING_LAKE_FOLDER_ID not configured');
      return;
    }
    
    const folder = DriveApp.getFolderById(meetingLakeFolderId);
    const files = folder.getFiles();
    
    let foundFile = null;
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName() === fileName) {
        foundFile = file;
        break;
      }
    }
    
    if (!foundFile) {
      Logger.log(`ERROR: File "${fileName}" not found in Meeting_Lake folder`);
      Logger.log('Available files:');
      const allFiles = folder.getFiles();
      let fileList = [];
      while (allFiles.hasNext()) {
        fileList.push(allFiles.next().getName());
      }
      if (fileList.length === 0) {
        Logger.log('  (No files found)');
      } else {
        fileList.forEach(name => Logger.log(`  - ${name}`));
      }
      return;
    }
    
    Logger.log(`Found file: ${foundFile.getName()}`);
    Logger.log(`File ID: ${foundFile.getId()}`);
    Logger.log(`Modified: ${foundFile.getLastUpdated()}`);
    
    processMeetingAudio(foundFile.getId());
    
    Logger.log('✓ Meeting processed successfully!');
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

/**
 * Test function: List all files in Meeting_Lake folder
 */
function testListMeetingLakeFiles() {
  try {
    Logger.log('=== Files in Meeting_Lake Folder ===');
    
    const meetingLakeFolderId = CONFIG.MEETING_LAKE_FOLDER_ID();
    if (!meetingLakeFolderId) {
      Logger.log('ERROR: MEETING_LAKE_FOLDER_ID not configured');
      Logger.log('Please set MEETING_LAKE_FOLDER_ID in your Config sheet');
      return;
    }
    
    Logger.log(`Folder ID: ${meetingLakeFolderId}`);
    
    const folder = DriveApp.getFolderById(meetingLakeFolderId);
    Logger.log(`Folder name: ${folder.getName()}`);
    Logger.log(`Folder URL: ${folder.getUrl()}`);
    
    const files = folder.getFiles();
    let fileCount = 0;
    const fileList = [];
    
    while (files.hasNext()) {
      const file = files.next();
      fileCount++;
      fileList.push({
        name: file.getName(),
        id: file.getId(),
        modified: file.getLastUpdated(),
        size: file.getSize(),
        type: file.getMimeType()
      });
    }
    
    if (fileCount === 0) {
      Logger.log('\nNo files found in Meeting_Lake folder');
      Logger.log('Upload a meeting audio file to test the system');
      return;
    }
    
    Logger.log(`\nFound ${fileCount} file(s):\n`);
    
    fileList.forEach((file, index) => {
      Logger.log(`${index + 1}. ${file.name}`);
      Logger.log(`   ID: ${file.id}`);
      Logger.log(`   Modified: ${file.modified}`);
      Logger.log(`   Size: ${(file.size / 1024 / 1024).toFixed(2)} MB`);
      Logger.log(`   Type: ${file.type}`);
      Logger.log('');
    });
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

/**
 * Test function: Check Meeting_Lake folder access
 */
function testMeetingLakeAccess() {
  try {
    Logger.log('=== Testing Meeting_Lake Access ===');
    
    const meetingLakeFolderId = CONFIG.MEETING_LAKE_FOLDER_ID();
    Logger.log(`1. MEETING_LAKE_FOLDER_ID from Config: ${meetingLakeFolderId || 'NOT SET'}`);
    
    if (!meetingLakeFolderId) {
      Logger.log('ERROR: MEETING_LAKE_FOLDER_ID not configured');
      Logger.log('Please set MEETING_LAKE_FOLDER_ID in your Config sheet');
      return;
    }
    
    Logger.log('2. Attempting to access folder...');
    const folder = DriveApp.getFolderById(meetingLakeFolderId);
    Logger.log(`   ✓ Folder found: ${folder.getName()}`);
    Logger.log(`   Folder URL: ${folder.getUrl()}`);
    
    Logger.log('3. Listing files...');
    const files = folder.getFiles();
    let fileCount = 0;
    while (files.hasNext()) {
      files.next();
      fileCount++;
    }
    Logger.log(`   Found ${fileCount} file(s) in folder`);
    
    Logger.log('\n=== Access Test Complete ===');
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

