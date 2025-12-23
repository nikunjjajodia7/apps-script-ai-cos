/**
 * Vertex AI Gemini Integration
 * Functions for calling Gemini models via Vertex AI API
 */

/**
 * Call Vertex AI API using UrlFetchApp
 * Note: This requires proper authentication setup in Google Cloud
 */
function callVertexAI(modelName, prompt, options = {}) {
  try {
    const projectId = CONFIG.VERTEX_AI_PROJECT_ID();
    const location = CONFIG.VERTEX_AI_LOCATION();
    
    if (!projectId) {
      throw new Error('VERTEX_AI_PROJECT_ID not configured');
    }
    
    const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/${modelName}:generateContent`;
    
    const payload = {
      contents: [{
        role: 'user',
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: {
        temperature: options.temperature || 0.7,
        topK: options.topK || 40,
        topP: options.topP || 0.95,
        maxOutputTokens: options.maxOutputTokens || 8192,
      }
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    
    if (response.getResponseCode() !== 200) {
      const errorText = response.getContentText();
      logError(ERROR_TYPE.API_ERROR, 'callVertexAI', `Vertex AI API error: ${errorText}`);
      throw new Error(`Vertex AI API error: ${errorText}`);
    }
    
    const result = JSON.parse(response.getContentText());
    return result.candidates[0].content.parts[0].text;
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'callVertexAI', error.toString(), null, error.stack);
    throw error;
  }
}

/**
 * Call Gemini Flash for quick tasks
 */
function callGeminiFlash(prompt, options = {}) {
  return callVertexAI(CONFIG.GEMINI_FLASH_MODEL, prompt, {
    ...options,
    temperature: options.temperature || 0.3, // Lower temperature for more consistent results
  });
}

/**
 * Call Gemini Pro for complex reasoning
 */
function callGeminiPro(prompt, options = {}) {
  return callVertexAI(CONFIG.GEMINI_PRO_MODEL, prompt, {
    ...options,
    temperature: options.temperature || 0.7,
  });
}

/**
 * Transcribe audio file with speaker diarization using Google Cloud Speech-to-Text API
 * Returns transcript with speaker labels
 */
function transcribeAudioWithDiarization(audioFileId) {
  try {
    Logger.log('Starting audio transcription with speaker diarization...');
    const file = DriveApp.getFileById(audioFileId);
    const fileName = file.getName();
    Logger.log(`File: ${fileName}`);
    
    const audioBlob = file.getBlob();
    const mimeType = audioBlob.getContentType();
    Logger.log(`MIME type: ${mimeType}`);
    
    // Check file size (Speech-to-Text has limits)
    const fileSize = audioBlob.getBytes().length;
    Logger.log(`File size: ${(fileSize / 1024 / 1024).toFixed(2)} MB`);
    
    if (fileSize > 60 * 1024 * 1024) { // 60MB limit for Speech-to-Text
      throw new Error('Audio file too large (max 60MB for Speech-to-Text API)');
    }
    
    const base64Audio = Utilities.base64Encode(audioBlob.getBytes());
    Logger.log('Audio encoded to base64');
    
    const projectId = CONFIG.VERTEX_AI_PROJECT_ID();
    
    if (!projectId) {
      throw new Error('VERTEX_AI_PROJECT_ID not configured in Config sheet');
    }
    
    Logger.log(`Project ID: ${projectId}`);
    
    // Speech-to-Text API endpoint
    const url = `https://speech.googleapis.com/v1/projects/${projectId}/locations/global:recognize`;
    Logger.log(`API URL: ${url}`);
    
    // Map MIME types to Speech-to-Text encoding
    let encoding = 'LINEAR16';
    let sampleRateHertz = 16000;
    
    if (mimeType.includes('webm') || mimeType.includes('opus')) {
      encoding = 'WEBM_OPUS';
      sampleRateHertz = 48000;
    } else if (mimeType.includes('mp3')) {
      encoding = 'MP3';
      sampleRateHertz = 44100;
    } else if (mimeType.includes('m4a') || mimeType.includes('aac')) {
      encoding = 'MPEG4_AAC';
      sampleRateHertz = 44100;
    } else if (mimeType.includes('wav')) {
      encoding = 'LINEAR16';
      sampleRateHertz = 16000;
    }
    
    // Build config object
    const config = {
      encoding: encoding,
      sampleRateHertz: sampleRateHertz,
      languageCode: CONFIG.SPEECH_TO_TEXT_LANGUAGE(),
      model: CONFIG.SPEECH_TO_TEXT_MODEL(),
      enableSpeakerDiarization: true,
      diarizationSpeakerCount: 10, // Maximum number of speakers to detect
      enableAutomaticPunctuation: true,
      enableWordTimeOffsets: true,
    };
    
    // Add alternative language codes for multilingual support (e.g., Hinglish)
    const alternativeLanguages = CONFIG.SPEECH_TO_TEXT_ALTERNATIVE_LANGUAGES();
    if (alternativeLanguages && alternativeLanguages.length > 0) {
      config.alternativeLanguageCodes = alternativeLanguages;
      Logger.log(`Multilingual mode enabled. Primary: ${config.languageCode}, Alternatives: ${alternativeLanguages.join(', ')}`);
    }
    
    const payload = {
      config: config,
      audio: {
        content: base64Audio
      }
    };
    
    Logger.log('Sending request to Speech-to-Text API...');
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    
    const responseCode = response.getResponseCode();
    Logger.log(`Response code: ${responseCode}`);
    
    if (responseCode !== 200) {
      const errorText = response.getContentText();
      Logger.log(`Error response: ${errorText}`);
      logError(ERROR_TYPE.API_ERROR, 'transcribeAudioWithDiarization', `Speech-to-Text API error (${responseCode}): ${errorText}`);
      throw new Error(`Speech-to-Text API error (${responseCode}): ${errorText}`);
    }
    
    const result = JSON.parse(response.getContentText());
    Logger.log('Response parsed successfully');
    
    if (!result.results || result.results.length === 0) {
      Logger.log('No transcription results returned');
      throw new Error('No transcription results returned from Speech-to-Text API');
    }
    
    // Format transcript with speaker labels
    let transcriptWithSpeakers = '';
    const speakerSegments = [];
    
    // Extract all words with speaker tags
    result.results.forEach((result, resultIndex) => {
      if (result.alternatives && result.alternatives[0] && result.alternatives[0].words) {
        result.alternatives[0].words.forEach(word => {
          if (word.speakerTag !== undefined) {
            speakerSegments.push({
              speaker: word.speakerTag,
              word: word.word,
              startTime: word.startTime,
              endTime: word.endTime
            });
          }
        });
      }
    });
    
    // Group words by speaker and format
    if (speakerSegments.length > 0) {
      let currentSpeaker = speakerSegments[0].speaker;
      let currentSegment = `[Speaker ${currentSpeaker}]: ${speakerSegments[0].word}`;
      
      for (let i = 1; i < speakerSegments.length; i++) {
        const segment = speakerSegments[i];
        if (segment.speaker === currentSpeaker) {
          currentSegment += ' ' + segment.word;
        } else {
          transcriptWithSpeakers += currentSegment + '\n\n';
          currentSpeaker = segment.speaker;
          currentSegment = `[Speaker ${currentSpeaker}]: ${segment.word}`;
        }
      }
      transcriptWithSpeakers += currentSegment;
    } else {
      // Fallback: use alternative transcript without speaker labels
      result.results.forEach((result, index) => {
        if (result.alternatives && result.alternatives[0] && result.alternatives[0].transcript) {
          transcriptWithSpeakers += result.alternatives[0].transcript;
          if (index < result.results.length - 1) {
            transcriptWithSpeakers += ' ';
          }
        }
      });
    }
    
    Logger.log(`Transcript with speaker diarization received (${transcriptWithSpeakers.length} characters)`);
    Logger.log(`Detected ${new Set(speakerSegments.map(s => s.speaker)).size} speakers`);
    
    return transcriptWithSpeakers;
    
  } catch (error) {
    Logger.log('ERROR in transcribeAudioWithDiarization: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
    logError(ERROR_TYPE.API_ERROR, 'transcribeAudioWithDiarization', error.toString(), null, error.stack);
    throw error;
  }
}

/**
 * Transcribe audio file using Gemini (multimodal)
 * Note: This requires the audio file to be accessible
 * @deprecated Use transcribeAudioWithDiarization() for speaker identification
 * Falls back to this method if Speech-to-Text is disabled or fails
 */
function transcribeAudio(audioFileId) {
  // Check if Speech-to-Text is enabled, use it if available
  if (CONFIG.SPEECH_TO_TEXT_ENABLED()) {
    try {
      Logger.log('Speech-to-Text enabled, attempting diarization...');
      return transcribeAudioWithDiarization(audioFileId);
    } catch (error) {
      Logger.log('Speech-to-Text failed, falling back to Gemini: ' + error.toString());
      // Continue to Gemini fallback
    }
  }
  
  try {
    Logger.log('Starting audio transcription with Gemini...');
    const file = DriveApp.getFileById(audioFileId);
    const fileName = file.getName();
    Logger.log(`File: ${fileName}`);
    
    const audioBlob = file.getBlob();
    const mimeType = audioBlob.getContentType();
    Logger.log(`MIME type: ${mimeType}`);
    
    // Check file size (Vertex AI has limits)
    const fileSize = audioBlob.getBytes().length;
    Logger.log(`File size: ${(fileSize / 1024 / 1024).toFixed(2)} MB`);
    
    if (fileSize > 20 * 1024 * 1024) { // 20MB limit
      throw new Error('Audio file too large (max 20MB)');
    }
    
    const base64Audio = Utilities.base64Encode(audioBlob.getBytes());
    Logger.log('Audio encoded to base64');
    
    const projectId = CONFIG.VERTEX_AI_PROJECT_ID();
    const location = CONFIG.VERTEX_AI_LOCATION();
    
    if (!projectId) {
      throw new Error('VERTEX_AI_PROJECT_ID not configured in Config sheet');
    }
    
    Logger.log(`Project ID: ${projectId}, Location: ${location}`);
    
    const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/${CONFIG.GEMINI_FLASH_MODEL}:generateContent`;
    Logger.log(`API URL: ${url}`);
    
    const payload = {
      contents: [{
        role: 'user',
        parts: [
          {
            text: `Transcribe this audio file accurately. Follow these guidelines:
- Preserve all spoken words exactly as spoken
- Include proper punctuation (periods, commas, question marks)
- Preserve numbers, dates, and times in their spoken form
- Keep names, email addresses, and technical terms exactly as spoken
- Do not add commentary, explanations, or corrections
- If audio quality is poor, transcribe what you can hear and indicate unclear parts with [unclear]
- Return only the transcribed text, nothing else`
          },
          {
            inlineData: {
              mimeType: mimeType,
              data: base64Audio
            }
          }
        ]
      }],
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 4096,
      }
    };
    
    Logger.log('Sending request to Vertex AI...');
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    
    const responseCode = response.getResponseCode();
    Logger.log(`Response code: ${responseCode}`);
    
    if (responseCode !== 200) {
      const errorText = response.getContentText();
      Logger.log(`Error response: ${errorText}`);
      logError(ERROR_TYPE.API_ERROR, 'transcribeAudio', `Audio transcription error (${responseCode}): ${errorText}`);
      throw new Error(`Audio transcription error (${responseCode}): ${errorText}`);
    }
    
    const result = JSON.parse(response.getContentText());
    Logger.log('Response parsed successfully');
    
    if (!result.candidates || !result.candidates[0] || !result.candidates[0].content || !result.candidates[0].content.parts || !result.candidates[0].content.parts[0]) {
      Logger.log('Unexpected response structure: ' + JSON.stringify(result));
      throw new Error('Unexpected response structure from Vertex AI');
    }
    
    const transcript = result.candidates[0].content.parts[0].text;
    Logger.log(`Transcript received (${transcript.length} characters)`);
    
    return transcript;
    
  } catch (error) {
    Logger.log('ERROR in transcribeAudio: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
    logError(ERROR_TYPE.API_ERROR, 'transcribeAudio', error.toString(), null, error.stack);
    throw error;
  }
}

/**
 * Parse voice command and extract structured data
 */
function parseVoiceCommand(transcript) {
  const prompt = `You are an expert task management assistant. Parse the following voice command transcript and extract all relevant information with high accuracy.

Voice command transcript: "${transcript}"

Extract the following information and return ONLY a valid JSON object (no markdown, no code blocks, no explanations):

{
  "task_name": "concise, clear task title (3-8 words, action-oriented)",
  "assignee_email": "full email address if mentioned explicitly, or null",
  "assignee_name": "full name if mentioned (first and last name if available), or null",
  "due_date": "date in YYYY-MM-DD format if mentioned, or null",
  "due_date_text": "exact original text about due date/timing (e.g., 'next Monday', 'by Friday', 'in 2 days'), or null",
  "due_time": "time if mentioned (HH:MM format in 24-hour), or null",
  "priority": "High, Medium, or Low based on urgency indicators (urgent, ASAP, important = High; normal = Medium; low priority, whenever = Low), or null if not mentioned",
  "project_tag": "project name or tag if mentioned, or null",
  "context": "all additional details, background, requirements, constraints, or important notes mentioned",
  "tone": "urgent, normal, or calm based on speaker's tone and urgency words",
  "confidence": 0.0-1.0 confidence score (be accurate: 0.9+ if very clear, 0.7-0.9 if mostly clear, 0.5-0.7 if some ambiguity, <0.5 if unclear)",
  "ambiguities": ["specific list of what is unclear or ambiguous (e.g., 'assignee name unclear', 'due date format ambiguous')"]
}

EXTRACTION GUIDELINES:

1. Task Name:
   - Extract the core action/objective (e.g., "Review Q4 report" not "I need to review the Q4 report")
   - Remove filler words like "please", "I want", "can you"
   - Keep it concise but descriptive

2. Assignee:
   - Look for: "assign to [name]", "give to [name]", "[name] should", "[name] will", "send to [name]"
   - Extract full name if possible (first + last)
   - If only first name mentioned, use just first name
   - Email: only if explicitly stated (e.g., "john@company.com")

3. Due Date:
   - Relative dates: "today" = today's date, "tomorrow" = tomorrow, "next Monday" = next Monday, "Friday" = this or next Friday
   - Absolute dates: "January 15th", "15th of January", "01/15/2024" → convert to YYYY-MM-DD
   - Preserve original text in due_date_text for reference
   - If time mentioned (e.g., "by 3 PM", "before noon"), extract in due_time field

4. Priority:
   - High: urgent, ASAP, immediately, critical, important, high priority, rush
   - Medium: normal, standard, regular (default if not mentioned)
   - Low: low priority, whenever, no rush, eventually

5. Project/Tag:
   - Look for: "for [project]", "related to [project]", "under [project]", project names mentioned

6. Context:
   - Include: why the task exists, background info, specific requirements, constraints, dependencies
   - Include: location, tools needed, people involved, budget constraints, etc.
   - Be comprehensive but concise

7. Tone:
   - Urgent: rushed speech, urgent words, time pressure mentioned
   - Normal: standard conversational tone
   - Calm: relaxed, no rush indicated

8. Confidence:
   - 0.9-1.0: All key info clear (task, assignee, date all clear)
   - 0.7-0.9: Most info clear, minor ambiguities
   - 0.5-0.7: Some key info missing or unclear
   - <0.5: Major ambiguities or unclear command

9. Ambiguities:
   - List specific unclear elements: "assignee name unclear", "due date ambiguous", "task scope vague"

EXAMPLES:

Input: "Please assign John Smith to review the quarterly financial report by next Friday, it's urgent"
Output: {
  "task_name": "Review quarterly financial report",
  "assignee_email": null,
  "assignee_name": "John Smith",
  "due_date": "2024-01-19",
  "due_date_text": "next Friday",
  "due_time": null,
  "priority": "High",
  "project_tag": null,
  "context": "Quarterly financial report needs review",
  "tone": "urgent",
  "confidence": 0.95,
  "ambiguities": []
}

Input: "I need someone to update the website homepage, maybe by end of week"
Output: {
  "task_name": "Update website homepage",
  "assignee_email": null,
  "assignee_name": null,
  "due_date": "2024-01-19",
  "due_date_text": "end of week",
  "due_time": null,
  "priority": "Medium",
  "project_tag": null,
  "context": "Website homepage needs updating",
  "tone": "normal",
  "confidence": 0.75,
  "ambiguities": ["assignee not specified"]
}

Now parse this voice command: "${transcript}"`;

  try {
    const response = callGeminiFlash(prompt);
    
    // Extract JSON from response (handle markdown code blocks if present)
    let jsonText = response.trim();
    if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```json\n?/, '').replace(/```$/, '');
    }
    if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```\n?/, '').replace(/```$/, '');
    }
    
    const parsed = JSON.parse(jsonText);
    
    // Process due date if it's text
    if (parsed.due_date_text && !parsed.due_date) {
      parsed.due_date = parseRelativeDate(parsed.due_date_text);
    }
    
    return parsed;
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'parseVoiceCommand', error.toString());
    return {
      task_name: transcript,
      assignee_email: null,
      assignee_name: null,
      due_date: null,
      context: '',
      tone: 'normal',
      confidence: 0.3,
      ambiguities: ['Failed to parse command']
    };
  }
}

/**
 * Parse relative date text to actual date
 */
function parseRelativeDate(dateText) {
  const today = new Date();
  const lowerText = dateText.toLowerCase();
  
  // Simple relative date parsing
  if (lowerText.includes('today')) {
    return Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (lowerText.includes('tomorrow')) {
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    return Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (lowerText.includes('next monday') || lowerText.includes('next week')) {
    const daysUntilMonday = (8 - today.getDay()) % 7 || 7;
    const nextMonday = new Date(today);
    nextMonday.setDate(today.getDate() + daysUntilMonday);
    return Utilities.formatDate(nextMonday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  
  // Try to parse as date
  try {
    const parsed = new Date(dateText);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  } catch (e) {
    // Ignore
  }
  
  return null;
}

/**
 * Generate task assignment email using Gemini Pro
 */
function generateAssignmentEmail(task) {
  const staff = getStaff(task.Assignee_Email);
  const assigneeName = staff ? staff.Name : task.Assignee_Email;
  
  const prompt = `You are a professional Chief of Staff assistant. Write a formal but friendly email assigning a task to a team member.

Task Name: ${task.Task_Name}
Assignee: ${assigneeName}
Due Date: ${task.Due_Date || 'Not specified'}
Context: ${task.Context_Hidden || task.Interaction_Log || 'No additional context'}

Write an email that:
1. Is polite but authoritative
2. Clearly states the task and expectations
3. Asks for confirmation of the deadline
4. Invites questions or concerns
5. Uses a professional but friendly tone

Return ONLY the email body text (no subject line, no signature - those will be added separately).`;

  try {
    const emailBody = callGeminiPro(prompt, { temperature: 0.7 });
    return emailBody.trim();
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'generateAssignmentEmail', error.toString(), task.Task_ID);
    // Fallback to template
    return getEmailTemplate('assignment', task);
  }
}

/**
 * Classify email reply type
 */
function classifyReplyType(replyContent) {
  const prompt = `You are an expert at analyzing email replies about task assignments. Classify this email reply into one of these categories:

1. ACCEPTANCE - The assignee accepts the task and agrees to the deadline. Examples: "I accept", "I'll do it", "Got it", "Will complete by deadline", "Acknowledged" (if accepting)
2. DATE_CHANGE - The assignee requests a different due date or mentions a date that's not feasible. Look for: "not feasible", "can't make that date", "need more time", "deadline is", "by [date]", "feasible as deadline", "propose [date]", "suggest [date]", "prefer [date]"
3. SCOPE_QUESTION - The assignee has questions about what needs to be done, needs clarification, or asks "what does this mean", "can you clarify", "I need more details"
4. ROLE_REJECTION - The assignee says this is not their job, responsibility, or role. Examples: "not my responsibility", "wrong person", "not my role", "should be assigned to"
5. OTHER - Something else that doesn't fit the above categories

IMPORTANT: If the email mentions a date that's different from the original deadline, or says a date is "not feasible" and proposes another date, classify as DATE_CHANGE.

Email reply: "${replyContent}"

Return ONLY a valid JSON object (no markdown, no code blocks):
{
  "type": "ACCEPTANCE|DATE_CHANGE|SCOPE_QUESTION|ROLE_REJECTION|OTHER",
  "confidence": 0.0-1.0,
  "extracted_date": "proposed date in YYYY-MM-DD format if DATE_CHANGE (convert dates like '10th Jan', 'Jan 10', '10/01/2025' to YYYY-MM-DD), or null",
  "reasoning": "brief explanation of why this classification was chosen"
}`;

  let response = null;
  try {
    response = callGeminiPro(prompt, { temperature: 0.2 });
    let jsonText = response.trim();
    
    // Remove markdown code blocks if present
    if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```json\n?/, '').replace(/```$/, '');
      jsonText = jsonText.replace(/^```\n?/, '').replace(/```$/, '');
    }
    
    // Try to extract JSON if wrapped in text
    const jsonMatch = jsonText.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      jsonText = jsonMatch[0];
    }
    
    const result = JSON.parse(jsonText);
    
    // Fallback: If classification failed but we detect a date, try DATE_CHANGE
    if (result.type === 'OTHER' && result.confidence < 0.7) {
      const datePatterns = [
        /(\d{1,2}(?:st|nd|rd|th)?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})/i,
        /((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4})/i,
        /(\d{1,2}\/\d{1,2}\/\d{4})/,
        /(\d{4}-\d{2}-\d{2})/,
        /(feasible|deadline|by|before|until).*?(\d{1,2}(?:st|nd|rd|th)?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*)/i,
      ];
      
      for (const pattern of datePatterns) {
        const match = replyContent.match(pattern);
        if (match) {
          Logger.log(`Fallback: Detected date pattern in OTHER classification: ${match[0]}`);
          result.type = 'DATE_CHANGE';
          result.confidence = 0.6;
          result.reasoning = 'Date detected in email content';
          break;
        }
      }
    }
    
    return result;
  } catch (error) {
    Logger.log(`ERROR in classifyReplyType: ${error.toString()}`);
    Logger.log(`Response was: ${response ? response.substring(0, 200) : 'null'}`);
    logError(ERROR_TYPE.API_ERROR, 'classifyReplyType', error.toString());
    
    // Fallback: Try to detect date change manually
    const datePatterns = [
      /(\d{1,2}(?:st|nd|rd|th)?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})/i,
      /((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4})/i,
      /(not feasible|can't make|need more time|deadline|by|before).*?(\d{1,2}(?:st|nd|rd|th)?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*)/i,
    ];
    
    for (const pattern of datePatterns) {
      const match = replyContent.match(pattern);
      if (match) {
        Logger.log(`Fallback detection: Found date change pattern: ${match[0]}`);
        return {
          type: 'DATE_CHANGE',
          confidence: 0.7,
          extracted_date: null, // Will be extracted by extractDateFromText
          reasoning: 'Date change detected via fallback pattern matching'
        };
      }
    }
    
    return {
      type: 'OTHER',
      confidence: 0.5,
      extracted_date: null,
      reasoning: 'Failed to classify reply - no patterns detected'
    };
  }
}

/**
 * Analyze meeting transcript and extract structured data
 */
function analyzeMeetingTranscript(transcript) {
  const prompt = `Analyze this meeting transcript and extract structured information.

Meeting Transcript:
"${transcript}"

Extract and return ONLY a valid JSON object:
{
  "executive_summary": "brief overview of meeting purpose and conclusions",
  "decisions_made": ["decision 1", "decision 2", ...],
  "action_items": [
    {
      "description": "action item description",
      "owner": "person responsible (name or email if mentioned)",
      "deadline": "deadline if mentioned (YYYY-MM-DD or null)"
    }
  ],
  "risks_sentiment": "notable risks or general sentiment/tone of meeting"
}`;

  try {
    const response = callGeminiPro(prompt, { 
      temperature: 0.7,
      maxOutputTokens: 8192 
    });
    
    let jsonText = response.trim();
    if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```json\n?/, '').replace(/```$/, '');
    }
    
    return JSON.parse(jsonText);
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'analyzeMeetingTranscript', error.toString());
    throw error;
  }
}

