/**
 * Vertex AI Gemini Integration
 * Functions for calling Gemini models via Vertex AI API
 */

// ============================================================
// MODEL DISCOVERY AND TESTING FUNCTIONS
// ============================================================

/**
 * List all available Gemini models in your Vertex AI project
 * Run this function from the Apps Script editor to see available models
 */
function listAvailableGeminiModels() {
  try {
    const projectId = CONFIG.VERTEX_AI_PROJECT_ID();
    const location = CONFIG.VERTEX_AI_LOCATION();
    
    if (!projectId) {
      Logger.log('ERROR: VERTEX_AI_PROJECT_ID not configured');
      return null;
    }
    
    Logger.log('=== Listing Available Gemini Models ===');
    Logger.log(`Project: ${projectId}`);
    Logger.log(`Location: ${location}`);
    Logger.log('');
    
    // List publisher models (Gemini models are publisher models from Google)
    const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json',
      },
      muteHttpExceptions: true,
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      Logger.log(`Error (${responseCode}): ${responseText}`);
      return null;
    }
    
    const result = JSON.parse(responseText);
    
    if (result.models) {
      // Filter for Gemini models only
      const geminiModels = result.models.filter(m => 
        m.name && m.name.toLowerCase().includes('gemini')
      );
      
      Logger.log(`Found ${geminiModels.length} Gemini models:`);
      Logger.log('');
      
      geminiModels.forEach(model => {
        const modelId = model.name.split('/').pop();
        Logger.log(`📦 ${modelId}`);
        if (model.displayName) Logger.log(`   Display: ${model.displayName}`);
        if (model.description) Logger.log(`   Desc: ${model.description.substring(0, 100)}...`);
      });
      
      return geminiModels;
    } else {
      Logger.log('No models found in response');
      return [];
    }
    
  } catch (error) {
    Logger.log('Error listing models: ' + error.toString());
    return null;
  }
}

/**
 * Test a specific model to see if it's available and working
 * @param {string} modelName - The model name (e.g., 'gemini-2.0-flash')
 * @returns {object} - Test result with success status and details
 */
function testGeminiModel(modelName) {
  modelName = modelName || 'gemini-1.5-flash';
  
  try {
    const projectId = CONFIG.VERTEX_AI_PROJECT_ID();
    const location = CONFIG.VERTEX_AI_LOCATION();
    
    const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/${modelName}:generateContent`;
    
    const payload = {
      contents: [{
        role: 'user',
        parts: [{ text: 'Respond with only the word "OK" if you can read this.' }]
      }],
      generationConfig: {
        maxOutputTokens: 10,
        temperature: 0,
      }
    };
    
    const startTime = new Date().getTime();
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    
    const endTime = new Date().getTime();
    const latency = endTime - startTime;
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const result = JSON.parse(response.getContentText());
      const responseText = result.candidates?.[0]?.content?.parts?.[0]?.text || '';
      
      return {
        success: true,
        model: modelName,
        latency: latency,
        response: responseText.trim(),
        message: `✅ AVAILABLE (${latency}ms)`
      };
    } else {
      const errorText = response.getContentText();
      let errorMessage = 'Unknown error';
      
      try {
        const errorJson = JSON.parse(errorText);
        errorMessage = errorJson.error?.message || errorText.substring(0, 100);
      } catch (e) {
        errorMessage = errorText.substring(0, 100);
      }
      
      return {
        success: false,
        model: modelName,
        errorCode: responseCode,
        message: `❌ NOT AVAILABLE: ${errorMessage}`
      };
    }
    
  } catch (error) {
    return {
      success: false,
      model: modelName,
      message: `❌ ERROR: ${error.toString()}`
    };
  }
}

/**
 * Test all known Gemini models and report which ones are available
 * Run this to discover what models you can use
 */
function testAllGeminiModels() {
  // Comprehensive list of Gemini models to test (latest first)
  const modelsToTest = [
    // Latest 2.5 models
    'gemini-2.5-pro-preview-06-05',
    'gemini-2.5-flash-preview-05-20',
    'gemini-2.5-pro',
    'gemini-2.5-flash',
    
    // 2.0 models
    'gemini-2.0-flash',
    'gemini-2.0-flash-exp',
    'gemini-2.0-flash-lite',
    'gemini-2.0-pro-exp',
    
    // Experimental models
    'gemini-exp-1206',
    'gemini-exp-1121',
    
    // 1.5 stable models (currently configured)
    'gemini-1.5-flash',
    'gemini-1.5-flash-002',
    'gemini-1.5-flash-001',
    'gemini-1.5-pro',
    'gemini-1.5-pro-002',
    'gemini-1.5-pro-001',
    
    // Legacy models
    'gemini-1.0-pro',
    'gemini-pro',
  ];
  
  Logger.log('╔════════════════════════════════════════════════════════════╗');
  Logger.log('║           GEMINI MODEL AVAILABILITY TEST                   ║');
  Logger.log('╚════════════════════════════════════════════════════════════╝');
  Logger.log('');
  Logger.log(`Project: ${CONFIG.VERTEX_AI_PROJECT_ID()}`);
  Logger.log(`Location: ${CONFIG.VERTEX_AI_LOCATION()}`);
  Logger.log(`Current Flash Model: ${CONFIG.GEMINI_FLASH_MODEL}`);
  Logger.log(`Current Pro Model: ${CONFIG.GEMINI_PRO_MODEL}`);
  Logger.log('');
  Logger.log('Testing models...');
  Logger.log('─────────────────────────────────────────────────────────────');
  
  const results = {
    available: [],
    unavailable: []
  };
  
  modelsToTest.forEach(model => {
    const result = testGeminiModel(model);
    
    if (result.success) {
      results.available.push({
        model: model,
        latency: result.latency
      });
      Logger.log(`${result.message} - ${model}`);
    } else {
      results.unavailable.push(model);
      Logger.log(`${result.message} - ${model}`);
    }
  });
  
  Logger.log('');
  Logger.log('═════════════════════════════════════════════════════════════');
  Logger.log('SUMMARY');
  Logger.log('═════════════════════════════════════════════════════════════');
  Logger.log('');
  Logger.log(`✅ AVAILABLE MODELS (${results.available.length}):`);
  results.available.forEach(r => {
    Logger.log(`   • ${r.model} (${r.latency}ms latency)`);
  });
  
  Logger.log('');
  Logger.log(`❌ UNAVAILABLE MODELS (${results.unavailable.length}):`);
  results.unavailable.forEach(m => {
    Logger.log(`   • ${m}`);
  });
  
  // Recommend best models
  Logger.log('');
  Logger.log('═════════════════════════════════════════════════════════════');
  Logger.log('RECOMMENDATIONS');
  Logger.log('═════════════════════════════════════════════════════════════');
  
  const flashModels = results.available.filter(r => r.model.includes('flash'));
  const proModels = results.available.filter(r => r.model.includes('pro'));
  
  if (flashModels.length > 0) {
    // Sort by version (newest first based on name)
    const bestFlash = flashModels[0].model;
    Logger.log(`🚀 Recommended Flash Model: ${bestFlash}`);
  }
  
  if (proModels.length > 0) {
    const bestPro = proModels[0].model;
    Logger.log(`🚀 Recommended Pro Model: ${bestPro}`);
  }
  
  Logger.log('');
  Logger.log('To update your models, run: updateToLatestModels()');
  Logger.log('Or manually update CONFIG.GEMINI_FLASH_MODEL and CONFIG.GEMINI_PRO_MODEL in Config.gs');
  
  return results;
}

/**
 * Find and set the best available models automatically
 * This will test models and update the Config sheet with the best available ones
 */
function updateToLatestModels() {
  Logger.log('Finding best available Gemini models...');
  Logger.log('');
  
  // Test Flash models (fastest first for quick tasks)
  const flashCandidates = [
    'gemini-2.5-flash-preview-05-20',
    'gemini-2.5-flash',
    'gemini-2.0-flash',
    'gemini-2.0-flash-lite',
    'gemini-1.5-flash-002',
    'gemini-1.5-flash',
  ];
  
  // Test Pro models (most capable for complex reasoning)
  const proCandidates = [
    'gemini-2.5-pro-preview-06-05',
    'gemini-2.5-pro',
    'gemini-2.0-pro-exp',
    'gemini-1.5-pro-002',
    'gemini-1.5-pro',
  ];
  
  let bestFlash = null;
  let bestPro = null;
  
  // Find best Flash model
  Logger.log('Testing Flash models...');
  for (const model of flashCandidates) {
    const result = testGeminiModel(model);
    if (result.success) {
      bestFlash = model;
      Logger.log(`✅ Found best Flash: ${model} (${result.latency}ms)`);
      break;
    } else {
      Logger.log(`❌ ${model}: Not available`);
    }
  }
  
  Logger.log('');
  
  // Find best Pro model
  Logger.log('Testing Pro models...');
  for (const model of proCandidates) {
    const result = testGeminiModel(model);
    if (result.success) {
      bestPro = model;
      Logger.log(`✅ Found best Pro: ${model} (${result.latency}ms)`);
      break;
    } else {
      Logger.log(`❌ ${model}: Not available`);
    }
  }
  
  Logger.log('');
  Logger.log('═════════════════════════════════════════════════════════════');
  Logger.log('RESULTS');
  Logger.log('═════════════════════════════════════════════════════════════');
  
  if (bestFlash) {
    Logger.log(`Best Flash Model: ${bestFlash}`);
    Logger.log(`  Current setting: ${CONFIG.GEMINI_FLASH_MODEL}`);
    
    if (bestFlash !== CONFIG.GEMINI_FLASH_MODEL) {
      Logger.log(`  ⚠️  You could upgrade to: ${bestFlash}`);
      Logger.log(`     Update CONFIG.GEMINI_FLASH_MODEL in Config.gs to use it`);
    } else {
      Logger.log(`  ✅ Already using the best available`);
    }
  } else {
    Logger.log('❌ No Flash model available!');
  }
  
  Logger.log('');
  
  if (bestPro) {
    Logger.log(`Best Pro Model: ${bestPro}`);
    Logger.log(`  Current setting: ${CONFIG.GEMINI_PRO_MODEL}`);
    
    if (bestPro !== CONFIG.GEMINI_PRO_MODEL) {
      Logger.log(`  ⚠️  You could upgrade to: ${bestPro}`);
      Logger.log(`     Update CONFIG.GEMINI_PRO_MODEL in Config.gs to use it`);
    } else {
      Logger.log(`  ✅ Already using the best available`);
    }
  } else {
    Logger.log('❌ No Pro model available!');
  }
  
  Logger.log('');
  Logger.log('═════════════════════════════════════════════════════════════');
  Logger.log('HOW TO UPDATE');
  Logger.log('═════════════════════════════════════════════════════════════');
  Logger.log('');
  Logger.log('To update your models, edit Config.gs and change:');
  Logger.log('');
  if (bestFlash && bestFlash !== CONFIG.GEMINI_FLASH_MODEL) {
    Logger.log(`  GEMINI_FLASH_MODEL: '${CONFIG.GEMINI_FLASH_MODEL}' → '${bestFlash}'`);
  }
  if (bestPro && bestPro !== CONFIG.GEMINI_PRO_MODEL) {
    Logger.log(`  GEMINI_PRO_MODEL: '${CONFIG.GEMINI_PRO_MODEL}' → '${bestPro}'`);
  }
  
  return {
    bestFlash: bestFlash,
    bestPro: bestPro,
    currentFlash: CONFIG.GEMINI_FLASH_MODEL,
    currentPro: CONFIG.GEMINI_PRO_MODEL
  };
}

/**
 * Quick test of current configured models
 * Use this to verify your current configuration is working
 */
function testCurrentModels() {
  Logger.log('Testing currently configured models...');
  Logger.log('');
  
  const flashResult = testGeminiModel(CONFIG.GEMINI_FLASH_MODEL);
  const proResult = testGeminiModel(CONFIG.GEMINI_PRO_MODEL);
  
  Logger.log(`Flash (${CONFIG.GEMINI_FLASH_MODEL}): ${flashResult.message}`);
  Logger.log(`Pro (${CONFIG.GEMINI_PRO_MODEL}): ${proResult.message}`);
  
  if (flashResult.success && proResult.success) {
    Logger.log('');
    Logger.log('✅ All configured models are working!');
  } else {
    Logger.log('');
    Logger.log('⚠️  Some models are not working. Run testAllGeminiModels() to find alternatives.');
  }
  
  return {
    flash: flashResult,
    pro: proResult
  };
}

// ============================================================
// CORE API FUNCTIONS
// ============================================================

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
 * Falls back to GEMINI_PRO_FALLBACK_MODEL if primary model is not available
 */
function callGeminiPro(prompt, options = {}) {
  try {
    return callVertexAI(CONFIG.GEMINI_PRO_MODEL, prompt, {
      ...options,
      temperature: options.temperature || 0.7,
    });
  } catch (error) {
    // Check if it's a model not found error (404)
    if (error.toString().includes('404') || error.toString().includes('not found')) {
      Logger.log(`${CONFIG.GEMINI_PRO_MODEL} not available, falling back to ${CONFIG.GEMINI_PRO_FALLBACK_MODEL}...`);
      return callVertexAI(CONFIG.GEMINI_PRO_FALLBACK_MODEL, prompt, {
        ...options,
        temperature: options.temperature || 0.7,
      });
    }
    throw error;
  }
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
    
    // Speech-to-Text API endpoint (standard Cloud Speech-to-Text)
    const url = `https://speech.googleapis.com/v1/speech:recognize`;
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
            text: `Transcribe this audio file accurately. This is a TASK ASSIGNMENT voice note, so pay EXTRA attention to NAMES and DATES.

CRITICAL FOR NAMES (MOST IMPORTANT):
- Pay MAXIMUM attention to proper nouns (names of people)
- When you hear a person's name, transcribe it CLEARLY even if pronunciation is unclear
- Common Indian/South Asian names you might hear: Anaaya, Nikunj, Priya, Rahul, Vikram, Shreya, Aarav, Arjun, Amit, Ankit, Deepak, Gaurav, Harsh, Karan, Manish, Neha, Pooja, Rajesh, Sachin, Suresh, Vinay
- For names with double vowels (Anaaya, Priya, etc.), preserve them: "aa", "ee", "oo"
- If a name sounds like "anaya", write it as "Anaya" (could be Anaaya)
- NEVER skip or omit names - they determine who gets assigned the task
- If you hear "assign to [name]", "give to [name]", "[name] should do", "[name] ko do" - capture the name precisely

CRITICAL FOR DATES AND TIMES:
- Preserve dates in their spoken form: "tomorrow", "next Monday", "by Friday", "15th January", "end of week"
- Preserve times: "by 3 PM", "before noon", "in the morning"
- Hindi date words: "kal" (tomorrow), "parson" (day after tomorrow), "agle hafte" (next week)

GENERAL RULES:
- Preserve all spoken words exactly as spoken
- Include proper punctuation (periods, commas, question marks)
- Keep email addresses and technical terms exactly as spoken
- Do not add commentary, explanations, or corrections
- If audio quality is poor, transcribe what you can hear and indicate unclear parts with [unclear]

Return ONLY the transcribed text, nothing else.`
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
        temperature: 0.2, // Lower temperature for more accurate transcription
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
 * Get staff list for context in AI prompts (enhanced with phonetic hints)
 */
function getStaffListForContext() {
  try {
    const staff = getSheetData(SHEETS.STAFF_DB);
    if (!staff || staff.length === 0) return '';
    
    const staffList = staff.map(s => {
      const name = s.Name || 'Unknown';
      const email = s.Email || 'No email';
      // Generate phonetic hints for first name
      const firstName = name.split(' ')[0];
      const phoneticHint = generatePhoneticHints(firstName);
      return `- ${name} (${email})${phoneticHint ? ` [sounds like: ${phoneticHint}]` : ''}`;
    }).join('\n');
    
    return `\n\nAVAILABLE STAFF MEMBERS (use EXACT spelling from this list):
${staffList}

NOTE: Match transcribed names to this list even if spelling differs slightly. The name in the transcript may be a phonetic approximation.`;
  } catch (error) {
    Logger.log(`Error getting staff list: ${error.toString()}`);
    return '';
  }
}

/**
 * Get projects list for context in AI prompts
 */
function getProjectsListForContext() {
  try {
    const projects = getSheetData(SHEETS.PROJECTS_DB);
    if (!projects || projects.length === 0) return '';
    
    const projectsList = projects.map(p => `- ${p.Project_Name || 'Unknown'} (Tag: ${p.Project_Tag || 'N/A'})`).join('\n');
    return `\n\nAVAILABLE PROJECTS:\n${projectsList}`;
  } catch (error) {
    Logger.log(`Error getting projects list: ${error.toString()}`);
    return '';
  }
}

/**
 * Get current date context for AI prompts
 * Provides today's date and calculated relative dates in dd-MM-yyyy format
 */
function getDateContextForPrompt() {
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const dayOfWeek = Utilities.formatDate(today, tz, 'EEEE');
  const dateFormat = 'dd-MM-yyyy'; // Indian date format
  const todayStr = Utilities.formatDate(today, tz, dateFormat);
  
  // Calculate various relative dates
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  const dayAfterTomorrow = new Date(today);
  dayAfterTomorrow.setDate(dayAfterTomorrow.getDate() + 2);
  
  // Current week Monday and Sunday
  const currentDay = today.getDay();
  const monday = new Date(today);
  monday.setDate(today.getDate() - (currentDay === 0 ? 6 : currentDay - 1));
  
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  
  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);
  
  // Next week
  const nextMonday = new Date(monday);
  nextMonday.setDate(monday.getDate() + 7);
  
  const nextSunday = new Date(nextMonday);
  nextSunday.setDate(nextMonday.getDate() + 6);
  
  const nextFriday = new Date(nextMonday);
  nextFriday.setDate(nextMonday.getDate() + 4);
  
  // This Friday (current or next depending on today)
  const thisFriday = currentDay <= 5 ? friday : nextFriday;
  
  // One/two weeks from today
  const oneWeek = new Date(today);
  oneWeek.setDate(today.getDate() + 7);
  
  const twoWeeks = new Date(today);
  twoWeeks.setDate(today.getDate() + 14);
  
  return {
    TODAY_DATE: todayStr,
    DAY_OF_WEEK: dayOfWeek,
    TOMORROW_DATE: Utilities.formatDate(tomorrow, tz, dateFormat),
    DAY_AFTER_TOMORROW_DATE: Utilities.formatDate(dayAfterTomorrow, tz, dateFormat),
    MONDAY_DATE: Utilities.formatDate(monday, tz, dateFormat),
    SUNDAY_DATE: Utilities.formatDate(sunday, tz, dateFormat),
    FRIDAY_DATE: Utilities.formatDate(friday, tz, dateFormat),
    NEXT_MONDAY_DATE: Utilities.formatDate(nextMonday, tz, dateFormat),
    NEXT_SUNDAY_DATE: Utilities.formatDate(nextSunday, tz, dateFormat),
    NEXT_FRIDAY_DATE: Utilities.formatDate(nextFriday, tz, dateFormat),
    THIS_FRIDAY_DATE: Utilities.formatDate(thisFriday, tz, dateFormat),
    ONE_WEEK_FROM_TODAY: Utilities.formatDate(oneWeek, tz, dateFormat),
    TWO_WEEKS_FROM_TODAY: Utilities.formatDate(twoWeeks, tz, dateFormat),
  };
}

/**
 * Generate phonetic hints for common name variations
 */
function generatePhoneticHints(name) {
  if (!name) return '';
  const lower = name.toLowerCase();
  
  const hints = [];
  
  // Common variations for Indian names
  if (lower.includes('aa')) hints.push(lower.replace(/aa/g, 'a'));
  if (lower.includes('ee')) hints.push(lower.replace(/ee/g, 'i'));
  if (lower.includes('oo')) hints.push(lower.replace(/oo/g, 'u'));
  
  // Specific name variations
  const nameVariations = {
    'anaaya': ['anaya', 'anaia', 'anayya', 'anaiah'],
    'nikunj': ['nikanj', 'neekunj', 'nikunge', 'nikunje'],
    'priya': ['pria', 'priyah', 'preya', 'priyaa'],
    'shreya': ['shrey', 'sreya', 'shriya', 'shrea'],
    'rahul': ['rahool', 'raul', 'rahull'],
    'vikram': ['vikrum', 'vickram', 'vikrm'],
    'aarav': ['arav', 'aaruv', 'aruv'],
    'arjun': ['arjoon', 'arjn', 'arjuna'],
    'amit': ['amith', 'ameet', 'ammit'],
    'ankit': ['ankith', 'ankeet', 'ankit'],
    'deepak': ['deepk', 'dipak', 'deepuk'],
    'gaurav': ['gorav', 'gauruv', 'gaorav'],
    'harsh': ['harash', 'harsch', 'harish'],
    'karan': ['karn', 'karun', 'krn'],
    'manish': ['maneesh', 'manesh', 'mansh'],
    'neha': ['neeha', 'neha', 'nehaa'],
    'pooja': ['puja', 'poojah', 'pujaa'],
    'rajesh': ['rajsh', 'rajeesh', 'rajeshh'],
    'sachin': ['sachinn', 'sachn', 'sacheen'],
    'suresh': ['suresh', 'sureesh', 'sureshh'],
    'vinay': ['vinaay', 'vinai', 'vinayy'],
    'michael': ['mike', 'mikel', 'michel'],
    'robert': ['bob', 'rob', 'bobby'],
    'james': ['jim', 'jimmy', 'jamie'],
    'william': ['will', 'bill', 'billy'],
    'elizabeth': ['liz', 'beth', 'lizzy'],
    'jennifer': ['jen', 'jenny', 'jenn'],
    'sarah': ['sara', 'sahra', 'sarra'],
    'jonathan': ['jon', 'john', 'jonny'],
  };
  
  if (nameVariations[lower]) {
    hints.push(...nameVariations[lower]);
  }
  
  // Check if any variation matches
  for (const [canonical, variations] of Object.entries(nameVariations)) {
    if (variations.includes(lower)) {
      hints.push(canonical);
    }
  }
  
  // Remove duplicates and the original name
  const uniqueHints = [...new Set(hints)].filter(h => h !== lower);
  
  return uniqueHints.length > 0 ? uniqueHints.slice(0, 4).join(', ') : '';
}

/**
 * Parse voice command and extract structured data
 */
function parseVoiceCommand(transcript) {
  // Get context about available staff, projects, and dates
  const staffContext = getStaffListForContext();
  const projectsContext = getProjectsListForContext();
  const dateContext = getDateContextForPrompt();
  
  // Try to load prompt from sheet, fallback to hardcoded
  let promptTemplate = null;
  try {
    const promptData = getPrompt('parseVoiceCommand', 'voice');
    if (promptData && promptData.Content) {
      promptTemplate = promptData.Content;
    }
  } catch (e) {
    Logger.log('Could not load prompt from sheet, using fallback: ' + e.toString());
  }
  
  // Fallback to hardcoded prompt if not found in sheet
  if (!promptTemplate) {
    promptTemplate = `You are an expert task management assistant with advanced name recognition and date parsing capabilities. Parse the following voice command transcript and extract all relevant information with high accuracy.

Voice command transcript: "{{TRANSCRIPT}}"

===========================================
CURRENT DATE AND TIME CONTEXT
===========================================
- Today's date: {{TODAY_DATE}}
- Today is: {{DAY_OF_WEEK}}
- Current week: Monday={{MONDAY_DATE}} through Sunday={{SUNDAY_DATE}}
- This Friday: {{THIS_FRIDAY_DATE}}
- Next week: {{NEXT_MONDAY_DATE}} through {{NEXT_SUNDAY_DATE}}
- Next Friday: {{NEXT_FRIDAY_DATE}}
- Timezone: Asia/Kolkata (IST)

{{STAFF_CONTEXT}}{{PROJECTS_CONTEXT}}

===========================================
CRITICAL INSTRUCTIONS FOR NAME/ASSIGNEE MATCHING
===========================================

1. ALWAYS scan the AVAILABLE STAFF MEMBERS list above FIRST before processing
2. When you hear ANY name in the transcript, immediately check for matches in the staff list
3. Use PHONETIC MATCHING - names that SOUND similar should match:
   - "anaya" heard → match to "Anaaya" in staff list
   - "neekunj" or "nikanj" heard → match to "Nikunj"
   - "shray" or "shreya" heard → match to "Shreya"
   - "priyah" or "pria" heard → match to "Priya"
   - "john" heard → match to "Jon" or "Jonathan"
   - "mike" heard → match to "Michael"

4. Common voice transcription errors to watch for:
   - Double vowels getting reduced: "Anaaya" transcribed as "Anaya" or "Anaia"
   - 'j' and 'g' sounds confused: "Nikunj" → "Nikung"
   - 'sh' and 's' confused: "Shreya" → "Sreya"
   - 'th' sounds: "Prathik" → "Pratik"
   - Silent letters dropped or added

5. PRIORITY: If you identify a name in the transcript, you MUST:
   a. Find the closest match in AVAILABLE STAFF MEMBERS
   b. Use the EXACT spelling from the staff list
   c. Only return null for assignee_name if NO name is mentioned at all

6. Look for these assignee indicators:
   - "assign to [name]", "give to [name]", "[name] should", "[name] will"
   - "send to [name]", "[name] can", "[name] needs to", "for [name]"
   - "[name] ko do" (Hindi), "[name] se karwa do", "[name] handle karega"

===========================================
CRITICAL INSTRUCTIONS FOR DUE DATE EXTRACTION
===========================================

TODAY IS: {{TODAY_DATE}} ({{DAY_OF_WEEK}})
DATE FORMAT: Always output dates as DD-MM-YYYY (e.g., 27-01-2025)

1. RELATIVE DATE CONVERSIONS (calculate from today's date above):
   - "today" = {{TODAY_DATE}}
   - "tomorrow" = {{TOMORROW_DATE}}
   - "day after tomorrow" = {{DAY_AFTER_TOMORROW_DATE}}
   - "this week" = by {{SUNDAY_DATE}}
   - "end of week" = {{FRIDAY_DATE}} or {{SUNDAY_DATE}}
   - "next week" = {{NEXT_MONDAY_DATE}} through {{NEXT_SUNDAY_DATE}}
   - "next Monday" = {{NEXT_MONDAY_DATE}}
   - "next Friday" = {{NEXT_FRIDAY_DATE}}
   
2. DAY NAME CONVERSIONS (relative to today being {{DAY_OF_WEEK}}):
   - If today is Monday-Thursday and someone says "Friday" = this Friday ({{THIS_FRIDAY_DATE}})
   - If today is Friday-Sunday and someone says "Friday" = next Friday ({{NEXT_FRIDAY_DATE}})
   - "this [day]" = upcoming occurrence in current week
   - "next [day]" = occurrence in the following week

3. TIME EXPRESSIONS:
   - "in X days" = add X days to {{TODAY_DATE}}
   - "in a week" = {{ONE_WEEK_FROM_TODAY}}
   - "in two weeks" = {{TWO_WEEKS_FROM_TODAY}}
   - "by month end" = last day of current month
   - "end of month" = last day of current month

4. ABSOLUTE DATES - Convert to DD-MM-YYYY:
   - "15th" or "the 15th" = 15-01-2025 (15th of current month, or next if passed)
   - "January 15" or "15th January" or "15 Jan" = 15-01-2025
   - "15/01" or "15-01" = assume DD/MM format = 15-01-2025
   - "2025-01-15" = convert to 15-01-2025

5. HINDI/HINGLISH DATE EXPRESSIONS:
   - "kal" = tomorrow = {{TOMORROW_DATE}}
   - "parson" = day after tomorrow = {{DAY_AFTER_TOMORROW_DATE}}
   - "agle hafte" = next week = {{NEXT_MONDAY_DATE}}
   - "is hafte" = this week = by {{SUNDAY_DATE}}
   - "mahine ke end tak" = end of month

6. PRESERVE ORIGINAL TEXT:
   - Always capture the exact spoken words in due_date_text field
   - Convert to DD-MM-YYYY format in due_date field

7. AMBIGUOUS DATES:
   - If unclear, prefer the NEARER date
   - Add ambiguity note if date interpretation is uncertain

===========================================
CRITICAL INSTRUCTIONS FOR PROJECT RECOGNITION
===========================================

1. Look for project mentions in the task name, context, or assignee context
2. Use the AVAILABLE PROJECTS list above to match project names
3. Consider partial matches and variations (e.g., "Q4 Report" → "Q4 Financial Report")
4. Project tags can be abbreviations or codes - match both name and tag
5. If project is mentioned indirectly (e.g., "for the website project"), extract it

===========================================
OUTPUT FORMAT
===========================================

Extract the following information and return ONLY a valid JSON object (no markdown, no code blocks, no explanations):

{
  "task_name": "concise, clear task title (3-8 words, action-oriented)",
  "assignee_email": "full email address if mentioned explicitly, or null",
  "assignee_name": "EXACT name from AVAILABLE STAFF MEMBERS list if matched, or the heard name if no match, or null",
  "due_date": "date in DD-MM-YYYY format (MUST be calculated from today={{TODAY_DATE}}), or null",
  "due_date_text": "exact original text about due date/timing as spoken, or null",
  "due_time": "time if mentioned (HH:MM format in 24-hour), or null",
  "project_tag": "project name or tag if mentioned, or null",
  "context": "all additional details, background, requirements, constraints mentioned",
  "tone": "urgent, normal, or calm",
  "confidence": 0.0-1.0 confidence score,
  "ambiguities": ["list of unclear elements"],
  "name_heard_as": "the raw name as it was transcribed before matching (for debugging)"
}

===========================================
EXTRACTION GUIDELINES
===========================================

1. Task Name:
   - Extract the core action/objective (e.g., "Review Q4 report" not "I need to review the Q4 report")
   - Remove filler words like "please", "I want", "can you", "I need", "we should"
   - Keep it concise but descriptive (3-8 words ideal)
   - Use action verbs: "Review", "Update", "Create", "Send", "Schedule", etc.

2. Tone (indicates urgency level):
   - urgent: ASAP, immediately, critical, important, rush, jaldi
   - normal: standard, regular (default if not mentioned)
   - calm: whenever, no rush, eventually

3. Context:
   - Include: why the task exists, background info, specific requirements, constraints, dependencies
   - Be comprehensive but concise

4. Confidence:
   - 0.9-1.0: All key info clear (task, assignee, date all clear)
   - 0.7-0.9: Most info clear, minor ambiguities
   - 0.5-0.7: Some key info missing or unclear
   - <0.5: Major ambiguities or unclear command

5. Ambiguities:
   - List specific unclear elements: "assignee name unclear", "due date ambiguous", "task scope vague"

===========================================
EXAMPLES WITH NAME AND DATE MATCHING
===========================================

EXAMPLE 1 - Name with phonetic variation:
Transcript: "Assign anaya to review the Q4 financial report by next Friday, it's urgent"
Today: 27-01-2025 (Monday)
Staff List: Anaaya Udhas (anaayaudhas92@gmail.com)

Output: {
  "task_name": "Review Q4 financial report",
  "assignee_email": null,
  "assignee_name": "Anaaya Udhas",
  "due_date": "31-01-2025",
  "due_date_text": "next Friday",
  "due_time": null,
  "project_tag": null,
  "context": "Q4 financial report needs review, marked urgent",
  "tone": "urgent",
  "confidence": 0.95,
  "ambiguities": [],
  "name_heard_as": "anaya"
}

EXAMPLE 2 - Relative date with Hindi:
Transcript: "Nikunj ko kal tak website update karna hai"
Today: 27-01-2025 (Monday)
Staff List: Nikunj Jajodia (nikunj@example.com)

Output: {
  "task_name": "Update website",
  "assignee_email": null,
  "assignee_name": "Nikunj Jajodia",
  "due_date": "28-01-2025",
  "due_date_text": "kal tak",
  "due_time": null,
  "project_tag": null,
  "context": "Website update needed",
  "tone": "normal",
  "confidence": 0.9,
  "ambiguities": [],
  "name_heard_as": "Nikunj"
}

EXAMPLE 3 - End of week:
Transcript: "I need someone to finish the presentation by end of week"
Today: 27-01-2025 (Monday)

Output: {
  "task_name": "Finish presentation",
  "assignee_email": null,
  "assignee_name": null,
  "due_date": "31-01-2025",
  "due_date_text": "end of week",
  "due_time": null,
  "project_tag": null,
  "context": "Presentation needs to be completed",
  "tone": "normal",
  "confidence": 0.7,
  "ambiguities": ["assignee not specified"],
  "name_heard_as": null
}

EXAMPLE 4 - Day after tomorrow in Hindi:
Transcript: "Priya ko parson tak report submit karni hai"
Today: 27-01-2025 (Monday)
Staff List: Priya Sharma (priya@example.com)

Output: {
  "task_name": "Submit report",
  "assignee_email": null,
  "assignee_name": "Priya Sharma",
  "due_date": "29-01-2025",
  "due_date_text": "parson tak",
  "due_time": null,
  "project_tag": null,
  "context": "Report needs to be submitted",
  "tone": "normal",
  "confidence": 0.9,
  "ambiguities": [],
  "name_heard_as": "Priya"
}

Now parse this voice command: "{{TRANSCRIPT}}"`;
  }
  
  // Replace template variables
  let prompt = promptTemplate
    .replace(/\{\{TRANSCRIPT\}\}/g, transcript)
    .replace(/\{\{STAFF_CONTEXT\}\}/g, staffContext)
    .replace(/\{\{PROJECTS_CONTEXT\}\}/g, projectsContext);
  
  // Replace all date context placeholders
  for (const [key, value] of Object.entries(dateContext)) {
    prompt = prompt.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'g'), value);
  }

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
    
    // Process due date if it's text but due_date is missing
    if (parsed.due_date_text && !parsed.due_date) {
      parsed.due_date = parseRelativeDate(parsed.due_date_text);
    }
    
    // Log the name matching for debugging
    if (parsed.name_heard_as && parsed.assignee_name) {
      Logger.log(`Voice name matching: heard "${parsed.name_heard_as}" → matched to "${parsed.assignee_name}"`);
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
 * Parse relative date text to actual date in dd-MM-yyyy format
 */
function parseRelativeDate(dateText) {
  const today = new Date();
  const tz = Session.getScriptTimeZone();
  const dateFormat = 'dd-MM-yyyy';
  const lowerText = dateText.toLowerCase();
  
  // Today
  if (lowerText.includes('today') || lowerText.includes('aaj')) {
    return Utilities.formatDate(today, tz, dateFormat);
  }
  
  // Tomorrow (English and Hindi)
  if (lowerText.includes('tomorrow') || lowerText.includes('kal')) {
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    return Utilities.formatDate(tomorrow, tz, dateFormat);
  }
  
  // Day after tomorrow (English and Hindi)
  if (lowerText.includes('day after tomorrow') || lowerText.includes('parson') || lowerText.includes('parso')) {
    const dayAfter = new Date(today);
    dayAfter.setDate(dayAfter.getDate() + 2);
    return Utilities.formatDate(dayAfter, tz, dateFormat);
  }
  
  // Next week / Next Monday (English and Hindi)
  if (lowerText.includes('next monday') || lowerText.includes('next week') || lowerText.includes('agle hafte') || lowerText.includes('agle week')) {
    const daysUntilMonday = (8 - today.getDay()) % 7 || 7;
    const nextMonday = new Date(today);
    nextMonday.setDate(today.getDate() + daysUntilMonday);
    return Utilities.formatDate(nextMonday, tz, dateFormat);
  }
  
  // End of week / This week (English and Hindi)
  if (lowerText.includes('end of week') || lowerText.includes('is hafte') || lowerText.includes('this week') || lowerText.includes('week end')) {
    const currentDay = today.getDay();
    const daysUntilFriday = (5 - currentDay + 7) % 7;
    const friday = new Date(today);
    friday.setDate(today.getDate() + (daysUntilFriday === 0 ? 0 : daysUntilFriday));
    return Utilities.formatDate(friday, tz, dateFormat);
  }
  
  // Next Friday
  if (lowerText.includes('next friday')) {
    const currentDay = today.getDay();
    // Calculate days until next Friday (at least 7 days from now)
    const daysUntilNextFriday = ((5 - currentDay + 7) % 7) + 7;
    const nextFriday = new Date(today);
    nextFriday.setDate(today.getDate() + daysUntilNextFriday);
    return Utilities.formatDate(nextFriday, tz, dateFormat);
  }
  
  // This Friday / Friday
  if (lowerText.includes('friday') && !lowerText.includes('next')) {
    const currentDay = today.getDay();
    let daysUntilFriday = (5 - currentDay + 7) % 7;
    // If today is Friday or later, go to next week's Friday
    if (daysUntilFriday === 0 && currentDay >= 5) {
      daysUntilFriday = 7;
    }
    const friday = new Date(today);
    friday.setDate(today.getDate() + daysUntilFriday);
    return Utilities.formatDate(friday, tz, dateFormat);
  }
  
  // In X days
  const inDaysMatch = lowerText.match(/in\s+(\d+)\s+days?/);
  if (inDaysMatch) {
    const days = parseInt(inDaysMatch[1], 10);
    const futureDate = new Date(today);
    futureDate.setDate(today.getDate() + days);
    return Utilities.formatDate(futureDate, tz, dateFormat);
  }
  
  // In a week / one week
  if (lowerText.includes('in a week') || lowerText.includes('in one week') || lowerText.includes('ek hafte mein')) {
    const oneWeek = new Date(today);
    oneWeek.setDate(today.getDate() + 7);
    return Utilities.formatDate(oneWeek, tz, dateFormat);
  }
  
  // In two weeks
  if (lowerText.includes('in two weeks') || lowerText.includes('in 2 weeks') || lowerText.includes('do hafte mein')) {
    const twoWeeks = new Date(today);
    twoWeeks.setDate(today.getDate() + 14);
    return Utilities.formatDate(twoWeeks, tz, dateFormat);
  }
  
  // End of month / Month end
  if (lowerText.includes('end of month') || lowerText.includes('month end') || lowerText.includes('mahine ke end')) {
    const endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
    return Utilities.formatDate(endOfMonth, tz, dateFormat);
  }
  
  // Try to parse DD-MM-YYYY format first
  const ddmmyyyyMatch = dateText.match(/(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (ddmmyyyyMatch) {
    const [_, day, month, year] = ddmmyyyyMatch;
    const parsed = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, tz, dateFormat);
    }
  }
  
  // Try to parse DD/MM/YYYY format
  const ddmmyyyySlashMatch = dateText.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (ddmmyyyySlashMatch) {
    const [_, day, month, year] = ddmmyyyySlashMatch;
    const parsed = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, tz, dateFormat);
    }
  }
  
  // Try to parse "15th January", "January 15", etc.
  const monthNames = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
  const monthAbbr = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  
  for (let i = 0; i < monthNames.length; i++) {
    const monthPattern = new RegExp(`(\\d{1,2})(?:st|nd|rd|th)?\\s*(?:of\\s*)?${monthNames[i]}|${monthNames[i]}\\s*(\\d{1,2})(?:st|nd|rd|th)?|` +
                                    `(\\d{1,2})(?:st|nd|rd|th)?\\s*${monthAbbr[i]}|${monthAbbr[i]}\\s*(\\d{1,2})(?:st|nd|rd|th)?`, 'i');
    const match = lowerText.match(monthPattern);
    if (match) {
      const day = parseInt(match[1] || match[2] || match[3] || match[4], 10);
      const year = today.getFullYear();
      const parsed = new Date(year, i, day);
      // If the date has passed, assume next year
      if (parsed < today) {
        parsed.setFullYear(year + 1);
      }
      return Utilities.formatDate(parsed, tz, dateFormat);
    }
  }
  
  // Try to parse as generic date
  try {
    const parsed = new Date(dateText);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, tz, dateFormat);
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
function classifyReplyType(replyContent, originalDueDate) {
  // Build context about the original deadline if available
  const dateContext = originalDueDate ? `\n\nOriginal task due date: ${originalDueDate}` : '';
  
  // Try to load prompt from sheet, fallback to hardcoded
  let promptTemplate = null;
  try {
    const promptData = getPrompt('classifyReplyType', 'email');
    if (promptData && promptData.Content) {
      promptTemplate = promptData.Content;
    }
  } catch (e) {
    Logger.log('Could not load prompt from sheet, using fallback: ' + e.toString());
  }
  
  // Fallback to hardcoded prompt if not found in sheet
  if (!promptTemplate) {
    promptTemplate = `You are an expert at analyzing email replies about task assignments. Classify this email reply into one of these categories:

1. ACCEPTANCE - The assignee accepts the task and agrees to the deadline. Examples: "I accept", "I'll do it", "Got it", "Will complete by deadline", "Acknowledged" (if accepting)
2. DATE_CHANGE - The assignee requests a different due date or mentions a date that's not feasible. Look for: "not feasible", "can't make that date", "need more time", "deadline is", "by [date]", "feasible as deadline", "propose [date]", "suggest [date]", "prefer [date]", "10th Jan", "January 10th", "10/01/2025", etc.
3. SCOPE_QUESTION - The assignee has questions about what needs to be done, needs clarification, or asks "what does this mean", "can you clarify", "I need more details"
4. ROLE_REJECTION - The assignee says this is not their job, responsibility, or role. Examples: "not my responsibility", "wrong person", "not my role", "should be assigned to"
5. OTHER - Something else that doesn't fit the above categories

IMPORTANT DATE EXTRACTION RULES:
- IGNORE dates in quoted email sections (text after "On [date] wrote:" or lines starting with ">")
- IGNORE dates in email signatures
- Extract ONLY the PROPOSED/NEW date from the actual reply text (the part the employee wrote)
- If the email mentions ANY date in the actual reply (not quoted text), and that date is different from the original deadline, classify as DATE_CHANGE
- Look for phrases like: "7th of Jan 2026", "10th Jan is feasible", "deadline is 10th Jan", "by 10th Jan", "10th Jan works", "can do it by 10th Jan", "I can do it by the 7th of Jan 2026"
- Convert ALL date formats to YYYY-MM-DD:
  * "7th of Jan 2026" or "7th of January 2026" → 2026-01-07
  * "10th Jan" or "10 Jan" → Assume current year if year not mentioned, or next year if date has passed
  * "Jan 10th" or "January 10" → Same as above
  * "10/01/2025" → 2025-01-10 (assume DD/MM/YYYY format)
  * "2025-01-10" → Keep as is
  * Relative dates like "next Monday", "in 2 weeks" → Calculate actual date
- If multiple dates are mentioned, extract the one that appears to be the NEW/PROPOSED deadline (usually the one mentioned with "by", "deadline", "feasible", "can do")
- If the email says "not feasible" or "can't make that date" and mentions another date, that other date is the proposed date
- PRIORITIZE dates that appear in sentences with words like "by", "deadline", "feasible", "can do", "I can"

Email reply (already cleaned - quoted text removed): "{{REPLY_CONTENT}}"{{DATE_CONTEXT}}

Return ONLY a valid JSON object (no markdown, no code blocks):
{
  "type": "ACCEPTANCE|DATE_CHANGE|SCOPE_QUESTION|ROLE_REJECTION|OTHER",
  "confidence": 0.0-1.0,
  "extracted_date": "proposed date in YYYY-MM-DD format if DATE_CHANGE (MUST be the NEW date being requested, not the original. Convert all formats: '10th Jan' → '2025-01-10', 'Jan 10' → '2025-01-10', '10/01/2025' → '2025-01-10', 'next Monday' → calculate actual date). If no clear date, return null",
  "reasoning": "brief explanation of why this classification was chosen and what date was extracted (if any)"
}`;
  }
  
  // Replace template variables
  const prompt = promptTemplate
    .replace(/\{\{REPLY_CONTENT\}\}/g, replyContent)
    .replace(/\{\{DATE_CONTEXT\}\}/g, dateContext);

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
 * Analyze task context and suggest appropriate actions
 * This replaces the simple "approve" button with context-aware actions
 */
function analyzeTaskContextForActions(taskId) {
  try {
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`Task ${taskId} not found`);
      return null;
    }
    
    // Get conversation history
    let conversationHistory = [];
    
    // Try to parse Conversation_History JSON field
    if (task.Conversation_History) {
      try {
        conversationHistory = JSON.parse(task.Conversation_History);
      } catch (e) {
        Logger.log('Could not parse Conversation_History, trying Interaction_Log');
      }
    }
    
    // Fallback: Extract from Interaction_Log if Conversation_History not available
    if (conversationHistory.length === 0 && task.Interaction_Log) {
      conversationHistory = extractMessagesFromInteractionLog(task.Interaction_Log);
    }
    
    // Also get latest employee reply if available
    if (task.Employee_Reply && !conversationHistory.some(m => m.content === task.Employee_Reply)) {
      conversationHistory.push({
        timestamp: task.Last_Updated || new Date().toISOString(),
        senderEmail: task.Assignee_Email,
        senderName: task.Assignee_Name || task.Assignee_Email,
        type: 'email_reply',
        content: task.Employee_Reply
      });
    }
    
    if (conversationHistory.length === 0) {
      Logger.log(`No conversation history found for task ${taskId}`);
      return getDefaultActionsForStatus(task.Status);
    }
    
    // Build conversation context
    const conversationText = conversationHistory
      .map((msg, idx) => {
        const sender = msg.senderName || msg.senderEmail || (msg.type === 'boss_message' ? 'Boss' : 'Employee');
        const timestamp = msg.timestamp ? new Date(msg.timestamp).toLocaleString() : `Message ${idx + 1}`;
        return `[${timestamp}] ${sender}: ${msg.content}`;
      })
      .join('\n\n');
    
    // Get current state
    const currentDueDate = task.Due_Date ? 
      Utilities.formatDate(new Date(task.Due_Date), Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy') : 
      'Not set';
    const proposedDate = task.Proposed_Date ? 
      Utilities.formatDate(new Date(task.Proposed_Date), Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy') : 
      'None';
    
    const prompt = `You are analyzing a task that requires action from the boss. Based on the conversation history and current state, determine what actions should be available.

Task Details:
- Task Name: "${task.Task_Name}"
- Current Status: "${task.Status}"
- Current Due Date: ${currentDueDate}
- Proposed Date: ${proposedDate}
- Employee: ${task.Assignee_Email}
- Approval State: ${task.Approval_State || 'none'}

Conversation History:
${conversationText || 'No conversation history yet'}

Latest Employee Message:
${task.Employee_Reply || 'None'}

Based on this context, analyze:
1. What is the employee asking for? (date change, clarification, etc.)
2. What actions should the boss be able to take?
3. Should the boss approve, reject, propose alternative, or message the employee?
4. What is the current approval state? (awaiting_boss, boss_approved, negotiating, etc.)

IMPORTANT: Look at the ENTIRE conversation, not just the latest message. Sometimes approval needs emerge across multiple messages.

Return ONLY a valid JSON object:
{
  "primaryAction": {
    "type": "APPROVE|REJECT|MESSAGE|PROPOSE_ALTERNATIVE|NONE",
    "label": "Approve Date Change",
    "description": "Approve the employee's requested date",
    "icon": "check",
    "confidence": 0.9,
    "actionId": "approve_date_change"
  },
  "secondaryActions": [
    {
      "type": "MESSAGE",
      "label": "Message Employee",
      "description": "Send a message explaining why the date cannot be changed",
      "icon": "message-square",
      "actionId": "message_employee_reject"
    },
    {
      "type": "PROPOSE_ALTERNATIVE",
      "label": "Propose Different Date",
      "description": "Suggest a compromise date",
      "icon": "calendar",
      "actionId": "propose_date"
    }
  ],
  "approvalState": "awaiting_boss|boss_approved|boss_rejected|boss_proposed|negotiating",
  "reasoning": "Employee requested date change to Jan 15. Boss should approve if feasible, or message to explain why not.",
  "requiresEmployeeConfirmation": true/false,
  "suggestedMessage": "Optional: Pre-filled message text if boss chooses to message"
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
    
    Logger.log(`Context analysis for ${taskId}:`);
    Logger.log(`  Primary Action: ${analysis.primaryAction.type} - ${analysis.primaryAction.label}`);
    Logger.log(`  Approval State: ${analysis.approvalState}`);
    Logger.log(`  Requires Employee Confirmation: ${analysis.requiresEmployeeConfirmation}`);
    
    return analysis;
    
  } catch (error) {
    Logger.log(`Error analyzing context: ${error.toString()}`);
    // Fallback to default actions based on status
    return getDefaultActionsForStatus(task.Status);
  }
}

/**
 * Analyze conversation history and update task state
 * This is the main function that derives state from conversation
 * Called after every email/message exchange
 * 
 * @param {string} taskId - The task ID
 * @param {Object} newMessage - Optional new message to add before analysis
 * @returns {Object} Analysis result with conversationState, pendingChanges, summary
 */
function analyzeConversationAndUpdateState(taskId, newMessage = null) {
  try {
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`Task ${taskId} not found`);
      return { success: false, error: 'Task not found' };
    }
    
    // Parse initial parameters
    let initialParams = {};
    if (task.Initial_Parameters) {
      try {
        initialParams = JSON.parse(task.Initial_Parameters);
      } catch (e) {
        Logger.log('Could not parse Initial_Parameters');
        initialParams = {
          dueDate: task.Due_Date,
          assignee: task.Assignee_Email,
          taskName: task.Task_Name,
          scope: task.Context_Hidden
        };
      }
    }
    
    // Parse conversation history
    let conversationHistory = [];
    if (task.Conversation_History) {
      try {
        conversationHistory = JSON.parse(task.Conversation_History);
      } catch (e) {
        Logger.log('Could not parse Conversation_History');
        conversationHistory = [];
      }
    }
    
    // Add new message if provided
    if (newMessage) {
      conversationHistory.push({
        id: newMessage.id || newMessage.messageId || `local_${Date.now()}`,
        messageId: newMessage.messageId || newMessage.id || `local_${Date.now()}`,
        timestamp: newMessage.timestamp || new Date().toISOString(),
        senderEmail: newMessage.senderEmail,
        senderName: newMessage.senderName || newMessage.senderEmail,
        type: newMessage.type || 'message',
        content: newMessage.content,
        metadata: newMessage.metadata || {}
      });
    }
    
    // Also include Employee_Reply if not already in history
    if (task.Employee_Reply && !conversationHistory.some(m => m.content === task.Employee_Reply)) {
      conversationHistory.push({
        timestamp: task.Last_Updated || new Date().toISOString(),
        senderEmail: task.Assignee_Email,
        senderName: task.Assignee_Name || task.Assignee_Email,
        type: 'employee_reply',
        content: task.Employee_Reply
      });
    }
    
    // If no conversation, default to active state
    if (conversationHistory.length === 0) {
      const result = {
        conversationState: CONVERSATION_STATE.ACTIVE,
        pendingChanges: [],
        summary: 'No conversation history yet',
        requiresAction: false
      };
      
      // Update task with state
      updateTask(taskId, {
        Conversation_State: result.conversationState,
        Pending_Changes: JSON.stringify(result.pendingChanges),
        AI_Summary: result.summary,
        Conversation_History: JSON.stringify(conversationHistory),
        Derived_Task_Name: task.Task_Name || '',
        Derived_Due_Date_Effective: task.Due_Date || '',
        Derived_Due_Date_Proposed: task.Proposed_Date || '',
        Derived_Scope_Summary: task.Context_Hidden || '',
        Derived_Field_Provenance: JSON.stringify({}),
        Derived_Last_Analyzed_At: new Date().toISOString()
      });
      
      return { success: true, data: result };
    }
    
    // Build conversation text for AI analysis
    const bossEmail = CONFIG.BOSS_EMAIL();
    const conversationText = conversationHistory
      .map((msg, idx) => {
        const isBoss = msg.senderEmail && msg.senderEmail.toLowerCase() === bossEmail.toLowerCase();
        const sender = isBoss ? 'Boss' : (msg.senderName || 'Employee');
        const timestamp = msg.timestamp ? new Date(msg.timestamp).toLocaleString() : `Message ${idx + 1}`;
        const msgId = msg.messageId || msg.id || '';
        return `[${timestamp}]${msgId ? ` [id:${msgId}]` : ''} ${sender}: ${msg.content}`;
      })
      .join('\n\n');
    
    // Determine last sender
    const lastMessage = conversationHistory[conversationHistory.length - 1];
    const lastSenderIsBoss = lastMessage && lastMessage.senderEmail && 
      lastMessage.senderEmail.toLowerCase() === bossEmail.toLowerCase();
    
    // Format current task state
    const tz = Session.getScriptTimeZone();
    const currentDueDate = task.Due_Date ? 
      Utilities.formatDate(new Date(task.Due_Date), tz, 'yyyy-MM-dd') : 'Not set';
    const initialDueDate = initialParams.dueDate ? 
      Utilities.formatDate(new Date(initialParams.dueDate), tz, 'yyyy-MM-dd') : 'Not set';
    
    // Relationship context (minimal): helps AI label requestedBy/awaitingFrom deterministically
    const staff = task.Assignee_Email ? getStaff(task.Assignee_Email) : null;
    const project = task.Project_Tag ? getProject(task.Project_Tag) : null;
    const relationshipContext = `RELATIONSHIP CONTEXT (use this to set requestedBy/awaitingFrom):
- Boss email: ${bossEmail || 'unknown'}
- Assignee email: ${task.Assignee_Email || 'unknown'}
- Assignee manager email (from Staff_DB): ${(staff && staff.Manager_Email) || 'unknown'}
- Project tag: ${task.Project_Tag || 'none'}
- Project team lead email (from Projects_DB): ${(project && project.Team_Lead_Email) || 'unknown'}`;

    const prompt = `You are analyzing a conversation between a boss and employee about a task. Your job is to produce an AI-derived "truth snapshot" of the task plus a typed list of pending items.

TASK INFORMATION:
- Task Name: "${task.Task_Name}"
- Assignee: ${task.Assignee_Name || task.Assignee_Email}
- Current Due Date: ${currentDueDate}
- Initial Due Date: ${initialDueDate}
- Current Status: ${task.Status}
- Scope/Description: ${task.Context_Hidden || 'Not provided'}

${relationshipContext}

CONVERSATION HISTORY (chronological order):
${conversationText}

LAST MESSAGE WAS FROM: ${lastSenderIsBoss ? 'Boss' : 'Employee'}

ANALYZE THE CONVERSATION AND DETERMINE:

1. CONVERSATION STATE - What is the current state of this conversation?
   - "active" - Normal operation, no pending items
   - "update_received" - Employee sent progress update, FYI only (no action needed)
   - "change_requested" - Employee is requesting a change (date, scope, role) that needs boss approval
   - "completion_pending" - Employee claims task is done, needs boss verification
   - "blocker_reported" - Employee reported a blocker or issue
   - "awaiting_employee" - Boss sent message/question, waiting for employee response
   - "awaiting_confirmation" - Boss approved a change, waiting for employee to confirm
   - "boss_proposed" - Boss proposed an alternative (date, approach, etc.)
   - "negotiating" - Active back-and-forth negotiation happening
   - "resolved" - Issue/request was resolved
   - "rejected" - Boss rejected a request (conversation may continue if employee replies)

2. PENDING ITEMS (typed) - What changes/decisions are pending (date/scope/task name/etc.)?
   Only include ACTIVE items that haven't been resolved yet.
   Each item MUST include:
   - parameter (dueDate/scope/taskName/assignee/etc.)
   - changeType (date_change/scope_change/scope_addition/etc.)
   - currentValue, proposedValue
   - requestedBy (boss/employee)
   - awaitingFrom (boss/employee/none)
   - requiresApproval (true/false)
   - status (pending/approved/rejected/confirmed) [minimal subset ok]
   - reasoning (short)

3. SUMMARY - One sentence summary of the current conversation state

4. REQUIRES ACTION - Does the boss need to take action right now?

5. TASK SNAPSHOT (derived truth)
   - taskName
   - dueDateEffective (nullable; ISO yyyy-MM-dd)
   - dueDateProposed (nullable; ISO yyyy-MM-dd)
   - scopeSummary (short)

6. PROVENANCE (per derived field)
   For each: sourceMessageId (use [id:...] from conversation), sourceSnippet, confidence (0-1), extractedAt (ISO)

IMPORTANT RULES:
- If boss rejected something and employee hasn't replied since, state should be "awaiting_employee" or "rejected"
- If employee sent an update (progress, FYI) without requesting anything, state is "update_received"
- If employee explicitly requested a date/scope/role change, state is "change_requested"
- If employee said the task is done/complete, state is "completion_pending"
- Look at the FULL conversation, not just the last message
- If a message is ambiguous (e.g., "push it by a bit") and you cannot extract an explicit date with high confidence, DO NOT change dueDateEffective; instead add a pending item that requests clarification.

Return ONLY valid JSON:
{
  "conversationState": "change_requested",
  "pendingChanges": [
    {
      "id": "change_1",
      "parameter": "dueDate",
      "changeType": "date_change",
      "currentValue": "2025-01-15",
      "proposedValue": "2025-01-25",
      "requestedBy": "employee",
      "awaitingFrom": "boss",
      "requiresApproval": true,
      "status": "pending",
      "reasoning": "Employee requested extension due to dependencies"
    }
  ],
  "summary": "Employee requested 10-day extension due to upstream dependencies.",
  "requiresAction": true,
  "actionSuggestion": "Approve or reject the date change request",
  "taskSnapshot": {
    "taskName": "Q1 Metrics Summary Deck",
    "dueDateEffective": "2026-01-10",
    "dueDateProposed": "2026-01-15",
    "scopeSummary": "Create a concise metrics summary deck and include a summary slide."
  },
  "provenance": {
    "taskName": { "sourceMessageId": "abc123", "sourceSnippet": "Rename this to: Q1 Metrics Summary Deck", "confidence": 0.8, "extractedAt": "2026-01-03T00:00:00.000Z" },
    "dueDateEffective": { "sourceMessageId": "def456", "sourceSnippet": "Approved. New due date is Jan 15.", "confidence": 0.8, "extractedAt": "2026-01-03T00:00:00.000Z" },
    "dueDateProposed": { "sourceMessageId": "ghi789", "sourceSnippet": "Can we move it to Jan 15?", "confidence": 0.8, "extractedAt": "2026-01-03T00:00:00.000Z" },
    "scopeSummary": { "sourceMessageId": "jkl012", "sourceSnippet": "Also add a summary slide with key metrics.", "confidence": 0.8, "extractedAt": "2026-01-03T00:00:00.000Z" }
  }
}`;

    const response = callGeminiPro(prompt, { temperature: 0.2 });
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
    
    // Validate and normalize conversation state
    const validStates = Object.values(CONVERSATION_STATE);
    if (!validStates.includes(analysis.conversationState)) {
      analysis.conversationState = CONVERSATION_STATE.ACTIVE;
    }

    // ------------------------------------------------------------------
    // Guardrails: preserve system-tracked pending changes / confirmations
    // ------------------------------------------------------------------
    // The model prompt focuses on "employee requested" changes, but the system
    // can also track boss-proposed changes that are awaiting employee confirmation.
    // Never allow the model to wipe out those pending changes or mark them resolved.
    let existingPendingChanges = [];
    try {
      existingPendingChanges = task.Pending_Changes ? JSON.parse(task.Pending_Changes) : [];
    } catch (e) {
      existingPendingChanges = [];
    }

    // Preserve existing pending changes if the model returned none.
    if (existingPendingChanges.length > 0 && (!analysis.pendingChanges || analysis.pendingChanges.length === 0)) {
      analysis.pendingChanges = existingPendingChanges;
    }

    // If we have a Pending_Decision awaiting employee, force the appropriate state.
    let pendingDecision = null;
    try {
      pendingDecision = task.Pending_Decision ? JSON.parse(task.Pending_Decision) : null;
    } catch (e) {
      pendingDecision = null;
    }

    const hasAwaitingEmployeePending =
      (pendingDecision && pendingDecision.awaitingFrom === 'employee') ||
      (analysis.pendingChanges || []).some(c => c && c.awaitingFrom === 'employee');

    if (hasAwaitingEmployeePending && (analysis.conversationState === CONVERSATION_STATE.RESOLVED || analysis.conversationState === CONVERSATION_STATE.ACTIVE)) {
      analysis.conversationState = CONVERSATION_STATE.AWAITING_EMPLOYEE;
      analysis.requiresAction = false;
      analysis.summary = analysis.summary || 'Awaiting employee confirmation on proposed change.';
    }
    
    // -----------------------------
    // Derived truth snapshot write
    // -----------------------------
    const nowIso = new Date().toISOString();
    const prevSnapshot = {
      taskName: task.Derived_Task_Name || task.Task_Name || '',
      dueDateEffective: task.Derived_Due_Date_Effective || task.Due_Date || null,
      dueDateProposed: task.Derived_Due_Date_Proposed || task.Proposed_Date || null,
      scopeSummary: task.Derived_Scope_Summary || task.Context_Hidden || ''
    };

    const extractedSnapshot = (analysis && analysis.taskSnapshot) ? analysis.taskSnapshot : {};
    const extractedProv = (analysis && analysis.provenance) ? analysis.provenance : {};
    let existingProv = {};
    try {
      existingProv = task.Derived_Field_Provenance ? JSON.parse(task.Derived_Field_Provenance) : {};
    } catch (e) {
      existingProv = {};
    }
    const mergedProv = Object.assign({}, existingProv || {}, extractedProv || {});

    // No-silent-overwrite guardrail: keep prior value if confidence is low/unknown
    function chooseField(fieldName, extractedValue, prevValue) {
      const prov = extractedProv && extractedProv[fieldName] ? extractedProv[fieldName] : null;
      const conf = prov && typeof prov.confidence === 'number' ? prov.confidence : null;
      if (extractedValue === undefined || extractedValue === null || extractedValue === '') return prevValue;
      if (conf !== null && conf < 0.6) return prevValue;
      return extractedValue;
    }

    const derivedTaskName = chooseField('taskName', extractedSnapshot.taskName, prevSnapshot.taskName);
    const derivedDueDateEffective = chooseField('dueDateEffective', extractedSnapshot.dueDateEffective, prevSnapshot.dueDateEffective);
    const derivedDueDateProposed = chooseField('dueDateProposed', extractedSnapshot.dueDateProposed, prevSnapshot.dueDateProposed);
    const derivedScopeSummary = chooseField('scopeSummary', extractedSnapshot.scopeSummary, prevSnapshot.scopeSummary);

    // Update task with analyzed state + derived snapshot
    updateTask(taskId, {
      Conversation_State: analysis.conversationState,
      Pending_Changes: JSON.stringify(analysis.pendingChanges || []),
      AI_Summary: analysis.summary || '',
      Conversation_History: JSON.stringify(conversationHistory),
      Derived_Task_Name: derivedTaskName,
      Derived_Due_Date_Effective: derivedDueDateEffective || '',
      Derived_Due_Date_Proposed: derivedDueDateProposed || '',
      Derived_Scope_Summary: derivedScopeSummary,
      Derived_Field_Provenance: JSON.stringify(mergedProv || {}),
      Derived_Last_Analyzed_At: nowIso,
      Last_Updated: new Date()
    });
    
    Logger.log(`Conversation analysis for ${taskId}:`);
    Logger.log(`  State: ${analysis.conversationState}`);
    Logger.log(`  Pending Changes: ${(analysis.pendingChanges || []).length}`);
    Logger.log(`  Requires Action: ${analysis.requiresAction}`);
    Logger.log(`  Summary: ${analysis.summary}`);
    
    return { 
      success: true, 
      data: {
        conversationState: analysis.conversationState,
        pendingChanges: analysis.pendingChanges || [],
        summary: analysis.summary,
        requiresAction: analysis.requiresAction,
        actionSuggestion: analysis.actionSuggestion,
        taskSnapshot: {
          taskName: derivedTaskName,
          dueDateEffective: derivedDueDateEffective || null,
          dueDateProposed: derivedDueDateProposed || null,
          scopeSummary: derivedScopeSummary
        },
        provenance: mergedProv || {}
      }
    };
    
  } catch (error) {
    Logger.log(`Error analyzing conversation: ${error.toString()}`);
    return { 
      success: false, 
      error: error.toString(),
      data: {
        conversationState: CONVERSATION_STATE.ACTIVE,
        pendingChanges: [],
        summary: 'Error analyzing conversation',
        requiresAction: false
      }
    };
  }
}

/**
 * Detect parameter changes from conversation history
 * Uses stored Conversation_State and Pending_Changes when available
 * Falls back to analyzeConversationAndUpdateState() if needed
 */
function detectParameterChanges(taskId, forceReanalyze = false) {
  try {
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`Task ${taskId} not found`);
      return { pendingChanges: [], showApprovals: false, conversationState: 'active', bossRejectedLast: false };
    }

    // Derived snapshot + provenance (for UI rendering consistency)
    const taskSnapshot = {
      taskName: task.Derived_Task_Name || task.Task_Name || '',
      dueDateEffective: task.Derived_Due_Date_Effective || task.Due_Date || '',
      dueDateProposed: task.Derived_Due_Date_Proposed || task.Proposed_Date || '',
      scopeSummary: task.Derived_Scope_Summary || task.Context_Hidden || ''
    };
    let provenance = {};
    try {
      provenance = task.Derived_Field_Provenance ? JSON.parse(task.Derived_Field_Provenance) : {};
    } catch (e) {
      provenance = {};
    }
    
    // Check if we have stored state and it's recent (less than 5 minutes old)
    const hasStoredState = task.Conversation_State && task.Pending_Changes;
    const lastUpdated = task.Last_Updated ? new Date(task.Last_Updated) : null;
    const isRecent = lastUpdated && (new Date() - lastUpdated < 5 * 60 * 1000); // 5 minutes
    
    // Use stored state if available and recent (unless force reanalyze)
    if (hasStoredState && isRecent && !forceReanalyze) {
      let pendingChanges = [];
      try {
        pendingChanges = JSON.parse(task.Pending_Changes);
      } catch (e) {
        pendingChanges = [];
      }
      
      const conversationState = task.Conversation_State;
      
      // Determine if approvals should be shown based on state
      const showApprovals = [
        CONVERSATION_STATE.CHANGE_REQUESTED,
        CONVERSATION_STATE.COMPLETION_PENDING,
        CONVERSATION_STATE.BLOCKER_REPORTED
      ].includes(conversationState);
      
      // Check if boss rejected last (for hiding approvals)
      const bossRejectedLast = conversationState === CONVERSATION_STATE.REJECTED || 
                               conversationState === CONVERSATION_STATE.AWAITING_EMPLOYEE;
      
      return {
        pendingChanges: pendingChanges,
        showApprovals: showApprovals && pendingChanges.length > 0,
        conversationState: conversationState,
        bossRejectedLast: bossRejectedLast,
        summary: task.AI_Summary || '',
        draftMessage: null,
        taskSnapshot: taskSnapshot,
        provenance: provenance,
        lastMessageTimestamp: task.Last_Message_Timestamp || '',
        lastMessageSender: task.Last_Message_Sender || '',
        lastMessageSnippet: task.Last_Message_Snippet || ''
      };
    }
    
    // Need to analyze - call the main analysis function
    Logger.log(`Analyzing conversation for task ${taskId} (forceReanalyze: ${forceReanalyze})`);
    const analysisResult = analyzeConversationAndUpdateState(taskId);
    
    if (!analysisResult.success) {
      return { 
        pendingChanges: [], 
        showApprovals: false, 
        conversationState: CONVERSATION_STATE.ACTIVE, 
        bossRejectedLast: false,
        error: analysisResult.error
      };
    }
    
    const analysis = analysisResult.data;
    
    // Determine if approvals should be shown based on state
    const showApprovals = [
      CONVERSATION_STATE.CHANGE_REQUESTED,
      CONVERSATION_STATE.COMPLETION_PENDING,
      CONVERSATION_STATE.BLOCKER_REPORTED
    ].includes(analysis.conversationState);
    
    // Check if boss rejected last
    const bossRejectedLast = analysis.conversationState === CONVERSATION_STATE.REJECTED || 
                             analysis.conversationState === CONVERSATION_STATE.AWAITING_EMPLOYEE;
    
    return {
      pendingChanges: analysis.pendingChanges || [],
      showApprovals: showApprovals && (analysis.pendingChanges || []).length > 0,
      conversationState: analysis.conversationState,
      bossRejectedLast: bossRejectedLast,
      summary: analysis.summary,
      draftMessage: analysis.actionSuggestion || null,
      taskSnapshot: analysis.taskSnapshot || taskSnapshot,
      provenance: analysis.provenance || provenance,
      lastMessageTimestamp: task.Last_Message_Timestamp || '',
      lastMessageSender: task.Last_Message_Sender || '',
      lastMessageSnippet: task.Last_Message_Snippet || ''
    };
    
  } catch (error) {
    Logger.log(`Error detecting parameter changes: ${error.toString()}`);
    return { 
      pendingChanges: [], 
      showApprovals: false, 
      conversationState: CONVERSATION_STATE.ACTIVE, 
      bossRejectedLast: false,
      error: error.toString()
    };
  }
}

/**
 * Extract messages from Interaction_Log format
 * Helper function to parse log entries into structured messages
 */
function extractMessagesFromInteractionLog(log) {
  const messages = [];
  const lines = log.split('\n');
  let currentMessage = null;
  
  for (const line of lines) {
    // Look for timestamp patterns: "2025-12-20 09:00 - ..."
    const timestampMatch = line.match(/^(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})/);
    if (timestampMatch) {
      if (currentMessage) {
        messages.push(currentMessage);
      }
      currentMessage = {
        timestamp: timestampMatch[1],
        content: line.substring(timestampMatch[0].length + 3).trim()
      };
    } else if (currentMessage && line.trim()) {
      currentMessage.content += '\n' + line.trim();
    }
  }
  
  if (currentMessage) {
    messages.push(currentMessage);
  }
  
  return messages;
}

/**
 * Fallback: Get default actions based on status
 */
function getDefaultActionsForStatus(status) {
  const defaults = {
    'review_date': {
      primaryAction: {
        type: 'APPROVE',
        label: 'Approve Date Change',
        description: 'Approve the employee\'s requested date change',
        icon: 'check',
        actionId: 'approve_date_change',
        confidence: 0.8
      },
      secondaryActions: [
        { 
          type: 'MESSAGE', 
          label: 'Message Employee', 
          description: 'Send a message to the employee',
          icon: 'message-square',
          actionId: 'message_employee'
        },
        { 
          type: 'REJECT', 
          label: 'Reject Date Change', 
          description: 'Reject the date change request',
          icon: 'x',
          actionId: 'reject_date_change'
        }
      ],
      approvalState: 'awaiting_boss',
      reasoning: 'Employee has requested a date change',
      requiresEmployeeConfirmation: true
    },
    'review_scope': {
      primaryAction: {
        type: 'MESSAGE',
        label: 'Provide Clarification',
        description: 'Clarify the task scope for the employee',
        icon: 'message-square',
        actionId: 'provide_clarification',
        confidence: 0.8
      },
      secondaryActions: [],
      approvalState: 'awaiting_boss',
      reasoning: 'Employee has questions about task scope',
      requiresEmployeeConfirmation: false
    },
    'review_role': {
      primaryAction: {
        type: 'APPROVE',
        label: 'Confirm Assignment',
        description: 'Confirm the task is correctly assigned',
        icon: 'check',
        actionId: 'override_role',
        confidence: 0.8
      },
      secondaryActions: [
        {
          type: 'MESSAGE',
          label: 'Reassign Task',
          description: 'Reassign the task to someone else',
          icon: 'user',
          actionId: 'reassign_task'
        }
      ],
      approvalState: 'awaiting_boss',
      reasoning: 'Employee has questioned task assignment',
      requiresEmployeeConfirmation: false
    },
    'completed': {
      primaryAction: {
        type: 'APPROVE',
        label: 'Approve Completion',
        description: 'Approve the task as completed',
        icon: 'check',
        actionId: 'approve_done',
        confidence: 0.8
      },
      secondaryActions: [
        {
          type: 'MESSAGE',
          label: 'Request Proof',
          description: 'Request proof of completion',
          icon: 'file-text',
          actionId: 'request_proof'
        }
      ],
      approvalState: 'awaiting_boss',
      reasoning: 'Employee has marked task as complete',
      requiresEmployeeConfirmation: false
    }
  };
  
  return defaults[status] || {
    primaryAction: { 
      type: 'NONE', 
      label: 'No action needed',
      description: 'No action required at this time',
      icon: 'info',
      actionId: 'none',
      confidence: 1.0
    },
    secondaryActions: [],
    approvalState: 'none',
    reasoning: 'No action needed for this status',
    requiresEmployeeConfirmation: false
  };
}

/**
 * Generate AI summary of employee review request
 * Summarizes what the employee is asking for in their reply
 */
function summarizeReviewRequest(reviewType, employeeReply, taskName, currentDueDate, proposedDate) {
  let prompt = '';
  
  switch (reviewType) {
    case 'DATE_CHANGE':
      prompt = `An employee has requested a date change for a task. Summarize their request in a clear, concise way.

Task: "${taskName}"
Current Due Date: ${currentDueDate || 'Not set'}
Proposed Date: ${proposedDate || 'Not specified in reply'}
Employee's Reply: "${employeeReply}"

Generate a brief summary (2-3 sentences) explaining:
1. What the employee is requesting
2. Why they need the change (if mentioned)
3. What date they're proposing (if mentioned)

Return ONLY the summary text (no markdown, no JSON, just plain text).`;
      break;
      
    case 'SCOPE_QUESTION':
      prompt = `An employee has questions about the scope of a task. Summarize their questions and concerns.

Task: "${taskName}"
Employee's Reply: "${employeeReply}"

Generate a brief summary (2-3 sentences) explaining:
1. What questions or concerns the employee has
2. What clarification they need
3. Any specific aspects they're unsure about

Return ONLY the summary text (no markdown, no JSON, just plain text).`;
      break;
      
    case 'ROLE_REJECTION':
      prompt = `An employee has indicated this task is not their responsibility. Summarize their concern.

Task: "${taskName}"
Employee's Reply: "${employeeReply}"

Generate a brief summary (2-3 sentences) explaining:
1. Why the employee believes this isn't their responsibility
2. Who they think it should be assigned to (if mentioned)
3. Any context they provided

Return ONLY the summary text (no markdown, no JSON, just plain text).`;
      break;
      
    default:
      prompt = `An employee has sent a reply about a task. Summarize their message.

Task: "${taskName}"
Employee's Reply: "${employeeReply}"

Generate a brief summary (2-3 sentences) explaining what the employee is communicating.

Return ONLY the summary text (no markdown, no JSON, just plain text).`;
  }
  
  try {
    const summary = callGeminiPro(prompt, { 
      temperature: 0.5,
      maxOutputTokens: 500 
    });
    return summary.trim();
  } catch (error) {
    Logger.log(`Error generating review summary: ${error.toString()}`);
    // Fallback summary
    return `Employee has sent a ${reviewType.toLowerCase().replace('_', ' ')} request regarding this task.`;
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

/**
 * Analyze manually created MoM document
 * Extracts action items and project knowledge
 */
function analyzeMoMDocument(momText) {
  // Try to load prompt from sheet, fallback to hardcoded
  let promptTemplate = null;
  try {
    const promptData = getPrompt('analyzeMoMDocument', 'mom');
    if (promptData && promptData.Content) {
      promptTemplate = promptData.Content;
    }
  } catch (e) {
    Logger.log('Could not load prompt from sheet, using fallback: ' + e.toString());
  }
  
  // Fallback to hardcoded prompt if not found in sheet
  if (!promptTemplate) {
    promptTemplate = `You are analyzing a manually created Minutes of Meeting (MoM) document. Extract the following information:

1. Action Items: List all action items mentioned in the document. For each action item, extract:
   - description: What needs to be done
   - owner: Who is responsible (name or email if mentioned)
   - deadline: When it's due (extract date if mentioned, format as YYYY-MM-DD or null)
   - priority: High, Medium, or Low (based on urgency indicators)

2. Projects: Identify all projects mentioned in the document. For each project, extract:
   - project_tag: The project identifier/tag (if mentioned, or infer from project name)
   - project_name: The full project name
   - summary: Brief summary of what was discussed about this project
   - key_points: Array of key points discussed (max 10 points)
   - decisions: Array of decisions made about the project (max 10 decisions)
   - include_full_context: Boolean - true if full document text should be stored

3. Executive Summary: A brief summary of the entire meeting (2-3 sentences)

Return your response as a JSON object with this structure:
{
  "action_items": [
    {
      "description": "string",
      "owner": "string or null",
      "deadline": "YYYY-MM-DD or null",
      "priority": "High|Medium|Low"
    }
  ],
  "projects": [
    {
      "project_tag": "string or null",
      "project_name": "string",
      "summary": "string",
      "key_points": ["string"],
      "decisions": ["string"],
      "include_full_context": boolean
    }
  ],
  "executive_summary": "string"
}

MoM document text:
"{{MOM_TEXT}}"

Return your response as a JSON object with this structure.`;
  }
  
  // Replace template variables
  const prompt = promptTemplate.replace(/\{\{MOM_TEXT\}\}/g, momText);

  try {
    const response = callGeminiPro(prompt, { temperature: 0.3 });
    
    // Extract JSON from response
    let jsonText = response.trim();
    if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```json\n?/, '').replace(/```$/, '');
    }
    if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```\n?/, '').replace(/```$/, '');
    }
    
    const parsed = JSON.parse(jsonText);
    
    // Ensure arrays exist
    if (!parsed.action_items) parsed.action_items = [];
    if (!parsed.projects) parsed.projects = [];
    if (!parsed.executive_summary) parsed.executive_summary = '';
    
    return parsed;
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'analyzeMoMDocument', error.toString());
    // Return empty structure on error
    return {
      action_items: [],
      projects: [],
      executive_summary: 'Failed to analyze MoM document'
    };
  }
}


