/**
 * VERIFICATION FUNCTION - Run this in Apps Script to verify deployment
 * This function checks if all new functions exist
 */
function verifyDeployment() {
  Logger.log('=== Deployment Verification ===');
  
  // Check if new functions exist
  const functionsToCheck = [
    'findStaffEmailByName',
    'findProjectTagByName',
    'classifyReplyType',
    'markFileAsUnclear',
    'extractDateFromText'
  ];
  
  const missing = [];
  const found = [];
  
  functionsToCheck.forEach(funcName => {
    try {
      if (typeof eval(funcName) === 'function') {
        found.push(funcName);
        Logger.log(`✅ ${funcName} - EXISTS`);
      } else {
        missing.push(funcName);
        Logger.log(`❌ ${funcName} - MISSING`);
      }
    } catch (e) {
      missing.push(funcName);
      Logger.log(`❌ ${funcName} - ERROR: ${e.toString()}`);
    }
  });
  
  // Check file line counts (approximate)
  Logger.log('\n=== File Verification ===');
  Logger.log('Note: Line counts may vary, but functions should exist');
  
  Logger.log(`\n✅ Found: ${found.length}/${functionsToCheck.length} functions`);
  if (missing.length > 0) {
    Logger.log(`❌ Missing: ${missing.join(', ')}`);
    Logger.log('\n⚠️  DEPLOYMENT INCOMPLETE - Some functions are missing!');
    Logger.log('Please run: clasp push --force');
  } else {
    Logger.log('\n✅ ALL FUNCTIONS DEPLOYED SUCCESSFULLY!');
  }
  
  // Test findStaffEmailByName if it exists
  if (found.includes('findStaffEmailByName')) {
    Logger.log('\n=== Testing findStaffEmailByName ===');
    try {
      const testResult = findStaffEmailByName('Test Name');
      Logger.log(`Test result: ${testResult || 'null (expected if Test Name not in STAFF_DB)'}`);
    } catch (e) {
      Logger.log(`Test error: ${e.toString()}`);
    }
  }
  
  // Test findProjectTagByName if it exists
  if (found.includes('findProjectTagByName')) {
    Logger.log('\n=== Testing findProjectTagByName ===');
    try {
      const testResult = findProjectTagByName('Test Project');
      Logger.log(`Test result: ${testResult || 'null (expected if Test Project not in PROJECTS_DB)'}`);
    } catch (e) {
      Logger.log(`Test error: ${e.toString()}`);
    }
  }
  
  Logger.log('\n=== Verification Complete ===');
}

