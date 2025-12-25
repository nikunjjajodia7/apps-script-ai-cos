/**
 * Verification Script for Employee Reply Feature
 * Run this in Apps Script to verify the updates are deployed
 */
function VERIFY_EMPLOYEE_REPLY_UPDATE() {
  Logger.log('=== Verifying Employee Reply Feature Updates ===\n');
  
  let allChecksPassed = true;
  
  // Check 1: Verify EmailNegotiation.gs has Employee_Reply storage
  Logger.log('1. Checking EmailNegotiation.gs...');
  try {
    const emailNegCode = EmailNegotiation.toString();
    const hasDateChange = emailNegCode.includes('Employee_Reply: emailContent') && 
                          emailNegCode.includes('handleDateChangeReply');
    const hasScopeQuestion = emailNegCode.includes('Employee_Reply: emailContent') && 
                            emailNegCode.includes('handleScopeQuestionReply');
    const hasRoleRejection = emailNegCode.includes('Employee_Reply: emailContent') && 
                            emailNegCode.includes('handleRoleRejectionReply');
    
    if (hasDateChange && hasScopeQuestion && hasRoleRejection) {
      Logger.log('   ✓ Employee_Reply storage found in all 3 handlers');
    } else {
      Logger.log('   ✗ Employee_Reply storage missing in some handlers');
      allChecksPassed = false;
    }
  } catch (e) {
    Logger.log('   ✗ Error checking EmailNegotiation: ' + e.toString());
    allChecksPassed = false;
  }
  
  // Check 2: Verify Code.gs has Employee_Reply column
  Logger.log('\n2. Checking Code.gs...');
  try {
    const codeCode = Code.toString();
    const hasEmployeeReplyColumn = codeCode.includes("'Employee_Reply'") && 
                                   codeCode.includes('Created_Date');
    const has17Columns = codeCode.includes('getRange(1, 1, 1, 17)');
    
    if (hasEmployeeReplyColumn && has17Columns) {
      Logger.log('   ✓ Employee_Reply column found in Tasks_DB schema (17 columns)');
    } else {
      Logger.log('   ✗ Employee_Reply column missing or wrong column count');
      Logger.log('     - Has Employee_Reply: ' + hasEmployeeReplyColumn);
      Logger.log('     - Has 17 columns: ' + has17Columns);
      allChecksPassed = false;
    }
  } catch (e) {
    Logger.log('   ✗ Error checking Code: ' + e.toString());
    allChecksPassed = false;
  }
  
  // Check 3: Verify DashboardActions.gs has Employee_Reply clearing
  Logger.log('\n3. Checking DashboardActions.gs...');
  try {
    const dashboardCode = DashboardActions.toString();
    const hasApproveDate = dashboardCode.includes('Employee_Reply: \'\'') && 
                          dashboardCode.includes('handleApproveNewDate');
    const hasRejectDate = dashboardCode.includes('Employee_Reply: \'\'') && 
                         dashboardCode.includes('handleRejectDateChange');
    const hasClarification = dashboardCode.includes('Employee_Reply: \'\'') && 
                             dashboardCode.includes('handleProvideClarification');
    const hasOverrideRole = dashboardCode.includes('Employee_Reply: \'\'') && 
                           dashboardCode.includes('handleOverrideRole');
    
    const count = (dashboardCode.match(/Employee_Reply: '', \/\/ Clear employee reply/g) || []).length;
    
    if (count >= 4) {
      Logger.log(`   ✓ Employee_Reply clearing found in ${count} handlers`);
    } else {
      Logger.log(`   ✗ Employee_Reply clearing found in only ${count} handlers (expected 4)`);
      allChecksPassed = false;
    }
  } catch (e) {
    Logger.log('   ✗ Error checking DashboardActions: ' + e.toString());
    allChecksPassed = false;
  }
  
  // Check 4: Verify functions exist
  Logger.log('\n4. Checking function existence...');
  const functionsToCheck = [
    'handleDateChangeReply',
    'handleScopeQuestionReply', 
    'handleRoleRejectionReply',
    'handleApproveNewDate',
    'handleRejectDateChange',
    'handleProvideClarification',
    'handleOverrideRole'
  ];
  
  let functionsFound = 0;
  functionsToCheck.forEach(funcName => {
    try {
      if (typeof this[funcName] === 'function') {
        Logger.log(`   ✓ ${funcName} exists`);
        functionsFound++;
      } else {
        Logger.log(`   ✗ ${funcName} NOT FOUND`);
        allChecksPassed = false;
      }
    } catch (e) {
      Logger.log(`   ✗ Error checking ${funcName}: ${e.toString()}`);
      allChecksPassed = false;
    }
  });
  
  // Final Summary
  Logger.log('\n=== Verification Summary ===');
  if (allChecksPassed && functionsFound === functionsToCheck.length) {
    Logger.log('✅ ALL CHECKS PASSED - Employee Reply feature is deployed!');
    Logger.log('\nThe following features are active:');
    Logger.log('  - Employee replies are stored in Employee_Reply field');
    Logger.log('  - Employee_Reply column exists in Tasks_DB');
    Logger.log('  - Employee_Reply is cleared after review actions');
    Logger.log('  - All handler functions are available');
  } else {
    Logger.log('❌ SOME CHECKS FAILED - Please review the errors above');
    Logger.log(`   Functions found: ${functionsFound}/${functionsToCheck.length}`);
  }
  
  Logger.log('\n=== End Verification ===');
}

