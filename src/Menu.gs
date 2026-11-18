/**
 * Gmail Cleanup – the spreadsheet menu.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🧹 Gmail cleanup')
    .addItem('Audit: start (from scratch)', 'startAudit')
    .addItem('Audit: resume', 'runBatch')
    .addItem('Audit: rebuild overview', 'buildOverview')
    .addItem('Audit: compact sender export', 'exportCompact')
    .addItem('Audit: stop', 'stopAudit')
    .addSeparator()
    .addItem('1. Prepare rules', 'prepareRules')
    .addItem('2. Build plan (dry run)', 'buildPlan')
    .addItem('3. Execute plan', 'executePlanFromMenu')
    .addSeparator()
    .addItem('Run weekly clean-up now', 'weeklyRun')
    .addItem('Activate weekly run', 'activateWeekly')
    .addItem('Stop weekly run', 'deactivateWeekly')
    .addToUi();
}
