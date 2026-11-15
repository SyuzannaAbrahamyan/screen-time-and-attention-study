/**
 * Google Apps Script Web App Endpoint
 * Receives JSON payloads from the browser-based SART & Stroop experiment
 * and logs participant summaries + trial-by-trial records to Google Sheets.
 */
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payload = JSON.parse(e.postData.contents);
  
  // 1. Append 1 summary row to the "Summary" sheet
  const summarySheet = ss.getSheetByName("Summary");
  const s = payload.summary;
  summarySheet.appendRow([
    s.participant_id,
    s.timestamp,
    s.daily_screen_time_hours,
    s.smds_score,
    s.smds_item_responses,
    s.sart_commission_errors,
    s.sart_omission_errors,
    s.sart_mean_rt_ms,
    s.sart_sd_rt_ms,
    s.stroop_match_rt_ms,
    s.stroop_mismatch_rt_ms,
    s.stroop_interference_cost_ms,
    s.stroop_match_acc,
    s.stroop_mismatch_acc
  ]);

  // 2. Batch-append all 104 trial rows to the "Detailed_Trials" sheet
  const trialsSheet = ss.getSheetByName("Detailed_Trials");
  const trials = payload.trials;
  if (trials && trials.length > 0) {
    const rows = trials.map(t => [
      t.participant_id,
      t.task,
      t.trial_num,
      t.stimulus,
      t.condition,
      t.response,
      t.correct,
      t.rt_ms
    ]);
    trialsSheet.getRange(trialsSheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
