const COUNTS_SHEET_NAME   = 'count';
const MAIN_SHEET_NAME     = 'Main';
const DIFF_SHEET_NAME     = 'Difference_Report';

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("PIV Tools")
    .addItem("Rebuild Difference Report", "rebuildDifferenceReport")
    .addToUi();
}

function saveScanData(
  partId, desc, loc, systemStock,
  areaType, subLocation,
  boxCount, qtyPerBox, looseQty, physicalQty,
  recorderName
) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(COUNTS_SHEET_NAME);
  const locationToSave = (areaType === 'Loc_area_1') ? loc : '';

  sheet.appendRow([
    new Date(),              // A Timestamp
    partId,                  // B Part_ID
    desc,                    // C Description
    areaType,                // D Area_Type
    locationToSave,          // E area_1_Location
    subLocation || "",       // F Sub_Location(Prod.)
    Number(systemStock),     // G System_Stock
    Number(boxCount),        // H Box_Count
    Number(qtyPerBox),       // I Qty_Per_Box
    Number(physicalQty),     // J Physical_Qty
    Number(looseQty) || 0,   // K Loose_Qty
    "",                      // L (reserved / blank)
    recorderName || ""       // M Recorder_Name
  ]);

  return `Saved: ${partId} | Physical Qty = ${physicalQty}`;
}

/* Difference report UNCHANGED */
function rebuildDifferenceReport() {
  const ss = SpreadsheetApp.getActive();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  const countSheet = ss.getSheetByName(COUNTS_SHEET_NAME);
  let diffSheet = ss.getSheetByName(DIFF_SHEET_NAME);

  if (!diffSheet) diffSheet = ss.insertSheet(DIFF_SHEET_NAME);
  else if (diffSheet.getLastRow() > 1)
    diffSheet.getRange(2,1,diffSheet.getLastRow()-1,diffSheet.getLastColumn()).clearContent();

  const mainVals = mainSheet.getDataRange().getValues();
  const countVals = countSheet.getDataRange().getValues();

  const sysMap = {};
  for (let i=1;i<mainVals.length;i++){
    const id = mainVals[i][0];
    if (!sysMap[id]) sysMap[id]={desc:mainVals[i][1],sys:0};
    sysMap[id].sys += Number(mainVals[i][3])||0;
  }

  const physMap = {};
  for (let i=1;i<countVals.length;i++){
    const id = countVals[i][1];
    const packed = Number(countVals[i][9])  || 0;  // J Physical_Qty
    const loose  = Number(countVals[i][10]) || 0;  // K Loose_Qty

    physMap[id] = (physMap[id] || 0) + packed + loose;
  }

  const ids=[...new Set([...Object.keys(sysMap),...Object.keys(physMap)])];
  const out=[["Part_ID","Description","system_stock","physical_total","difference"]];
  ids.forEach(id=>{
    const s=sysMap[id]?.sys||0;
    const p=physMap[id]||0;
    out.push([id,sysMap[id]?.desc||"",s,p,p-s]);
  });

  diffSheet.getRange(1,1,out.length,5).setValues(out);
}
