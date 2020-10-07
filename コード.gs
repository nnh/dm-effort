/**
* Obtain data to create a pie chart
* @param none
* @return none
*/
function getDataCreateChart(){
  const ui = SpreadsheetApp.getUi();
  const inputSheets = getInputSheet();
  const outputSheets = getOutputSheet();
  const colNames = setColumnName();
  const colIndex = getColumnIndex(colNames);
  const inputStartDateAddress = 'A2';
  const startDate = outputSheets.conditionPreference.getRange(inputStartDateAddress).getValue();
  if (!(dateCheck(startDate))){
    ui.alert(outputSheets.conditionPreference.getName() + 'シートの' + inputStartDateAddress + 'セルに集計開始日を入力して再実行してください');
    return;
  }
  const inputEndDateAddress = 'B2';
  const endDate = outputSheets.conditionPreference.getRange(inputEndDateAddress).getValue();
  if (!(dateCheck(endDate))){
    ui.alert(outputSheets.conditionPreference.getName() + 'シートの' + inputStartDateAddress + 'セルに集計終了日を入力して再実行してください');
    return;
  }
  /** Copy the working hour and aggregate category */
  getWorkingHoursRate(inputSheets, outputSheets, colIndex);
  getCategoryForAggregate(inputSheets, outputSheets, colIndex);
  /** Extract the effort between start and end dates */
  const matome1 = new classSetValuesFromRanges(inputSheets.dataSource);
  const matome1FilterInfo = {startColIdx:colIndex.dataSourceYmd, startDate:startDate, endColIdx:colIndex.dataSourceYmd, endDate:endDate};
  const dataSourceValues = matome1.filterTargetPeriod(matome1FilterInfo);
  /** Adding formulas to datasource sheet */
  const temp = dataSourceValues.map(function(targetRow, idx){
    const temp = targetRow;
    const rowNum = idx + 1;
    const businessHoursSheetName = outputSheets.businessHours.getName();
    const categorySheetName = outputSheets.category.getName();
    /** Example of a formula...'=filter(businessHours!F:F,(businessHours!A:A=A2)*(businessHours!D:D<=B2)*((businessHours!E:E>=B2)+(businessHours!E:E="")))' */
    const workingHoursRateFormula = '=filter(' + businessHoursSheetName + '!' + colNames.businessHoursPer + ':' + colNames.businessHoursPer + ',' + 
                                           '(' + businessHoursSheetName + '!' + colNames.businessHoursName + ':' + colNames.businessHoursName + '=' + colNames.businessHoursName + rowNum + ')' + '*' + 
                                           '(' + businessHoursSheetName + '!' + colNames.businessHoursStart + ':' + colNames.businessHoursStart + '<=' + colNames.businessHoursTimePerDay + rowNum + ')' + '*' + 
                                           '((' + businessHoursSheetName + '!' + colNames.businessHoursEnd + ':' + colNames.businessHoursEnd + '>=' + colNames.businessHoursTimePerDay + rowNum + ')' + '+' + 
                                            '(' + businessHoursSheetName + '!' + colNames.businessHoursEnd + ':' + colNames.businessHoursEnd + '="")))';
    const categoryFormula = '=vlookup(' + colNames.dataSourceProtocolId + rowNum + ',' + categorySheetName + '!' + colNames.dmOfficeItem + ':' + colNames.dmOfficeCategory + ',2,false)';
    if (rowNum > 1){
      temp.push(workingHoursRateFormula);
      temp.push('=' + colNames.dataSourceEffort + rowNum + '*' + colNames.dataSourcePer + rowNum);
      temp.push(categoryFormula);
    } else {
      /** A first row is a header */ 
      temp.push('勤務割合');
      temp.push('Effort%（補正）');
      temp.push('カテゴリー');
    }
    return temp;
  });
  const outputDataSourceValues = new classSetValues(temp);
  outputDataSourceValues.copyFromArrayToRange(outputSheets.dataSource);
  outputSheets.summary.getRange('A1').setValue('集計期間：'+ Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy/MM/dd') + '〜' + Utilities.formatDate(endDate, 'Asia/Tokyo', 'yyyy/MM/dd'));
  SpreadsheetApp.flush();
  ui.alert('処理が完了しました。');
}
/**
* Output work hour information
* @param {Object} input sheet object
* @param {Object} output sheet object
* @param {Object} an associative array of array index
* @return none
*/
function getWorkingHoursRate(inputSheets, outputSheets, colIndex){
  const kinmujikan = new classSetValuesFromRanges(inputSheets.businessHours);
  /** Set the percentage of shortened hours for a 40-hour work week as 100% */
  const businessHoursValues = kinmujikan.targetValues.map(function(x, idx){
    var temp = x;
    var effortPer;
    if (idx > 0){ 
      effortPer = (x[colIndex.businessHoursTimePerDay] * x[colIndex.businessHoursDaysPerWeek]) / 40;
    } else {
      effortPer = '勤務割合';
    }
    temp.push(effortPer);
    return temp;
  });
  const outputBusinessHours = new classSetValues(businessHoursValues);
  outputBusinessHours.copyFromArrayToRange(outputSheets.businessHours);  
}
/**
* Output the aggregate category
* @param {Object} input sheet object
* @param {Object} output sheet object
* @param {Object} an associative array of array index
* @return none
*/
function getCategoryForAggregate(inputSheets, outputSheets, colIndex){
  /** Remove the header */
  const protocolIdValues = new classSetValuesFromRanges(inputSheets.protocolId).targetValues.filter((x, idx) => idx != 0);
  const protocolIdKeyValue = protocolIdValues.map(x => [x[colIndex.protocolIdProtocolId], x[colIndex.protocolIdOrganization]]);
  const dmOfficeValues = new classSetValuesFromRanges(inputSheets.dmOffice).targetValues.filter((x, idx) => idx != 0);
  const dmOfficeKeyValues = dmOfficeValues.map(x => [x[colIndex.dmOfficeItem], x[colIndex.dmOfficeCategory]]);
  /** Remove the blank row */ 
  const categoryValues = protocolIdKeyValue.concat(dmOfficeKeyValues).filter(x => x[0] != '');
  /** Set header */
  const headerValues = [['items', 'category']];
  const outputCategory = new classSetValues(headerValues.concat(categoryValues));
  outputCategory.copyFromArrayToRange(outputSheets.category); 
}
/** Class Column Handling */
class classGetColumnInfo{
  /** 
  * Returns a column number from a column name
  * @param {string}
  */
  constructor(columnName){
    this.columnNumber = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(columnName + '1').getColumn();
  }
  /** 
  * Returns the array index from the column number
  * @return {number}
  */
  getArrayIndex(){
    var temp = this.columnNumber;
    temp--;
    return temp;
  }
  /** 
  * Returns the column number for query from the column number
  * @return {string} e.g. 'Col2'
  */
  createQueryColumnName(){
    return 'Col' + this.columnNumber;
  }  
}
/** Class Output values */
class classSetValues{
  constructor(targetValues){
    this.targetValues = targetValues;
  }
  /** 
  * Extract the values based on the conditions of the argument
  * @param {Object} Key: startColIdx Value: Column number to be extracted
  * @param {Object} Key: startDate Value: First date of extractione
  * @param {Object} Key: endColIdx Value: Column number to be extracted
  * @param {Object} Key: endDate Value: End date of extractione
  * @return {Array}
  */
  filterTargetPeriod(filterInfo){
    const temp = this.targetValues.filter((x, idx) => idx == 0 || (x[filterInfo.startColIdx] >= filterInfo.startDate && x[filterInfo.endColIdx] <= filterInfo.endDate));
    return temp;
  }
  /** 
  * Output values to an output sheet
  * @param {Object} Output Sheet
  */
  copyFromArrayToRange(targetSheet){
    targetSheet.clear();
    targetSheet.getRange(1, 1, this.targetValues.length, this.targetValues[0].length).setValues(this.targetValues);
  }
}
/**
* Class Output values from input sheet
* @extends classSetValues
* @param {Object} Input Sheet
*/
class classSetValuesFromRanges extends classSetValues{
  constructor(inputSheet){
    super(null);
    this.targetValues = inputSheet.getRange(1, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues();
  }
}
/**
* Returns true if the argument is a date, false otherwise.
* @param {string}
* @return {boolean}
*/
function dateCheck(targetValue){
  return Date.prototype.isPrototypeOf(targetValue);
}
/**
* Create an Associative Array of Column Name Constants
* @param none
* @return {Object} Key: Name(e.g. 'dataSourceName'), Value: Column name (e.g. 'A')
*/
function setColumnName(){
  const colName = {};
  colName.dataSourceName = 'A';
  colName.dataSourceYmd = 'B';
  colName.dataSourceEffort = 'C';
  colName.dataSourceProtocolId = 'D';
  colName.dataSourcePer = 'E';
  colName.protocolIdProtocolId = 'A';
  colName.protocolIdOrganization = 'G';
  colName.dmOfficeItem = 'A';
  colName.dmOfficeCategory = 'B';
  colName.businessHoursName = 'A';
  colName.businessHoursTimePerDay = 'B';
  colName.businessHoursDaysPerWeek = 'C';
  colName.businessHoursStart = 'D';
  colName.businessHoursEnd = 'E';
  colName.businessHoursPer = 'F';
  return colName;
}
/**
* Returns an associative array of array index from an associative array of column names
* @param {Object} Key: Name(e.g. 'dataSourceName'), Value: Column name (e.g. 'A')
* @return {Object} Key: Name(e.g. 'dataSourceName'), Value: Array index (e.g. '0')
*/
function getColumnIndex(colName){
  const colIndex = {};
  Object.keys(colName).forEach(x => colIndex[x] = new classGetColumnInfo(colName[x]).getArrayIndex());
  return colIndex;
}
/**
* Get the sheet object
* @param none
* @return {Object} Key: Sheet name, Value: Sheet object
*/
function getInputSheet(){
  const inputFileId = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('wk').getRange('B2').getValue();
  const ss = SpreadsheetApp.openById(inputFileId);
  const sheetObj = {};
  sheetObj.dataSource = ss.getSheetByName('Matome1');
  sheetObj.protocolId = ss.getSheetByName('wk_Protocol_ID');
  sheetObj.dmOffice = ss.getSheetByName('option_DM_Office');
  sheetObj.businessHours = ss.getSheetByName('業務時間');
  return sheetObj;
}
function getOutputSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetObj = {};
  sheetObj.conditionPreference = ss.getSheetByName('input');
  sheetObj.dataSource = ss.getSheetByName('dataSource');
  sheetObj.category = ss.getSheetByName('category');
  sheetObj.businessHours = ss.getSheetByName('businessHours');
  sheetObj.summary = ss.getSheetByName('summary');
  return sheetObj;
}
