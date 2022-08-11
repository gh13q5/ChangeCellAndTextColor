/**
 *  이름이 지정된 범위에 #REF! 잘못된 참조가 발생했을 때 사용
 *  반드시 필요한 코드 XXXXXX
 */

/**
var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = spreadSheet.getActiveSheet();
var activateCell = sheet.getActiveCell();

function onEdit(e) {
  if(activateCell.getA1Notation() === sheet.getRange('A2').getA1Notation()){
    removeDeadReferences();
  }
}
*/

/** For Reset */
function removeAllNamedRanges(){
  var namedRanges = SpreadsheetApp.getActive().getNamedRanges();
  for (var i = 0; i < namedRanges.length; i++) {
    namedRanges[i].remove();
  }
}

/** 특정 name의 이름이 지정된 범위 삭제 */
function removeDeadReferences(){
  var sheetNamedRanges = sheet.getNamedRanges();
  if (sheetNamedRanges.length){
    for (i = 0; i < sheetNamedRanges.length; i++){
      try{
        nowNamedRange = sheetNamedRanges[i].getRange();
      }
      catch (error){
        nowNamedRange = null;
      }

      nowNamedRangeA1Notation = (nowNamedRange != null ? nowNamedRange.getA1Notation() : false);
      if (nowNamedRangeA1Notation == false || nowNamedRangeA1Notation.slice(0,1) === '#'
      || nowNamedRangeA1Notation.slice(-1) === '!' || nowNamedRangeA1Notation.indexOf('REF') > -1){
        sheetNamedRanges[i].remove();
      }
    }
  }
}
