/** @OnlyCurrentDoc */

var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = spreadSheet.getActiveSheet();
var activateCell = sheet.getActiveCell();

/** Named Range Info */
var firstMainTheme = ['FirstMainTheme', '19'];
var firstSub1Theme = ['FirstSub1Theme', '127'];
var firstSub2Theme = ['FirstSub2Theme', '21'];
var firstSub3Theme = ['FirstSub3Theme', '9'];
var firstBorderAll = ['FirstBorderAll', '51'];
var firstBorderTop = ['FirstBorderTop', '6'];
var firstBorderMiddle = ['FirstBorderMiddle', '6'];
var firstBorderBottom = ['FirstBorderBottom', '12'];
var firstBorderMerge = ['FirstBorderMerge', '18'];
var firstBorderHorizon = ['FirstBorderHorizon', '2'];

var secondMainTheme = ['SecondMainTheme', '19'];
var secondSub1Theme = ['SecondSub1Theme', '127'];
var secondSub2Theme = ['SecondSub2Theme', '21'];
var secondSub3Theme = ['SecondSub3Theme', '9'];
var secondBorderAll = ['SecondBorderAll', '51'];
var secondBorderTop = ['SecondBorderTop', '6'];
var secondBorderMiddle = ['SecondBorderMiddle', '6'];
var secondBorderBottom = ['SecondBorderBottom', '12'];
var secondBorderMerge = ['SecondBorderMerge', '18'];
var secondBorderHorizon = ['SecondBorderHorizon', '2'];

var thirdMainTheme = ['ThirdMainTheme', '19'];
var thirdSub1Theme = ['ThirdSub1Theme', '127'];
var thirdSub2Theme = ['ThirdSub2Theme', '21'];
var thirdSub3Theme = ['ThirdSub3Theme', '9'];
var thirdBorderAll = ['ThirdBorderAll', '51'];
var thirdBorderTop = ['ThirdBorderTop', '6'];
var thirdBorderMiddle = ['ThirdBorderMiddle', '6'];
var thirdBorderBottom = ['ThirdBorderBottom', '12'];
var thirdBorderMerge = ['ThirdBorderMerge', '18'];
var thirdBorderHorizon = ['ThirdBorderHorizon', '2'];

var fourthMainTheme = ['FourthMainTheme', '19'];
var fourthSub1Theme = ['FourthSub1Theme', '127'];
var fourthSub2Theme = ['FourthSub2Theme', '21'];
var fourthSub3Theme = ['FourthSub3Theme', '9'];
var fourthBorderAll = ['FourthBorderAll', '51'];
var fourthBorderTop = ['FourthBorderTop', '6'];
var fourthBorderMiddle = ['FourthBorderMiddle', '6'];
var fourthBorderBottom = ['FourthBorderBottom', '12'];
var fourthBorderMerge = ['FourthBorderMerge', '18'];
var fourthBorderHorizon = ['FourthBorderHorizon', '2'];


function onEdit(e) {
  switch(activateCell.getA1Notation()){
    case (spreadSheet.getRangeByName('FirstSub3ThemeColor').getA1Notation()):
      setCellColor('FirstSub3ThemeColor', firstSub3Theme);
      break;
    case (spreadSheet.getRangeByName('FirstSub2ThemeColor').getA1Notation()):
      setCellColor('FirstSub2ThemeColor', firstSub2Theme);
      break;
    case (spreadSheet.getRangeByName('FirstSub1ThemeColor').getA1Notation()):
      setCellColor('FirstSub1ThemeColor', firstSub1Theme);
      break;
    case (spreadSheet.getRangeByName('FirstMainThemeColor').getA1Notation()):
      setCellColor('FirstMainThemeColor', firstMainTheme);
      setOutBorderColor('FirstMainThemeColor', firstBorderAll);
      setPartialBorderColor('FirstMainThemeColor', firstBorderTop, firstBorderMiddle, firstBorderBottom);
      setInBorderColor('FirstMainThemeColor', firstBorderAll, firstBorderMerge, firstBorderHorizon);
      break;
    case (spreadSheet.getRangeByName('FirstDefaultTextColor').getA1Notation()):
      setTextColor('FirstDefaultTextColor', firstBorderAll);
      setTextColor('FirstDefaultTextColor', firstBorderTop);
      setTextColor('FirstMainTextColor', firstMainTheme);
      break;
    case (spreadSheet.getRangeByName('FirstMainTextColor').getA1Notation()):
      setTextColor('FirstMainTextColor', firstMainTheme);
      break;
    case (spreadSheet.getRangeByName('SecondSub3ThemeColor').getA1Notation()):
      setCellColor('SecondSub3ThemeColor', secondSub3Theme);
      break;
    case (spreadSheet.getRangeByName('SecondSub2ThemeColor').getA1Notation()):
      setCellColor('SecondSub2ThemeColor', secondSub2Theme);
      break;
    case (spreadSheet.getRangeByName('SecondSub1ThemeColor').getA1Notation()):
      setCellColor('SecondSub1ThemeColor', secondSub1Theme);
      break;
    case (spreadSheet.getRangeByName('SecondMainThemeColor').getA1Notation()):
      setCellColor('SecondMainThemeColor', secondMainTheme);
      setOutBorderColor('SecondMainThemeColor', secondBorderAll);
      setPartialBorderColor('SecondMainThemeColor', secondBorderTop, secondBorderMiddle, secondBorderBottom);
      setInBorderColor('SecondMainThemeColor', secondBorderAll, secondBorderMerge, secondBorderHorizon);
      break;
    case (spreadSheet.getRangeByName('SecondDefaultTextColor').getA1Notation()):
      setTextColor('SecondDefaultTextColor', secondBorderAll);
      setTextColor('SecondDefaultTextColor', secondBorderTop);
      setTextColor('SecondMainTextColor', secondMainTheme);
      break;
    case (spreadSheet.getRangeByName('SecondMainTextColor').getA1Notation()):
      setTextColor('SecondMainTextColor', secondMainTheme);
      break;
    case (spreadSheet.getRangeByName('ThirdSub3ThemeColor').getA1Notation()):
      setCellColor('ThirdSub3ThemeColor', thirdSub3Theme);
      break;
    case (spreadSheet.getRangeByName('ThirdSub2ThemeColor').getA1Notation()):
      setCellColor('ThirdSub2ThemeColor', thirdSub2Theme);
      break;
    case (spreadSheet.getRangeByName('ThirdSub1ThemeColor').getA1Notation()):
      setCellColor('ThirdSub1ThemeColor', thirdSub1Theme);
      break;
    case (spreadSheet.getRangeByName('ThirdMainThemeColor').getA1Notation()):
      setCellColor('ThirdMainThemeColor', thirdMainTheme);
      setOutBorderColor('ThirdMainThemeColor', thirdBorderAll);
      setPartialBorderColor('ThirdMainThemeColor', thirdBorderTop, thirdBorderMiddle, thirdBorderBottom);
      setInBorderColor('ThirdMainThemeColor', thirdBorderAll, thirdBorderMerge, thirdBorderHorizon);
      break;
    case (spreadSheet.getRangeByName('ThirdDefaultTextColor').getA1Notation()):
      setTextColor('ThirdDefaultTextColor', thirdBorderAll);
      setTextColor('ThirdDefaultTextColor', thirdBorderTop);
      setTextColor('ThirdMainTextColor', thirdMainTheme);
      break;
    case (spreadSheet.getRangeByName('ThirdMainTextColor').getA1Notation()):
      setTextColor('ThirdMainTextColor', thirdMainTheme);
      break;
    case (spreadSheet.getRangeByName('FourthSub3ThemeColor').getA1Notation()):
      setCellColor('FourthSub3ThemeColor', fourthSub3Theme);
      break;
    case (spreadSheet.getRangeByName('FourthSub2ThemeColor').getA1Notation()):
      setCellColor('FourthSub2ThemeColor', fourthSub2Theme);
      break;
    case (spreadSheet.getRangeByName('FourthSub1ThemeColor').getA1Notation()):
      setCellColor('FourthSub1ThemeColor', fourthSub1Theme);
      break;
    case (spreadSheet.getRangeByName('FourthMainThemeColor').getA1Notation()):
      setCellColor('FourthMainThemeColor', fourthMainTheme);
      setOutBorderColor('FourthMainThemeColor', fourthBorderAll);
      setPartialBorderColor('FourthMainThemeColor', fourthBorderTop, fourthBorderMiddle, fourthBorderBottom);
      setInBorderColor('FourthMainThemeColor', fourthBorderAll, fourthBorderMerge, fourthBorderHorizon);
      break;
    case (spreadSheet.getRangeByName('FourthDefaultTextColor').getA1Notation()):
      setTextColor('FourthDefaultTextColor', fourthBorderAll);
      setTextColor('FourthDefaultTextColor', fourthBorderTop);
      setTextColor('FourthMainTextColor', fourthMainTheme);
      break;
    case (spreadSheet.getRangeByName('FourthMainTextColor').getA1Notation()):
      setTextColor('FourthMainTextColor', fourthMainTheme);
      break;
    default:
      break;
  }
}

function getRangeList(rangeName) {
  var rangeList = new Array();
  for(var i = 1; i <= parseInt(rangeName[1]); i++){
    var range = spreadSheet.getRangeByName(rangeName[0] + i).getA1Notation();
    rangeList.push(range);
  }
  return rangeList;
}

/**
 * Cell
 */

function setCellColor(rangeColor, rangeName) {
  var color = spreadSheet.getRangeByName(rangeColor).getBackground();
  var rangeList = getRangeList(rangeName);
  sheet.getRangeList(rangeList).setBackground(color);
}

/**
 * Text
 */

function setTextColor(rangeColor, rangeName) {
  var color = spreadSheet.getRangeByName(rangeColor).getBackground();
  var rangeList = getRangeList(rangeName);
  sheet.getRangeList(rangeList).setFontColor(color);
}

/**
 * Border
 */

function setInBorderColor(rangeColor, allRange, mergeRange, horizonRange) {
  /** 내부 테두리들 모두 설정 후, 필요없는 테두리 제거 및 변경 */
  var color = spreadSheet.getRangeByName(rangeColor).getBackground();
  var allRangeList = getRangeList(allRange);
  var mergeRangeList = getRangeList(mergeRange);
  var horizonRangeList = getRangeList(horizonRange);

  sheet.getRangeList(allRangeList).setBorder(null, null, null, null, true, null, color, SpreadsheetApp.BorderStyle.DASHED);
  sheet.getRangeList(allRangeList).setBorder(null, null, null, null, null, true, color, null);
  sheet.getRangeList(mergeRangeList).setBorder(null, null, null, null, null, false, color, null);

  sheet.getRangeList(horizonRangeList).setBorder(null, null, null, null, true, null, color, null);
  sheet.getRangeList(horizonRangeList).setBorder(null, null, null, null, null, true, color, SpreadsheetApp.BorderStyle.DASHED);
}

function setOutBorderColor(rangeColor, rangeName) {
  var color = spreadSheet.getRangeByName(rangeColor).getBackground();
  var rangeList = getRangeList(rangeName);
  sheet.getRangeList(rangeList).setBorder(true, true, true, true, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function setPartialBorderColor(rangeColor, topRange, middleRange, bottomRange) {
  var color = spreadSheet.getRangeByName(rangeColor).getBackground();
  var topRangeList = getRangeList(topRange);
  var middleRangeList = getRangeList(middleRange);
  var bottomRangeList = getRangeList(bottomRange);

  sheet.getRangeList(topRangeList).setBorder(true, true, null, true, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRangeList(middleRangeList).setBorder(null, true, null, true, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRangeList(bottomRangeList).setBorder(null, true, true, true, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}
