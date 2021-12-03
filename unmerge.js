function onInstall(e){
  onOpen(e)
}

function onOpen(e){

  var ui = SpreadsheetApp.getUi().createAddonMenu();
  
 ui.addItem('Unmerge', 'unmerge')
      .addToUi();
}



function unmerge() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  lastRow = spreadsheet.getLastRow();
  lastCol = spreadsheet.getLastColumn();
  bValue =  spreadsheet.getRange('B1').getValue()
  spreadsheet.getRange('B1').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A1:B1').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B1'));
  spreadsheet.getActiveRange().mergeAcross();
  
  var vrange = validateRange(spreadsheet,lastRow,lastCol)
  Logger.log(vrange)
  var range = spreadsheet.getRange(1,vrange[1],1,1)
  var lc = range.getA1Notation().match(/([A-Z]+)/)[0];
  strCon = "A1:" +lc + vrange[0];
  Logger.log(strCon)
  Logger.log(range.getA1Notation().match(/([A-Z]+)/)[0]); // Logs "A"
  spreadsheet.getRange(strCon).activate();
  //rangeList.activate();
  spreadsheet.getRange(1,1,vrange[0],vrange[1]).breakApart();
  spreadsheet.getRange('B1').setValue(bValue)
};

function validateRange(spreadsheet,lastRow,lastCol){
  var arr = new Array()
var mrg = spreadsheet.getRange(1, 1, 50, 50).getMergedRanges()
  for (var i = 0; i < mrg.length; i++) {
  //Logger.log(mrg[i].getA1Notation());
    arr.push(mrg[i].getA1Notation().split(":")[1])

}
//Logger.log(arr)
var arr2 = new Array()
var arr3 = new Array()
arr.forEach(function(a1){
  arr2.push(a1[0])
  //arr3.push(a1.match(/([A-Z]+)/)[0])
  arr3.push(a1.match(/([0-9]+)/)[0])

})
var arr22 = new Array()
arr2.forEach(function(a1){

  arr22.push(colLetToNum(a1))

})

//Logger.log(arr2)
//Logger.log(arr3)
//Logger.log(arr22)

var lrow = arr3.sort((a,b)=>a-b)[arr3.length - 1];
var lcol = arr22.sort((a,b)=>a-b)[arr22.length - 1];
var newArr = new Array()
var maxRow = Math.max(lastRow,lrow)
var maxCol = Math.max(lastCol,lcol)
newArr.push(maxRow)
newArr.push(maxCol)
return newArr;

}





function colLetToNum(letter){
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++){
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}
