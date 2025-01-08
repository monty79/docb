

var oDocument = Api.GetDocument();

function escapeRegExp(str){
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function SearchAndReplaceInRange(oRange, searchText, replaceText, matchCase){
  oDocument.Search(searchText, matchCase).forEach(function(item) {
    let oRangeFind = oRange.IntersectWith(item);
    if (oRangeFind !== null) {
      let rawText = oRangeFind.GetText();
      if (rawText == searchText) {
        oRangeFind.Select();
        let oParagraphNew = Api.CreateParagraph();
        oParagraphNew.AddText(replaceText);
        oDocument.InsertContent([oParagraphNew], true, { "KeepTextOnly": true });
      }
    }
  });
}

function SearchAndReplace(searchText, replaceText, matchCase) {
  oDocument.SearchAndReplace({ "searchString": searchText, "replaceString": replaceText, "matchCase": matchCase });
}

function SearchAndReplaceInRow(oRow, searchText, replaceText, matchCase) {
  oRow.Search(searchText, matchCase).forEach(function(item) {
    item.Select();
    let oParagraphNew = Api.CreateParagraph();
    oParagraphNew.AddText(replaceText);
    oDocument.InsertContent([oParagraphNew], true, { "KeepTextOnly": true });
  });
}

function DuplicateRow(oRow,isBefore){
  oRow.AddRows(1, isBefore);
  let oRowNew = isBefore ? oRow.GetPrevious() : oRow.GetNext();
  for (let i = 0, n = oRow.GetCellsCount(); i < n; ++i) {
    let oSource = oRow.GetCell(i).GetContent();
    let oTarget = oRowNew.GetCell(i).GetContent();
    for (let j = 0, m = oSource.GetElementsCount(); j < m; ++j) {
      oTarget.Push(oSource.GetElement(j).Copy());
     }
     oTarget.RemoveElement(0);
  }
  return oRowNew;
}

function GetBookmarkRow(bookmarkName){
  return oDocument.GetBookmarkRange(bookmarkName).GetParagraph(0).GetParentTableCell().GetParentRow();
}

// Используется для учета неразрывных пробелов
let mapFindText = {};
// Используется для определения таблиц
let mapVarName = {};
{
  let sText = oDocument.GetRange().GetText();
  Argument.abapData.VALUES.forEach(function(lineValues){
    if ( lineValues.FIND_TEXT ) {
      if ( !mapFindText[lineValues.FIND_TEXT] ) {
        let regexp = new RegExp(escapeRegExp(lineValues.FIND_TEXT).replace(/ /g, "[  ]"), "ig");
        mapFindText[lineValues.FIND_TEXT] = new Set( sText.match(regexp) );
      }
    }
    if ( lineValues.VAR_NAME ) {
      let setVarNum = mapVarName[lineValues.VAR_NAME];
      if ( !setVarNum ) { setVarNum = new Set(); }
      setVarNum.add(lineValues.VAR_NUM);
      mapVarName[lineValues.VAR_NAME] = setVarNum;
    }
  });
}

{
  let lineValuesPrev = {};
  let oRange = null;
  let oRow = null;
  let sMode = null;
  Argument.abapData.VALUES.forEach(function(lineValues){
    if ( lineValuesPrev.VAR_NAME !== lineValues.VAR_NAME ) {
      if ( lineValues.VAR_NAME ) {
        if ( mapVarName[lineValues.VAR_NAME].size > 1 ) {
//        Таблица
          sMode = "T";
          oRange = null;
          try {
            oRow = DuplicateRow( GetBookmarkRow( lineValues.VAR_NAME ),true);
          } catch (err) {
            oRow = null;
          }
        } else {
//        Не таблица
          sMode = "N";
          try {
            oRange = oDocument.GetBookmarkRange(lineValues.VAR_NAME);
          } catch (err) {
            oRange = null;
          }
          oRow = null;
        }
      } else {
//      Весь документ
        sMode = "D";
        oRange = null;
        oRow = null;
      }
    }

    if ( lineValuesPrev.VAR_NAME !== lineValues.VAR_NAME ||
         lineValuesPrev.VAR_NUM  !== lineValues.VAR_NUM ) {

    }

    if ( lineValues.FIND_TEXT ) {
      mapFindText[ lineValues.FIND_TEXT ].forEach(function(findText){
        if ( sMode === "N" && oRange ) {
          SearchAndReplaceInRange(oRange, findText, lineValues.VALUE, true);
        } else if ( sMode === "T" && oRow ) {
          SearchAndReplaceInRow(oRow, findText, lineValues.VALUE, true);
        } else if (sMode === "D" ){
          SearchAndReplace(findText, lineValues.VALUE, true );
        }
      });
    } else {
      if ( lineValues.VAL_TYPE === "M" ) {
        let extMacro = extMacros[ lineValues.VALUE ];
        try{
          if ( extMacro ) { extMacro(); }
        } catch (err) {}
      } else if ( lineValues.VAL_TYPE === "S" || lineValues.VAL_TYPE === "" ) {
        if ( oRange ) {
          oRange.Select();
          let oParagraphNew = Api.CreateParagraph();
          oParagraphNew.AddText(lineValues.VALUE);
          oDocument.InsertContent([oParagraphNew], true, { "KeepTextOnly": true });
        }
      }
    }
    lineValuesPrev = lineValues;
  });
}

