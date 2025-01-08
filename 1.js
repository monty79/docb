

let GetRanges;

function __Get__GetRanges(aRanges) {
  let oRangeSheet = (aRanges.length > 0) && aRanges[0].GetWorksheet();
  let EmptyToNull = a => a.length ? a : null;
  return (name, oSheet = oRangeSheet ) =>
    !IsNameEmpty(name)
    ? EmptyToNull( GetDefNameRanges( name, oSheet ) )  ||
      name.split(",").reduce((r,n) => {
                        let oRange = CullRange( n.includes("!") || !oSheet
                                                ? Api.GetRange(n)
                                                : oSheet.GetRange(n) );
                        if (oRange) {
                          r.push(oRange);
                        }
                        return r;
                      }, [])
    : aRanges;
}

function ForEachRange(aNames, fCallBack, oSheet) {
  if (aNames && aNames.length > 0) {
    for (let i = 0; i < aNames.length; ++i) {
      GetRanges(aNames[i], oSheet).forEach( fCallBack, aNames[i] );
    }
  } else {
    GetRanges(undefined, oSheet).forEach( fCallBack, undefined );
  }
}

let GetSheets;

function __Get__GetSheets(aRanges) {
  return name => !IsNameEmpty(name)
                 ? [ ( name !== ":Active" ? Api.GetSheet(name) : Api.GetActiveSheet() ) ]
                 : ( aRanges.length ? aRanges.map(oRange => oRange.GetWorksheet()) : [ Api.GetActiveSheet() ] );
}

function ForEachSheet(aNames, fCallBack) {
  if (aNames && aNames.length > 0) {
    if ( aNames.find(n => n === ":All") ) {
      Api.GetSheets().forEach( fCallBack );
    } else {
      for (let i = 0; i < aNames.length; ++i) {
        GetSheets(aNames[i]).forEach( fCallBack );
      }
    }
  } else {
    GetSheets().forEach( fCallBack );
  }
}

function IsNameEmpty(name) {
  return !( name && name.trim() );
}

function CreateColor(color) {
  let c = +color;
  if (c<0) {
    c = 16777216 + c;
  }
  let r = c % 256;
  c = (c - r) / 256;
  let g = c % 256;
  c = (c - g) / 256;
  let b = c % 256;

  return Api.CreateColorFromRGB(r,g,b);
}

function GetUsedRangeFromRange(oRange) {
  return oRange && Intersect(oRange, oRange.GetWorksheet().GetUsedRange());
}

function CopySheet(sheet,after,name) {
  Api.asc_copyWorksheet(after.GetIndex() + 1, [name], [sheet.GetIndex()]);
  return Api.GetActiveSheet();
}

function GetLastSheet(){
  let s = Api.GetSheets();
  return s[ s.length - 1 ];
}

function escapeRegExp(str){
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function Intersect(r1,r2){
  let r = null;
//  try {
//    r = Api.Intersect(r1, r2);
//    if ( r == {} ) { r = null }
//  } catch (err) { r = null; }
  let w = r1.GetWorksheet();
  if ( w.GetIndex() === r2.GetWorksheet().GetIndex() ) {
    let r1_x1 = r1.GetCol();
    let r1_y1 = r1.GetRow();
    let r1_x2 = r1_x1 + r1.GetCols().GetCount() - 1;
    let r1_y2 = r1_y1 + r1.GetRows().GetCount() - 1;
    let r2_x1 = r2.GetCol();
    let r2_y1 = r2.GetRow();
    let r2_x2 = r2_x1 + r2.GetCols().GetCount() - 1;
    let r2_y2 = r2_y1 + r2.GetRows().GetCount() - 1;
    let r_x1 = r1_x1 > r2_x1 ? r1_x1 : r2_x1;
    let r_y1 = r1_y1 > r2_y1 ? r1_y1 : r2_y1;
    let r_x2 = r1_x2 < r2_x2 ? r1_x2 : r2_x2;
    let r_y2 = r1_y2 < r2_y2 ? r1_y2 : r2_y2;
    if ( r_x1 <= r_x2 && r_y1 <= r_y2 ) {
      r = w.GetRange( w.GetCells(r_y1, r_x1), w.GetCells(r_y2, r_x2) );
    }
  }
  return r;
}

function RangeSetIndent(oRange, indent) {
  let s = Api.GetSelection();
  oRange.Select();
  Api.asc_setCellIndent(indent);
  s.Select();
}

function RangeClearValue(r) {
  if (r) {
    const constEmptyValue = {toString:(()=>"")};
    r.SetValue(constEmptyValue);
  }
}

function RangeSetValue(r,v) {
  if (r) {
    if (!v || v === "'") {
      RangeClearValue(r);
    } else if (v[0] !== "'") {
      let f = r.GetNumberFormat();
      r.SetValue(v);
      r.SetNumberFormat(f);
    } else {
      GetUsedRangeFromRange(r).ForEach( function (c) {
        if ("@" === c.GetNumberFormat() ) {
          c.SetValue(v.slice(1));
        } else {
          c.SetValue(v);
        }
      });
    }
  }
}

function ReplaceInRange(range,textOld,textNew,MatchCase=false) {
  let r = GetUsedRangeFromRange(range);
  if (r) {
    let reg = new RegExp( escapeRegExp(textOld), "g" + (!MatchCase ? "i" : "") );
    r.ForEach( function(cell){
      let oldValue = cell.GetText();
      let newValue = oldValue.replace(reg,textNew);
      if ( oldValue !== newValue ) {
        RangeSetValue(cell, newValue);
      }
    });
  }
}

function ForEachCellInUsedRange(oRange, fCallBack) {
  oRange = GetUsedRangeFromRange( oRange );
  let result;
  if (oRange) {
    let oSheet = oRange.GetWorksheet();
    let col1 = oRange.GetCol();
    let row1 = oRange.GetRow();
    let col2 = col1 + oRange.GetCols().GetCount();
    let row2 = row1 + oRange.GetRows().GetCount();
    for (let row = row1; row < row2 && !result; ++row) {
      for (let col = col1; col < col2 && !result; ++col) {
        let oCell = oSheet.GetCells(row,col);
        result = fCallBack(oCell);
      }
    }
  }
  return result;
}

function FindFirstCellInRange(range,text,MatchCase=false) {
  let reg = new RegExp( escapeRegExp(text), (!MatchCase ? "i" : "") );
  return ForEachCellInUsedRange(range, oCell => oCell.GetText().match(reg) && oCell )
      || null;
}

function GetSheetIgnoringCase(name) {
  name = name.toUpperCase().replace(/^'(.*)'$/,"$1");
  return Api.GetSheets().find(function(s){
           return s.GetName().toUpperCase() === name;
         });
}

function SplitDefName(name) {
  let aName = name.match(/(^[^!]*)!(.*)$/);
  return aName ? { sheet: aName[1], range: aName[2] }
               : { sheet: null,     range: name };
}

function GetDefNameRanges(name, sheet) {
  let dn = SplitDefName(name);
  return GetSheetDefNameRanges((dn.sheet ? GetSheetIgnoringCase(dn.sheet) : sheet || Api.GetActiveSheet()), dn.range );
}

function GetSheetDefNameRanges(sheet, name) {
  let sAddressAll = sheet.GetDefName(name).GetRefersTo();
  return sAddressAll ? sAddressAll.split(',').reduce(function(aRanges,sAddress){
                         let oRange = CullRange( Api.GetRange(sAddress) );
                         if (oRange) {
                           aRanges.push(oRange);
                         }
                         return aRanges;
                       }, [])
                     : [];
}

function CullRange(oRange) {
  try {
    oRange.GetRow();
    return oRange;
  } catch (err) {
    return null;
  }
}

function SheetGetRange(oSheet, nRow, nCol, nRowsCount = 1, nColsCount = 1) {
  nRow = +nRow;
  nCol = +nCol;
  nRowsCount = +nRowsCount;
  nColsCount = +nColsCount;
  return oSheet.GetRange( oSheet.GetCells(nRow, nCol), oSheet.GetCells(nRow + nRowsCount - 1, nCol + nColsCount - 1) );
}

function CloneRangeRows(oRange,countClones) {
  let rows = oRange.GetRows().GetCount();
  let oRangeClone = oRange.GetCells();
  for (let i = 0; i < countClones; ++i) {
    oRangeClone.SetOffset(rows,0);
    oRangeClone.Insert("down");
    oRange.Copy(oRangeClone);
//    RangeCopy(oRange,oRangeClone);
  }
}

// Не используется так как сдвигает лишнее, но как запасной вариант.
function CloneRangeCols_v(oRange,countClones) {
  let cols = oRange.GetCols().GetCount();
  let oRangeFrom = oRange.GetCells();
  for (let i = 0; i < countClones; ++i) {
    let oRangeTo = oRangeFrom;
    oRangeFrom.Insert("right");
    oRangeFrom.SetOffset(0,cols);
    oRangeFrom.Copy(oRangeTo);
  }
}

function CloneRangeCols(oRange,countClones) {
  let cols = oRange.GetCols().GetCount();
  let oRangeClone = oRange.GetCells();
  for (let i = 0; i < countClones; ++i) {
    oRangeClone.SetOffset(0,cols);
    oRangeClone.Insert("right");
//    oRange.Copy(oRangeClone);
    RangeCopy(oRange,oRangeClone);
  }
}

function RangeCopy(oRangeFrom, oRangeTo){
//  oRangeFrom = GetUsedRangeFromRange(oRangeFrom);
  let r1 = oRangeTo.GetRow() - 1; // смена точки отсчета на 0
  let c1 = oRangeTo.GetCol() - 1; // смена точки отсчета на 0
  let adrs =  oRangeFrom.GetAddress();
  let r2 = adrs.match(/^[A-Z]+:[A-Z]+$/) ? NaN : r1 - 1 + oRangeFrom.GetRows().GetCount();
  let c2 = adrs.match(/^\d+:\d+$/)       ? NaN : c1 - 1 + oRangeFrom.GetCols().GetCount();
  oRangeFrom.qc.move(oRangeTo.qc.Ma.He(r1, c1, r2, c2).kb, true, oRangeTo.qc.Ma);
}

function RangeRowsResize(range, rowsNew) {
  if ( range.GetRows().GetCount() === rowsNew ) {
    return range.GetCells();
  }
  let adrsOld = range.GetAddress();
  let aSubstr;
  rowsNew = +rowsNew;
  aSubstr = adrsOld.match(/^([A-Z]*)(\d+)\:([A-Z]*)(\d+)$/);
  if (aSubstr) {
    return range.GetWorksheet().GetRange( aSubstr[1] + aSubstr[2] + ":" + aSubstr[3] + (+aSubstr[2] - 1 + rowsNew) );
  }
  aSubstr = adrsOld.match(/^([A-Z]*)(\d+)$/);
  if (aSubstr) {
    return range.GetWorksheet().GetRange( aSubstr[1] + aSubstr[2] + ":" + aSubstr[1] + ( +aSubstr[2] - 1 + rowsNew) );
  }
  aSubstr = adrsOld.match(/^([A-Z]+):([A-Z]+)$/);
  if (aSubstr) {
    return range.GetWorksheet().GetRange( aSubstr[1] + "1" + ":" + aSubstr[2] +  rowsNew );
  }
}

function RangeBringToRows(oRange, nLines) {
  let nRows = oRange.GetRows().GetCount();
  let oRangeTmp = oRange.GetCells();
  if ( nLines > nRows ) {
    oRangeTmp.SetOffset(nRows,0);
    RangeRowsResize(oRangeTmp, nLines - nRows).Insert("down");
  } else if (nLines < nRows) {
    oRangeTmp.SetOffset(nLines,0);
    RangeRowsResize(oRangeTmp, nRows - nLines).Delete("up");
  }
  return RangeRowsResize(oRange, nLines);
}

function CopyRangeToRange(fromRange,toRange,copyHeight = false) {
  let r = RangeBringToRows(toRange, fromRange.GetRows().GetCount());
  fromRange.Copy(toRange);
  if (copyHeight) {
    let fromSheet = fromRange.GetWorksheet();
    let toSheet = toRange.GetWorksheet();
    let fromRow = fromRange.GetRow();
    let toRow = toRange.GetRow();
    for (i = r.GetRows().GetCount() - 1 ; i >= 0; --i) {
      toSheet.GetRows(toRow + i).SetRowHeight( fromSheet.GetRows(fromRow + i).GetRowHeight() );
    }
  }
  return r;
}

function InsertStringAsTableIntoRangeFromCell(oIntoRange,oFromCell,sTable) {
  let n = 0;
  if (oFromCell) {
    let oLine = oFromCell.GetCells(1,1);
    let aLines = sTable.split("\r\n");
    let nLastColl = oIntoRange.GetCol() + oIntoRange.GetCols().GetCount() - 1;
    if ( !aLines[aLines.length - 1 ] ) {
      aLines.pop();
    }
    n = aLines.length;
//    RangeBringToRows(oIntoRange, 1);
    if (aLines.length > 1) {
      CloneRangeRows(RangeRowsResize(oIntoRange, 1), aLines.length - 1 );
    }
    if (aLines.length > 0) {
      aLines.forEach(function(sLine){
        let oCell = oLine.GetCells(1,1);
        sLine.split("\t").forEach(function(sValue){
          if (oCell.GetCol() <= nLastColl ) {
            RangeSetValue(oCell, sValue);
            oCell.SetOffset(0, 1);
          }
        });
        oLine.SetOffset(1, 0);
      });
    } else {
      RangeClearValue(oLine);
    }
  }
  return n;
}

function AddSheetAndIndex(arr, oSheet, rangeIndex = null) {
  let i = oSheet.GetIndex();
  let j = arr.findIndex(s => i === s.oSheet.GetIndex());
  if (j < 0) {
    if ( rangeIndex == null ) {
      arr.push( { oSheet: oSheet, aRangeIndex: null } );
    } else {
      arr.push( { oSheet: oSheet, aRangeIndex: [rangeIndex] } );
    }
  } else if ( rangeIndex == null ) {
    arr[j].aRangeIndex = null;
  } else {
    let aRI = arr[j].aRangeIndex;
    if ( aRI && !aRI.find(ri => ri === rangeIndex) ) {
      aRI.push(rangeIndex);
    }
  }
}

{ //Все объявления сделаны, теперь начинается работа. На всякий случай производится изоляция от макросов.
  let listVarNames = {};
  {
    Argument.abapData.VALUES.forEach(function(lineValues){
      if ( lineValues.VAR_NAME ) {
        let lineVarName = listVarNames[lineValues.VAR_NAME];
        if ( lineVarName ) {
          if ( lineVarName.VarNumLast !== lineValues.VAR_NUM ) {
            lineVarName.VarNumLast = lineValues.VAR_NUM;
            ++lineVarName.VarNumCount;
          }
        } else {
          listVarNames[lineValues.VAR_NAME] = { VarNumFirst: lineValues.VAR_NUM,
                                                VarNumLast:  lineValues.VAR_NUM,
                                                VarNumCount: 1 };
        }
      }
    });
  }

  let delVarNames = {};
  {
    let lineValuesPrev = {};
    let oRangeBase;
    let oRange;
    let aRangesBase = [];
    let aRanges = [];
    Argument.abapData.VALUES.forEach(function(lineValues){
      if ( lineValues.VAL_TYPE === "D" ) {
        if (lineValues.VAR_NAME) {
          let dn = SplitDefName(lineValues.VAR_NAME);
          if (!delVarNames[dn.range]) {
            delVarNames[dn.range] = [];
          }
          AddSheetAndIndex( delVarNames[dn.range], dn.sheet ? GetSheetIgnoringCase(dn.sheet) : Api.GetActiveSheet() );
        }
      } else if ( lineValuesPrev.VAR_NAME !== lineValues.VAR_NAME ) {
        if ( lineValues.VAR_NAME ) {
          aRangesBase = GetDefNameRanges(lineValues.VAR_NAME);
          aRanges = Array.from( aRangesBase );
          let lineVarName = listVarNames[lineValues.VAR_NAME];
          if ( lineVarName.VarNumCount > 1 ) {
            aRanges.forEach(function(oRange){
              CloneRangeRows(oRange,lineVarName.VarNumCount - 1);
            });
          }
        } else {
          aRangesBase = [ Api.GetActiveSheet().GetCells() ];
          aRanges = Array.from( aRangesBase );
        }
      } else if ( lineValuesPrev.VAR_NUM  !== lineValues.VAR_NUM ) {
        aRangesBase.forEach( (oRangeBase, i) => {
          let oRange = aRanges[i];
          oRangeBase.SetOffset(oRange.GetRow() - oRangeBase.GetRow() + oRange.GetRows().GetCount(),0);
        });
        aRanges = Array.from( aRangesBase );
      }
      if ( lineValues.FIND_TEXT ) {
        aRanges.forEach(function(oRange, rangeIndex) {
          switch (lineValues.VAL_TYPE) {
            case "S":
  //            oRange.Replace(lineValues.FIND_TEXT,lineValues.VALUE,"xlPart","xlByRows","xlNext",false,true);
              ReplaceInRange(oRange,lineValues.FIND_TEXT,lineValues.VALUE);
              break;
            case "R":
//              RangeSetValue(FindFirstCellInRange(oRange, lineValues.FIND_TEXT), lineValues.VALUE);
              InsertStringAsTableIntoRangeFromCell(oRange, FindFirstCellInRange(oRange, lineValues.FIND_TEXT), lineValues.VALUE.split("\r\n")[0] );
              break;
            case "T":
              let n = InsertStringAsTableIntoRangeFromCell(oRange, FindFirstCellInRange(oRange, lineValues.FIND_TEXT), lineValues.VALUE);
              if (n > 1) {
//                oRange.SetOffset(n-1, 0);
                aRanges[rangeIndex] = RangeRowsResize(oRange, oRange.GetRows().GetCount() - 1 + n );
              }
              break;
          }
        });
      } else {
        switch (lineValues.VAL_TYPE) {
          case "":
          case "S":
          case "R":
          case "T":
            aRanges.forEach(function(oRange) {
              InsertStringAsTableIntoRangeFromCell(oRange, oRange.GetCells(1,1), lineValues.VALUE.split("\r\n")[0] );
            });
            break;
          case "V":
            let dn = SplitDefName(lineValues.VALUE);
            if (!delVarNames[dn.range]) {
              delVarNames[dn.range] = [];
            }
            let oSheetDN = dn.sheet ? GetSheetIgnoringCase(dn.sheet) : null;
            aRanges.forEach(function(oRange,rangeIndex) {
              let oSheet = oSheetDN || oRange.GetWorksheet();
              let oRangeV = GetSheetDefNameRanges(oSheet, lineValues.VALUE)[rangeIndex];
              if ( oRangeV ) {
                AddSheetAndIndex( delVarNames[dn.range], oSheet, rangeIndex );
                aRanges[rangeIndex] = CopyRangeToRange(oRangeV, oRange);
              }
            });
            break;
          case "M":
            if (lineValues.VALUE) {
              let aValues = lineValues.VALUE.split("\t");
              let extMacro = extMacros[ aValues[0] ];
              try{
                if ( extMacro ) {
                  aValues.shift();
                  for(let i = aValues.length - 1; i >= 0; --i) {
                    aValues[i] = '"' + aValues[i] + '"';
                    try {
                      aValues[i] = JSON.parse(aValues[i]);
                    } catch (err) {}
                  }
                  GetRanges = __Get__GetRanges(aRanges);
                  GetSheets = __Get__GetSheets(aRanges);
                  extMacro.apply(aRanges, aValues);
                }
              } catch (err) {}
            }
            break;
        }
      }
      lineValuesPrev = lineValues;
    });
  }

  Object.keys(delVarNames).forEach(function(name){
    delVarNames[name].forEach(function(item) {
      let aRanges = GetSheetDefNameRanges(item.oSheet, name);
      if (item.aRangeIndex) {
        item.aRangeIndex.forEach( (rangeIndex) => {
          let oRange = aRanges[rangeIndex];
          if (oRange) {
            oRange.Delete("up");
          }
        });
      } else {
        aRanges.forEach( (oRange) => {
          oRange.Delete("up");
        });
      }
    });
  });
}
