function (name, FormTop){
    let aSheetsPageBreaks = [];

    ForEachRange([name], function(oRange) {
        ForEachCellInUsedRange(oRange, function(oCell) {
            try {
                let shtNameCurrent = oCell.GetValue();
                if (shtNameCurrent) {
                    let oSheetControl = oCell.GetWorksheet();
                    let row = oCell.GetRow();
                    let mainSheetName = oSheetControl.GetCells(row, 3).GetValue();
                    if (mainSheetName) {
                        let oCurrentSheet = Api.GetSheet(shtNameCurrent);       //Лист, с которого копируем данные (текущий лист)
                        let oMainSheet = Api.GetSheet(mainSheetName);           //Лист, в который копируем данные (основной лист)
                        let oCurrentForm = GetRanges("FORM", oCurrentSheet)[0]; //Весь формуляр текущего листа
                        let oRangeMainForm = GetRanges("FORM", oMainSheet)[0];  //Весь формуляр основного листа
                        let mainLastRow = oRangeMainForm.GetRows().GetCount();  //Последняя строка формуляра основного листа
                        if (FormTop) {
                            let seqnr = +oSheetControl.GetCells(row, 5).Value;     //Порядковый номер ставки для основного листа
                            if (seqnr === 1) {
                                //Удалим шапку текущего листа (если добавляем лист с первой налоговой ставкой)
                                GetRanges("FORM_TOP", oCurrentSheet).forEach( r => { r.Delete("up"); });
                            } else if (seqnr > 1) {
                                // Добавим разрыв страницы в конце формуляра основного листа (точнее запоминаем, где надо проставить, а проставлять будем позже),
                                // если добавляем лист со второй или последующими налоговыми ставками.
                                let mainPenultRow = mainLastRow - 1;            //Предпоследняя строка формуляра основного листа
                                let sheetIndex = oMainSheet.GetIndex();
                                let sheetPageBreaks = aSheetsPageBreaks.find( line => line.oSheet.GetIndex() === sheetIndex );
                                if (sheetPageBreaks) {
                                    sheetPageBreaks.aPageBreaks.push(mainPenultRow);
                                } else {
                                    aSheetsPageBreaks.push( {oSheet: oMainSheet, aPageBreaks: [mainPenultRow]} );
                                }
                            }
                        }
                        //Копируем данные с текущего листа в конец основного
                        let oMainLastRow = oMainSheet.GetRows(mainLastRow);
                        oMainLastRow.Insert("down");
                        oMainLastRow.Insert("down");
                        CopyRangeToRange(oCurrentForm, oMainLastRow, true);
                        //Удалим текущий лист
                        oCurrentSheet.Delete();
                    }
                }
            } catch (err) {}
        });
    });


    //Добавим разрывы страниц в нужных местах
    aSheetsPageBreaks.forEach(line => {
        let oDefNamePrintArea = line.oSheet.GetDefNames().find(dn => dn.GetName() === "Print_Area");
        let oRangePrintArea = oDefNamePrintArea.GetRefersToRange();
        let row = oRangePrintArea.GetRow();
        let col = oRangePrintArea.GetCol();
        let rowEnd = row + oRangePrintArea.GetRows().GetCount();
        let cols = oRangePrintArea.GetCols().GetCount();
        let sheetName = line.oSheet.GetName();
        if (line.aPageBreaks[line.aPageBreaks.length - 1] < rowEnd) {
            line.aPageBreaks.push(rowEnd);
        }
        let sAddr = line.aPageBreaks.reduce((str, rowPageBreak) => {
          let res = (str ? str + "," : "") + "'" + sheetName + "'!" + SheetGetRange(row, col, rowPageBreak - 1, cols).GetAddress(true,true,"xlA1", false);
          row = rowPageBreak;
          return res;
        }, "");
        oDefNamePrintArea.SetRefersTo(sAddr);
    });

    //Финальное переименование листов
    ForEachRange([name], function(oRange) {
        ForEachCellInUsedRange(oRange, function(oCell) {
            try {
                let shtNameCurrent = oCell.GetValue();
                if (shtNameCurrent) {
                    let oSheetControl = oCell.GetWorksheet();
                    let row = oCell.GetRow();
                    let finalSheetName = oSheetControl.GetCells(row, 4).GetValue();
                    if (finalSheetName) {
                        Api.GetSheet(shtNameCurrent).SetName( finalSheetName );
                    }
                }
            } catch (err) {}
        });
    });
}
