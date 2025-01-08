function() {
    const tableNumber = 2; //номер таблицы в документе (считая от 0)
    const firstRowNumber = 4; //номер первой строки с данными (считая от 0)
    const columnMergeNumbers = [0, 1, 2, 3, 9, 10, 11, 12];  //номера стоблцов для объединения (считая от 0)
    const columnCount = 13; //общее количество столбцов
	const compareColNum = 1; //номер столбца(с 0), при одинаковом значении в котором объединять ячейки

    let lvPrevVal = [];
    for (let h = 0; h<= columnCount - 1; ++h) {
        lvPrevVal[h] = "[NO DATA]";
    }

    let k = 0;
    let i = firstRowNumber;

    let Table1 = Api.GetDocument().GetAllTables()[tableNumber];

    let rnum = Table1.GetRowsCount();

    while(i < rnum - 1) {
        let j = i + 1;
        let ErrorNumber = 0;

        while (ErrorNumber === 0) {
            //если у соседних строк одинаковые табельники

            let par_i = Table1.GetCell(i, compareColNum).GetContent().GetElement(0);
            let pernr_i = par_i.GetText();

            let pernr_j = "[NO MORE ROWS]";

            if (j < rnum){


                let par_j = Table1.GetCell(j, compareColNum).GetContent().GetElement(0);
                pernr_j = par_j.GetText();

                if (pernr_i == pernr_j) {
                    for (let k = 0; k <= columnCount - 1; ++k){
                        if (columnMergeNumbers.indexOf(k) != -1){
                            if (lvPrevVal[k] == "[NO DATA]") {
                                lvPrevVal[k] = Table1.GetCell(i, k).GetContent().GetElement(0).GetText();
                            }

                            if (Table1.GetCell(j, k).GetContent().GetElement(0).GetText() == lvPrevVal[k]){
                                Table1.GetCell(j, k).GetContent().GetElement(0).AddText("[MERGE]");

                            }
                            else{
                                lvPrevVal[k] = Table1.GetCell(j, k).GetContent().GetElement(0).GetText();
                            }
                        }

                    }

                    j = j + 1;

                }
            }
            if(j == rnum || pernr_i != pernr_j){

                    for (k = 0; k<= columnCount - 1; ++k){

                        if (columnMergeNumbers.indexOf(k) != -1){

                        //if (k < 4 || k > 7){
                            for (let m = i + 1; m <= j - 1; ++m){

                                let cell1 = Table1.GetCell(m - 1, k);
                                let cell2 = Table1.GetCell(m ,k);

                                if (cell2.GetContent().GetElement(0).GetText().indexOf("[MERGE]", 0) > -1){

                                    cell2.GetContent().RemoveAllElements();

                                    //соединить cell1 и cell2

                                    let CellArray = [];
                                    CellArray[0] = cell1;
                                    CellArray[1] = cell2;

                                    Table1.MergeCells(CellArray);


                                    while (Table1.GetCell(m - 1, k).GetContent().GetElementsCount() > 1){
                                        Table1.GetCell(m - 1, k).GetContent().RemoveElement(1);
                                    }


                                }
                            }
                        }

                }

            k = 0;
            i = j;

            for (m = 0; m<= columnCount - 1; ++m){
                lvPrevVal[m] = "[NO DATA]";
            }

            ErrorNumber = 1;

            }


        }
    }
}
