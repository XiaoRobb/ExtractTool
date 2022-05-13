function extract(filePathList, position,frontValue, endValue) {
    //当前没有book
    //
    let targetBook = wps.EtApplication().ActiveWorkbook
    if (!targetBook) {
        return "当前没有打开任何文档"
    }
    //当前book新建表单，并激活新建的表单
    let targetSheet = wps.EtApplication().Worksheets.Add()
    let rowIndex = 1;
    filePathList.forEach((filePath) => {
        let sourceBook = wps.EtApplication().Workbooks.Open(filePath)
        let sourceSheet = sourceBook.Activesheet
        let startRow = 0;
        let endRow = 0;
        if(endValue == "末尾") {
            endRow = 1;
        }
        if(startRow != 0 && endRow != 0){
            targetSheet.Cells.Item(rowIndex, 1).Formula = sourceBook.Name
            targetSheet.Cells.Item(rowIndex, 2).Formula = sourceSheet.Range(position).Formula
        }
    })

    //提取



    //
    // let curSheet = wps.EtApplication().ActiveSheet;
    // curSheet
    // if (curSheet) {
    //     curSheet.Cells.Item(1, 1).Formula = "Hello, wps加载项!" + curSheet.Cells.Item(1, 1).Formula
    // }
    console.log(filePathList, frontValue, endValue)
    return "提取成功！"
}

function extractTemp(filePathList, position, startRowStr, endRowStr) {
    let startRow = parseInt(startRowStr) + 1
    let endRow = parseInt(endRowStr) - 1
    //当前没有book
    //
    let targetBook = wps.EtApplication().ActiveWorkbook
    if (!targetBook) {
        return "当前没有打开任何文档"
    }
    //当前book新建表单，并激活新建的表单
    let targetSheet = wps.EtApplication().Worksheets.Add()
    let rowIndex = 1;
    filePathList.forEach((filePath) => {
        let sourceBook = wps.EtApplication().Workbooks.Open(filePath)
        console.log(sourceBook)
        let sourceSheet = sourceBook.Worksheets.Item(1)
        console.log(sourceSheet)
        targetSheet.Cells.Item(rowIndex, 1).Formula = sourceBook.Name
        targetSheet.Cells.Item(rowIndex, 2).Formula = sourceSheet.Range(position).Formula
        let len = endRow - startRow
        let maxColumn = "E"
        console.log("A"+startRow, ",", endRow+maxColumn)
        sourceSheet.Range("A"+startRow,maxColumn+endRow).Copy(targetSheet.Range("C"+rowIndex))
        targetSheet.Range("A"+rowIndex, "A"+(rowIndex+len)).Merge()
        targetSheet.Range("B"+rowIndex, "B"+(rowIndex+len)).Merge()
        rowIndex += len+1
        sourceBook.Close()
    })

    //提取



    //
    // let curSheet = wps.EtApplication().ActiveSheet;
    // curSheet
    // if (curSheet) {
    //     curSheet.Cells.Item(1, 1).Formula = "Hello, wps加载项!" + curSheet.Cells.Item(1, 1).Formula
    // }
    return "提取成功！"
}

export default {
    extract,
    extractTemp
}