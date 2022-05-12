
function extract(location, frontValue, endValue) {

    let curSheet = wps.EtApplication().ActiveSheet;
    if (curSheet) {
        curSheet.Cells.Item(1, 1).Formula = "Hello, wps加载项!" + curSheet.Cells.Item(1, 1).Formula
    }
    console.log(location, frontValue, endValue)
}

export default {
    extract
}