async function jsonifyRange(excel){
    let currentlySelectedRange = excel.workbook.getSelectedRange();
    currentlySelectedRange.load('values');
    await excel.sync();
    
}