async function jsonifyRange(excel){
    // get currently selected range
    let selRange = excel.workbook.getSelectedRange();
    // load the values property
    selRange.load(['values']);
    await excel.sync();
    // transform the to json (note if i dont place selRange.values in a variable it doesn't work bug maybe ???)
    let rangeValues = selRange.values;
    // output to JADE console
    Jade.print(JSON.stringify(rangeValues), 'JSON VALUES');
  }