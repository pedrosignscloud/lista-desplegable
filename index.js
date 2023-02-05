function seleccionMultiple(e){
    if (!e) {
        return;
    }
    let activeCell = SpreadsheetApp.getActiveRange();
    let column = activeCell.getColumn();
    let row = activeCell.getRow();
    let value = e.value;
    let previousValue = e.oldValue;
    let resultCell = activeCell.offset(0,1);
    let accumulatedValue = resultCell.getValue();

    let duplicateIndex = accumulatedValue.toString().indexOf(value);

    if (column === 2 && row > 2) {
        if (!value) {
            accumulatedValue = "";
        } else if (!previousValue) {
            accumulatedValue = value;
        } else if (duplicateIndex !== -1) {
            accumulatedValue = accumulatedValue;
        } else if (value !== previousValue) {
            accumulatedValue = accumulatedValue + "," + value;
        }
        resultCell.setValue(accumulatedValue);
    }
}

function onEdit(e){
    seleccionMultiple(e);
}
