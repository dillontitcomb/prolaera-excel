//assume inputs of 'A5'/'E23'/'F88' or [5,0]/[13,4]/[3,6]

exports.convertCellType = function(cellValue) {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  if (typeof cellValue === 'string') {
    let col;
    for (let i = 0; i < alphabet.length; i++) {
      console.log(cellValue[0]);

      if (alphabet[i] == cellValue[0]) {
        col = i;
        break;
      }
    }
    let row = parseInt(cellValue[1] + 1);
    return [row, col];
  } else {
    //[4][6] aka G5
    let row = cellValue[0] + 1;
    let col = alphabet[cellValue[1]];
    return `${col}${row}`;
  }
};
