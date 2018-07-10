//convert excel cell to array of array value or vice versa

exports.convertCellType = function(cellValue) {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  if (typeof cellValue === 'string') {
    let col;
    for (let i = 0; i < alphabet.length; i++) {
      if (alphabet[i] == cellValue[0]) {
        col = i;
        break;
      }
    }
    let row = parseInt(cellValue[1] + 1);
    return [row, col];
  } else {
    let row = cellValue[0] + 1;
    let col = alphabet[cellValue[1]];
    return `${col}${row}`;
  }
};
