import Sheet from './Sheet';

export interface RangeOptions {
  a1?: string;
  row?: number;
  col?: number;
  numRows?: number;
  numColumns?: number;
}
export interface RangeComputed {
  row: number;
  col: number;
  numRows: number;
  numColumns: number;
}

export default class Range {
  __sheet: Sheet;
  criteria: RangeOptions;
  value: any | undefined;
  values: any[] = [];
  rangeValues: any[] = [];
  rangeComputed: RangeComputed;

  constructor(values: any[], criteria: RangeOptions, __sheet: Sheet) {
    this.__sheet = __sheet;
    this.values = values;
    this.criteria = criteria;

    if (criteria.a1) {
      this.rangeComputed = computeRangeFromA1Notation(this.values, criteria.a1);
    } else {
      this.rangeComputed = {
        row: criteria.row ? criteria.row - 1 : 0,
        col: criteria.col ? criteria.col - 1 : 0,
        numRows: criteria.numRows || 1,
        numColumns: criteria.numColumns || 1,
      };
    }
    this.rangeValues = getValuesWithCriteria(this.values, this.rangeComputed);
  }

  activate() {
    return this;
  }

  activateAsCurrentCell() {
    return this;
  }

  setValue(value: string) {
    this.value = value;

    // Update Sheet
    if (this.__sheet) {
      const rc = this.rangeComputed;
      this.__sheet.rows[rc.row + 1][rc.col] = value;
    }

    return this;
  }

  getValue() {
    if (this.__sheet) {
      const rc = this.rangeComputed;
      return this.__sheet.rows[rc.row + 1][rc.col];
    }

    return this.value;
  }

  setValues(values: string[][]) {
    const rc = { ...this.rangeComputed };
    // If range is 1x1, then we need to set the range to the size of the values
    if (rc.numRows === 1 && rc.numColumns === 1) {
      rc.numRows = values.length;
      rc.numColumns = values[0].length;
    }
    // For any other range size, the values must be the same size as the range
    if (values.length !== rc.numRows) {
      throw new Error(
        `The number of rows in the data does not match the number of rows in the range. The data has ${values.length} but the range has ${rc.numRows}.`
      );
    }
    if (values[0].length !== rc.numColumns) {
      throw new Error(
        `The number of columns in the data does not match the number of columns in the range. The data has ${values[0].length} but the range has ${rc.numColumns}.`
      );
    }
    // Update locally cached values
    for (let row = 0; row < rc.numRows; row++) {
      this.values[rc.row + row].splice(rc.col, rc.numColumns, ...values[row])
    }
    this.rangeValues = getValuesWithCriteria(this.values, this.rangeComputed);

    if (!this.__sheet) {
      return this;
    }

    // Update Sheet value
    for (let row = 0; row < rc.numRows; row++) {
      this.__sheet.rows[rc.row + row].splice(rc.col, rc.numColumns, ...values[row])
    }

    return this;
  }

  getValues() {
    return this.rangeValues;
  }

  getDisplayValues() {
    return this.getValues().map(row => row.map(toString));
  }

  // @TODO: All of these...
  setFontWeight(weight: string) {
    return this;
  }
  setFontWeights(weights: string[][]) {
    return this;
  }
  setNumberFormat(format: string) {
    return this;
  }
  setDataValidation(rule: any) {
    return this;
  }
  setBackground(color: string) {
    return this;
  }
  setBackgrounds(colors: string[][]) {
    return this;
  }
  setFontColor(color: string) {
    return this;
  }
  setFontColors(colors: string[][]) {
    return this;
  }
  setNote(note: string) {
    return this;
  }
  setNotes(notes: string[][]) {
    return this;
  }
}

/**
 * The "meat" of the Range slicing...
 */
function getValuesWithCriteria(values: any[][], c: RangeComputed): any[][] {
  const startRow = c.row;
  const endRow = c.row + c.numRows - 1;
  const startCol = c.col;
  const endCol = c.col + c.numColumns - 1;

  // Not enough rows in our data?
  if (values.length < endRow + 1) {
    values = new Array(endRow + 1).fill([]).map((emptyValue, index) => {
      return values[index] || emptyValue;
    });
  }

  return values.slice(startRow, endRow + 1).map(function (i: any[]) {
    if (i.length < endCol + 1) {
      const numCols = endCol + 1;
      // Not enough cols in our data?
      i = new Array(numCols <= 0 ? endCol : numCols).fill('').map((emptyValue, index) => {
        return i[index] || emptyValue;
      });
    }
    return i.slice(startCol, endCol + 1);
  });
}

function letterToColumn(letter: string) {
  let column = 0,
    length = letter.length;

  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }

  return column;
}

/**
 * @link https://stackoverflow.com/a/58545538
 */
function computeRangeFromA1Notation(values: any[][], textRange: string): RangeComputed {
  let startRow: number, startCol: number, endRow: number, endCol: number;
  let range = textRange.split(':');
  let ret = cellToRoWCol(range[0]);
  startRow = ret[0];
  startCol = ret[1];
  if (startRow == -1) {
    startRow = 0;
  }
  if (startCol == -1) {
    startCol = 0;
  }

  if (range[1]) {
    ret = cellToRoWCol(range[1]);
    endRow = ret[0];
    endCol = ret[1];
    if (endRow == -1) {
      endRow = values.length - 1;
    }
    if (endCol == -1) {
      endCol = values[0].length - 1;
    }
  } else {
    // only one cell
    endRow = startRow;
    endCol = startCol;
  }
  return {
    row: startRow,
    col: startCol,
    numRows: 1 + endRow - startRow,
    numColumns: 1 + endCol - startCol,
  }
}

function cellToRoWCol(cell: string): [number, number] {
  // returns row & col from A1 notation
  let row = Number(cell.replace(/[^0-9]+/g, ''));
  const letter = cell.replace(/[^a-zA-Z]+/g, '').toUpperCase();
  let column = letterToColumn(letter);

  row = row - 1;
  column--;

  return [row, column];
}

function toString(value: any) {
  if (value === true) {
    return 'TRUE';
  }
  if (value === false) {
    return 'FALSE';
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value);
}
