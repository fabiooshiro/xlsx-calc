'use strict';
const utils = require('./utils');
const error = require('./errors');

// +---------------------+
// | FORMULAS REGISTERED |
// +---------------------+
let formulas = {
  FLOOR: Math.floor,
  '_xlfn.FLOOR.MATH': Math.floor,
  'FLOOR.MATH': Math.floor,
  ABS: Math.abs,
  SQRT: Math.sqrt,
  VLOOKUP: vlookup,
  MAX: max,
  SUM: sum,
  MIN: min,
  CONCATENATE: concatenate,
  PMT: pmt,
  COUNTA: counta,
  IRR: irr,
  'NORM.INV': normsInv,
  '_xlfn.NORM.INV': normsInv,
  STDEV: stDeviation,
  AVERAGE: avg,
  EXP: EXP,
  LN: Math.log,
  '_xlfn.VAR.P': var_p,
  'VAR.P': var_p,
  '_xlfn.COVARIANCE.P': covariance_p,
  'COVARIANCE.P': covariance_p,
  TRIM: trim,
  LEN: len,
  ISBLANK: is_blank,
  HLOOKUP: hlookup,
  INDEX: index,
  MATCH: match,
  SUMPRODUCT: sumproduct,
  ISNUMBER: isnumber,
  TODAY: today,
  ISERROR: iserror,
  TIME: time,
  DAY: day,
  MONTH: month,
  YEAR: year,
  RIGHT: right,
  LEFT: left,
  IFS: ifs,
  ROUND: round,
  CORREL: correl, // missing test
  SUMIF: sumif, // missing test,
  CHOOSE: choose,
  SUBSTITUTE: substitute,
  CEILING: ceiling,
  FILTER: throwErrors(FILTER),
  DATEDIF: datediff,
  EOMONTH: eomonth,
};

function choose(option) {
  return arguments[option];
}

function sumif() {
  let elementToSum = arguments[1];
  let sumResult = 0;

  [].slice.call(arguments)[0][0].forEach((elt, key) => {
    if (elt !== null) {
      //if the element is not a string but a number, number has no replace function, so converting to string.
      elt = elt.toString();
      if (elt.replace(/\'/g, '') === elementToSum) {
        if (!isNaN([].slice.call(arguments)[2][0][key])) {
          sumResult += [].slice.call(arguments)[2][0][key];
        }
      }
    }
  });
  return sumResult;
}

function correl(a, b) {
  a = getArrayOfNumbers(a);
  b = getArrayOfNumbers(b);

  if (a.length !== b.length) {
    return 'N/D';
  }
  var inv_n = 1.0 / (a.length - 1);
  var avg_a = sum.apply(this, a) / a.length;
  var avg_b = sum.apply(this, b) / b.length;
  var s = 0.0;
  var sa = 0;
  var sb = 0;
  for (var i = 0; i < a.length; i++) {
    s += (a[i] - avg_a) * (b[i] - avg_b);

    sa += Math.pow(a[i], 2);
    sb += Math.pow(b[i], 2);
  }

  sa = Math.sqrt(sa / inv_n);
  sb = Math.sqrt(sb / inv_n);

  return s / (inv_n * sa * sb);
}

function round(value, decimalPlaces) {
  if (arguments.length === 0) throw new Error('Err:511');
  if (arguments.length === 1) return Math.round(value);
  let roundMeasure = Math.pow(10, decimalPlaces);
  return Math.round(roundMeasure * value) / roundMeasure;
}

function today() {
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  return today;
}

function isnumber(x) {
  return !isNaN(x);
}

function sumproduct() {
  var parseNumber = function (string) {
      if (string === undefined || string === '' || string === null) {
        return 0;
      }
      if (!isNaN(string)) {
        return parseFloat(string);
      }
      return 0;
    },
    consistentSizeRanges = function (matrixArray) {
      var getRowCount = function (matrix = []) {
          return matrix.length;
        },
        getColCount = function (matrix = []) {
          return matrix[0] && matrix[0].length;
        },
        rowCount = getRowCount(matrixArray[0]),
        colCount = getColCount(matrixArray[0]);

      for (var i = 1; i < matrixArray.length; i++) {
        if (getRowCount(matrixArray[i]) !== rowCount || getColCount(matrixArray[i]) !== colCount) {
          return false;
        }
      }
      return true;
    };

  if (!arguments || arguments.length === 0 || !arguments[0]) {
    throw Error('#VALUE!');
  }
  if (!consistentSizeRanges(arguments)) {
    throw Error('#VALUE!');
  }
  // throw error if any of the cells passed in arguments is in error
  for (var i = 0; i < arguments.length; i++) {
    var row = arguments[i];
    if (Array.isArray(row)) {
      for (var j = 0; j < row.length; j++) {
        var col = row[j];
        if (Array.isArray(col)) {
          for (var k = 0; k < col.length; k++) {
            var cell = col[k];
            if (cell instanceof Error) {
              throw cell;
            }
          }
        } else {
          var cell = col;
          if (cell instanceof Error) {
            throw cell;
          }
        }
      }
    } else {
      var cell = row;
      if (cell instanceof Error) {
        throw cell;
      }
    }
  }

  var arrays = arguments.length + 1;
  var result = 0;
  var product;
  var k;
  var _i;
  var _ij;
  for (var i = 0; i < arguments[0].length; i++) {
    if (!(arguments[0][i] instanceof Array)) {
      product = 1;
      for (k = 1; k < arrays; k++) {
        _i = parseNumber(arguments[k - 1][i]);

        product *= _i;
      }
      result += product;
    } else {
      for (var j = 0; j < arguments[0][i].length; j++) {
        product = 1;
        for (k = 1; k < arrays; k++) {
          _ij = parseNumber(arguments[k - 1][i][j]);

          product *= _ij;
        }
        result += product;
      }
    }
  }
  return result;
}

function match_less_than_or_equal(matrix, lookupValue) {
  var index;
  var indexValue;
  for (var idx = 0; idx < matrix.length; idx++) {
    if (matrix[idx] === lookupValue) {
      return idx + 1;
    } else if (matrix[idx] < lookupValue) {
      if (!indexValue) {
        index = idx + 1;
        indexValue = matrix[idx];
      } else if (matrix[idx] > indexValue) {
        index = idx + 1;
        indexValue = matrix[idx];
      }
    }
  }
  if (!index) {
    throw Error('#N/A');
  }
  return index;
}

function match_exactly_string(matrix, lookupValue) {
  for (var idx = 0; idx < matrix.length; idx++) {
    lookupValue = lookupValue.replace(/\?/g, '.');
    if (Array.isArray(matrix[idx])) {
      if (matrix[idx].length === 1 && typeof matrix[idx][0] === 'string') {
        if (matrix[idx][0].toLowerCase() === lookupValue.toLowerCase()) {
          return idx + 1;
        }
      }
    } else if (typeof matrix[idx] === 'string') {
      if (matrix[idx].toLowerCase() === lookupValue.toLowerCase()) {
        return idx + 1;
      }
    }
  }
  throw Error('#N/A');
}

function match_exactly_non_string(matrix, lookupValue) {
  for (var idx = 0; idx < matrix.length; idx++) {
    if (Array.isArray(matrix[idx])) {
      if (matrix[idx].length === 1) {
        if (matrix[idx][0] === lookupValue) {
          return idx + 1;
        }
      }
    } else if (matrix[idx] === lookupValue) {
      return idx + 1;
    }
  }
  throw Error('#N/A');
}

// +---------------------+
// | THE IMPLEMENTATIONS |
// +---------------------+

function match_greater_than_or_equal(matrix, lookupValue) {
  var index;
  var indexValue;
  for (var idx = 0; idx < matrix.length; idx++) {
    if (Array.isArray(matrix[idx]) && matrix[idx].length > 0) {
      // For array in column
      if (matrix[idx][0] === lookupValue) {
        return idx + 1;
      } else if (matrix[idx][0] > lookupValue) {
        if (!indexValue) {
          index = idx + 1;
          indexValue = matrix[idx][0];
        } else if (matrix[idx][0] < indexValue) {
          index = idx + 1;
          indexValue = matrix[idx][0];
        }
      }
    } else {
      // For array in row
      if (matrix[idx] === lookupValue) {
        return idx + 1;
      } else if (matrix[idx] > lookupValue) {
        if (!indexValue) {
          index = idx + 1;
          indexValue = matrix[idx];
        } else if (matrix[idx] < indexValue) {
          index = idx + 1;
          indexValue = matrix[idx];
        }
      }
    }
  }
  if (!index) {
    throw Error('#N/A');
  }
  return index;
}

function match(lookupValue, matrix, matchType) {
  if (Array.isArray(matrix) && matrix.length === 1 && Array.isArray(matrix[0])) {
    matrix = matrix[0];
  }
  if (!lookupValue && !matrix) {
    throw Error('#N/A');
  }
  if (arguments.length === 2) {
    matchType = 1;
  }
  if (!(matrix instanceof Array)) {
    throw Error('#N/A');
  }
  if (matchType === 0) {
    if (typeof lookupValue === 'string') {
      return match_exactly_string(matrix, lookupValue);
    } else {
      return match_exactly_non_string(matrix, lookupValue);
    }
  } else if (matchType === 1) {
    return match_less_than_or_equal(matrix, lookupValue);
  } else if (matchType === -1) {
    return match_greater_than_or_equal(matrix, lookupValue);
  } else {
    throw Error('#N/A');
  }
}

function index(matrix, row_num, column_num) {
  if (row_num <= matrix.length) {
    var row = matrix[row_num - 1];
    if (Array.isArray(row)) {
      if (!column_num) {
        return row;
      } else if (column_num <= row.length) {
        return row[column_num - 1];
      }
    } else {
      return matrix[row_num];
    }
  }
  throw Error('#REF!');
}

// impl ported from https://github.com/FormulaPages/hlookup
function hlookup(needle, table, index, exactmatch) {
  if (typeof needle === 'undefined' || (0, is_blank)(needle)) {
    throw Error('#N/A');
  }

  index = index || 0;
  let row = table[0],
    i,
    searchingFor;

  if (typeof needle === 'string') {
    searchingFor = needle.toLowerCase();
    for (i = 0; i < row.length; i++) {
      if (
        (exactmatch && row[i] === searchingFor) ||
        row[i].toLowerCase().indexOf(searchingFor) !== -1
      ) {
        return index < table.length + 1 ? table[index - 1][i] : table[0][i];
      }
    }
  } else {
    searchingFor = needle;
    for (i = 0; i < row.length; i++) {
      if ((exactmatch && row[i] === searchingFor) || row[i] === searchingFor) {
        return index < table.length + 1 ? table[index - 1][i] : table[0][i];
      }
    }
  }
  throw Error('#N/A');
}

function len(a) {
  return ('' + a).length;
}

function trim(a) {
  return ('' + a).trim();
}

function is_blank(a) {
  // console.log(a)
  return !a;
}

function covariance_p(a, b) {
  a = getArrayOfNumbers(a);
  b = getArrayOfNumbers(b);
  if (a.length != b.length) {
    return 'N/D';
  }
  var inv_n = 1.0 / a.length;
  var avg_a = sum.apply(this, a) / a.length;
  var avg_b = sum.apply(this, b) / b.length;
  var s = 0.0;
  for (var i = 0; i < a.length; i++) {
    s += (a[i] - avg_a) * (b[i] - avg_b);
  }
  return s * inv_n;
}

function getArrayOfNumbers(range) {
  var arr = [];
  for (var i = 0; i < range.length; i++) {
    var arg = range[i];
    if (Array.isArray(arg)) {
      var matrix = arg;
      for (var j = matrix.length; j--; ) {
        if (typeof matrix[j] == 'number') {
          arr.push(matrix[j]);
        } else if (Array.isArray(matrix[j])) {
          for (var k = matrix[j].length; k--; ) {
            if (typeof matrix[j][k] == 'number') {
              arr.push(matrix[j][k]);
            }
          }
        }
        // else {
        //   wtf is that?
        // }
      }
    } else {
      if (typeof arg == 'number') {
        arr.push(arg);
      }
    }
  }
  return arr;
}

function var_p() {
  var average = avg.apply(this, arguments);
  var s = 0.0;
  var c = 0;
  for (var i = 0; i < arguments.length; i++) {
    var arg = arguments[i];
    if (Array.isArray(arg)) {
      var matrix = arg;
      for (var j = matrix.length; j--; ) {
        for (var k = matrix[j].length; k--; ) {
          if (matrix[j][k] !== null && matrix[j][k] !== undefined) {
            s += Math.pow(matrix[j][k] - average, 2);
            c++;
          }
        }
      }
    } else {
      s += Math.pow(arg - average, 2);
      c++;
    }
  }
  return s / c;
}

function EXP(n) {
  return Math.pow(Math.E, n);
}

function avg() {
  var aux = counta.apply(this, arguments);
  if (aux == 0) {
    throw Error('#DIV/0!');
  }
  return sum.apply(this, arguments) / aux;
}

function stDeviation() {
  var array = getArrayOfNumbers(arguments);

  function _mean(array) {
    return (
      array.reduce(function (a, b) {
        return a + b;
      }) / array.length
    );
  }
  var mean = _mean(array),
    dev = array.map(function (itm) {
      return (itm - mean) * (itm - mean);
    });
  return Math.sqrt(
    dev.reduce(function (a, b) {
      return a + b;
    }) /
      (array.length - 1),
  );
}

/// Original C++ implementation found at http://www.wilmott.com/messageview.cfm?catid=10&threadid=38771
/// C# implementation found at http://weblogs.asp.net/esanchez/archive/2010/07/29/a-quick-and-dirty-implementation-of-excel-norminv-function-in-c.aspx
/*
 *     Compute the quantile function for the normal distribution.
 *
 *     For small to moderate probabilities, algorithm referenced
 *     below is used to obtain an initial approximation which is
 *     polished with a final Newton step.
 *
 *     For very large arguments, an algorithm of Wichura is used.
 *
 *  REFERENCE
 *
 *     Beasley, J. D. and S. G. Springer (1977).
 *     Algorithm AS 111: The percentage points of the normal distribution,
 *     Applied Statistics, 26, 118-121.
 *
 *      Wichura, M.J. (1988).
 *      Algorithm AS 241: The Percentage Points of the Normal Distribution.
 *      Applied Statistics, 37, 477-484.
 */
function normsInv(p, mu, sigma) {
  if (p < 0 || p > 1) {
    throw 'The probality p must be bigger than 0 and smaller than 1';
  }
  if (sigma < 0) {
    throw 'The standard deviation sigma must be positive';
  }

  if (p == 0) {
    return -Infinity;
  }
  if (p == 1) {
    return Infinity;
  }
  if (sigma == 0) {
    return mu;
  }

  var q, r, val;

  q = p - 0.5;

  /*-- use AS 241 --- */
  /* double ppnd16_(double *p, long *ifault)*/
  /*      ALGORITHM AS241  APPL. STATIST. (1988) VOL. 37, NO. 3
            Produces the normal deviate Z corresponding to a given lower
            tail area of P; Z is accurate to about 1 part in 10**16.
    */
  if (Math.abs(q) <= 0.425) {
    /* 0.075 <= p <= 0.925 */
    r = 0.180625 - q * q;
    val =
      (q *
        (((((((r * 2509.0809287301226727 + 33430.575583588128105) * r + 67265.770927008700853) * r +
          45921.953931549871457) *
          r +
          13731.693765509461125) *
          r +
          1971.5909503065514427) *
          r +
          133.14166789178437745) *
          r +
          3.387132872796366608)) /
      (((((((r * 5226.495278852854561 + 28729.085735721942674) * r + 39307.89580009271061) * r +
        21213.794301586595867) *
        r +
        5394.1960214247511077) *
        r +
        687.1870074920579083) *
        r +
        42.313330701600911252) *
        r +
        1);
  } else {
    /* closer than 0.075 from {0,1} boundary */

    /* r = min(p, 1-p) < 0.075 */
    if (q > 0) r = 1 - p;
    else r = p;

    r = Math.sqrt(-Math.log(r));
    /* r = sqrt(-log(r))  <==>  min(p, 1-p) = exp( - r^2 ) */

    if (r <= 5) {
      /* <==> min(p,1-p) >= exp(-25) ~= 1.3888e-11 */
      r += -1.6;
      val =
        (((((((r * 7.7454501427834140764e-4 + 0.0227238449892691845833) * r +
          0.24178072517745061177) *
          r +
          1.27045825245236838258) *
          r +
          3.64784832476320460504) *
          r +
          5.7694972214606914055) *
          r +
          4.6303378461565452959) *
          r +
          1.42343711074968357734) /
        (((((((r * 1.05075007164441684324e-9 + 5.475938084995344946e-4) * r +
          0.0151986665636164571966) *
          r +
          0.14810397642748007459) *
          r +
          0.68976733498510000455) *
          r +
          1.6763848301838038494) *
          r +
          2.05319162663775882187) *
          r +
          1);
    } else {
      /* very close to  0 or 1 */
      r += -5;
      val =
        (((((((r * 2.01033439929228813265e-7 + 2.71155556874348757815e-5) * r +
          0.0012426609473880784386) *
          r +
          0.026532189526576123093) *
          r +
          0.29656057182850489123) *
          r +
          1.7848265399172913358) *
          r +
          5.4637849111641143699) *
          r +
          6.6579046435011037772) /
        (((((((r * 2.04426310338993978564e-15 + 1.4215117583164458887e-7) * r +
          1.8463183175100546818e-5) *
          r +
          7.868691311456132591e-4) *
          r +
          0.0148753612908506148525) *
          r +
          0.13692988092273580531) *
          r +
          0.59983220655588793769) *
          r +
          1);
    }

    if (q < 0.0) {
      val = -val;
    }
  }

  return mu + sigma * val;
}

function irr(range, guess) {
  var min = -2.0;
  var max = 1.0;
  var n = 0;
  do {
    var guest = (min + max) / 2;
    var NPV = 0;
    for (var i = 0; i < range.length; i++) {
      var arg = range[i];
      NPV += arg[0] / Math.pow(1 + guest, i);
    }
    if (NPV > 0) {
      if (min === max) {
        max += Math.abs(guest);
      }
      min = guest;
    } else {
      max = guest;
    }
    n++;
  } while (Math.abs(NPV) > 0.000001 && n < 100000);
  //console.log(n);
  return guest;
}

function counta() {
  var r = 0;
  for (var i = arguments.length; i--; ) {
    var arg = arguments[i];
    if (Array.isArray(arg)) {
      var matrix = arg;
      for (var j = matrix.length; j--; ) {
        for (var k = matrix[j].length; k--; ) {
          if (matrix[j][k] !== null && matrix[j][k] !== undefined) {
            r++;
          }
        }
      }
    } else {
      if (arg !== null && arg !== undefined) {
        r++;
      }
    }
  }
  return r;
}

function pmt(rate_per_period, number_of_payments, present_value, future_value, type) {
  type = type || 0;
  future_value = future_value || 0;
  if (rate_per_period != 0.0) {
    // Interest rate exists
    var q = Math.pow(1 + rate_per_period, number_of_payments);
    return (
      -(rate_per_period * (future_value + q * present_value)) /
      ((-1 + q) * (1 + rate_per_period * type))
    );
  } else if (number_of_payments != 0.0) {
    // No interest rate, but number of payments exists
    return -(future_value + present_value) / number_of_payments;
  }
  return 0;
}

function concatenate() {
  var r = '';
  for (var i = 0; i < arguments.length; i++) {
    var arg = arguments[i];
    if (arg === null || arg === undefined) continue;
    r += arg;
  }
  return r;
}

function sum() {
  var r = 0;
  for (var i = arguments.length; i--; ) {
    var arg = arguments[i];
    if (Array.isArray(arg)) {
      var matrix = arg;
      for (var j = matrix.length; j--; ) {
        for (var k = matrix[j].length; k--; ) {
          if (!isNaN(matrix[j][k])) {
            r += +matrix[j][k];
          }
        }
      }
    } else {
      r += +arg;
    }
  }
  return r;
}

function max() {
  var max = null;
  for (var i = arguments.length; i--; ) {
    var arg = arguments[i];
    if (Array.isArray(arg)) {
      var arr = arg;
      for (var j = arr.length; j--; ) {
        var col = arr[j];
        if (Array.isArray(col)) {
          for (var k = col.length; k--; ) {
            if (max == null || (col[k] != null && max < col[k])) {
              max = col[k];
            }
          }
        } else if (max == null || (col != null && max < col)) {
          max = col;
        }
      }
    } else if (!isNaN(arg) && (max == null || (arg != null && max < arg))) {
      max = arg;
    }
  }
  return max;
}

function min() {
  var min = null;
  for (var i = arguments.length; i--; ) {
    var arg = arguments[i];
    if (Array.isArray(arg)) {
      var arr = arg;
      for (var j = arr.length; j--; ) {
        var col = arr[j];
        if (Array.isArray(col)) {
          for (var k = col.length; k--; ) {
            if (min == null || (col[k] != null && min > col[k])) {
              min = col[k];
            }
          }
        } else if (min == null || (col != null && min > col)) {
          min = col;
        }
      }
    } else if (!isNaN(arg) && (min == null || (arg != null && min > arg))) {
      min = arg;
    }
  }
  return min;
}

function vlookup(key, matrix, return_index, range_lookup = true) {
  var val = null;
  for (var i = 0; i < matrix.length; i++) {
    if (matrix[i][0] < key) {
      val = matrix[i][return_index - 1];
    } else if (matrix[i][0] == key) {
      return matrix[i][return_index - 1];
    }
  }
  if (range_lookup && val) {
    return val;
  }
  throw Error('#N/A');
}

function iserror() {
  // if an error is catched before getting there, true will be returned from the catch block
  // if we get here then it's not an error
  return false;
}

function time(hours, minutes, seconds) {
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  return (((hours * 60 + minutes) * 60 + seconds) * 1000) / MS_PER_DAY;
}

function day(date) {
  if (!date.getDate) {
    throw Error('#VALUE!');
  }
  var day = date.getUTCDate();
  if (isNaN(day)) {
    throw Error('#VALUE!');
  }
  return day;
}

function month(date) {
  if (!date.getMonth) {
    throw Error('#VALUE!');
  }
  var month = date.getUTCMonth();
  if (isNaN(month)) {
    throw Error('#VALUE!');
  }
  return month + 1;
}

function year(date) {
  if (!date.getFullYear) {
    throw Error('#VALUE!');
  }
  var year = date.getUTCFullYear();
  if (isNaN(year)) {
    throw Error('#VALUE!');
  }
  return year;
}

function datediff(date1, date2, unit) {
  date1 = new Date(date1);
  date2 = new Date(date2);

  if (!date1 || !date2 || date1 == 'Invalid Date' || date2 == 'Invalid Date') {
    throw Error('#VALUE!');
  }

  unit = unit.replace(/[^DMY]/gi, '');

  switch (unit) {
    case 'M':
      return date2.getMonth() - date1.getMonth() + 12 * (date2.getFullYear() - date1.getFullYear());
    case 'Y':
      return Math.abs(date2.getUTCFullYear() - date1.getUTCFullYear());
    case 'D':
    default:
      var timeDiff = Math.abs(date2 - date1);
      return Math.ceil(timeDiff / (1000 * 3600 * 24));
  }
}

function eomonth(date, months) {
  date = new Date(date);
  if (!date || date == 'Invalid Date') {
    throw Error('#VALUE!');
  }
  months = months || 0;
  var endofmonth = new Date(date.getUTCFullYear(), date.getUTCMonth() + months + 1, 0);
  endofmonth.setUTCHours(0);
  endofmonth.setUTCMinutes(0);
  endofmonth.setUTCSeconds(0);
  return endofmonth;
}

function right(text, number) {
  number = number === undefined ? 1 : parseFloat(number);

  if (isNaN(number)) {
    throw Error('#VALUE!');
  }
  if (text === undefined || text === null) {
    text = '';
  } else {
    text = '' + text;
  }
  return text.substring(text.length - number);
}

function left(text, number) {
  number = number === undefined ? 1 : parseFloat(number);

  if (isNaN(number)) {
    throw Error('#VALUE!');
  }
  if (text === undefined || text === null) {
    text = '';
  } else {
    text = '' + text;
  }
  return text.substring(0, number);
}

function ifs(/*_cond1, _val1, _cond2, _val2, _cond3, _val3, ... */) {
  for (var i = 0; i + 1 < arguments.length; i += 2) {
    var cond = arguments[i];
    var val = arguments[i + 1];
    if (cond) {
      return val;
    }
  }
  throw Error('#N/A');
}

function escapeRegExp(str) {
  return str.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, '\\$1');
}

function substitute(text, old_text, new_text, occurrence) {
  if (occurrence <= 0) {
    throw Error('#VALUE!');
  }
  if (!text || !old_text || (!new_text && new_text !== '')) {
    return text;
  } else if (occurrence === undefined) {
    return text.replace(new RegExp(escapeRegExp(old_text), 'g'), new_text);
  } else {
    var index = 0;
    var i = 0;
    while (text.indexOf(old_text, index) > 0) {
      index = text.indexOf(old_text, index + 1);
      i++;
      if (i === occurrence) {
        return text.substring(0, index) + new_text + text.substring(index + old_text.length);
      }
    }
  }
}

function ceiling(number, significance) {
  return Math.ceil(number / significance) * significance;
}
/**
 * Filters an array based on a Boolean (True/False) array.
 *
 * Category: Lookup and reference
 *
 * @param {*} array The array, or range to filter. E.g. [[1,2,3],[4,5,6]]
 * @param {*} include A boolean array whose height or width is the same as the array. E.g. [[true, false, true]] OR [[true],[false]]
 * @param {*} if_empty Optional. The value to return if all values in the included array are empty (filter returns nothing). E.g. "No results"
 * @returns
 */
function FILTER(array, include, if_empty) {
  // correct types
  if (!array || !include) {
    return error.na;
  }

  if (!(array instanceof Array)) {
    return error.na;
  }

  if (!(include instanceof Array)) {
    return error.na;
  }

  // array lengths must be greater than 0 and symmetrical
  if (array.length === 0) {
    return error.na;
  }

  if (include.length === 0) {
    return error.na;
  }

  for (let i = 0; i < array.length; i++) {
    if (!(array[i] instanceof Array)) {
      return error.na;
    }

    if (array[i].length === 0) {
      return error.na;
    }

    if (array[i].length !== array[0].length) {
      return error.na;
    }
  }

  for (let i = 0; i < include.length; i++) {
    if (!(include[i] instanceof Array)) {
      return error.na;
    }

    if (include[i].length === 0) {
      return error.na;
    }

    if (include[i].length !== include[0].length) {
      return error.na;
    }
  }

  const arrayWidth = array[0].length;
  const arrayHeight = array.length;
  const includeWidth = include[0].length;
  const includeHeight = include.length;

  // include array must have same width or height as array (and generally not both)
  if (arrayWidth !== includeWidth && arrayHeight !== includeHeight) {
    return error.na;
  }

  if (
    arrayHeight > 1 &&
    arrayWidth > 1 &&
    ((arrayWidth === includeWidth && includeHeight !== 1) ||
      (arrayHeight === includeHeight && includeWidth !== 1))
  ) {
    return error.na;
  }

  if (
    arrayHeight > 1 &&
    arrayWidth === 1 &&
    (includeWidth !== 1 || (includeHeight !== 1 && includeHeight !== arrayHeight))
  ) {
    return error.na;
  }

  // filter
  const result = [];
  for (let i = 0; i < arrayHeight; i++) {
    const row = [];
    for (let j = 0; j < arrayWidth; j++) {
      const value =
        getMatrixValueForFilter(include, i, j) ||
        getMatrixValueForFilter(include, 0, j) ||
        getMatrixValueForFilter(include, i, 0);
      const bool = utils.parseBool(value);
      if (bool === true) row.push(array[i][j]);
      else if (bool instanceof Error)
        return utils.addEmptyValuesToArray([[bool]], arrayWidth, arrayHeight);
    }
    if (row.length > 0) result.push(row);
  }

  if (result.length === 0) {
    if (if_empty != null) {
      return utils.addEmptyValuesToArray([[if_empty]], arrayWidth, arrayHeight);
    }

    return utils.addEmptyValuesToArray([[error.calc]], arrayWidth, arrayHeight);
  }

  return utils.addEmptyValuesToArray(result, arrayWidth, arrayHeight);
}

function getMatrixValueForFilter(matrix, i, j) {
  if (matrix[i]) {
    if (matrix[i][j] === false) {
      return 'FALSE';
    }
    return matrix[i][j];
  }
}

function throwErrors(someFormula) {
  return function () {
    const result = someFormula.apply(this, arguments);
    if (result instanceof Error) {
      throw result;
    }
    return result;
  };
}

function XLOOKUP(
  lookup_value,
  lookup_array_ref,
  return_array_ref,
  if_not_found,
  match_mode = 0,
  search_mode = 1,
) {
  const ctx = this;
  const workbook = ctx.wb;
  const sheet_name = ctx.sheet_name;

  lookup_value = anytype2value(lookup_value);

  let lookup_array = getRangeValues(lookup_array_ref, ctx);

  let return_array = getRangeValues(return_array_ref, ctx);

  lookup_array = flattenArray(lookup_array);
  return_array = flattenArray(return_array);

  if (lookup_array.length !== return_array.length) {
    throw error.na;
  }

  let index = -1;
  if (search_mode === 1) {
    index = searchLookupArray(lookup_value, lookup_array, match_mode);
  } else if (search_mode === -1) {
    index = searchLookupArrayReverse(lookup_value, lookup_array, match_mode);
  } else {
    throw error.value;
  }

  if (index !== -1) {
    let result = return_array[index];
    result = anytype2value(result);
    return result;
  } else {
    if (if_not_found !== undefined) {
      return if_not_found;
    } else {
      throw error.na;
    }
  }
}

function getRangeValues(range_ref, ctx) {
  const workbook = ctx.wb;
  const sheet_name = ctx.sheet_name;

  if (typeof range_ref === 'string') {
    const { sheetName, range } = parseRangeReference(range_ref, sheet_name);

    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      throw error.ref;
    }

    const cells = getCellsInRange(sheet, range);
    const values = cells.map(cell => (cell.v !== undefined ? cell.v : null));
    return values;
  } else {
    return range_ref;
  }
}

function parseRangeReference(range_ref, current_sheet_name) {
  const parts = range_ref.split('!');
  let sheetName = '';
  let range = '';

  if (parts.length === 2) {
    sheetName = parts[0].replace(/^'|'$/g, '');
    range = parts[1];
  } else {
    sheetName = current_sheet_name;
    range = range_ref;
  }

  return { sheetName, range };
}

function getCellsInRange(sheet, range) {
  let rangeRef;

  if (/^[A-Za-z]+:[A-Za-z]+$/.test(range)) {
    const colStart = range.split(':')[0];
    const colEnd = range.split(':')[1];
    const maxRow = getMaxRow(sheet);
    rangeRef = {
      s: { c: XLSX.utils.decode_col(colStart), r: 0 },
      e: { c: XLSX.utils.decode_col(colEnd), r: maxRow },
    };
  } else if (/^\d+:\d+$/.test(range)) {
    const rowStart = parseInt(range.split(':')[0], 10) - 1;
    const rowEnd = parseInt(range.split(':')[1], 10) - 1;
    const maxCol = getMaxCol(sheet);
    rangeRef = {
      s: { c: 0, r: rowStart },
      e: { c: maxCol, r: rowEnd },
    };
  } else {
    rangeRef = XLSX.utils.decode_range(range);
  }

  const cells = [];
  for (let R = rangeRef.s.r; R <= rangeRef.e.r; ++R) {
    for (let C = rangeRef.s.c; C <= rangeRef.e.c; ++C) {
      const cell_address = { c: C, r: R };
      const cell_ref = XLSX.utils.encode_cell(cell_address);
      const cell = sheet[cell_ref] || {};
      cells.push(cell);
    }
  }

  return cells;
}

function getMaxRow(sheet) {
  let maxRow = 0;
  for (const cellRef in sheet) {
    if (cellRef[0] === '!') continue;
    const cellAddress = XLSX.utils.decode_cell(cellRef);
    if (cellAddress.r > maxRow) {
      maxRow = cellAddress.r;
    }
  }
  return maxRow;
}

function getMaxCol(sheet) {
  let maxCol = 0;
  for (const cellRef in sheet) {
    if (cellRef[0] === '!') continue;
    const cellAddress = XLSX.utils.decode_cell(cellRef);
    if (cellAddress.c > maxCol) {
      maxCol = cellAddress.c;
    }
  }
  return maxCol;
}

function flattenArray(array) {
  const result = [];
  (function flatten(arr) {
    arr.forEach(el => {
      if (Array.isArray(el)) {
        flatten(el);
      } else {
        result.push(el);
      }
    });
  })(array);
  return result;
}

function anytype2value(val) {
  if (val && val.v !== undefined) {
    return val.v;
  }
  return val;
}

function searchLookupArray(lookup_value, lookup_array, match_mode) {
  for (let i = 0; i < lookup_array.length; i++) {
    const lookup_item = anytype2value(lookup_array[i]);
    if (matchValues(lookup_value, lookup_item, match_mode)) {
      return i;
    }
  }
  return -1;
}

function searchLookupArrayReverse(lookup_value, lookup_array, match_mode) {
  for (let i = lookup_array.length - 1; i >= 0; i--) {
    const lookup_item = anytype2value(lookup_array[i]);
    if (matchValues(lookup_value, lookup_item, match_mode)) {
      return i;
    }
  }
  return -1;
}

function matchValues(lookup_value, current_value, match_mode) {
  switch (match_mode) {
    case 0:
      return lookup_value === current_value;
    case -1:
      if (lookup_value === current_value) return true;
      if (typeof lookup_value === 'number' && typeof current_value === 'number') {
        return lookup_value > current_value;
      }
      return false;
    case 1:
      if (lookup_value === current_value) return true;
      if (typeof lookup_value === 'number' && typeof current_value === 'number') {
        return lookup_value < current_value;
      }
      return false;
    case 2:
      const regex = new RegExp(
        '^' + lookup_value.replace(/\*/g, '.*').replace(/\?/g, '.') + '$',
        'i',
      );
      return regex.test(current_value);
    default:
      throw error.value;
  }
}

module.exports = formulas;
