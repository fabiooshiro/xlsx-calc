"use strict";

// +---------------------+
// | FORMULAS REGISTERED |
// +---------------------+
let formulas = {
    'FLOOR': Math.floor,
    '_xlfn.FLOOR.MATH': Math.floor,
    'ABS': Math.abs,
    'SQRT': Math.sqrt,
    'VLOOKUP': vlookup,
    'MAX': max,
    'SUM': sum,
    'MIN': min,
    'CONCATENATE': concatenate,
    'IF': _if,
    'PMT': pmt,
    'COUNTA': counta,
    'IRR': irr,
    'NORM.INV': normsInv,
    '_xlfn.NORM.INV': normsInv,
    'STDEV': stDeviation,
    'AVERAGE': avg,
    'EXP': EXP,
    'LN': Math.log,
    '_xlfn.VAR.P': var_p,
    'VAR.P': var_p,
    '_xlfn.COVARIANCE.P': covariance_p,
    'COVARIANCE.P': covariance_p,
    'TRIM': trim,
    'LEN': len,
    'ISBLANK': is_blank
};

function len(a) {
    return ('' + a).length;
}

function trim(a) {
    return ('' + a).trim();
}

function is_blank(a) {
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
            for (var j = matrix.length; j--;) {
                if (typeof(matrix[j]) == 'number') {
                    arr.push(matrix[j]);
                }
                else if (Array.isArray(matrix[j])) {
                    for (var k = matrix[j].length; k--;) {
                        if (typeof(matrix[j][k]) == 'number') {
                            arr.push(matrix[j][k]);
                        }
                    }
                }
                // else {
                //   wtf is that?
                // }
            }
        }
        else {
            if (typeof(arg) == 'number') {
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
            for (var j = matrix.length; j--;) {
                for (var k = matrix[j].length; k--;) {
                    if (matrix[j][k] !== null && matrix[j][k] !== undefined) {
                        s += Math.pow(matrix[j][k] - average, 2);
                        c++;
                    }
                }
            }
        }
        else {
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
    return sum.apply(this, arguments) / counta.apply(this, arguments);
}

function stDeviation() {
    var array = getArrayOfNumbers(arguments);

    function _mean(array) {
        return array.reduce(function(a, b) {
            return a + b;
        }) / array.length;
    }
    var mean = _mean(array),
        dev = array.map(function(itm) {
            return (itm - mean) * (itm - mean);
        });
    return Math.sqrt(dev.reduce(function(a, b) {
        return a + b;
    }) / (array.length - 1));
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
        throw "The probality p must be bigger than 0 and smaller than 1";
    }
    if (sigma < 0) {
        throw "The standard deviation sigma must be positive";
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
    if (Math.abs(q) <= .425) { /* 0.075 <= p <= 0.925 */
        r = .180625 - q * q;
        val =
            q * (((((((r * 2509.0809287301226727 +
                            33430.575583588128105) * r + 67265.770927008700853) * r +
                        45921.953931549871457) * r + 13731.693765509461125) * r +
                    1971.5909503065514427) * r + 133.14166789178437745) * r +
                3.387132872796366608) / (((((((r * 5226.495278852854561 +
                        28729.085735721942674) * r + 39307.89580009271061) * r +
                    21213.794301586595867) * r + 5394.1960214247511077) * r +
                687.1870074920579083) * r + 42.313330701600911252) * r + 1);
    }
    else { /* closer than 0.075 from {0,1} boundary */

        /* r = min(p, 1-p) < 0.075 */
        if (q > 0)
            r = 1 - p;
        else
            r = p;

        r = Math.sqrt(-Math.log(r));
        /* r = sqrt(-log(r))  <==>  min(p, 1-p) = exp( - r^2 ) */

        if (r <= 5) { /* <==> min(p,1-p) >= exp(-25) ~= 1.3888e-11 */
            r += -1.6;
            val = (((((((r * 7.7454501427834140764e-4 +
                                .0227238449892691845833) * r + .24178072517745061177) *
                            r + 1.27045825245236838258) * r +
                        3.64784832476320460504) * r + 5.7694972214606914055) *
                    r + 4.6303378461565452959) * r +
                1.42343711074968357734) / (((((((r *
                                1.05075007164441684324e-9 + 5.475938084995344946e-4) *
                            r + .0151986665636164571966) * r +
                        .14810397642748007459) * r + .68976733498510000455) *
                    r + 1.6763848301838038494) * r +
                2.05319162663775882187) * r + 1);
        }
        else { /* very close to  0 or 1 */
            r += -5;
            val = (((((((r * 2.01033439929228813265e-7 +
                                2.71155556874348757815e-5) * r +
                            .0012426609473880784386) * r + .026532189526576123093) *
                        r + .29656057182850489123) * r +
                    1.7848265399172913358) * r + 5.4637849111641143699) *
                r + 6.6579046435011037772) / (((((((r *
                            2.04426310338993978564e-15 + 1.4215117583164458887e-7) *
                        r + 1.8463183175100546818e-5) * r +
                    7.868691311456132591e-4) * r + .0148753612908506148525) * r + .13692988092273580531) * r +
                .59983220655588793769) * r + 1);
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
            NPV += arg[0] / Math.pow((1 + guest), i);
        }
        if (NPV > 0) {
            if (min === max) {
                max += Math.abs(guest);
            }
            min = guest;
        }
        else {
            max = guest;
        }
        n++;
    } while (Math.abs(NPV) > 0.000001 && n < 100000);
    //console.log(n);
    return guest;
}

function counta() {
    var r = 0;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var matrix = arg;
            for (var j = matrix.length; j--;) {
                for (var k = matrix[j].length; k--;) {
                    if (matrix[j][k] !== null && matrix[j][k] !== undefined) {
                        r++;
                    }
                }
            }
        }
        else {
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
        return -(rate_per_period * (future_value + (q * present_value))) / ((-1 + q) * (1 + rate_per_period * (type)));

    }
    else if (number_of_payments != 0.0) {
        // No interest rate, but number of payments exists
        return -(future_value + present_value) / number_of_payments;
    }
    return 0;
}

function _if(condition, _then, _else) {
    if (condition) {
        return _then;
    }
    else {
        return _else;
    }
}

function concatenate() {
    var r = '';
    for (var i = 0; i < arguments.length; i++) {
        r += arguments[i];
    }
    return r;
}

function sum() {
    //console.log('sum...');
    var r = 0;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var matrix = arg;
            for (var j = matrix.length; j--;) {
                for (var k = matrix[j].length; k--;) {
                    if (!isNaN(matrix[j][k])) {
                        r += +matrix[j][k];
                    }
                }
            }
        }
        else {
            r += +arg;
        }
    }
    //console.log('end sum.');
    return r;
}

function max() {
    var max = null;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var arr = arg;
            for (var j = arr.length; j--;) {
                max = max == null || max < arr[j] ? arr[j] : max;
            }
        }
        else if (!isNaN(arg)) {
            max = max == null || max < arg ? arg : max;
        }
        else {
            console.log('WTF??', arg);
        }
    }
    return max;
}

function min() {
    var result = null;
    for (var i = arguments.length; i--;) {
        var arg = arguments[i];
        if (Array.isArray(arg)) {
            var arr = arg;
            for (var j = arr.length; j--;) {
                result = result == null || result > arr[j] ? arr[j] : result;
            }
        }
        else if (!isNaN(arg)) {
            result = result == null || result > arg ? arg : result;
        }
        else {
            console.log('WTF??', arg);
        }
    }
    return result;
}

function vlookup(key, matrix, return_index) {
    for (var i = 0; i < matrix.length; i++) {
        if (matrix[i][0] == key) {
            return matrix[i][return_index - 1];
        }
    }
    throw Error('#N/A');
}

module.exports = formulas;
