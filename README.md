<pre>
 _  _  __    ____  _  _     ___   __   __     ___ 
( \/ )(  )  / ___)( \/ )   / __) / _\ (  )   / __)
 )  ( / (_/\\___ \ )  (   ( (__ /    \/ (_/\( (__ 
(_/\_)\____/(____/(_/\_)   \___)\_/\_/\____/ \___)</pre>
<div style="clear: both"></div>

![alt text](https://travis-ci.org/fabiooshiro/xlsx-calc.svg?branch=master "Build status")

# Installation
With [npm](https://www.npmjs.org/package/xlsx-calc):
```sh
npm install xlsx-calc
```

# How to use

Read the workbook with the great <a href="https://github.com/SheetJS/js-xlsx">js-xlsx</a> lib.
```js
var XLSX = require('xlsx');
var workbook = XLSX.readFile('test.xlsx');

// change some cell value
workbook.Sheets['Sheet1'].A1.v = 42;

// recalc the workbook
var XLSX_CALC = require('xlsx-calc');
XLSX_CALC(workbook);
```

## formulajs integration

`npm install --save formulajs`

```js
var XLSX_CALC = require('xlsx-calc');

// load your calc functions lib
var formulajs = require('formulajs');

// import your calc functions lib
XLSX_CALC.import_functions(formulajs);

var workbook = {Sheets: {Sheet1: {}}};

// use it
workbook.Sheets.Sheet1.A5 = {f: 'BETA.DIST(2, 8, 10, true, 1, 3)'};
XLSX_CALC(workbook);

// see the result -> 0.6854705810117458
console.log(workbook.Sheets.Sheet1.A5.v);
```

# How to contribute

Read the <a href="https://github.com/fabiooshiro/xlsx-calc/blob/master/test/1-basic-test.js">basic-tests.js</a>.

Run tests
```sh
$ npm run test-w
```

Run webpack
```sh
$ npm run dev
```

write some test like:
```js
//(...)
describe('HELLO', function() {
    it('says: Hello, World!', function() {
        workbook.Sheets['Sheet1'].A1.f = 'HELLO("World")';
        XLSX_CALC(workbook);
        assert.equal(workbook.Sheets['Sheet1'].A1.v, "Hello, World!");
    });
});
//(...)
```

Register your formula/function in the src/formulas.js file
below the commentary "FORMULAS REGISTERED"

```js
  // +---------------------+
  // | FORMULAS REGISTERED |
  // +---------------------+
  var formulas = {
    'FLOOR': Math.floor,
    'COUNTA': counta,
    'IRR': irr,
    'HELLO': hello // <---- Your contribution!!
  };
```
Write the implementation function below the commentary "THE IMPLEMENTATIONS".

```js
// +---------------------+
// | THE IMPLEMENTATIONS |
// +---------------------+
function hello(name) {
  return name;
}
```

If everything is OK you will see the mocha out:

```sh
  1) XLSX_CALC HELLO says: Hello, World!:

      AssertionError: "World" == "Hello, World!"
      + expected - actual

      -World
      +Hello, World!
      
      at Context.<anonymous> (test/basic-test.js:510:20)
```

So end with the correct implementation:

```js
// +---------------------+
// | THE IMPLEMENTATIONS |
// +---------------------+
function hello(name) {
  return "Hello, " + name + "!";
}
```
Now in terminal:

```sh
  HELLO
    ✓ says: Hello, World!

  79 passing (75ms)
```

> Give me the balloon watermelon!

Create a pull request

Thx!

# MIT LICENSE

Copyright 2017, fabiooshiro

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
