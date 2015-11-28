<pre>
 _  _  __    ____  _  _     ___   __   __     ___ 
( \/ )(  )  / ___)( \/ )   / __) / _\ (  )   / __)
 )  ( / (_/\\___ \ )  (   ( (__ /    \/ (_/\( (__ 
(_/\_)\____/(____/(_/\_)   \___)\_/\_/\____/ \___)</pre>
<div style="clear: both"></div>

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
# How to contribute

Read the <a href="https://github.com/fabiooshiro/xlsx-calc/blob/master/test/basic-test.js">basic-tests.js</a>.

Run the mocha
```sh
$ mocha -w
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

Register your formula/function in the xlsx_Fx variable found inside <a href="https://github.com/fabiooshiro/xlsx-calc/blob/master/xlsx-calc.js">xlsx-calc.js</a> 
below the commentary "FORMULAS REGISTERED"

```js
  // +---------------------+
  // | FORMULAS REGISTERED |
  // +---------------------+
  var xlsx_Fx = {
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