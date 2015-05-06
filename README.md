# node-spreadsheet-column

Convert spreadsheet columns coordinates to numeric coordinates and reverse too.

This is [Node](http://nodejs.org) version of my PHP library [coord2coord](https://github.com/malenkiki/coord2coord).

## Why?

Sometimes, when you develop some code about some CSV or something like that, you have reference given by some people using spreadsheet value (column A, AA, Z…) or index value (1, 2, 3) or even index value starting to zero (0, 1, 2…).

So, this little library deals with that. So, you can easily switch from one system to another.

## How?

You can play with it using two ways:

 - By using library to use it into your code
 - Into the futur, by using CLI script

## Install it!

Simple. Just use npm:

```
npm install --save node-spreadsheet-column
```

## Use it!

Example is better than long blahblah, so:

```js
var Spreadsheetcolumn = require('node-spreadsheet-column');
var sc = new SpreadsheetColumn();

console.log(sc.fromInt(1)); // "A"
console.log(sc.fromInt(26)); // "Z"
console.log(sc.fromInt(27)); // "AA"
console.log(sc.fromInt(28)); // "AB"

console.log(sc.fromStr('A')); // 1
console.log(sc.fromStr('Z')); // 26
console.log(sc.fromStr('AA')); // 27
console.log(sc.fromStr('AB')); // 28
```

Simple, isn’t it? You can even use zero-indexed way too, just tell it when you create object:

```js
var Spreadsheetcolumn = require('node-spreadsheet-column');
var scz = new SpreadsheetColumn({zero: true});

console.log(scz.fromInt(0)); // "A"
console.log(scz.fromInt(25)); // "Z"
console.log(scz.fromInt(26)); // "AA"
console.log(scz.fromInt(27)); // "AB"

console.log(scz.fromStr('A')); // 0
console.log(scz.fromStr('Z')); // 25
console.log(scz.fromStr('AA')); // 26
console.log(scz.fromStr('AB')); // 27
```

## Test it!

Clone current source code, then go into the source’s directory, and install dependencies:

```
npm install
```

And then run unit tests:

```
npm test
```

## Near futur

Soon, CLI version will be available, like its PHP brother.

## License?

It is under MIT License.
