# node-spreadsheet-column

Convert spreadsheet columns coordinates to numeric coordinates and reverse too.

This is [Node](http://nodejs.org) version of my PHP library [coord2coord](https://github.com/malenkiki/coord2coord).

## Why?

Sometimes, when you develop some code about some CSV or something like that, you have reference given by some people using spreadsheet value (column A, AA, Z…) or index value (1, 2, 3) or even index value starting to zero (0, 1, 2…).

So, this little library deals with that. So, you can easily switch from one system to another.

## How?

You can play with it using two ways:

 - By using library to use it into your code
 - By using CLI script

## Install it!

Simple. Just use npm.

To use it into your project:

```
npm install --save spreadsheet-column
```

To use it as CLI application:

```
npm install -g spreadsheet-column
```


## Use it into your code!

Example is better than long blahblah, so:

```js
var Spreadsheetcolumn = require('spreadsheet-column');
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
var Spreadsheetcolumn = require('spreadsheet-column');
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

A new feature allows you to have collection of results from a string having digits and/or letters:

```js
var Spreadsheetcolumn = require('spreadsheet-column');
var sc = new SpreadsheetColumn();
console.log(sc.fromAny('V 4 L iD'));
// with custom separator too:
console.log(sc.fromAny('V;4;L;iD'));
```

Previous 2 instructions will return the same following result:

```js
[
    {
        original: "V",
        converted: 22
    },
    {
        original: 4,
        converted: "D"
    },
    {
        original: "L",
        converted: 12
    },
    {
        original: "ID",
        converted: 238
    }
]
```

Note: Letters are converted to uppercases.

Note 2: A group having both letters and digits throws an error. Example: `V4 LI D` has its first group invalid.

## Use it as CLI!

Very simple:

```
$ spreadsheet-column A B Z 45 AA AB
```

Gives you converting results for each values passed as arguments.

You can choose output format: JSON or formated output.

You can use 0 or 1 as starting point for number indexes.

Look at the help message, by running CLI without argument nor options.


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
