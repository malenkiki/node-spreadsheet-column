/*
The MIT License (MIT)

Copyright (c) 2015 Michel Petit

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/


/**
 * Implements new SpreadSheetColumn object. 
 * 
 * @class
 * @param {{zero: Boolean}} Options
 * @return {Object}
 */
function SpreadsheetColumn(opt){


    this.fromZero = false;
    this.arr = [];



    /**
     * Constructor, takes options and does some basic stuff.
     *
     * @private
     * @param {{zero: Boolean}} Options
     */
    this.init = function(opt){
        if(opt && typeof(opt.zero) === 'boolean'){
            this.fromZero = opt.zero;
        }

        for(var code = 97; code < (97 + 26); code++){
            this.arr.push(String.fromCharCode(code));
        }
    };




    /**
     * Gets letters by its numeric index into alphabet.
     *
     * @private 
     * @param {Number} idx Numeric index of the letter
     * @param {String} Corresponding letter
     */
    this.getLetterByIndex = function(idx){
        return this.arr[idx - 1];
    };




    /**
     * @private 
     */
    this.getIndexByLetter = function(letter){
        return this.arr.indexOf(letter) + 1;
    };




    /**
     * Converts numeric ASCII code to its letter only if this letter is ASCII
     * character into the range [A-Z].
     *
     * @private 
     * @param {Number} code ASCII code
     * @return {String} Character [A-Z] or void string
     */
    this.letter = function(code) {
        var l = String.fromCharCode(code);

        if(l.match(/^[A-Z]$/g)){
            return l;
        } else {
            return '';
        }
    };




    /**
     * Converts an integer to string column name using letters.
     *
     * Example:
     *
     *     var sc = new SpreadsheetColumn();
     *     sc.fromInt(0); // Error
     *     sc.fromInt(1); // A
     *     sc.fromInt(26); // Z
     *     sc.fromInt(27); // AA
     *
     *     var scz = new SpreadsheetColumn({zero: true});
     *     scz.fromInt(0); // A
     *     scz.fromInt(1); // B
     *     scz.fromInt(26); // AA
     *     scz.fromInt(27); // AB
     *
     * @public
     * @param {!Number} n Number to convert into ASCII letter column name
     * @return {String} Column’s name using letters
     * @throws {Error} Exception if input number is not an integer.
     * @throws {Error} Exception if input number is negative.
     * @throws {Error} Exception if SpreadsheetColumn instance is set with `fromZero` to `false` and input param is zero.
     */
    this.fromInt = function(n) {

        if(typeof(n) !== 'number' || (typeof(n) === 'number' && (parseInt(n) - n) !== 0)){
            throw 'Numeric coordinate must be integer!';
        }

        if(n < 0){
            throw 'Negative numeric coordinate is not allowed!';
        }
        if(!this.fromZero && n === 0){
            throw 'Numeric coordinate cannot be zero if starting point is one!' ;
        }

        // Thanks to http://stackoverflow.com/a/3302991
        if(!this.fromZero){
            n = n - 1;
        }
        var nMod = n % 26;
        var l = this.letter(65 + nMod);
        var nIntDiv = parseInt(n / 26);
        if (nIntDiv > 0) {
            return this.fromInt(this.fromZero ? nIntDiv - 1 : nIntDiv) + l;
        } else {
            return l;
        }

    };




    /**
     * Converts alphabetic column name into its numeric version.
     *
     * Example:
     * 
     *     var sc = new SpreadsheetColumn();
     *     sc.fromStr("A"); // 1
     *     sc.fromStr("Z"); // 26
     *     sc.fromStr("AA"); // 27
     * 
     *     var scz = new SpreadsheetColumn({zero: true});
     *     scz.fromStr("A"); // 0
     *     scz.fromStr("Z"); // 25
     *     scz.fromStr("AA"); // 26
     *
     * @public 
     * @param {String} s String to parse
     * @throws {Error} Exception if parameter is not a String
     * @throws {Error} Exception if string does not contain ASCII letters (upper or lower cases)
     */
    this.fromStr = function(s){
        if(typeof(s) !== 'string'){
            throw 'Spreadsheet coordinate must be a string.';
        }

        if(!s.match(/^[A-Z]+$/i)){
            throw 'Spreadsheet coordinates must be a not null string containing ASCII letters.' ;
        }

        s = s.toLowerCase();
        var out = null;
        if(s.length > 1){
            var prov = s.split('').reverse();
            var cnt = 0;
            for(k in prov){
                var l = prov[k];
                cnt += this.getIndexByLetter(l) * Math.pow(this.arr.length, k);
            }
            out = cnt;
        } else {
            out = this.getIndexByLetter(s);
        }
        return this.fromZero ? out - 1 : out;
    };




    /**
     * Convert a string containing letters and digits to their converted values.
     *
     * This accepts any string containing ASCII letters and/or digits, separated
     * by any other characters. It splits the string into groups of letters and
     * digits, then it converts each group to numeric or alphabetic column
     * coordinates. 
     *
     * Example:
     *
     *     var sc = new SpreadsheetColumn();
     *     console.log(sc.fromAny('A Z 3;aa'));
     *     // gives:
     *     [
     *          {
     *              orginal: "A",
     *              converted: 1
     *          },
     *          {
     *              orginal: "Z",
     *              converted: 26
     *          },
     *          {
     *              orginal: 3,
     *              converted: "C"
     *          },
     *          {
     *              orginal: "AA",
     *              converted: 27
     *          }
     *     ]
     *
     *     var scz = new SpreadsheetColumn({zero: true});
     *     console.log(sc.fromAny('A Z 3;aa'));
     *     // gives:
     *     [
     *          {
     *              orginal: "A",
     *              converted: 0
     *          },
     *          {
     *              orginal: "Z",
     *              converted: 25
     *          },
     *          {
     *              orginal: 3,
     *              converted: "D"
     *          },
     *          {
     *              orginal: "AA",
     *              converted: 26
     *          }
     *     ]
     *
     * __Warning:__ if a group has letters AND digits, then this raises an error.
     *
     * __Note:__ Letters are converted to uppercases.
     *
     * @public 
     * @param {String} thing Input string to parse
     * @return {Object[]}
     * @throws {Error} Exception if param is not a String
     * @throws {Error} Exception if string does not contain valid characters
     * @throws {Error} Exception if at least one group has digits and letters together
     */
    this.fromAny = function(thing){
        if(typeof(thing) !== 'string'){
            throw 'Autodetection is done only with string of letters and digits.';
        }

        if(thing.match(/^[^0-9a-z]+$/gi) || thing.trim().length === 0){
            throw 'Autodetection cannot be done, there are some invalid characters';
        }

        var out = [];

        var columns = thing.split(/[^0-9a-z]+/gi);

        for(var i in columns){
            var item = columns[i];
            var checkings = item.split('');
            var type;

            for(var j in checkings){
                var c = checkings[j];
                var forcedToDigit = parseInt(c);

                if(j === 0 || typeof(type) !== 'boolean'){
                    type = isNaN(forcedToDigit);
                }

                if(isNaN(forcedToDigit) !== type){
                    throw 'Column having both digits and letters is not correct. Aborting';
                }
            }

            if(type){
                out.push({
                    original: item.toUpperCase(),
                    converted: this.fromStr(item)
                });
            } else {
                out.push({
                    original: parseInt(item),
                    converted: this.fromInt(parseInt(item))
                });
            }
            
            type = null;
        }



        return out;
    };

    this.init(opt);
};

module.exports = SpreadsheetColumn;
