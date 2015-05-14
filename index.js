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
     * @private 
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
     * @private 
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
     * @public
     * @param {!Number} Number Number to convert
     * @return {String} Column’s name using letters
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
     * @public 
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
     * @public 
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
