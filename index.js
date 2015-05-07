function SpreadsheetColumn(opt){


    this.fromZero = false;
    this.arr = [];

    this.init = function(opt){
        if(opt && typeof(opt.zero) === 'boolean'){
            this.fromZero = opt.zero;
        }

        for(var code = 97; code < (97 + 26); code++){
            this.arr.push(String.fromCharCode(code));
        }
    };

    this.getLetterByIndex = function(idx){
        return this.arr[idx - 1];
    };

    this.getIndexByLetter = function(letter){
        return this.arr.indexOf(letter) + 1;
    };

    this.letter = function(code) {
        var l = String.fromCharCode(code);

        if(l.match(/^[A-Z]$/g)){
            return l;
        } else {
            return '';
        }
    };

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
