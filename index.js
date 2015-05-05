function SpreadsheetColumn(opt){


    this.fromZero = false;
    this.arr = [];

    this.init = function(opt){
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

        var result = '';
        var count = 0;

        for(var i = 0; count < n; i++) {

            var name = '';
            var colN = i;

            while(colN > 0 && colN % 27 !== 0) {  
                var c = this.letter(65 + ((colN) % 27) - 1);
                name = c + name;
                colN = colN / 27;
            }

            result = name;

            if (i % 27 !== 0) {
                count++;
            }
        }

        return result;
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
    
    this.init(opt);
};


/*
var sc = new SpreadsheetColumn();
console.log(sc.fromInt(1));
console.log(sc.fromInt(2));
console.log(sc.fromInt(25));
console.log(sc.fromInt(26));
console.log(sc.fromInt(27));
console.log(sc.fromInt(28));

console.log(sc.fromStr("A"));
console.log(sc.fromStr("B"));
console.log(sc.fromStr("Y"));
console.log(sc.fromStr("Z"));
console.log(sc.fromStr("AA"));
console.log(sc.fromStr("AB"));
*/
