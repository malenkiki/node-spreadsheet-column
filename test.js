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

var SpreadsheetColumn = require('./index.js');
var assert = require("assert")

var sc = new SpreadsheetColumn();
var scz = new SpreadsheetColumn({zero: true});

describe('Converting from numeric index starting from one', function(){
    it('should throw error if not an integer', function(){
        assert.throws(function(){sc.fromInt("a");});
        assert.throws(function(){sc.fromInt(3.14);});
    });
    
    it('should throw error if it is a negative integer', function(){
        assert.throws(function(){sc.fromInt(-42);});
    });
    
    it('should throw error if zero', function(){
        assert.throws(function(){sc.fromInt(0);});
    });


    it('should return one-character length string', function(){
        assert.equal(sc.fromInt(1).length, 1);
        assert.equal(sc.fromInt(17).length, 1);
        assert.equal(sc.fromInt(26).length, 1);
    });

    it('should return two-characters length string', function(){
        assert.equal(sc.fromInt(27).length, 2);
        assert.equal(sc.fromInt(300).length, 2);
        assert.equal(sc.fromInt(676).length, 2);
    });

    it('should return three-characters length string', function(){
        assert.equal(sc.fromInt(900).length, 3);
        assert.equal(sc.fromInt(12000).length, 3);
        assert.equal(sc.fromInt(17576).length, 3);
    });

    it('should return "A" string for integer 1', function(){
        assert.equal(sc.fromInt(1), 'A');
    });

    it('should return "B" string for integer 2', function(){
        assert.equal(sc.fromInt(2), "B");
    });

    it('should return "Y" string for integer 25', function(){
        assert.equal(sc.fromInt(25), "Y");
    });

    it('should return "Z" string for integer 26', function(){
        assert.equal(sc.fromInt(26), "Z");
    });

    it('should return "AA" string for integer 27', function(){
        assert.equal(sc.fromInt(27), "AA");
    });

    it('should return "AB" string for integer 28', function(){
        assert.equal(sc.fromInt(28), "AB");
    });

    it('should return "BI" string for integer 61', function(){
        assert.equal(sc.fromInt(61), "BI");
    });

    it('should return "ZZ" string for integer 702', function(){
        assert.equal(sc.fromInt(702), "ZZ");
    });

    it('should return "AAA" string for integer 703', function(){
        assert.equal(sc.fromInt(703), "AAA");
    });

    it('should return "AAB" string for integer 704', function(){
        assert.equal(sc.fromInt(704), "AAB");
    });

    it('should return "ADO" string for integer 795', function(){
        assert.equal(sc.fromInt(795), "ADO");
    });

    it('should return "AEO" string for integer 821', function(){
        assert.equal(sc.fromInt(821), "AEO");
    });
});

describe('Converting from numeric index starting from zero', function(){
    it('should throw error if not an integer', function(){
        assert.throws(function(){scz.fromInt("a");});
        assert.throws(function(){scz.fromInt(3.14);});
    });
    
    it('should throw error if it is a negative integer', function(){
        assert.throws(function(){scz.fromInt(-42);});
    });
    
    it('should not throw error if zero', function(){
        assert.doesNotThrow(function(){scz.fromInt(0);});
    });


    it('should return "A" string for integer 0', function(){
        assert.equal(scz.fromInt(0), 'A');
    });

    it('should return "B" string for integer 1', function(){
        assert.equal(scz.fromInt(1), "B");
    });

    it('should return "Y" string for integer 24', function(){
        assert.equal(scz.fromInt(24), "Y");
    });

    it('should return "Z" string for integer 25', function(){
        assert.equal(scz.fromInt(25), "Z");
    });

    it('should return "AA" string for integer 26', function(){
        assert.equal(scz.fromInt(26), "AA");
    });

    it('should return "AB" string for integer 27', function(){
        assert.equal(scz.fromInt(27), "AB");
    });

    it('should return "BI" string for integer 60', function(){
        assert.equal(scz.fromInt(60), "BI");
    });

    it('should return "ZZ" string for integer 701', function(){
        assert.equal(scz.fromInt(701), "ZZ");
    });

    it('should return "AAA" string for integer 702', function(){
        assert.equal(scz.fromInt(702), "AAA");
    });

    it('should return "AAB" string for integer 703', function(){
        assert.equal(scz.fromInt(703), "AAB");
    });

    it('should return "ADO" string for integer 794', function(){
        assert.equal(scz.fromInt(794), "ADO");
    });

    it('should return "AEO" string for integer 820', function(){
        assert.equal(scz.fromInt(820), "AEO");
    });
});


describe('Converting from string to integer', function(){
    it('should throw error if argument is not a string', function(){
        assert.throws(function(){sc.fromStr(1);});
        assert.throws(function(){sc.fromStr(2.0);});
        assert.throws(function(){sc.fromStr({three: 3});});
        assert.throws(function(){sc.fromStr([4]);});
        assert.throws(function(){sc.fromStr(function(){var foo = 'I does nothing';});});
    });

    it('should throw error if argument is null string', function(){
        assert.throws(function(){sc.fromStr('');});
    });

    it('should throw error if argument is non only ASCII letter string', function(){
        assert.throws(function(){sc.fromStr('éçà:&?,()£$ ô.');});
    });

    it('should return integer value for a valid string', function(){
        assert.equal(typeof(sc.fromStr('BAR')), 'number');
        assert.equal(parseInt(sc.fromStr('BAR')) - sc.fromStr('BAR'), 0);
    });

    it('should return same value for same string in lower or upper cases', function(){
        assert.equal(sc.fromStr('BAR'), sc.fromStr('bar'));
        assert.equal(sc.fromStr('FOO'), sc.fromStr('foo'));
    });
});

describe('Converting from string to integer one-indexed', function(){
    it('should return 1 for "A" or "a"', function(){
        assert.equal(sc.fromStr('A'), 1);
        assert.equal(sc.fromStr('a'), 1);
    });
    
    it('should return 2 for "B" or "b"', function(){
        assert.equal(sc.fromStr('B'), 2);
        assert.equal(sc.fromStr('b'), 2);
    });
    
    it('should return 25 for "Y" or "y"', function(){
        assert.equal(sc.fromStr('Y'), 25);
        assert.equal(sc.fromStr('y'), 25);
    });
    
    it('should return 26 for "Z" or "z"', function(){
        assert.equal(sc.fromStr('Z'), 26);
        assert.equal(sc.fromStr('z'), 26);
    });
    
    it('should return 27 for "AA" or "aa" or "Aa" or "aA"', function(){
        assert.equal(sc.fromStr('AA'), 27);
        assert.equal(sc.fromStr('aa'), 27);
        assert.equal(sc.fromStr('Aa'), 27);
        assert.equal(sc.fromStr('aA'), 27);
    });
    
    it('should return 702 for "ZZ" or "zz" or "Zz" or "zZ"', function(){
        assert.equal(sc.fromStr('ZZ'), 702);
        assert.equal(sc.fromStr('zz'), 702);
        assert.equal(sc.fromStr('zZ'), 702);
        assert.equal(sc.fromStr('Zz'), 702);
    });

    it('should return 703 for "AAA" or "aaa" or "AAa" or…', function(){
        assert.equal(sc.fromStr('AAA'), 703);
        assert.equal(sc.fromStr('AAa'), 703);
        assert.equal(sc.fromStr('AaA'), 703);
        assert.equal(sc.fromStr('aAA'), 703);
        assert.equal(sc.fromStr('aaA'), 703);
        assert.equal(sc.fromStr('Aaa'), 703);
    });

    it('should return 704 for "AAB" or "aab" or "AAb" or…', function(){
        assert.equal(sc.fromStr('AAB'), 704);
        assert.equal(sc.fromStr('AAb'), 704);
        assert.equal(sc.fromStr('AaB'), 704);
        assert.equal(sc.fromStr('aAB'), 704);
        assert.equal(sc.fromStr('aaB'), 704);
        assert.equal(sc.fromStr('Aab'), 704);
    });

});

describe('Converting from string to integer zero-indexed', function(){
    it('should return 0 for "A" or "a"', function(){
        assert.equal(scz.fromStr('A'), 0);
        assert.equal(scz.fromStr('a'), 0);
    });
    
    it('should return 1 for "B" or "b"', function(){
        assert.equal(scz.fromStr('B'), 1);
        assert.equal(scz.fromStr('b'), 1);
    });
    
    it('should return 24 for "Y" or "y"', function(){
        assert.equal(scz.fromStr('Y'), 24);
        assert.equal(scz.fromStr('y'), 24);
    });
    
    it('should return 25 for "Z" or "z"', function(){
        assert.equal(scz.fromStr('Z'), 25);
        assert.equal(scz.fromStr('z'), 25);
    });
    
    it('should return 26 for "AA" or "aa" or "Aa" or "aA"', function(){
        assert.equal(scz.fromStr('AA'), 26);
        assert.equal(scz.fromStr('aa'), 26);
        assert.equal(scz.fromStr('Aa'), 26);
        assert.equal(scz.fromStr('aA'), 26);
    });
    
    it('should return 701 for "ZZ" or "zz" or "Zz" or "zZ"', function(){
        assert.equal(scz.fromStr('ZZ'), 701);
        assert.equal(scz.fromStr('zz'), 701);
        assert.equal(scz.fromStr('zZ'), 701);
        assert.equal(scz.fromStr('Zz'), 701);
    });

    it('should return 702 for "AAA" or "aaa" or "AAa" or…', function(){
        assert.equal(scz.fromStr('AAA'), 702);
        assert.equal(scz.fromStr('AAa'), 702);
        assert.equal(scz.fromStr('AaA'), 702);
        assert.equal(scz.fromStr('aAA'), 702);
        assert.equal(scz.fromStr('aaA'), 702);
        assert.equal(scz.fromStr('Aaa'), 702);
    });

    it('should return 703 for "AAB" or "aab" or "AAb" or…', function(){
        assert.equal(scz.fromStr('AAB'), 703);
        assert.equal(scz.fromStr('AAb'), 703);
        assert.equal(scz.fromStr('AaB'), 703);
        assert.equal(scz.fromStr('aAB'), 703);
        assert.equal(scz.fromStr('aaB'), 703);
        assert.equal(scz.fromStr('Aab'), 703);
    });

});

describe('Converting automagically from any type together', function(){
    it('should throw error if parameter is not a string', function(){
        assert.throws(function(){sc.fromAny(1);});
    });
    
    it('should throw error if parameter is a string without any letter or digit', function(){
        assert.throws(function(){sc.fromAny(':;; =+');});
        assert.throws(function(){sc.fromAny('');});
        assert.throws(function(){sc.fromAny('         ');});
    });

    it('should throw error if string has not valid column letter and/or digit indexes', function(){
        assert.throws(function(){sc.fromAny('A1 5Y U2 i9')});
    });

    it('should return an object if string is valid', function(){
        assert.equal(typeof(sc.fromAny('V 4 LI D ST R 1 NG')), 'object');
    });

    var result = [
        {
            original: 'V',
            converted: 22
        }, 
        {
            original: 4,
            converted: 'D'
        }, 
        {
            original: 'LI',
            converted: 321
        }, 
        {
            original: 'D',
            converted: 4
        }, 
        {
            original: 'ST',
            converted: 514
        }, 
        {
            original: 'R',
            converted: 18
        }, 
        {
            original: 1,
            converted: 'A'
        }, 
        {
            original: 'NG',
            converted: 371
        }
    ];
    
    var resultz = [
        {
            original: 'V',
            converted: 21
        }, 
        {
            original: 4,
            converted: 'E'
        }, 
        {
            original: 'LI',
            converted: 320
        }, 
        {
            original: 'D',
            converted: 3
        }, 
        {
            original: 'ST',
            converted: 513
        }, 
        {
            original: 'R',
            converted: 17
        }, 
        {
            original: 1,
            converted: 'B'
        }, 
        {
            original: 'NG',
            converted: 370
        }
    ];

    it('should return an array of objects as result of convertions from values into the string', function(){
        assert.deepEqual( sc.fromAny('V 4 LI D ST R 1 NG'), result);
        assert.deepEqual( scz.fromAny('V 4 LI D ST R 1 NG'), resultz);
        assert.deepEqual( sc.fromAny('V-4-LI;D,ST:R,1 NG'), result);
        assert.deepEqual( scz.fromAny('V 4 LI D ST.R/1 NG'), resultz);
    });

    it('should return same result object if source string differs only by cases', function(){
        assert.deepEqual(sc.fromAny('V 4 LI D ST R 1 NG'), sc.fromAny('v 4 li d st r 1 ng'));
        assert.deepEqual(scz.fromAny('V 4 LI D ST R 1 NG'), scz.fromAny('v 4 li d st r 1 ng'));
    });

    it('should have different result for the same source string, but with too diffrent instance: one from zero, other from one', function(){
        assert.notDeepEqual(sc.fromAny('V 4 LI D ST R 1 NG'), scz.fromAny('V 4 LI D ST R 1 NG'));
    });
});
