var SpreadsheetColumn = require('./index.js');
var assert = require("assert")

var sc = new SpreadsheetColumn();
var scz = new SpreadsheetColumn({zero: true});

describe('Converting from numeric index starting from one', function(){
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
