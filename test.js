var SpreadsheetColumn = require('./index.js');
var assert = require("assert")

var sc = new SpreadsheetColumn();

describe('Converting from numeric index', function(){
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

    //TODO I must enhance perf here : too slow!
    it('should return three-characters length string', function(){
        assert.equal(sc.fromInt(900).length, 3);
        //assert.equal(sc.fromInt(12000).length, 3);
        //assert.equal(sc.fromInt(17576).length, 3);
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
