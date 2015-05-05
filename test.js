var SpreadsheetColumn = require('./index.js');
var assert = require("assert")

var sc = new SpreadsheetColumn();

describe('Converting from numeric index', function(){
    it('should return one-character length string', function(){
        assert.equal(1, sc.fromInt(1).length);
        assert.equal(1, sc.fromInt(17).length);
        assert.equal(1, sc.fromInt(26).length);
    });

    it('should return two-characters length string', function(){
        assert.equal(2, sc.fromInt(27).length);
        assert.equal(2, sc.fromInt(300).length);
        assert.equal(2, sc.fromInt(676).length);
    });

    //TODO I must enhance perf here : too slow!
    it('should return three-characters length string', function(){
        assert.equal(3, sc.fromInt(900).length);
        //assert.equal(3, sc.fromInt(12000).length);
        //assert.equal(3, sc.fromInt(17576).length);
    });

    it('should return "A" string for integer 1', function(){
        assert.equal("A", sc.fromInt(1));
    });

    it('should return "B" string for integer 2', function(){
        assert.equal("B", sc.fromInt(2));
    });

    it('should return "Y" string for integer 25', function(){
        assert.equal("Y", sc.fromInt(25));
    });

    it('should return "Z" string for integer 26', function(){
        assert.equal("Z", sc.fromInt(26));
    });

    it('should return "AA" string for integer 27', function(){
        assert.equal("AA", sc.fromInt(27));
    });

    it('should return "AB" string for integer 28', function(){
        assert.equal("AB", sc.fromInt(28));
    });

    it('should return "BI" string for integer 61', function(){
        assert.equal("BI", sc.fromInt(61));
    });

    it('should return "ZZ" string for integer 702', function(){
        assert.equal("ZZ", sc.fromInt(702));
    });

    it('should return "AAA" string for integer 703', function(){
        assert.equal("AAA", sc.fromInt(703));
    });

    it('should return "ADO" string for integer 795', function(){
        assert.equal("ADO", sc.fromInt(795));
    });

    it('should return "AEO" string for integer 821', function(){
        assert.equal("AEO", sc.fromInt(821));
    });
});
