var path = require("path");
var fs = require("fs");
var assert = require("assert");
var js2o = require("../")({ pretty : true });

var expected = fs.readFileSync(__dirname + '/test.xls', { encoding : 'utf8' });

var doc1 = {
    "sheet1" : [
        { col1 : 'asdf', col2 : 'sdf3', col3 : 'sdf34', date1 : new Date('8/28/2014 12:37:19') }
        , { col1 : 'asdf', col2 : 'sdf3', col3 : 'sdf34', date1 : new Date('8/28/2014 12:38:10') }
    ]
    , "sheet2" : [
        { col1 : 'asdf', col2 : 'sdf3', col3 : 'sdf34' }
        , { col1 : 'asdf', col2 : 'sdf3', col3 : 'sdf34' }
    ]
};


var result = js2o(doc1);
console.log(result);
assert.equal(expected, result);
