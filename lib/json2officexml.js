var xmlbuilder = require('xmlbuilder');
    
var j2o = exports;

var XMLBANNEDCHARS = /[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/;


var ExcelOfficeXmlWriter = j2o.ExcelOfficeXmlWriter = function(options) {

};

ExcelOfficeXmlWriter.prototype.writeDoc = function(doc) {
    if (!doc) return;
    return _writeExcelDoc(this, doc);
};

j2o.createExcelOfficeXmlWriter = function(path, options) {
    return new ExcelOfficeXmlWriter(options);
};

function _isoDateString(d){  
    function pad(n){return n<10 ? '0'+n : n}  
    return d.getUTCFullYear()+'-'
    + pad(d.getUTCMonth()+1)+'-'  
    + pad(d.getUTCDate())+'T'  
    + pad(d.getUTCHours())+':'  
    + pad(d.getUTCMinutes())+':'  
    + pad(d.getUTCSeconds())+".000"  ;//'Z'  
}

function _writeExcelDoc(writer, inobj) {
    var XMLHDR = { 'version': '1.0'};
    var doc = xmlbuilder.create('ss:Workbook').att("xmlns:ss","urn:schemas-microsoft-com:office:spreadsheet");
    var child = doc;//doc.ele('ss:Workbook', XMLHDR).att("xmlns:ss","urn:schemas-microsoft-com:office:spreadsheet");
    var o;

    if (Array.isArray(inobj)) {
        o = { Export : inobj };
    }
    else {
        o = inobj;
    }

    Object.keys(o).forEach(function (sheetTitle) {
        var rows = o[sheetTitle];

        //get columns titles based on key's from the first record in the rows array
        var columns = Object.keys(rows[0]);

        child = child.ele("ss:Worksheet").att("ss:Name", sheetTitle).ele("ss:Table");
        columns.forEach(function(columnTitle) {
            child = child.ele("ss:Column").att("ss:AutoFitWidth", "1").up();
        });
        child = child.ele("ss:Row");
        columns.forEach(function(columnTitle){
            child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "String").txt(columnTitle).up().up();
        });
        child = child.up();
        rows.forEach(function (record) {
            child = child.ele("ss:Row");
            Object.keys(record).forEach(function (columnTitle) {
                var val = record[columnTitle];

                if (typeof val !== 'function') {
                    if (typeof val === 'object') {
                        if (val instanceof Date) {
                            child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "DateTime").raw(_isoDateString((val))).up().up();                    
                        } else {
                            if (val instanceof Array) { }
                        } 
                    } else {
                        if ((typeof val) === 'boolean') {
                        } else if ((typeof val) === 'number') {
                            child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "Number").txt(val).up().up();                    
                        } else if (val !== undefined){
                          		    //chr = str.match(chars);
    			            var str = val.split('\u000b').join(' ');
    			            child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "String").txt(str).up().up();                    
                        }
                    }
                }
            });

            child = child.up();
        }); // rows.forEach
        child = child.up().up();
    }); // sheets.forEach
    return child.doc();
}

