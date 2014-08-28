var xmlbuilder = require('xmlbuilder');
    
var XMLBANNEDCHARS = /[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/;

module.exports = function (options) {
    return new ExcelOfficeXmlWriter(options);
}

function ExcelOfficeXmlWriter (options) {
    var self = this;

    self.options = options;
};

function _isoDateString(d){  
    return d.toISOString().split('Z')[0];
}

ExcelOfficeXmlWriter.prototype.writeDoc = function (obj) {
    var self = this;

    var XMLHDR = { 'version': '1.0'};
    var doc = xmlbuilder.create('ss:Workbook')
    	.att("xmlns", "urn:schemas-microsoft-com:office:spreadsheet")
        .att("xmlns:o", "urn:schemas-microsoft-com:office:office")
        .att("xmlns:x", "urn:schemas-microsoft-com:office:excel")
        .att("xmlns:html", "http://www.w3.org/TR/REC-html140")
        .att("xmlns:ss","urn:schemas-microsoft-com:office:spreadsheet");

    var child = doc;
    var o;

    if (Array.isArray(obj)) {
        o = { Export : obj };
    }
    else {
        o = obj;
    }

    Object.keys(o).forEach(function (sheetTitle) {
        var rows = o[sheetTitle];

        if (!Array.isArray(rows)) {
            rows = [rows];
        }

        //get columns titles based on key's from the first record in the rows array
        var columns = Object.keys(rows[0]);

        child = child.ele("Worksheet").att("ss:Name", sheetTitle).ele("ss:Table");
        columns.forEach(function(columnTitle) {
            child = child.ele("Column").att("ss:AutoFitWidth", "1").up();
        });
        child = child.ele("Row");
        columns.forEach(function(columnTitle){
            child = child.ele("Cell").ele("Data").att("ss:Type", "String").txt(columnTitle).up().up();
        });
        child = child.up();
        rows.forEach(function (record) {
            child = child.ele("Row");
            Object.keys(record).forEach(function (columnTitle) {
                var val = record[columnTitle];

                if (typeof val !== 'function') {
                    if (typeof val === 'object') {
                        if (val instanceof Date) {
                            child = child.ele("Cell").ele("Data").att("ss:Type", "DateTime").raw(_isoDateString(val)).up().up();                    
                        } else {
                            if (val instanceof Array) { }
                        } 
                    } else {
                        if ((typeof val) === 'boolean') {
                        } else if ((typeof val) === 'number') {
                            child = child.ele("Cell").ele("Data").att("ss:Type", "Number").txt(val).up().up();                    
                        } else if (val !== undefined){
                                    //chr = str.match(chars);
                            var str = val.split('\u000b').join(' ');
                            child = child.ele("Cell").ele("Data").att("ss:Type", "String").txt(str).up().up();                    
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

