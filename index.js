var xmlbuilder = require('xmlbuilder');
    
var XMLBANNEDCHARS = /[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/;

module.exports = function (options) {
    var writer =  new ExcelOfficeXmlWriter(options);

    return function Write (doc) {
        return writer.writeDoc(doc).toString(options);
    }
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
        o = obj || {};
    }

    if (typeof o !== 'object') {
        return child.doc();
    }

    child = child.ele("ss:Styles")
        .ele("ss:Style").att("ss:ID", "DateTime")
            .ele("ss:NumberFormat").att("ss:Format", "mm/dd/yyyy_hh:mm:ss")
            .up()
        .up()
    .up();

    Object.keys(o).forEach(function (sheetTitle) {
        var rows = o[sheetTitle];
        var columns;

        if (!Array.isArray(rows) && rows) {
            rows = [rows];
        }

        if (!rows || !rows.length) {
            return;
        }

        //get columns titles based on key's from the first record in the rows array
        if (rows[0] && typeof rows[0] !== 'object') {
            columns = [sheetTitle];
            //make this an array of objects
            rows = rows.map(function (row) {
                var tmp = {};
                tmp[sheetTitle] = row;
                return tmp;
            });
        }
        else {
            columns = Object.keys(rows[0] || {});
        }

        child = child.ele("ss:Worksheet").att("ss:Name", sheetTitle).ele("ss:Table");
        columns.forEach(function(columnTitle, columnIndex) {
        columnIndex += 1;
            child = child.ele("ss:Column").att("ss:Index",columnIndex).att("ss:AutoFitWidth", "1").up();
        });
        child = child.ele("ss:Row");
        columns.forEach(function(columnTitle){
            child = child.ele("ss:Cell").ele("ss:Data").att("ss:Type", "String").txt(columnTitle).up().up();
        });
        child = child.up();
        rows.forEach(function (record) {
            child = child.ele("ss:Row");
            columns.forEach(function (columnTitle, columnIndex) {
            columnIndex += 1;
                var val = record[columnTitle];

                if (typeof val !== 'function') {
                    if (val && typeof val === 'object') {
                        if (val instanceof Date) {
                            child = child.ele("ss:Cell").att("ss:Index", columnIndex).att("ss:StyleID","DateTime").ele("Data").att("ss:Type", "DateTime").raw(_isoDateString(val)).up().up();
                        } else {
                            if (val instanceof Array) { }
                        } 
                    } else {
                        if ((typeof val) === 'boolean') {
                        } else if ((typeof val) === 'number') {
                            child = child.ele("ss:Cell").att("ss:Index", columnIndex).ele("ss:Data").att("ss:Type", "Number").txt(val).up().up();
                        } else if (val !== undefined && val !== null){
                                    //chr = str.match(chars);
                            var str = val.split('\u000b').join(' ');
                            child = child.ele("ss:Cell").att("ss:Index", columnIndex).ele("ss:Data").att("ss:Type", "String").txt(str).up().up(); 
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

