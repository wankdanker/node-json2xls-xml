
# json2officexml v0.0.5.1

Minimalist JSON to XLS (Excel Office XML) writer written by Pierre Metrailler <pierre@shockfish.com>. This fork contains minor changes in functionality and in the overall code format.

## Installation

```bash
[~] ➔ npm install https://github.com/tobius/node-json2officexml/tarball/master
```

## Usage

```javascript
var json2oxml = require('json2officexml');
var writer = json2oxml.createExcelOfficeXmlWriter();
var doc = {
    'sheets': [
        {
            name: 'Foo',
            columns: ['first', 'last', 'age', 'weight', 'birth'],
            rows: [
                { firstname : 'John', lastname: 'Doo'},
                { firstname : 'Foo', lastname: 'Bar', age: 23, weight: 25.7876, birth : new Date()}
            ]
        },
        {
            name: 'Bar',
            columns: ['first', 'last', 'age', 'weight', 'birth'],
            rows: [
                { firstname : 'Rene', lastname: 'Malin'},
                { firstname : 'Foo', lastname: 'foobar', age: 73, weight: 22225.33, birth : new Date()}
            ]
        }
    ]
};
console.log(writer.writeDoc(doc).toString({ pretty: true }));
```
