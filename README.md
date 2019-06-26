# json2xls-xml

Convert Javascript objects to XLS (XML format), based on json2officexml.

## Installation

```
[~] ➔ npm install json2xls-xml
```

## Usage

```javascript
var j2xls = require('json2xls-xml')({ pretty : true });

var doc = {
    Foo : {
       columns: ['firstname', 'lastname'],
       rows: [
        { firstname : 'John', lastname: 'Doo'},
        { firstname : 'Foo', lastname: 'Bar', age: 23, weight: 25.7876, birth : new Date()}
       ]
    },
    Bar : {
        columns:  ['firstname', 'lastname'],,
        rows: [
           { firstname : 'Rene', lastname: 'Malin'},
           { firstname : 'Foo', lastname: 'foobar', age: 73, weight: 22225.33, birth : new Date()}
        ],
    },
};

console.log(j2xls(doc));
```
