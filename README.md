# json2xls-xml

Convert Javascript objects to XLS (XML format), based on json2officexml.

## Installation

```
[~] ➔ npm install json2xls-xml
```

## Usage

```javascript
var writer = require('json2xls-xml')();

var doc = {
    Foo : [
        { firstname : 'John', lastname: 'Doo'}
        , { firstname : 'Foo', lastname: 'Bar', age: 23, weight: 25.7876, birth : new Date()}
    ]
    , Bar : [
        { firstname : 'Rene', lastname: 'Malin'}
        , { firstname : 'Foo', lastname: 'foobar', age: 73, weight: 22225.33, birth : new Date()}
    ]
};

console.log(writer.writeDoc(doc).toString({ pretty: true }));
```
