# xlsx-rows

Parses an *.xlsx file into rows

### Usage

All you have to do is pass it an Excel (`*.xlsx`) file and you get rows of information:

``` js
  var xlsxRows = require('xlsx-rows');

  var rows = xlsxRows('my-workbook.xlsx');
  console.dir(rows); // yay rows of things!
```

#### Author: [Charlie Robbins](http://github.com/indexzero)
#### License: MIT