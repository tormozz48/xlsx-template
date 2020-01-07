# xlsx-template
Fill workbook by given data into places marked by special placeholders.

[![Build Status](https://travis-ci.org/tormozz48/xlsx-template.svg?branch=master)](https://travis-ci.org/tormozz48/xlsx-template)

## Install

Install package via npm:
```
npm install xlsx-template
```

## Usage Example
Assume there file `./template-simple.xlsx` xlsx file in your working directory with given content:
| # | A | B | C | D | E |
| - | - | - | - | - | - |
| 1 |   |   |   |   |   |
| 2 |   |   |   |   |   |

```js
const {XLSXPopulateTemplate} = require('../../dist');
const xlsxPopulateTemplate = new XLSXPopulateTemplate();

await xlsxPopulateTemplate.loadTemplate('./template-simple.xlsx');
xlsxPopulateTemplate.applyData({
    data: {
        strVal: 'Some String value',
        numberVal: 3.14159,
        dateVal: new Date(2019, 0, 6),
        linkVal: {text: 'some link', ref: 'http://github.com'}
    }
});
await xlsxPopulateTemplate.toFile('./output.xlsx');
```

Code above will create new file `./output.xlsx` and replace all placeholders from `./template-simple.xlsx` with provided data.
It also will set valid format to cells according to used placeholders.

| # | A | B | C | D | E |
| - | - | - | - | - | - |
| 1 | Some String value  | 3,14  |  06/01/2019 | [some link](http://github.com)  | str(data.strVal) |
| 2 |   | 3,14159  | 06-01-2019 |   |   |


## Placeholders

## Development
