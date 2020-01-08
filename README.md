# xlsx-template
Fill workbook by given data into places marked by special placeholders.

[![Build Status](https://travis-ci.org/tormozz48/xlsx-template.svg?branch=master)](https://travis-ci.org/tormozz48/xlsx-template)

## Install

Install package via npm:
```
npm install xlsx-template
```

## Usage Example
Assume there file `./template-simple.xlsx` XLSX file in your working directory with given content:

| # | A | B | C | D | E |
| - | - | - | - | - | - |
| 1 | str(data.strVal) | number(data.numberVal 0.00) | date(data.dateVal dd/mm/yyyy) | link(data.linkVal) | {str(data.strVal)} |
| 2 |   | number(data.numberVal) | date(data.dateVal) | | |

```js
const {XLSXPopulateTemplate} = require('xlsx-template');
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

### str()

Paste given value to cell marked with placeholder and use default cell string formatter.

| # | A |
| - | - |
| 1 | str(data.foo) |

```js
xlsxPopulateTemplate.applyData({data: {foo: 'Hello World'}})
```

| # | A |
| - | - |
| 1 | Hello World |

### number()

Paste given value to cell and use number formatter. Optionally use second argument of number placeholder function to specify
number format which should be applied to cell value.

| # | A |
| - | - |
| 1 | number(data.foo, 0.00) |

```js
xlsxPopulateTemplate.applyData({data: {foo: 3.14159}})
```

| # | A |
| - | - |
| 1 | 3.14 |

### date()

### link()

### {} - raw formatter

## API

#### .loadTemplate(templatePath)

Async function to load XLSX workbook from buffer of file or event create it from scratch.
Load XLSX workbook from file:
```js
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate('./template-simple.xlsx');
```

Load XLSX workbook from buffer:
```js
const buffer = await fs.readFile('./template-simple.xlsx');
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate(buffer);
```

Create XLSX workbook from scratch (empty workbook)
```js
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate();
```

#### .applyData(data)

Fill matched template placeholders with given data.
```js
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate();
xlsxPopulateTemplate.applyData({foo: 'bar'})
```

#### .toBuffer()

Serializes current workbook into buffer.

```js
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate();
const buffer = await xlsxPopulateTemplate.toBuffer();
```

#### .toFile(filePath)

Saves XLSX workbook into file with given filePath on local filesystem.

```js
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate();
await xlsxPopulateTemplate.toFile('./my-awesome.xlsx');
```

#### .workbook

Getter which simply returns XLSX workbook instance of [xlsx-populate2](https://www.npmjs.com/package/xlsx-populate2) package.
It is usefull for direct manipulation with workbook or it inner parts like cells, rows and columns.

```js
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate();

const workbook = xlsxPopulateTemplate.workbook;
workbook.sheet('Sheet1').cell('A1').value('Foo');
```

## Development
