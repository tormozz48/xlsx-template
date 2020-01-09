# xlsx-template
Fill workbook by given data into places marked by special placeholders.

[![Build Status](https://travis-ci.org/tormozz48/xlsx-template.svg?branch=master)](https://travis-ci.org/tormozz48/xlsx-template)
[![codecov](https://codecov.io/gh/tormozz48/xlsx-template/branch/master/graph/badge.svg)](https://codecov.io/gh/tormozz48/xlsx-template)

## Install

Install package via npm:
```
npm install @tormozz48/xlsx-template
```

## Usage Example
Assume there `./template-simple.xlsx` XLSX file in your working directory with given content:

| # | A | B | C | D | E |
| --- | --- | --- | --- | --- | --- |
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
        dateVal: new Date(2020, 0, 6),
        linkVal: {text: 'some link', ref: 'http://github.com'}
    }
});
await xlsxPopulateTemplate.toFile('./output.xlsx');
```

Code above will create new file `./output.xlsx` with replacing all placeholders from `./template-simple.xlsx` on provided data.
It also will set valid format to cells according to used placeholders.

| # | A | B | C | D | E |
| --- | --- | --- | --- | --- | --- |
| 1 | Some String value | 3,14 | 06/01/2020 | [some link](http://github.com) | str(data.strVal) |
| 2 |   | 3,14159 | 06-01-2020 |   |   |

### Apply values given by arrays.

It is possible to fill range of rows by values given as array. For example array of objects such as:
```js
[
    {
        strVal: 'Github',
        numberVal: 134.1222,
        dateVal: new Date(2020, 1, 1),
        linkVal: {text: 'github', ref: 'https://github.com'}
    },
    {
        strVal: 'Facebook',
        numberVal: 4352.232,
        dateVal: new Date(2020, 2, 23),
        linkVal: {text: 'facebook', ref: 'https://facebook.com'}
    },
    {
        strVal: 'Google',
        numberVal: 733.2122321,
        dateVal: new Date(2020, 3, 7),
        linkVal: {text: 'google', ref: 'https://google.com'}
    }
]
```

applied to following template:

| # | A | B | C | D |
| --- | --- | --- | --- | --- |
| 1 | str(data[i].strVal) | number(data[i].numberVal 0.00) | date(data[i].dateVal dd/mm/yyyy) | link(data[i].linkVal) |

will fill first three rows by corresponded values from data array:

| # | A | B | C | D |
| --- | --- | --- | --- | --- |
| 1 | Github | 134,12 |	01/02/2020 | [github](https://github.com) |
| 2 | Facebook | 4352,23 |	23/03/2020 | [facebook](https://facebook.com) |
| 3 | Google |	733,21 | 07/04/2020 | [google](https://google.com) |

## Placeholders

### str()

Paste given value to a cell marked with placeholder and use default cell string formatter.

| # | A |
| --- | --- |
| 1 | str(data.foo) |

```js
xlsxPopulateTemplate.applyData({data: {foo: 'Hello World'}});
```

| # | A |
| --- | --- |
| 1 | Hello World |

### number()

Paste given value to a cell and use number formatter. Optionally use second argument of number placeholder function to specify
number format which should be applied to cell value.

| # | A |
| --- | --- |
| 1 | number(data.foo, 0.00) |

```js
xlsxPopulateTemplate.applyData({data: {foo: 3.14159}});
```

| # | A |
| --- | --- |
| 1 | 3.14 |

### date()

Paste given value to a cell and use date formatter. Optionally use second argument of date placeholder function to specify
date format which should be applied to cell value. Default date format is `dd-mm-yyyy`.

| # | A |
| --- | --- |
| 1 | date(data.foo dd/mm/yyyy) |

```js
xlsxPopulateTemplate.applyData({data: {foo: new Date(2020, 0, 9)}});
```

| # | A |
| --- | --- |
| 1 | 09/01/2020 |

### link()

Create link value in cell with placeholder. Needs to receive data item as object with fields: `text` and `ref` where `text` - is
link text representation and `ref` - is link reference url.

| # | A |
| --- | --- |
| 1 | link(data.foo) |

```js
xlsxPopulateTemplate.applyData({data: {foo: {text: 'Github', ref: 'https://github.com'}}});
```

| # | A |
| --- | --- |
| 1 | [Github](https://github.com) |


### {} - raw formatter

A special formatter which allows to simply expand inner placeholder without appliyng it. It is useful when
there you need to fill template with some part of needed data at first stage and then fill rest of data at second stage.

| # | A |
| --- | --- |
| 1 | {str(data.foo)} |

```js
xlsxPopulateTemplate.applyData({});
```
After first call it simply expand inner placeholder `str(data.foo)` so it will be ready to use on next call.

| # | A |
| --- | --- |
| 1 | str(data.foo) |

```js
xlsxPopulateTemplate.applyData({data: {foo: 'Hello World'}});
```

| # | A |
| --- | --- |
| 1 | Hello World |

### array item placeholders

To fill multiple rows below by corresponded array items you should use `[i]` to mark data node
that should be applied as array of items.

| # | A |
| --- | --- |
| 1 | str(data[i].name) |

```js
xlsxPopulateTemplate.applyData({data: [
    {name: 'Github'},
    {name: 'Facebook'},
    {name: 'Google'}
]);
```

| # | A |
| --- | --- |
| 1 | Github |
| 2 | Facebook |
| 3 | Google |

## API

#### .loadTemplate(templatePath)

Async function to load XLSX workbook from buffer of file or event create it from scratch.
Loads XLSX workbook from file:
```js
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate('./template-simple.xlsx');
```

Loads XLSX workbook from buffer:
```js
const buffer = await fs.readFile('./template-simple.xlsx');
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate(buffer);
```

Creates XLSX workbook from scratch (empty workbook)
```js
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate();
```

#### .applyData(data)

Fills matched template placeholders with given data.
```js
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate();
xlsxPopulateTemplate.applyData({foo: 'bar'})
```

#### .toBuffer()

Async function which serializes current workbook into buffer.

```js
const xlsxPopulateTemplate = new XLSXPopulateTemplate();
await xlsxPopulateTemplate.loadTemplate();
const buffer = await xlsxPopulateTemplate.toBuffer();
```

#### .toFile(filePath)

Async function which saves XLSX workbook into file with given filePath on local filesystem.

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

Some useful commands for development:

* `npm run build` - compiles typescript code into javascript code.
* `npm run clean` - cleans js dist folder with compiled javascript code.
* `npm run format` - performs code formatting via prettier tool.
* `npm run lint` - runs tslint syntax checker.
* `npm run test` - runs mocha tests.
* `npm run test -watch` - runs mocha tests in "watch" mode. Launch tests on every code change.
* `npm run test -cov` - runs mocha tests with coverage calculation.
