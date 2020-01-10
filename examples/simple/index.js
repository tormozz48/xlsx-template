const path = require('path');
const {XLSXPopulateTemplate} = require('../../dist');

(async () => {
    const xlsxPopulateTemplate = new XLSXPopulateTemplate();

    await xlsxPopulateTemplate.loadTemplate(path.join(__dirname,  './template-simple.xlsx'));
    xlsxPopulateTemplate.applyData({
        data: {
            strVal: 'Some String value',
            numberVal: 3.14159,
            dateVal: new Date(2019, 0, 6, 16, 30),
            linkVal: {text: 'some link', ref: 'http://github.com'}
        }
    });

    await xlsxPopulateTemplate.toFile(path.join(__dirname, 'output.xlsx'));
})();