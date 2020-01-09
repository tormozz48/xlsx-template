const path = require('path');
const _ = require('lodash');
const {XLSXPopulateTemplate} = require('../../dist');

(async () => {
    const xlsxPopulateTemplate = new XLSXPopulateTemplate();

    await xlsxPopulateTemplate.loadTemplate(path.join(__dirname,  './template-array.xlsx'));
    const data = _.fill(Array(5000), 0).map((item, index) => ({
        strVal: `Item: ${index}`,
        numberVal: index,
        dateVal: new Date(2020, 1, 1),
        linkVal: {text: `github ${index}`, ref: `https://github.com/${index}`}
    }));

    console.time('start');
    xlsxPopulateTemplate.applyData({data});
    console.timeEnd('start');

    await xlsxPopulateTemplate.toFile(path.join(__dirname, 'output.xlsx'));
})();