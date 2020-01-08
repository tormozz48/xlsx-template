const path = require('path');
const {XLSXPopulateTemplate} = require('../../dist');

(async () => {
    const xlsxPopulateTemplate = new XLSXPopulateTemplate();

    await xlsxPopulateTemplate.loadTemplate(path.join(__dirname,  './template-array.xlsx'));
    xlsxPopulateTemplate.applyData({
        data: [
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
    });

    await xlsxPopulateTemplate.toFile(path.join(__dirname, 'output.xlsx'));
})();