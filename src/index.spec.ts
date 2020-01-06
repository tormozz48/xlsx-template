import 'mocha';
import {expect} from 'chai';

import {XLSXPopulateTemplate} from './index';

describe('xlsx-template', () => {
    describe('single cells', () => {
        let xlsxPopulateTemplate;

        beforeEach(async () => {
            xlsxPopulateTemplate = new XLSXPopulateTemplate();
            await xlsxPopulateTemplate.loadTemplate();
        });

        it('should replace string placeholder', async () => {
            xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('str(data.foo)');
            xlsxPopulateTemplate.applyData({data: {foo: 'bar'}});

            const buffer = await xlsxPopulateTemplate.toBuffer();
            await xlsxPopulateTemplate.loadTemplate(buffer);
            const value = xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value();

            expect(value).to.equal('bar');
        });

        it('should replace multiple string placeholders', async () => {
            xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('str(data.foo1)');
            xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A2').value('str(data.foo2)');
            xlsxPopulateTemplate.applyData({
                data: {
                    foo1: 'bar1',
                    foo2: 'bar2',
                },
            });

            const buffer = await xlsxPopulateTemplate.toBuffer();
            await xlsxPopulateTemplate.loadTemplate(buffer);
            const value1 = xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value();
            const value2 = xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A2').value();

            expect(value1).to.equal('bar1');
            expect(value2).to.equal('bar2');
        });

        it('should leave cell empty if there no appropriate value in data', async () => {
            xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('str(data.foo)');
            xlsxPopulateTemplate.applyData({data: {}});

            const buffer = await xlsxPopulateTemplate.toBuffer();
            await xlsxPopulateTemplate.loadTemplate(buffer);
            const value = xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value();

            expect(value).to.equal(undefined);
        });

        it('should apply date format if date formatter is set', async () => {
            xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('date(data.foo dd/mmm/yyyy)');
            xlsxPopulateTemplate.applyData({data: {foo: new Date(2020, 0, 6)}});

            const buffer = await xlsxPopulateTemplate.toBuffer();
            await xlsxPopulateTemplate.loadTemplate(buffer);
            const cell = xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1');

            expect(cell.value()).to.equal(43836);
            expect(cell.style('numberFormat')).to.equal('dd/mmm/yyyy');
        });

        it('should apply default date format if format was not provided', async () => {
            xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('date(data.foo)');
            xlsxPopulateTemplate.applyData({data: {foo: new Date(2020, 0, 6)}});

            const buffer = await xlsxPopulateTemplate.toBuffer();
            await xlsxPopulateTemplate.loadTemplate(buffer);
            const cell = xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1');

            expect(cell.value()).to.equal(43836);
            expect(cell.style('numberFormat')).to.equal('dd-mm-yyyy');
        });

        it('should number format if number formatter is set', async () => {
            xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('number(data.foo 0.00)');
            xlsxPopulateTemplate.applyData({data: {foo: '3.14159'}});

            const buffer = await xlsxPopulateTemplate.toBuffer();
            await xlsxPopulateTemplate.loadTemplate(buffer);
            const cell = xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1');

            expect(cell.value()).to.equal(3.14159);
            expect(cell.style('numberFormat')).to.equal('0.00');
        });

        it('should set General format if format was not provided', async () => {
            xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('number(data.foo)');
            xlsxPopulateTemplate.applyData({data: {foo: '3.14159'}});

            const buffer = await xlsxPopulateTemplate.toBuffer();
            await xlsxPopulateTemplate.loadTemplate(buffer);
            const cell = xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1');

            expect(cell.value()).to.equal(3.14159);
            expect(cell.style('numberFormat')).to.equal('General');
        });

        it('should format value as link if link formatter is set', async () => {
            xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('link(data.foo)');
            xlsxPopulateTemplate.applyData({data: {foo: {text: 'some-link', ref: 'http://foo.bar'}}});

            const buffer = await xlsxPopulateTemplate.toBuffer();
            await xlsxPopulateTemplate.loadTemplate(buffer);
            const cell = xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1');

            expect(cell.value()).to.equal('some-link');
            expect(cell.hyperlink()).to.equal('http://foo.bar');
        });

        it('should simply expand values from {} placeholders', async () => {
            xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('{str(data.foo)}');
            xlsxPopulateTemplate.applyData({});

            const buffer = await xlsxPopulateTemplate.toBuffer();
            await xlsxPopulateTemplate.loadTemplate(buffer);
            const value = xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value();

            expect(value).to.equal('str(data.foo)');
        });
    });

    describe('cell ranges (iteration)', () => {

    });
});
