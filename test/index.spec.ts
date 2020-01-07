import 'mocha';
import {expect} from 'chai';

import {XLSXPopulateTemplate} from '../src/index';

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

        it('should set default link ref "#" if link reference was not set', async () => {
            xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('link(data.foo)');
            xlsxPopulateTemplate.applyData({data: {foo: {text: 'some-link'}}});

            const buffer = await xlsxPopulateTemplate.toBuffer();
            await xlsxPopulateTemplate.loadTemplate(buffer);
            const cell = xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1');

            expect(cell.value()).to.equal('some-link');
            expect(cell.hyperlink()).to.equal('#');
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
        let xlsxPopulateTemplate;

        beforeEach(async () => {
            xlsxPopulateTemplate = new XLSXPopulateTemplate();
            await xlsxPopulateTemplate.loadTemplate();
        });

        describe('text cells', () => {
            it('should create filled cell range by array of simple values', async () => {
                xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('str(data.foo[i])');
                xlsxPopulateTemplate.applyData({
                    data: {
                        foo: ['A1 value', 'A2 value', 'A3 value'],
                    },
                });

                const buffer = await xlsxPopulateTemplate.toBuffer();
                await xlsxPopulateTemplate.loadTemplate(buffer);
                const sheet = xlsxPopulateTemplate.workbook.sheet('Sheet1');

                expect(sheet.cell('A1').value()).to.equal('A1 value');
                expect(sheet.cell('A2').value()).to.equal('A2 value');
                expect(sheet.cell('A3').value()).to.equal('A3 value');
            });

            it('should create filled cell range by array of objects', async () => {
                xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('str(data[i].foo)');
                xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('B1').value('str(data[i].bar)');
                xlsxPopulateTemplate.applyData({
                    data: [
                        {foo: 'foo1', bar: 'bar1'},
                        {foo: 'foo2', bar: 'bar2'},
                    ],
                });

                const buffer = await xlsxPopulateTemplate.toBuffer();
                await xlsxPopulateTemplate.loadTemplate(buffer);
                const sheet = xlsxPopulateTemplate.workbook.sheet('Sheet1');

                expect(sheet.cell('A1').value()).to.equal('foo1');
                expect(sheet.cell('A2').value()).to.equal('foo2');
                expect(sheet.cell('B1').value()).to.equal('bar1');
                expect(sheet.cell('B2').value()).to.equal('bar2');
            });
        });

        describe('number cells', () => {
            it('should create filled cell range by array of simple values', async () => {
                xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('number(data.foo[i] 0.00)');
                xlsxPopulateTemplate.applyData({
                    data: {
                        foo: [123.45, 234.56],
                    },
                });

                const buffer = await xlsxPopulateTemplate.toBuffer();
                await xlsxPopulateTemplate.loadTemplate(buffer);
                const sheet = xlsxPopulateTemplate.workbook.sheet('Sheet1');

                expect(sheet.cell('A1').value()).to.equal(123.45);
                expect(sheet.cell('A2').value()).to.equal(234.56);
                expect(sheet.cell('A1').style('numberFormat')).to.equal('0.00');
                expect(sheet.cell('A2').style('numberFormat')).to.equal('0.00');
            });

            it('should create filled cell range by array of objects', async () => {
                xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('number(data[i].foo)');
                xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('B1').value('number(data[i].bar)');
                xlsxPopulateTemplate.applyData({
                    data: [
                        {foo: 11.789, bar: 99.1234},
                        {foo: 12.987, bar: 98.4321},
                    ],
                });

                const buffer = await xlsxPopulateTemplate.toBuffer();
                await xlsxPopulateTemplate.loadTemplate(buffer);
                const sheet = xlsxPopulateTemplate.workbook.sheet('Sheet1');

                expect(sheet.cell('A1').value()).to.equal(11.789);
                expect(sheet.cell('A2').value()).to.equal(12.987);
                expect(sheet.cell('B1').value()).to.equal(99.1234);
                expect(sheet.cell('B2').value()).to.equal(98.4321);

                ['A1', 'A2', 'B1', 'B2'].forEach((item) => {
                    expect(sheet.cell(item).style('numberFormat')).to.equal('General');
                });
            });
        });

        describe('date cells', () => {
            it('should create filled cell range by array of simple values', async () => {
                xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('date(data.foo[i] dd/mm/yyyy)');
                xlsxPopulateTemplate.applyData({
                    data: {
                        foo: [
                            new Date(2020, 0, 7),
                            new Date(2020, 1, 8),
                        ],
                    },
                });

                const buffer = await xlsxPopulateTemplate.toBuffer();
                await xlsxPopulateTemplate.loadTemplate(buffer);
                const sheet = xlsxPopulateTemplate.workbook.sheet('Sheet1');

                expect(sheet.cell('A1').value()).to.equal(43837);
                expect(sheet.cell('A2').value()).to.equal(43869);
                expect(sheet.cell('A1').style('numberFormat')).to.equal('dd/mm/yyyy');
                expect(sheet.cell('A2').style('numberFormat')).to.equal('dd/mm/yyyy');
            });

            it('should create filled cell range by array of objects', async () => {
                xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('date(data[i].foo)');
                xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('B1').value('date(data[i].bar)');
                xlsxPopulateTemplate.applyData({
                    data: [
                        {foo: new Date(2020, 0, 7), bar: new Date(2020, 4, 12)},
                        {foo: new Date(2020, 1, 8), bar: new Date(2020, 5, 24)},
                    ],
                });

                const buffer = await xlsxPopulateTemplate.toBuffer();
                await xlsxPopulateTemplate.loadTemplate(buffer);
                const sheet = xlsxPopulateTemplate.workbook.sheet('Sheet1');

                expect(sheet.cell('A1').value()).to.equal(43837);
                expect(sheet.cell('A2').value()).to.equal(43869);
                expect(sheet.cell('B1').value()).to.equal(43963);
                expect(sheet.cell('B2').value()).to.equal(44006);

                ['A1', 'A2', 'B1', 'B2'].forEach((item) => {
                    expect(sheet.cell(item).style('numberFormat')).to.equal('dd-mm-yyyy');
                });
            });
        });

        describe('link cells', () => {
            it('should create filled cell range by array of objects', async () => {
                xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('A1').value('link(data[i].foo)');
                xlsxPopulateTemplate.workbook.sheet('Sheet1').cell('B1').value('link(data[i].bar)');
                xlsxPopulateTemplate.applyData({
                    data: [
                        {
                            foo: {text: 'some-link11', ref: 'http://link11'},
                            bar: {text: 'some-link12', ref: 'http://link12'},
                        },
                        {
                            foo: {text: 'some-link21', ref: 'http://link21'},
                            bar: {text: 'some-link22', ref: 'http://link22'},
                        },
                    ],
                });

                const buffer = await xlsxPopulateTemplate.toBuffer();
                await xlsxPopulateTemplate.loadTemplate(buffer);
                const sheet = xlsxPopulateTemplate.workbook.sheet('Sheet1');

                expect(sheet.cell('A1').value()).to.equal('some-link11');
                expect(sheet.cell('A1').hyperlink()).to.equal('http://link11');
                expect(sheet.cell('A2').value()).to.equal('some-link21');
                expect(sheet.cell('A2').hyperlink()).to.equal('http://link21');

                expect(sheet.cell('B1').value()).to.equal('some-link12');
                expect(sheet.cell('B1').hyperlink()).to.equal('http://link12');
                expect(sheet.cell('B2').value()).to.equal('some-link22');
                expect(sheet.cell('B2').hyperlink()).to.equal('http://link22');
            });
        });
    });
});
