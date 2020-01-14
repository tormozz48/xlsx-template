import _ from 'lodash';
import XlsxPopulate from 'xlsx-populate2';
import Workbook from 'xlsx-populate2/lib/Workbook';
import Cell from 'xlsx-populate2/lib/Cell';

export class XLSXPopulateTemplate {
    private readonly DEFAULT_LINK_COLOR: string = '0563c1';
    private readonly MATCHER_STR: string = '^str\\((.+)\\)$';
    private readonly MATCHER_RAW: string = '^\\{(.+)\\}$';

    private wb: Workbook;

    public get workbook(): Workbook {
        return this.wb;
    }

    /**
     * Load xlsx workbook from file or buffer and parse it
     * @param {(string|ArrayBuffer|Uint8Array|Buffer|Blob)} template
     * @returns {Promise<void>}
     * @memberOf XLSXPopulateTemplate
     */
    public async loadTemplate(template: string|ArrayBuffer|Uint8Array|Buffer|Blob): Promise<void> {
        if (typeof template === 'undefined') {
            this.wb = await XlsxPopulate.fromBlankAsync();
        } else if (typeof template === 'string') {
            this.wb = await XlsxPopulate.fromFileAsync(template);
        } else {
            this.wb = await XlsxPopulate.fromDataAsync(template);
        }
    }

    /**
     * Apply given data to current workbook
     * @param {Object} data
     * @memberOf XLSXPopulateTemplate
     */
    public applyData(data: any) {
        if (!this.wb) {
            throw new Error('XLSX workbook was not loaded');
        }

        data = data || {};

        this.applyStringCells(data)
            .applyNumberCells(data)
            .applyDateCells(data)
            .applyLinkCells(data)
            .applyRawCells();

        this.applySheetStrTitles(data)
            .applySheetRawTitles();
    }

    /**
     * Serialize xlsx workbook to Buffer
     * @returns {Promise<Buffer>}
     * @memberOf XLSXPopulateTemplate
     */
    public toBuffer(): Promise<Buffer> {
        return this.wb.outputAsync();
    }

    /**
     * Save xlsx workbook to file with given filePath
     * @param {string} filePath
     * @returns {Promise<void>}
     * @memberOf XLSXPopulateTemplate
     */
    public toFile(filePath: string): Promise<void> {
        return this.wb.toFileAsync(filePath);
    }

    /**
     * Finds and replaces placeholders like str(foo) by values of {foo} of given data object
     * @private
     * @param {Object} data
     * @returns {XLSXPopulateTemplate}
     * @memberOf XLSXPopulateTemplate
     */
    private applyStringCells(data: any): XLSXPopulateTemplate {
        const cellMatcher: RegExp = new RegExp(this.MATCHER_STR, 'g');
        const placeholderMatcher: RegExp = new RegExp(this.MATCHER_STR);

        this.wb.find(cellMatcher).forEach((cell: Cell) => {
            const [, placeholder] = cell.value().match(placeholderMatcher);
            this.fillCells({cell, data, placeholder, isLink: false});
        });

        return this;
    }

    /**
     * Finds and replaces placeholders like number(foo) by values of {foo} of given data object
     * Also has advanced format agruments for specifying number format
     * @private
     * @param {Object} data
     * @returns {XLSXPopulateTemplate}
     * @memberOf XLSXPopulateTemplate
     */
    private applyNumberCells(data: any): XLSXPopulateTemplate {
        const REG_EXP: string = '^number\\((\\S+)\\s?(\\S+)?\\)$';
        const cellMatcher: RegExp = new RegExp(REG_EXP, 'g');
        const placeholderMatcher: RegExp = new RegExp(REG_EXP);

        this.wb.find(cellMatcher).forEach((cell: Cell) => {
            const [, placeholder, format = '0'] = cell.value().match(placeholderMatcher);
            this
                .fillCells({cell, data, placeholder, isLink: false})
                .forEach((upCell) => upCell.style('numberFormat', format));
        });

        return this;
    }

    /**
     * Finds and replaces placeholders like date(foo) by values of {foo} of given data object
     * Also has advanced format agruments for specifying date format
     * @private
     * @param {Object} data
     * @returns {XLSXPopulateTemplate}
     * @memberOf XLSXPopulateTemplate
     */
    private applyDateCells(data: any): XLSXPopulateTemplate {
        const REG_EXP: string = '^date\\((\\S+)\\s?([\\S|\\s]+)?\\)$';
        const cellMatcher: RegExp = new RegExp(REG_EXP, 'g');
        const placeholderMatcher: RegExp = new RegExp(REG_EXP);

        this.wb.find(cellMatcher).forEach((cell: Cell) => {
            const [, placeholder, format] = cell.value().match(placeholderMatcher);
            this
                .fillCells({cell, data, placeholder, isLink: false})
                .forEach((upCell) => upCell.style('numberFormat', format || 'dd-mm-yyyy'));
        });

        return this;
    }

    /**
     *
     * @private
     * @param {Object} data
     * @returns {XLSXPopulateTemplate}
     * @memberOf XLSXPopulateTemplate
     */
    private applyLinkCells(data: any): XLSXPopulateTemplate {
        const REG_EXP: string = '^link\\((.+)\\)$';
        const cellMatcher: RegExp = new RegExp(REG_EXP, 'g');
        const placeholderMatcher: RegExp = new RegExp(REG_EXP);

        this.wb.find(cellMatcher).forEach((cell: Cell) => {
            const [, placeholder] = cell.value().match(placeholderMatcher);
            this
                .fillCells({cell, data, placeholder, isLink: true})
                .forEach((upCell) => upCell.style({
                    fontColor: this.DEFAULT_LINK_COLOR,
                    underline: true,
                }));
        });

        return this;
    }

    /**
     * Simply drop outer "{" and "}" and return released placeholders for further processing
     * @private
     * @param {Object} data
     * @returns {XLSXPopulateTemplate}
     * @memberOf XLSXPopulateTemplate
     */
    private applyRawCells(): XLSXPopulateTemplate {
        const cellMatcher: RegExp = new RegExp(this.MATCHER_RAW, 'g');
        const placeholderMatcher: RegExp = new RegExp(this.MATCHER_RAW);

        this.wb.find(cellMatcher, (match: string) => {
            const [, placeholder] = match.match(placeholderMatcher);
            return placeholder;
        });

        return this;
    }

    private fillCells({cell, data, placeholder, isLink = false}): Cell[] {
        const [arrPath, slugPath = ''] = placeholder.split('[i]');
        let dataSet = _.get(data, arrPath, ['']);
        if (!_.isArray(dataSet)) {
            dataSet = [dataSet];
        }

        if (dataSet.length === 0) {
            cell.value('');
        }

        return dataSet.map((item, i: number) => {
            const updatedCell: Cell = cell.relativeCell(i, 0);
            const keyPath: string = `[${i}]${slugPath}`;
            if (isLink) {
                updatedCell.value(_.get(dataSet, `${keyPath}.text`, ''));
                updatedCell.hyperlink(_.get(dataSet, `${keyPath}.ref`, '#'));
            } else {
                updatedCell.value(_.get(dataSet, keyPath, ''));
            }
            return updatedCell;
        });
    }

    private applySheetStrTitles(data) {
        const strMatcher: RegExp = new RegExp(this.MATCHER_STR);
        this.wb.sheets().forEach((sheet) => {
            const sheetName: string = sheet.name();
            const strMatch: string[] = sheetName.match(strMatcher);
            if (strMatch) {
                const [, placeholder] = strMatch;
                sheet.name(_.get(data, placeholder, sheet.name()));
            }
        });

        return this;
    }

    private applySheetRawTitles() {
        const rawMatcher: RegExp = new RegExp(this.MATCHER_RAW);
        this.wb.sheets().forEach((sheet) => {
            const sheetName: string = sheet.name();
            const rawMatch: string[] = sheetName.match(rawMatcher);
            if (rawMatch) {
                const [, placeholder] = rawMatch;
                sheet.name(placeholder);
            }
        });

        return this;
    }
}
