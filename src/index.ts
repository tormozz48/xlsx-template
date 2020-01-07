import _ from 'lodash';
import XlsxPopulate from 'xlsx-populate2';
import Workbook from 'xlsx-populate2/lib/Workbook';
import Cell from 'xlsx-populate2/lib/Cell';

export class XLSXPopulateTemplate {
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
        data = data || {};

        this.applyStringCells(data)
            .applyNumberCells(data)
            .applyDateCells(data)
            .applyLinkCells(data)
            .applyRawCells();
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
        const REG_EXP: string = '^str\\((.+)\\)$';
        const cellMatcher: RegExp = new RegExp(REG_EXP, 'g');
        const placeholderMatcher: RegExp = new RegExp(REG_EXP);

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
            const [, placeholder, format] = cell.value().match(placeholderMatcher);
            const updatedCells = this.fillCells({cell, data, placeholder, isLink: false});
            if (format) {
                updatedCells.forEach((upCell) => upCell.style('numberFormat', format));
            }
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
        const REG_EXP: string = '^date\\((\\S+)\\s?(\\S+)?\\)$';
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
            this.fillCells({cell, data, placeholder, isLink: true});
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
        const REG_EXP: string = '^\\{(.+)\\}$';
        const cellMatcher: RegExp = new RegExp(REG_EXP, 'g');
        const placeholderMatcher: RegExp = new RegExp(REG_EXP);

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
}
