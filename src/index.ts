import _ from 'lodash';
import XlsxPopulate from 'xlsx-populate2';
import Workbook from 'xlsx-populate2/lib/Workbook';

export class XLSXPopulateTemplate {
    private wb: Workbook;

    public get workbook(): Workbook {
        return this.wb;
    }

    public async loadTemplate(template: string|ArrayBuffer|Uint8Array|Buffer|Blob): Promise<void> {
        if (typeof template === 'undefined') {
            this.wb = await XlsxPopulate.fromBlankAsync();
        } else if (typeof template === 'string') {
            this.wb = await XlsxPopulate.fromFileAsync(template);
        } else {
            this.wb = await XlsxPopulate.fromDataAsync(template);
        }
    }

    public applyData(data: any) {
        data = data || {};

        this.applyStringCells(data);
        this.applyNumberCells(data);
        this.applyDateCells(data);
        this.applyLinkCells(data);

        this.applyRawCells();
    }

    public toBuffer(): Promise<Buffer> {
        return this.wb.outputAsync();
    }

    public toFile(filePath: string): Promise<void> {
        return this.wb.toFileAsync(filePath);
    }

    private applyStringCells(data: any): XLSXPopulateTemplate {
        const REG_EXP = '^str\\((.+)\\)$';
        const cellMatcher = new RegExp(REG_EXP, 'g');
        const placeholderMatcher = new RegExp(REG_EXP);

        this.wb.find(cellMatcher, (match) => {
            const [, placeholder] = match.match(placeholderMatcher);
            return _.get(data, placeholder, '');
        });

        return this;
    }

    private applyNumberCells(data: any): XLSXPopulateTemplate {
        const REG_EXP = '^number\\((\\S+)\\s?(\\S+)?\\)$';
        const cellMatcher = new RegExp(REG_EXP, 'g');
        const placeholderMatcher = new RegExp(REG_EXP);

        this.wb.find(cellMatcher).forEach((cell) => {
            const [, placeholder, format] = cell.value().match(placeholderMatcher);
            cell.value(_.get(data, placeholder, ''));
            if (format) {
                cell.style('numberFormat', format);
            }
        });

        return this;
    }

    private applyDateCells(data: any): void {
        const REG_EXP = '^date\\((\\S+)\\s?(\\S+)?\\)$';
        const cellMatcher = new RegExp(REG_EXP, 'g');
        const placeholderMatcher = new RegExp(REG_EXP);

        this.wb.find(cellMatcher).forEach((cell) => {
            const [, placeholder, format] = cell.value().match(placeholderMatcher);
            cell.value(_.get(data, placeholder, ''));
            cell.style('numberFormat', format || 'dd-mm-yyyy');
        });
    }

    private applyLinkCells(data: any): void {
        const REG_EXP = '^link\\((.+)\\)$';
        const cellMatcher = new RegExp(REG_EXP, 'g');
        const placeholderMatcher = new RegExp(REG_EXP);

        this.wb.find(cellMatcher).forEach((cell) => {
            const [, placeholder] = cell.value().match(placeholderMatcher);
            cell.value(_.get(data, `${placeholder}.text`, ''));
            cell.hyperlink(_.get(data, `${placeholder}.ref`));
        });
    }

    private applyRawCells(): void {
        const REG_EXP = '^\\{(.+)\\}$';
        const cellMatcher = new RegExp(REG_EXP, 'g');
        const placeholderMatcher = new RegExp(REG_EXP);

        this.wb.find(cellMatcher, (match) => {
            const [, placeholder] = match.match(placeholderMatcher);
            return placeholder;
        });
    }
}
