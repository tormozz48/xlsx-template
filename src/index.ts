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

    private applyStringCells(data: any): void {
        this.wb.find(/^str\((.+)\)$/g, (match) => {
            const [, placeholder] = match.match(/^str\((.+)\)$/);
            return _.get(data, placeholder, '');
        });
    }

    private applyNumberCells(data: any): void {
        const cells = this.wb.find(/^number\((.+)\)$/g);

        cells.forEach((cell) => {
            const [, placeholder, format] = cell.value().match(/^number\((\S+)\s?(\S+)?\)$/);
            cell.value(_.get(data, placeholder, ''));
            if (format) {
                cell.style('numberFormat', format);
            }
        });
    }

    private applyDateCells(data: any): void {
        const cells = this.wb.find(/^date\((.+)\)$/g);

        cells.forEach((cell) => {
            const [, placeholder, format = 'dd-mm-yyyy'] = cell.value().match(/^date\((\S+)\s?(\S+)?\)$/);
            cell.value(_.get(data, placeholder, ''));
            cell.style('numberFormat', format);
        });
    }

    private applyLinkCells(data: any): void {
        const cells = this.wb.find(/^link\((.+)\)$/g);

        cells.forEach((cell) => {
            const [, placeholder] = cell.value().match(/^link\((.+)\)$/);
            cell.value(_.get(data, `${placeholder}.text`, ''));
            cell.hyperlink(_.get(data, `${placeholder}.ref`));
        });
    }

    private applyRawCells(): void {
        this.wb.find(/^\{(.+)\}$/g, (match) => {
            const [, placeholder] = match.match(/^\{(.+)\}$/);
            return placeholder;
        });
    }
}
