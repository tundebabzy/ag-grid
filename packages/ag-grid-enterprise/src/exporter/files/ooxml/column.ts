import {ExcelOOXMLTemplate, ExcelColumn} from 'ag-grid-community';

const getExcelCellWidth = (width: number): number => Math.max(Math.ceil((width - 12) / 7 + 1), 10);

const column: ExcelOOXMLTemplate = {
    getTemplate(config: ExcelColumn) {
        const {min, max, s, width, hidden, bestFit} = config;

        return {
            name: 'col',
            properties: {
                rawMap: {
                    min,
                    max,
                    width: !width ? 10 : getExcelCellWidth(width),
                    style: s,
                    hidden: hidden ? '1' : '0',
                    bestFit: bestFit ? '1' : '0',
                    customWidth: width !== 10 ? '1' : '0'
                }
            }
        };
    }
};

export default column;