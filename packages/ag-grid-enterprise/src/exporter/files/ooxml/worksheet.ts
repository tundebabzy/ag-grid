import {ExcelOOXMLTemplate, ExcelWorksheet, ExcelRow, ExcelCell, _} from 'ag-grid-community';
import column from './column';
import row from './row';
import mergeCell from './mergeCell';

const getMergedCells = (rows: ExcelRow[]): string[] => {
    const mergedCells: string[] = [];

    rows.forEach((currentRow, rowIdx) => {
        const cells = currentRow.cells;
        let merges = 0;
        cells.forEach((currentCell, cellIdx) => {
            const start = getExcelColumnName(cellIdx + merges + 1);
            const outputRow = rowIdx + 1;

            if (currentCell.mergeAcross) {
                merges += currentCell.mergeAcross;
                const end = getExcelColumnName(cellIdx + merges + 1);

                mergedCells.push(`${start}${outputRow}:${end}${outputRow}`);
            }

            currentCell.ref = `${start}${outputRow}`;
        });
    });

    return mergedCells;
};

const getExcelColumnName = (colIdx: number): string => {
    const startCode = 65;
    const tableWidth = 26;
    const fromCharCode = String.fromCharCode;

    const pos = Math.floor(colIdx / tableWidth);
    const tableIdx = colIdx % tableWidth;

    if (!pos || colIdx === tableWidth) return fromCharCode(startCode + colIdx - 1);
    if (!tableIdx) return getExcelColumnName(pos - 1) + 'Z';
    if (pos < tableWidth) return fromCharCode(startCode + pos - 1) + fromCharCode(startCode + tableIdx - 1);

    return getExcelColumnName(pos) + fromCharCode(startCode + tableIdx - 1);
};

const worksheet: ExcelOOXMLTemplate = {
    getTemplate(config: ExcelWorksheet) {
        const {table} = config;
        const {rows, columns} = table;

        const mergedCells = getMergedCells(rows);

        const children = [].concat(
            columns.length ? {
                name: 'cols',
                children: _.map(columns, column.getTemplate)
            } : []
        ).concat(
            rows.length ? {
                name: 'sheetData',
                children: _.map(rows, row.getTemplate)
            } : []
        ).concat(
            mergedCells.length ? {
                name: 'mergeCells',
                properties: {
                    rawMap: {
                        count: mergedCells.length
                    }
                },
                children: _.map(mergedCells, mergeCell.getTemplate)
            } : []
        );

        return {
            name: "worksheet",
            properties: {
                prefixedAttributes:[{
                    prefix: "xmlns:",
                    map: {
                        r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    }
                }],
                rawMap: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                }
            },
            children
        };
    }
};

export default worksheet;