import React, { useState } from "react";
import * as XLSX from 'xlsx';

function App() {
    const [inputJson, setInputJson] = useState([]);
    const [jsonLoad, setJSONLoad] = useState(false);

    const columnHeaders = [
        "COMPANY NAME",
        "TICKER",
        "PRIMARY ISIN",
        "COUNTRY",
        "MEETING DATE",
        "RECORD DATE",
        "MEETING TYPE",
        "PROPONENT",
        "PROPOSAL NUMBER",
        "PROPOSAL TEXT",
        "MANAGEMENT RECOMMENDATION",
        "VOTE INSTRUCTION",
        "GOLDMAN SACHS ASSET MANAGEMENT RATIONALE"
    ];

    // Normalize header row for comparison
    const normalizeText = (text) => text.toString().trim().replace(/\r\n|\r|\n/g, ' ').replace(/\s+/g, ' ');
    const normalizeRow = (row) => row.map(cell => normalizeText(cell));

    const mergeRows = (rows) => {
        const mergedRows = [];
        let currentRow = [];
    
        // Columns to check if they are empty
        const emptyCheckColumns = [2, 4, 5, 10, 11];
    
        // Function to check if a row is empty in specified columns
        const isRowEmptyInColumns = (row) => {
            return emptyCheckColumns.some(index => row[index] === '');
        };
    
        // Function to check if a row is completely empty
        const isRowCompletelyEmpty = (row) => {
            return row.every(cell => cell === '' || cell === undefined || cell === null);
        };
    
        rows.forEach(row => {
            if (isRowCompletelyEmpty(row)) {
                // Skip the row if it is completely empty
                return;
            }
    
            if (isRowEmptyInColumns(row)) {
                // Merge cells with the previous row if any of the specified columns are empty
                currentRow = currentRow.map((cell, index) => {
                    if (index < row.length) {
                        return (cell || '') + ' ' + (row[index] || '');
                    }
                    return cell;
                });
            } else {
                // Push the current row to mergedRows if it's not empty
                if (currentRow.length > 0) {
                    mergedRows.push(currentRow);
                }
                currentRow = row;
            }
        });
    
        // Push the last row if it's not empty
        if (currentRow.length > 0) {
            mergedRows.push(currentRow);
        }
    
        return mergedRows;
    };
    

    const onFileUpload = (event) => {
        setJSONLoad(true);
        const file = event.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                // Process all sheets in the workbook
                const sheetsData = workbook.SheetNames.map(sheetName => {
                    const sheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                    // Normalize columnHeaders
                    const normalizedColumnHeaders = normalizeRow(columnHeaders);

                    // Find the header row index
                    let headerIndex = -1;
                    for (let i = 0; i < jsonData.length; i++) {
                        if (JSON.stringify(normalizeRow(jsonData[i])) === JSON.stringify(normalizedColumnHeaders)) {
                            headerIndex = i;
                            break;
                        }
                    }

                    // Default to the entire jsonData if no header found
                    const processedData = headerIndex !== -1 ? jsonData.slice(headerIndex + 1) : jsonData;

                    // Merge rows to handle split data
                    const mergedRows = mergeRows(processedData);

                    // Determine the number of columns based on columnHeaders
                    const maxColumns = columnHeaders.length;

                    // Ensure all cells are filled with " " instead of null or empty
                    const rows = mergedRows.map(rowData => 
                        Array.from({ length: maxColumns }, (_, index) => 
                            rowData[index] === null || rowData[index] === undefined || rowData[index] === "" ? " " : rowData[index]
                        )
                    );

                    return {
                        name: sheetName,
                        details: rows
                    };
                });

                setInputJson(sheetsData);
                console.log(sheetsData); // Log the processed data
                setJSONLoad(false);
            };
            reader.readAsArrayBuffer(file);
        }
    };

    return (
        <div className="App">
            <header className="App-header">
                <h3>Format 1 - Excel (P1)</h3>
                <input type="file" onChange={onFileUpload} accept=".xlsx, .xls" />
            </header>
        </div>
    );
}

export default App;
