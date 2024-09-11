import React, { useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
 
function Format3() {
  const [inputJson, setInputJson] = useState({});
  const [jsonLoad, setJSONLoad] = useState(false);
  const [splitHeaders, setSplitHeaders] = useState(
    "Company, Meeting Type, Meeting Date, Resolution, Proposal, Proposal Type, Vote Cast, Reason"
  );
  const [exportJson, setExportJson] = useState({});
  const [uploadedFileName, setUploadedFileName] = useState("");
 
  const columnHeaders = splitHeaders.split(", ");
 
  const mergeNextRows = (data) => {
    if (!Array.isArray(data) || data.length === 0) {
      throw new Error('Data is not defined or invalid.');
    }
 
    let processedData = [];
 
    for (let i = 0; i < data.length; i++) {
      let currentRow = [...data[i]];
 
      // Check if the current row has a value in index[2]
      if (currentRow[2] && currentRow[2] !== '') {
        // Check the next two rows for values in index[2]
        let nextRow1 = data[i + 1] || [];
        let nextRow2 = data[i + 2] || [];
 
        if ((!nextRow1[2] || nextRow1[2] === '') && (!nextRow2[2] || nextRow2[2] === '')) {
          // Merge next rows' index[0] and index[1] into the current row's index[0] and index[1]
          currentRow[0] = (currentRow[0] || '') + ' ' + (nextRow1[0] || '') + ' ' + (nextRow2[0] || '');
          currentRow[1] = (currentRow[1] || '') + ' ' + (nextRow1[1] || '') + ' ' + (nextRow2[1] || '');
 
          // Trim and keep only the necessary merged data
          currentRow[0] = currentRow[0].trim();
          currentRow[1] = currentRow[1].trim();
 
          // Replace merged columns in nextRow1 and nextRow2 with " "
          data[i + 1] = [' ', ' ', ...nextRow1.slice(2)];
          data[i + 2] = [' ', ' ', ...nextRow2.slice(2)];
        }
      }
 
      // Add the processed row to the result
      processedData.push(currentRow);
    }
 
    return processedData;
  };
 
 
  const processRowsFurther = (rows) => {
    if (!Array.isArray(rows) || rows.length === 0) {
      throw new Error('Rows are not defined or invalid.');
    }
 
    // Remove empty rows by ensuring every cell is a non-empty string
    let cleanedRows = rows.filter(row => row.some(cell => typeof cell === 'string' && cell.trim() !== ''));
 
    for (let i = 0; i < cleanedRows.length; i++) {
      let currentRow = [...cleanedRows[i]];
 
      // Check for index[0] value and fill next rows if they are missing values
      if (typeof currentRow[0] === 'string' && currentRow[0].trim() !== '') {
        for (let j = i + 1; j < cleanedRows.length; j++) {
          if (!cleanedRows[j][0] || (typeof cleanedRows[j][0] === 'string' && cleanedRows[j][0].trim() === '')) {
            cleanedRows[j][0] = currentRow[0]; // Copy the value from the current row
          } else {
            break; // Stop copying if a non-empty value is found
          }
        }
      }
 
      // Check for index[1] value and fill next rows if they are missing values
      if (typeof currentRow[1] === 'string' && currentRow[1].trim() !== '') {
        for (let j = i + 1; j < cleanedRows.length; j++) {
          if (!cleanedRows[j][1] || (typeof cleanedRows[j][1] === 'string' && cleanedRows[j][1].trim() === '')) {
            cleanedRows[j][1] = currentRow[1]; // Copy the value from the current row
          } else {
            break; // Stop copying if a non-empty value is found
          }
        }
      }
 
      // Check for index[2] value and fill next rows if they are missing values
      if (typeof currentRow[2] === 'string' && currentRow[2].trim() !== '') {
        for (let j = i + 1; j < cleanedRows.length; j++) {
          if (!cleanedRows[j][2] || (typeof cleanedRows[j][2] === 'string' && cleanedRows[j][2].trim() === '')) {
            cleanedRows[j][2] = currentRow[2]; // Copy the value from the current row
          } else {
            break; // Stop copying if a non-empty value is found
          }
        }
      }
 
      // Check for index[3] value and fill next rows if they are missing values
      if (typeof currentRow[3] === 'string' && currentRow[3].trim() !== '') {
        for (let j = i + 1; j < cleanedRows.length; j++) {
          if (!cleanedRows[j][3] || (typeof cleanedRows[j][3] === 'string' && cleanedRows[j][3].trim() === '')) {
            cleanedRows[j][3] = currentRow[3]; // Copy the value from the current row
          } else {
            break; // Stop copying if a non-empty value is found
          }
        }
      }
    }
 
    return cleanedRows;
  };
 
 
  const replaceLineBreaks = (rows) => {
    if (!Array.isArray(rows) || rows.length === 0) {
      throw new Error('Rows are not defined or invalid.');
    }
 
    return rows.map(row => {
      return row.map(cell => {
        if (typeof cell === 'string') {
          // Replace \r\n with a space if it exists
          return cell.replace(/\r\n/g, ' ');
        }
        return cell; // Leave non-string cells unchanged
      });
    });
  };
  // Usage example:
 
 
 
 
  const onFileUpload = (event) => {
    setJSONLoad(true);
    const file = event.target.files[0];
    setUploadedFileName(file?.name?.split(".")[0]);
 
    if (file) {
      const reader = new FileReader();
 
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
 
          let rawData = [];
          let allData = {};
 
          workbook.SheetNames.forEach((sheetName) => {
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
 
            rawData = rawData.concat(jsonData);
 
            console.log("Raw Data for sheet:", sheetName, jsonData);
 
            const processedRows = mergeNextRows(jsonData);
  const furtherProcessedRows = processRowsFurther(processedRows);
            const finalProcessedRows = replaceLineBreaks(furtherProcessedRows);
 
            console.log("Processed Rows for sheet:", sheetName, finalProcessedRows);
 
            allData[sheetName] = finalProcessedRows;
          });
 
          console.log("Raw Combined JSON Data:", rawData);
 
          setInputJson({ rawData, allSheets: allData });
          setExportJson({ rawData, allSheets: allData });
 
        } catch (error) {
          console.error("Error processing file:", error);
          alert("An error occurred while processing the file. Please check the console for details.");
        } finally {
          setJSONLoad(false);
        }
      };
 
      reader.readAsArrayBuffer(file);
    }
  };
 
  const convertJSONToExcel = () => {
    const workbook = new ExcelJS.Workbook();
 
    Object.keys(exportJson.allSheets).forEach((sheetName) => {
      const worksheet = workbook.addWorksheet(sheetName);
      worksheet.addRow(columnHeaders);
      exportJson.allSheets[sheetName].forEach((dataRow) => {
        const row = columnHeaders.map((_, index) => dataRow[index] || "");
        worksheet.addRow(row);
      });
    });
 
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Formatted_EXCEL.xlsx";
      a.click();
      URL.revokeObjectURL(url);
    });
  };
 
  return (
    <div>
      <input type="file" onChange={onFileUpload} />
      <button onClick={convertJSONToExcel}>Convert to Excel</button>
      {jsonLoad && <div>Loading...</div>}
    </div>
  );
}
 
export default Format3;