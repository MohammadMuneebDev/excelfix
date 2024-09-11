import React, { useState } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

function Format3() {
  const [inputJson, setInputJson] = useState([]);
  const [jsonLoad, setJSONLoad] = useState(false);
  const [splitHeaders, setSplitHeaders] = useState(
    "Company, Meeting Type, Meeting Date, Resolution, Proposal, Proposal Type, Vote Cast, Reason"
  );
  const [exportJson, setExportJson] = useState([]);
  const [formatLoading, setFormatLoading] = useState(false);
  const [uploadedFileName, setUploadedFileName] = useState("");
  const [debugData, setDebugData] = useState([]); // State for debug data

  const columnHeaders = splitHeaders.split(", ");
  const mergeAndIgnoreHeaders = (data) => {
    return data
      .map((set) => {
        if (set.length === 0) return set;
  
        const filteredSet = set.filter(
          (row) => !row.every((cell, index) => cell === columnHeaders[index])
        );
  
        if (filteredSet.length === 0) return [];
  
        let sectionStartIndex = -1;
        const mergedSet = [];
  
        filteredSet.forEach((row, i) => {
          if (row[2]) { // Check for "Meeting Date"
            if (sectionStartIndex !== -1) {
              mergeRows(filteredSet, sectionStartIndex, i, mergedSet);
            }
            sectionStartIndex = i;
            mergedSet.push(row);
          } else {
            mergedSet.push(row);
          }
        });
  
        // Handle any remaining section after the last date row
        if (sectionStartIndex !== -1 && filteredSet.length > sectionStartIndex + 1) {
          mergeRows(filteredSet, sectionStartIndex, filteredSet.length, mergedSet);
        }
  
        return mergedSet;
      })
      .filter((set) => set.length > 0);
  };
  
  const mergeRows = (filteredSet, start, end, mergedSet) => {
    const section = filteredSet.slice(start + 1, end);
    
    // Check if we have enough rows to consider merging
    if (section.length < 3) {
      mergedSet.push(...section); // Not enough rows to merge
      return;
    }
  
    const rowsToMerge = section.slice(0, 3);
    
    // Check if the fourth row has data
    if (section[3] && (section[3][0] || section[3][1])) {
      // Fourth row has data, do not merge
      mergedSet.push(filteredSet[start], ...section);
      return;
    }
  
    let mergedFirstColumn = (filteredSet[start][0] || "").toString();
    let mergedSecondColumn = (filteredSet[start][1] || "").toString();
  
    rowsToMerge.forEach(secRow => {
      if (secRow[0]) mergedFirstColumn += ` ${secRow[0]}`;
      if (secRow[1]) mergedSecondColumn += ` ${secRow[1]}`;
    });
  
    filteredSet[start][0] = mergedFirstColumn.trim() || " ";
    filteredSet[start][1] = mergedSecondColumn.trim() || " ";
    
    rowsToMerge.forEach(secRow => {
      secRow[0] = " ";
      secRow[1] = " ";
    });
  
    mergedSet.push(filteredSet[start]);
    mergedSet.push(...section.slice(3));
  };

  const onFileUpload = (event) => {
    setJSONLoad(true);
    const file = event.target.files[0];
    setUploadedFileName(file?.name?.split(".")[0]);

    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        let allData = [];

        workbook.SheetNames.forEach((sheetName) => {
          const sheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
          const processedRows = mergeAndIgnoreHeaders(jsonData); // Process rows with merging logic

          allData.push({
            "sheet name": sheetName,
            "details": processedRows
          });
        });

        setInputJson(allData); // Store the collected data in state
        setJSONLoad(false);
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const convertJSONToExcel = () => {
    const workbook = new ExcelJS.Workbook();

    // Loop through each sheet's data in the exportJson
    Object.keys(exportJson).forEach((sheetName) => {
      const worksheet = workbook.addWorksheet(sheetName); // Create a worksheet for each sheet name

      const dataSet = exportJson[sheetName]; // Get the data set for the current sheet

      // Add the column headers to each sheet
      worksheet.addRow(columnHeaders);

      // Add rows of data to the worksheet
      dataSet.forEach((dataRow) => {
        const row = columnHeaders.map((_, index) => dataRow[index] || ""); // Map each cell in the row
        worksheet.addRow(row); // Add each row to the worksheet
      });
    });

    // Write the workbook to a buffer and trigger the download
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Formatted_EXCEL.xlsx"; // Fixed download filename to be in quotes
      a.click();
      URL.revokeObjectURL(url);
    });
  };

  return (
    <div>
      <input type="file" onChange={onFileUpload} />
      <button onClick={convertJSONToExcel}>Convert to Excel</button>
    </div>
  );
}

export default Format3;
