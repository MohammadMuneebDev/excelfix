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
          allData = allData.concat(jsonData);
        });
        setInputJson(allData);
        setJSONLoad(false);
      };
      reader.readAsArrayBuffer(file);
    }
  };

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

  const shiftValuesInRows = (data) => {
    return data.map((subArray) => {
      if (Array.isArray(subArray)) {
        return subArray.map((row) => {
          // Log the row before any manipulation
          console.log("Before manipulation:", row);
  
          if (Array.isArray(row)) {
            // Check if the row length is between 3 and 5 and has non-empty values between indices [0] and [4]
            const hasValuesToShift = row.length > 2 && row.length <= 5 && row.slice(0, 5).some(cell => String(cell).trim() !== "");
  
            if (hasValuesToShift) {
              // Create a new row with empty strings and preserve the length
              const newRow = Array(row.length).fill("");
  
              // Initialize index for placing values
              let newIndex = 3;
  
              // Shift values from index [0] to [4] starting from index [3]
              row.forEach((cell, index) => {
                if (index < 5 && String(cell).trim() !== "") {
                  newRow[newIndex++] = cell; // Place the cell in the new row
                } else if (index >= 5) {
                  newRow[index] = cell; // Preserve cells beyond index 4
                }
              });
  
              // Log the row after manipulation
              console.log("After manipulation:", newRow);
  
              return newRow;
            }
  
            return row; // Return the original row if no shifting is needed
          }
  
          return row; // Return non-array values as-is
        });
      }
  
      return subArray; // Return non-array items as-is
    });
  };
  
  
  
  


  const formatJson = () => {
    setFormatLoading(true);
    let formattedData = [];
    let currentSet = [];
    let headersFound = false;

    inputJson.forEach((row) => {
      const rowLowerCase = row.map((cell) =>
        cell ? String(cell).toLowerCase() : ""
      );

      if (
        rowLowerCase.every(
          (value, index) => value === columnHeaders[index]?.toLowerCase()
        )
      ) {
        if (currentSet.length > 0) {
          formattedData.push([...currentSet]);
        }

        headersFound = true;
        currentSet = [];
      }

      if (headersFound) {
        currentSet.push(row);
      }
    });

    if (currentSet.length > 0) {
      formattedData.push([...currentSet]);
    }

    const structuredData = mergeAndIgnoreHeaders(formattedData);
    // console.log(structuredData);
    const news=shiftValuesInRows(structuredData);
    console.log(news);
    
    setExportJson(news);
    setDebugData(news); // Set the debug data state
    setFormatLoading(false);
  };

  const convertJSONToExcel = () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");

    worksheet.addRow(columnHeaders);

    exportJson.forEach((dataSet) => {
      dataSet.forEach((dataRow) => {
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
      a.download = `Formatted_EXCEL.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    });
  };

  return (
    <div style={{ padding: "20px" }}>
      <h3>Format 1 - Excel (P1)</h3>
      <input type="file" onChange={onFileUpload} accept=".xlsx, .xls" />
      {jsonLoad && <h4 style={{ color: "red" }}>Fetching Details ...</h4>}
      {inputJson?.length > 0 && !jsonLoad && (
        <h4 style={{ color: "green" }}>Data Fetched Successfully!</h4>
      )}
      <br />
      <br />
      <div style={{ marginBottom: "20px" }}>
        <label>
          Add Split Headers{" "}
          <small style={{ color: "red" }}> ( Add "," and "--space--" )</small>
        </label>
        <input
          style={{ width: "500px" }}
          type="text"
          onChange={(e) => setSplitHeaders(e.target.value)}
          value={splitHeaders}
          placeholder="Number, Proposal Text, Proponent, Mgmt"
        />
      </div>
      <br />
      {inputJson?.length > 0 && exportJson?.length === 0 && (
        <button
          style={{
            padding: "10px 20px",
            backgroundColor: "#007bff",
            color: "#fff",
            border: "none",
            borderRadius: "5px",
          }}
          onClick={() => formatJson()}
        >
          Format JSON
        </button>
      )}
      <br />
      {exportJson?.length > 0 && (
        <button
          style={{
            padding: "10px 20px",
            backgroundColor: "#28a745",
            color: "#fff",
            border: "none",
            borderRadius: "5px",
          }}
          onClick={() => convertJSONToExcel()}
        >
          Export to Excel
        </button>
      )}
      {formatLoading && <p>Loading...</p>}

      {/* Debug section */}
      {debugData.length > 0 && (
        <div style={{ marginTop: "20px" }}>
          <h4>Debug Data:</h4>
          <pre style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
            {JSON.stringify(debugData, null, 2)}
          </pre>
        </div>
      )}
    </div>
  );
}

export default Format3;
