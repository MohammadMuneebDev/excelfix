import React, { useState } from 'react';
import * as XLSX from 'xlsx';

function LastFormat() {
  const [mergedData, setMergedData] = useState([]);

  const handleFileUpload = (event) => {
    const files = event.target.files;
    const allSheetsData = [];

    Array.from(files).forEach((file) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        workbook.SheetNames.forEach((sheetName) => {
          const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
          console.log(`Data from sheet ${sheetName}:`, sheet); // Log each sheet's data
          allSheetsData.push(sheet);
        });

        const merged = processSheetsAndMerge(allSheetsData);
        console.log('Merged Data:', merged); // Log merged data
        setMergedData(merged);
      };
      reader.readAsArrayBuffer(file);
    });
  };

  const processSheetsAndMerge = (sheets) => {
    // Define the static heading row to be used
    const staticHeadingRow = [
      "Company",
      "Meeting Type",
      "Meeting Date",
      "Resolution",
      "Proposal",
      "Proposal Type",
      "Vote Cast",
      "Reason",
    ];
  
    const mergedData = [];
    let headingAdded = false;
  
    sheets.forEach((sheet, index) => {
      // Skip the first sheet
      if (index === 0) {
        console.log(`Skipping sheet ${index + 1}`);
        return;
      }
  
      // Remove only the first three rows
      const processedSheet = sheet.slice(3);
      console.log(`Processed Sheet ${index + 1} (after removing first three rows):`, processedSheet);
  
      // Add the static heading row only once
      if (!headingAdded) {
        mergedData.push(staticHeadingRow);
        headingAdded = true;
      }
  
      // Add actual data from each sheet
      const dataStartIndex = processedSheet.findIndex(
        (row) => Array.isArray(row) && row.length > 0
      );
  
      if (dataStartIndex !== -1) {
        const actualData = processedSheet.slice(dataStartIndex);
        console.log(`Actual Data from Sheet ${index + 1}:`, actualData);
        mergedData.push(...actualData);
      }
    });
  
    return mergedData;
  };

  const handleDownload = () => {
    console.log('Handle download called');
    try {
      if (mergedData.length === 0) {
        console.warn('No data to download');
        return;
      }
      const worksheet = XLSX.utils.aoa_to_sheet(mergedData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "MergedData");

      console.log('Writing file');
      XLSX.writeFile(workbook, "MergedData.xlsx");
      console.log('File written');
    } catch (error) {
      console.error('Error during file download:', error);
    }
  };

  return (
    <div className="App">
      <h1>Excel Sheet Merger</h1>
      <input type="file" onChange={handleFileUpload} multiple />
      {mergedData.length > 0 && (
        <>
          <button onClick={handleDownload}>Download Merged Excel</button>
          <table border="1">
            <thead>
              <tr>
                {mergedData[0].map((header, index) => (
                  <th key={index}>{header}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {mergedData.slice(1).map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {row.map((cell, cellIndex) => (
                    <td key={cellIndex}>{cell}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}
    </div>
  );
}

export default LastFormat;
