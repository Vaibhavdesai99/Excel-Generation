// Function to fetch data from the backend and append it to Table 6
async function fetchDataAndPopulateTable() {
  try {
    // Make a GET request to fetch data from the backend
    const response = await fetch("/odata/v4/postdata/ExcelData");
    const responseData = await response.json();
    const data = responseData.value; // Extract the array of data from the response
    console.log(data);

    // Get the workbook
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");

    // Define border style
    const borderStyle = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    // Add empty cells for margin
    for (let i = 0; i < 3; i++) {
      sheet.addRow([]); // Add empty row
    }

    //------------------------------HEADING PART----------------------------------------------------
    const heading = "Mainline Welding Report";
    const headingRow = sheet.addRow([""]); // Add a row with an empty cell
    const headingCell = headingRow.getCell(6); // Get the cell in the second column
    headingCell.value = heading; // Set the heading text to the cell
    headingCell.font = { bold: true, size: 25 }; // Apply styling to the heading cell

    // Apply border to columns 2 to 17 (B to Q) for the heading row
    for (let i = 2; i <= 17; i++) {
      headingRow.getCell(i).border = {
        top: { style: "thin" },
        left: i === 2 ? { style: "thin" } : null, // Apply left border only to column B (index 2)
        bottom: { style: "thin" },
        right: i === 17 ? { style: "thin" } : null, // Apply right border only to column Q (index 17)
      };
    }

    // Apply border to row 4 from columns B to Q

    // Hide gridlines (row and column borders) for the heading section

    // ----------------------END OF HEADING PART -----------------------------------------------------

    // ---------------Project name , client name , execution partner , contractor , DOC.NO , Report genration Date---------

    // Add project details rows
    const projectDetails = [
      { key: "Project Name", value: "Project X" },
      { key: "Client Name", value: "Client Y" },
      { key: "Execution Partner", value: "Execution Partner" },
      { key: "Contractor", value: "Contractor A" },
      { key: "DOC.NO", value: "12345" },
      { key: "Report Generation Date", value: "2024-03-08" },
    ];

    // Add project details rows with key-value pairs
    projectDetails.forEach((detail) => {
      const detailRow = sheet.addRow(["", detail.key, detail.value]); // Adjusted to add an empty cell in column A
      detailRow.getCell(2).border = {
        // Apply border only to the cell in column C
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      // Check if the key is "Project Name" or "Client Name" and set its font to bold
      if (
        detail.key === "Project Name" ||
        detail.key === "Client Name" ||
        detail.key === "Execution Partner" ||
        detail.key === "Contractor" ||
        detail.key === "DOC.NO" ||
        detail.key === "Report Generation Date"
      ) {
        detailRow.getCell(2).font = { bold: true }; // Apply bold font to the cell in column B
      }
    });
    // Apply border to the entire table
    for (let rowNumber = 3; rowNumber <= 11; rowNumber++) {
      for (let i = 2; i <= 17; i++) {
        const cell = sheet.getRow(rowNumber).getCell(i);
        cell.border = {
          top: rowNumber === 3 ? { style: "thin" } : null, // Apply top border only to row 3
          left: i === 2 ? { style: "thin" } : null, // Apply left border only to column B (index 2)
          bottom: rowNumber === 11 ? { style: "thin" } : null, // Apply bottom border only to row 11
          right: i === 17 ? { style: "thin" } : null, // Apply right border only to column Q (index 17)
        };
      }
    }
    //----------------------- Set font bold for specific cells

    //--------------- Set column widths for specific columns-----------------------------------
    const columnWidths = [
      { col: 2, width: 17 }, // Project Name column (assuming it's the first column)
      { col: 3, width: 17 }, // Client Name column (assuming it's the second column)
      // Add more entries for other columns if needed
    ];

    // Set column widths
    columnWidths.forEach((column) => {
      const columnObject = sheet.getColumn(column.col);
      columnObject.width = column.width;
    });
    // -----------END OF Project name , client name , execution partner , contractor , DOC.NO , Report genration Date---------

    //  ---------------------MAIN TABLE DYNAMIC DATA --------------------------------------------
    // Add headers with borders and styling
    const headerRow = sheet.addRow([
      "",
      "ID",
      "Sr No",
      "Pipe Unique ID",
      "Heat No",
      "Coating No",
      "Length",
      "Bend Detail",
      "Weld Joint No",
      "Fit-up Inspection",
      "Welder ID",
      "Root Pass Welder No",
      "Hot Pass Welder No",
      "Filler Pass Welder No",
      "Cap Pass Welder No",
      "WPS No",
      "Remarks",
    ]);
    headerRow.eachCell((cell, colNumber) => {
      if (colNumber !== 1) {
        // Skip applying border to the first column (A)
        cell.border = borderStyle;
      }
      cell.font = { bold: true };
      // Add more styling if needed
    });

    // Merge cells B12 and B13-----ID----------- after setting headers and before inserting data rows
    sheet.mergeCells("B12:B13");
    const mergedCellB = sheet.getCell("B12");
    mergedCellB.value = "ID"; // Customize this as needed
    // Example styling for the merged cell
    mergedCellB.alignment = { horizontal: "center", vertical: "center" };
    mergedCellB.border = borderStyle;
    mergedCellB.font = { bold: true };

    // Merge cells C12 and C13-----SR.NO----------- after setting headers and before inserting data rows
    sheet.mergeCells("C12:C13");
    const mergedCellc = sheet.getCell("C12");
    mergedCellc.value = "SR.NO"; // Customize this as needed
    // Example styling for the merged cell
    mergedCellc.alignment = { horizontal: "center", vertDcal: "center" };
    mergedCellc.border = borderStyle;
    mergedCellc.font = { bold: true };

    // Merge cells D12 and D13-----SR.NO----------- after setting headers and before inserting data rows
    sheet.mergeCells("D12:D13");
    const mergedCellD = sheet.getCell("D12");
    mergedCellc.value = "Pipe No/Unique id"; // Customize this as needed
    // Example styling for the merged cell
    mergedCellD.alignment = { horizontal: "center", vertical: "center" };
    mergedCellD.border = borderStyle;
    mergedCellD.font = { bold: true };

    // Merge cells E12 and E13-----SR.NO----------- after setting headers and before inserting data rows
    sheet.mergeCells("E12:E13");
    const mergedCellE = sheet.getCell("E12");
    mergedCellE.value = "Heat No"; // Customize this as needed
    // Example styling for the merged cell
    mergedCellE.alignment = { horizontal: "center", vertical: "center" };
    mergedCellE.border = borderStyle;
    mergedCellE.font = { bold: true };

    // Merge cells F12 and F13-----SR.NO----------- after setting headers and before inserting data rows
    sheet.mergeCells("F12:F13");
    const mergedCellF = sheet.getCell("F12");
    mergedCellF.value = "Coating No"; // Customize this as needed
    // Example styling for the mergeF cell
    mergedCellF.alignment = { horizontal: "center", vertical: "center" };
    mergedCellF.border = borderStyle;
    mergedCellF.font = { bold: true };

    // Merge cells G12 and G13-----SR.NO----------- after setting headers and before inserting data rows
    sheet.mergeCells("G12:G13");
    const mergedCellG = sheet.getCell("G12");
    mergedCellG.value = "Length"; // Customize this as needed
    // Example styling for the merged cell
    mergedCellG.alignment = { horizontal: "center", vertical: "center" };
    mergedCellG.border = borderStyle;
    mergedCellG.font = { bold: true };

    // Merge cells H12 and H13-----SR.NO----------- after setting headers and before inserting data rows
    sheet.mergeCells("H12:H13");
    const mergedCellH = sheet.getCell("H12");
    mergedCellG.value = "Bend Detail"; // Customize this as needed
    // Example styling for the merged cell
    mergedCellH.alignment = { horizontal: "center", vertical: "center" };
    mergedCellH.border = borderStyle;
    mergedCellH.font = { bold: true };

    // Merge cells I12 and I13-----SR.NO----------- after setting headers and before inserting data rows
    sheet.mergeCells("I12:I13");
    const mergedCellI = sheet.getCell("I12");
    mergedCellG.value = "Weld Joint No"; // Customize this as needed
    // Example styling for the merged cell
    mergedCellI.alignment = { horizontal: "center", vertical: "center" };
    mergedCellI.border = borderStyle;
    mergedCellI.font = { bold: true };

    // Merge cells J12 and J13-----SR.NO----------- after setting headers and before inserting data rows
    sheet.mergeCells("J12:J13");
    const mergedCellJ = sheet.getCell("J12");
    mergedCellG.value = "Fit-up Inspection"; // Customize this as needed
    // Example styling for the merged cell
    mergedCellJ.alignment = { horizontal: "center", vertical: "center" };
    mergedCellJ.border = borderStyle;
    mergedCellJ.font = { bold: true };

    // Merge cells K12 andK13-----SR.NO----------- after setting headers and before inserting data rows
    sheet.mergeCells("K12:K13");
    const mergedCellK = sheet.getCell("K12");
    mergedCellG.value = "Welder ID"; // Customize this as needed
    // Example styling for the merged cell
    mergedCellK.alignment = { horizontal: "center", vertical: "center" };
    mergedCellK.border = borderStyle;
    mergedCellK.font = { bold: true };
    // Add data rows with borders and styling

    // Merge cells P12 andP13-----SR.NO----------- after setting headers and before inserting data rows
    sheet.mergeCells("P12:P13");
    const mergedCellP = sheet.getCell("P12");
    mergedCellG.value = "WPS NO"; // Customize this as needed
    // Example styling for the merged cell
    mergedCellP.alignment = { horizontal: "center", vertical: "center" };
    mergedCellP.border = borderStyle;
    mergedCellP.font = { bold: true };
    // Add data rows with borders and styling

    // Merge cells Q12 andQ13-----SR.NO----------- after setting headers and before inserting data rows
    sheet.mergeCells("Q12:Q13");
    const mergedCellQ = sheet.getCell("Q12");
    mergedCellG.value = "RemarkQ"; // Customize this as needed
    // Example styling for the merged cell
    mergedCellQ.alignment = { horizontal: "center", vertical: "center" };
    mergedCellQ.border = borderStyle;
    mergedCellQ.font = { bold: true };
    // Add data rows with borders and styling
    //-----------------------------------------------------------------------------------------------------
    data.forEach((item, index) => {
      // Calculate the row number to start from 14
      const rowNumber = 14 + index;
      const row = sheet.getRow(rowNumber);

      // Populate the row with data
      row.values = [
        "",
        item.id || "",
        item.srno || "",
        item.pipeuniqueid || "",
        item.heatno || "",
        item.coatingno || "",
        item.length || "",
        item.benddetail || "",
        item.weldjointno || "",
        item.fitupinspection || "",
        item.welderid || "",
        item.rootpasswelderno || "",
        item.hotpasswelderno || "",
        item.fillerpasswelderno || "",
        item.cappasswelderno || "",
        item.wpsno || "",
        item.remarks || "",
      ];

      // Apply the border style and any additional styling
      row.eachCell((cell, colNumber) => {
        if (colNumber !== 1) {
          // Skip applying border to the first column (A)
          cell.border = borderStyle;
        }
        // Apply more styling as needed
        cell.font = { bold: true };
      });
    });

    // Increase the height of row 12
    sheet.getRow(12).height = 25; // Adjust the number to set your desired height

    // Your loop for setting common width for columns 2 to 17
    const commonWidth = 11;
    for (let col = 2; col <= 17; col++) {
      sheet.getColumn(col).width = commonWidth;
    }

    // Step 1: Set specific headings in cells L13 to O13
    sheet.getCell("L13").value = "Root Pass";
    sheet.getCell("M13").value = "Hot Pass";
    sheet.getCell("N13").value = "Filler Pass";
    sheet.getCell("O13").value = "Cap Pass";

    // Apply alignment and optional styling to these heading cells
    ["L13", "M13", "N13", "O13"].forEach((cellRef) => {
      sheet.getCell(cellRef).alignment = { horizontal: "center" };
      // Add any desired styling here
    });

    // Step 2: Merge cells L12 to O12 for "Welder No" heading
    sheet.mergeCells("L12:O12");
    const mergedCell = sheet.getCell("L12");
    mergedCell.value = "Welder No";
    mergedCell.alignment = { horizontal: "center", vertical: "center" };
    // Apply a bottom border to this cell to visually separate it from the headings below
    mergedCell.border = { bottom: { style: "thin" } };

    // Set the width of each column individually
    sheet.getColumn(2).width = 8; // For "ID"
    sheet.getColumn(3).width = 10; // For "Sr No"
    sheet.getColumn(4).width = 13; // For "Pipe Unique ID"
    sheet.getColumn(5).width = 10; // For "Heat No"
    sheet.getColumn(6).width = 12; // For "Coating No"
    sheet.getColumn(7).width = 8; // For "Length"
    sheet.getColumn(8).width = 12; // For "Bend Detail"
    sheet.getColumn(9).width = 13; // For "Weld Joint No"
    sheet.getColumn(10).width = 13; // For "Fit-up Inspection"
    sheet.getColumn(11).width = 10; // For "Welder ID"
    sheet.getColumn(12).width = 9; // For "Root Pass Welder No"
    sheet.getColumn(13).width = 9; // For "Hot Pass Welder No"
    sheet.getColumn(14).width = 9; // For "Filler Pass Welder No"
    sheet.getColumn(15).width = 9; // For "Cap Pass Welder No"
    sheet.getColumn(16).width = 8; // For "WPS No"
    sheet.getColumn(17).width = 12; // For "Remarks"

    // ------------END OF MIAN TABLE DYNAMIC DATA---------------------------------------

    //--------------------- Add a row for Progress -------------------------
    const progressRow = sheet.addRow(["", "Progress"]); // Add an empty cell in column A
    progressRow.getCell(2).font = { bold: true }; // Set font of the cell in column B to bold
    for (let i = 2; i <= 17; i++) {
      // Apply border to columns B to Q
      progressRow.getCell(i).border = {
        top: { style: "thin" },
        left: i === 2 ? { style: "thin" } : null, // Apply left border only to column B
        bottom: { style: "thin" },
        right: i === 17 ? { style: "thin" } : null, // Apply right border only to column Q
      };
    }
    //----------------END OF PROGRESS----------------------------------------

    // -------------------Total Joints Completed------------------
    const Row = sheet.addRow([]); // Add an empty row
    Row.getCell(7).value = "Total Joints Completed"; // Set the value of the cell in column G
    Row.getCell(7).font = { bold: true }; // Set font of the cell in column G to bold
    for (let i = 2; i <= 17; i++) {
      // Apply border to columns B to Q
      Row.getCell(i).border = {
        top: { style: "thin" },
        left: i === 2 ? { style: "thin" } : null, // Apply left border only to column B
        bottom: { style: "thin" },
        right: i === 17 ? { style: "thin" } : null, // Apply right border only to column Q
      };
    }
    // Write workbook to buffer
    const buffer = await workbook.xlsx.writeBuffer();

    // Create a Blob from the buffer
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    // Use native browser functionality to save the file
    if (window.navigator.msSaveOrOpenBlob) {
      window.navigator.msSaveOrOpenBlob(blob, `${Date.now()}_feedback.xlsx`);
    } else {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      document.body.appendChild(a);
      a.href = url;
      a.download = `${Date.now()}_feedback.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    }
  } catch (error) {
    console.error("Error fetching data:", error);
  }
}

// Call the function to convert HTML to Excel on button click
document
  .getElementById("convertToExcel")
  .addEventListener("click", fetchDataAndPopulateTable);
