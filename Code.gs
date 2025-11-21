/**
 * @OnlyCurrentDoc
 * This script manages multiple production applications with a shared project management system.
 */

// An object to namespace all functions related to the "Materials" sheet.
const materialsSheet = {
  NAME: 'Materials',

  getSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(this.NAME);
    if (!sheet) {
      console.error(`Sheet '${this.NAME}' not found. Please ensure a sheet with this exact name exists in the current spreadsheet.`);
    }
    return sheet;
  },

  getData: function(category = 'FABRICATION') {
    const sheet = this.getSheet();
    if (!sheet) {
      return [];
    }
    const range = sheet.getRange('A2:P' + sheet.getLastRow());
    const values = range.getValues();

    const filteredValues = values
      .filter(row => {
        const name = row[1];
        const primaryCategory = row[4];
        return name && name.toString().trim() !== "" && primaryCategory && primaryCategory.toString().toUpperCase().includes(category);
      });

    if (category === 'PRINT') {
        return filteredValues.map(row => {
            const name = row[1].toString().trim(); // Column B
            const type = row[6] ? row[6].toString().trim().toUpperCase() : 'SHEET'; // Column G (Material Type)
            const width = parseFloat(row[7]) || 0; // Column H (inches)
            const length = parseFloat(row[8]) || 0; // Column I (feet for ROLL, inches for SHEET)
            let sheetCost = row[9]; // Column J

            if (sheetCost && typeof sheetCost === 'string') {
                const cleanedCost = parseFloat(sheetCost.replace(/[^0-9.-]+/g,""));
                sheetCost = isNaN(cleanedCost) ? 0 : cleanedCost;
            } else if (typeof sheetCost !== 'number') {
                sheetCost = 0;
            }
            
            // Calculate linear foot cost for ROLL materials
            // For ROLL: length is already in feet, so cost per linear foot = unit cost / length
            let costLinFt = 0;
            if (type === 'ROLL' && length > 0) {
                costLinFt = sheetCost / length;
            }
            
            // Return: [name, type, width, height, costSheet, costLinFt]
            return [name, type, width, length, sheetCost, costLinFt];
        });
    } else { // 'FABRICATION' and default
        return filteredValues.map(row => {
            const name = row[1].toString().trim();
            let unitCost = row[9];
            if (unitCost && typeof unitCost === 'string') {
                const cleanedCost = parseFloat(unitCost.replace(/[^0-9.-]+/g,""));
                unitCost = isNaN(cleanedCost) ? 0 : cleanedCost;
            } else if (typeof unitCost !== 'number') {
                unitCost = 0;
            }
            return {
              name: name,
              unitCost: unitCost
            };
        });
    }
  }
};

// An object to namespace all functions related to the "Personnel" sheet.
const personnelSheet = {
  NAME: 'Personnel',

  getSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(this.NAME);
    if (!sheet) {
      console.error(`Sheet '${this.NAME}' not found. Please ensure a sheet with this exact name exists in the current spreadsheet.`);
    }
    return sheet;
  },

  getData: function() {
    const sheet = this.getSheet();
    if (!sheet) {
      return [];
    }
    const range = sheet.getRange('B2:C' + sheet.getLastRow());
    const values = range.getValues();
    
    return values
      .filter(row => row[0] && row[0].toString().trim() !== "")
      .map(row => {
        const name = row[0].toString().trim();
        let projectRate = row[1];

        if (typeof projectRate !== 'number') {
          projectRate = parseFloat(projectRate) || 0;
        }

        return {
          name: name,
          projectRate: projectRate
        };
      });
  }
};

// An object to namespace all functions related to the fabrication application.
const fabricationApp = {
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('FabricationIndex')
        .setWidth(750)
        .setHeight(850);
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'Fabrication Details');
  },

  getMaterials: function() {
    try {
      return materialsSheet.getData('FABRICATION');
    } catch (e) {
      console.error("Error in fabricationApp.getMaterials: " + e.toString());
      return [];
    }
  },

  getPersonnel: function() {
    try {
      return personnelSheet.getData();
    } catch (e) {
      console.error("Error in fabricationApp.getPersonnel: " + e.toString());
      return [];
    }
  },

  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, 'FabricationLog');
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('FabricationIndex')
          .setWidth(750)
          .setHeight(850);
      
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace(
        '<script>',
        `<script>window.editFormData = ${JSON.stringify(formData)};`
      );
      
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent)
          .setWidth(750)
          .setHeight(850);
      
      const ui = SpreadsheetApp.getUi();
      ui.showModalDialog(modifiedOutput, 'Edit Fabrication Details');
    } else {
      this.showDialog();
    }
  },

  addToProject: function(fabricationData) {
    try {
      return projectSheet.addProjectItem(fabricationData, 'FAB', 'FabricationLog');
    } catch (e) {
      console.error("Error in fabricationApp.addToProject: " + e.toString());
      return {
        success: false,
        message: `Error adding to project: ${e.toString()}`,
        rowNumber: null,
        logId: null
      };
    }
  }
};

// An object to namespace all functions related to the apparel application.
const apparelApp = {
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('ApparelIndex')
        .setWidth(750)
        .setHeight(850);
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'Apparel / Screen Printing');
  },

  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, 'ApparelLog');
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('ApparelIndex')
          .setWidth(750)
          .setHeight(850);
      
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace(
        '<script>',
        `<script>window.editFormData = ${JSON.stringify(formData)};`
      );
      
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent)
          .setWidth(750)
          .setHeight(850);
      
      const ui = SpreadsheetApp.getUi();
      ui.showModalDialog(modifiedOutput, 'Edit Apparel');
    } else {
      this.showDialog();
    }
  },

  addToProject: function(apparelData) {
    try {
      return projectSheet.addProjectItem(apparelData, 'APP', 'ApparelLog');
    } catch (e) {
      console.error("Error in apparelApp.addToProject: " + e.toString());
      return {
        success: false,
        message: `Error adding to project: ${e.toString()}`,
        rowNumber: null,
        logId: null
      };
    }
  }
};

// An object to namespace all functions related to the printing estimate application.
const printingApp = {
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('PrintingIndex')
        .setWidth(750)
        .setHeight(850);
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'PrintCut Estimate');
  },

  getMaterials: function() {
    try {
      return materialsSheet.getData('PRINT');
    } catch (e) {
      console.error("Error in printingApp.getMaterials: " + e.toString());
      return [];
    }
  },

  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, 'PrintingLog');
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('PrintingIndex')
          .setWidth(750)
          .setHeight(850);
      
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace(
        '<script>',
        `<script>window.editFormData = ${JSON.stringify(formData)};`
      );
      
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent)
          .setWidth(750)
          .setHeight(850);
      
      const ui = SpreadsheetApp.getUi();
      ui.showModalDialog(modifiedOutput, 'Edit PrintCut Estimate');
    } else {
      this.showDialog();
    }
  },

  addToProject: function(printingData) {
    try {
      return projectSheet.addProjectItem(printingData, 'PRT', 'PrintingLog');
    } catch (e) {
      console.error("Error in printingApp.addToProject: " + e.toString());
      return {
        success: false,
        message: `Error adding to project: ${e.toString()}`,
        rowNumber: null,
        logId: null
      };
    }
  }
};

// An object to namespace all functions related to project data management.
const projectSheet = {
  getActiveSheet: function() {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  },

  logFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = spreadsheet.getSheetByName(logSheetName);
      
      if (!logSheet) {
        logSheet = spreadsheet.insertSheet(logSheetName);
        logSheet.hideSheet();
        logSheet.getRange(1, 1, 1, 4).setValues([
          ['LogID', 'ProjectRow', 'Timestamp', 'FormData']
        ]);
      }
      
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      
      const formDataWithRow = {
        ...formData,
        originalRowNumber: projectRowNumber
      };
      const formDataJson = JSON.stringify(formDataWithRow);
      
      const lastLogRow = logSheet.getLastRow();
      const nextLogRow = lastLogRow + 1;
      
      logSheet.getRange(nextLogRow, 1, 1, 4).setValues([
        [logId, projectRowNumber, timestamp, formDataJson]
      ]);
      
      return logId;
      
    } catch (error) {
      console.error('Error logging form data:', error);
      return null;
    }
  },

  updateLogFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = spreadsheet.getSheetByName(logSheetName);
      
      if (!logSheet) {
        return this.logFormData(formData, projectRowNumber, logIdPrefix, logSheetName);
      }
      
      const dataRange = logSheet.getDataRange();
      const values = dataRange.getValues();
      
      let existingRowIndex = -1;
      for (let i = 1; i < values.length; i++) {
        if (values[i][1] === projectRowNumber) {
          existingRowIndex = i + 1;
          break;
        }
      }
      
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      
      const formDataWithRow = {
        ...formData,
        originalRowNumber: projectRowNumber
      };
      const formDataJson = JSON.stringify(formDataWithRow);
      
      if (existingRowIndex > 0) {
        logSheet.getRange(existingRowIndex, 1, 1, 4).setValues([
          [logId, projectRowNumber, timestamp, formDataJson]
        ]);
      } else {
        const lastLogRow = logSheet.getLastRow();
        const nextLogRow = lastLogRow + 1;
        logSheet.getRange(nextLogRow, 1, 1, 4).setValues([
          [logId, projectRowNumber, timestamp, formDataJson]
        ]);
      }
      
      return logId;
      
    } catch (error) {
      console.error('Error updating log data:', error);
      return null;
    }
  },

  getLoggedFormData: function(logId, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = spreadsheet.getSheetByName(logSheetName);
      
      if (!logSheet) {
        return null;
      }
      
      const dataRange = logSheet.getDataRange();
      const values = dataRange.getValues();
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === logId) {
          const formDataJson = values[i][3];
          return JSON.parse(formDataJson);
        }
      }
      
      return null;
      
    } catch (error) {
      console.error('Error retrieving form data:', error);
      return null;
    }
  },

  createEditInstruction: function(logId) {
    return "Edit";
  },

  getNextFabricationId: function() {
    const sheet = this.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let maxNumber = 0;
    
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][1];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('F')) {
        const numberPart = cellValue.substring(1);
        const number = parseInt(numberPart);
        if (!isNaN(number) && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
    
    const nextNumber = maxNumber + 1;
    return `F${nextNumber.toString().padStart(2, '0')}`;
  },

  getNextApparelId: function() {
    const sheet = this.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let maxNumber = 0;
    
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][1];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('AP')) {
        const numberPart = cellValue.substring(2);
        const number = parseInt(numberPart);
        if (!isNaN(number) && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
    
    const nextNumber = maxNumber + 1;
    return `AP${nextNumber.toString().padStart(2, '0')}`;
  },

  getNextPrintingId: function() {
    const sheet = this.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let maxNumber = 0;
    
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][1];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('PR')) {
        const numberPart = cellValue.substring(2);
        const number = parseInt(numberPart);
        if (!isNaN(number) && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
    
    const nextNumber = maxNumber + 1;
    return `PR${nextNumber.toString().padStart(2, '0')}`;
  },

  updateProjectItem: function(itemData, logIdPrefix, logSheetName) {
    try {
      const sheet = this.getActiveSheet();
      
      if (!itemData || typeof itemData !== 'object') {
        throw new Error('Invalid item data provided');
      }

      const { description, quantity, dimensions, totalPrice, formData, originalRowNumber } = itemData;
      
      const rowNum = parseInt(originalRowNumber);
      if (!rowNum || isNaN(rowNum) || rowNum < 1) {
        throw new Error(`Invalid original row number for update: ${originalRowNumber}`);
      }
      
      const maxRows = sheet.getMaxRows();
      if (rowNum > maxRows) {
        throw new Error(`Row number ${rowNum} exceeds sheet maximum rows ${maxRows}`);
      }
      
      const logId = this.updateLogFormData(formData, rowNum, logIdPrefix, logSheetName);
      
      let rowData;
      let editColumnIndex;
      
      if (logIdPrefix === 'FAB') {
        const existingId = sheet.getRange(rowNum, 2).getValue() || this.getNextFabricationId();
        rowData = [
          '',
          existingId,
          description || '',
          dimensions || '',
          '',
          totalPrice || 0,
          'Edit'
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(rowNum, 1, 1, 6);
        range.setValues([rowData.slice(0, 6)]);
        
        const priceCell = sheet.getRange(rowNum, 6);
        priceCell.setNumberFormat('$#,##0.00');
        
      } else if (logIdPrefix === 'APP') {
        const existingId = sheet.getRange(rowNum, 2).getValue() || this.getNextApparelId();
        const unitPrice = quantity && quantity > 0 ? (totalPrice / quantity) : 0;
        
        rowData = [
          '',
          existingId,
          description || '',
          quantity || '',
          unitPrice,
          totalPrice || 0,
          'Edit'
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(rowNum, 1, 1, 6);
        range.setValues([rowData.slice(0, 6)]);
        
        const unitPriceCell = sheet.getRange(rowNum, 5);
        const totalPriceCell = sheet.getRange(rowNum, 6);
        unitPriceCell.setNumberFormat('$#,##0.00');
        totalPriceCell.setNumberFormat('$#,##0.00');
        
      } else if (logIdPrefix === 'PRT') {
        const existingId = sheet.getRange(rowNum, 2).getValue() || this.getNextPrintingId();
        const unitPrice = quantity && quantity > 0 ? (totalPrice / quantity) : 0;
        
        rowData = [
          '',
          existingId,
          description || '',
          quantity || '',
          unitPrice,
          totalPrice || 0,
          'Edit'
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(rowNum, 1, 1, 6);
        range.setValues([rowData.slice(0, 6)]);
        
        const unitPriceCell = sheet.getRange(rowNum, 5);
        const totalPriceCell = sheet.getRange(rowNum, 6);
        unitPriceCell.setNumberFormat('$#,##0.00');
        totalPriceCell.setNumberFormat('$#,##0.00');
      }
      
      if (logId) {
        const editCell = sheet.getRange(rowNum, editColumnIndex);
        editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Production > Edit Selected Item\n\nLast updated: ${new Date().toLocaleString()}`);
      }
      
      return {
        success: true,
        message: `Item updated in row ${rowNum}`,
        rowNumber: rowNum,
        logId: logId,
        isUpdate: true
      };
      
    } catch (error) {
      console.error('Error updating project item:', error);
      return {
        success: false,
        message: `Error updating item: ${error.message}`,
        rowNumber: null,
        logId: null,
        isUpdate: false
      };
    }
  },

  addProjectItem: function(itemData, logIdPrefix, logSheetName) {
    try {
      let originalRowNumber = null;
      
      if (itemData.originalRowNumber) {
        originalRowNumber = itemData.originalRowNumber;
      } else if (itemData.formData && itemData.formData.originalRowNumber) {
        originalRowNumber = itemData.formData.originalRowNumber;
      }
      
      if (originalRowNumber && originalRowNumber > 0) {
        return this.updateProjectItem({
          ...itemData,
          originalRowNumber: originalRowNumber
        }, logIdPrefix, logSheetName);
      }
      
      const sheet = this.getActiveSheet();
      
      if (!itemData || typeof itemData !== 'object') {
        throw new Error('Invalid item data provided');
      }

      const { description, quantity, dimensions, totalPrice, formData } = itemData;
      
      const lastRow = sheet.getLastRow();
      const nextRow = lastRow + 1;
      
      const logId = this.logFormData(formData, nextRow, logIdPrefix, logSheetName);
      
      const editInstruction = logId ? this.createEditInstruction(logId) : 'Edit';
      
      let rowData;
      let editColumnIndex;
      
      if (logIdPrefix === 'FAB') {
        const fabricationId = this.getNextFabricationId();
        rowData = [
          '',
          fabricationId,
          description || '',
          dimensions || '',
          '',
          totalPrice || 0,
          editInstruction
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(nextRow, 1, 1, rowData.length);
        range.setValues([rowData]);
        
        const priceCell = sheet.getRange(nextRow, 6);
        priceCell.setNumberFormat('$#,##0.00');
        
      } else if (logIdPrefix === 'APP') {
        const apparelId = this.getNextApparelId();
        const unitPrice = quantity && quantity > 0 ? (totalPrice / quantity) : 0;
        
        rowData = [
          '',
          apparelId,
          description || '',
          quantity || '',
          unitPrice,
          totalPrice || 0,
          editInstruction
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(nextRow, 1, 1, rowData.length);
        range.setValues([rowData]);
        
        const unitPriceCell = sheet.getRange(nextRow, 5);
        const totalPriceCell = sheet.getRange(nextRow, 6);
        unitPriceCell.setNumberFormat('$#,##0.00');
        totalPriceCell.setNumberFormat('$#,##0.00');
        
      } else if (logIdPrefix === 'PRT') {
        const printingId = this.getNextPrintingId();
        const unitPrice = quantity && quantity > 0 ? (totalPrice / quantity) : 0;
        
        rowData = [
          '',
          printingId,
          description || '',
          quantity || '',
          unitPrice,
          totalPrice || 0,
          editInstruction
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(nextRow, 1, 1, rowData.length);
        range.setValues([rowData]);
        
        const unitPriceCell = sheet.getRange(nextRow, 5);
        const totalPriceCell = sheet.getRange(nextRow, 6);
        unitPriceCell.setNumberFormat('$#,##0.00');
        totalPriceCell.setNumberFormat('$#,##0.00');
      }
      
      if (logId) {
        const editCell = sheet.getRange(nextRow, editColumnIndex);
        editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Production > Edit Selected Item`);
        editCell.setBackground('#e3f2fd');
        editCell.setFontColor('#1976d2');
        editCell.setFontWeight('bold');
      }
      
      return {
        success: true,
        message: `Item added to row ${nextRow}`,
        rowNumber: nextRow,
        logId: logId,
        isUpdate: false
      };
      
    } catch (error) {
      console.error('Error adding project item:', error);
      return {
        success: false,
        message: `Error adding item: ${error.message}`,
        rowNumber: null,
        logId: null,
        isUpdate: false
      };
    }
  }
};

// ============================================================================
// NICH DOCS FUNCTIONALITY
// ============================================================================

/**
 * An object to namespace all functions related to NICH Docs functionality
 */
const nichDocs = {
  
  /**
   * Creates the Profit & Loss sheet based on logged data
   * Writes only to specific mapped cells in the existing PL sheet
   */
  createProfitLoss: function() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      
      // Get the PL sheet (must already exist)
      const plSheet = spreadsheet.getSheetByName('PL');
      if (!plSheet) {
        SpreadsheetApp.getUi().alert('Error: PL sheet not found. Please create the PL sheet template first.');
        return;
      }
      
      // Get the Main sheet for revenue data
      const mainSheet = spreadsheet.getSheetByName('Main');
      if (!mainSheet) {
        SpreadsheetApp.getUi().alert('Error: Main sheet not found.');
        return;
      }
      
      // Collect revenue data from Main sheet
      const revenueData = this.collectRevenueFromMain(mainSheet);
      
      // Collect cost data from log sheets
      const costData = this.collectCostData();
      
      // Write revenue to mapped cells
      this.writeRevenueToPL(plSheet, revenueData);
      
      // Write costs to mapped cells
      this.writeCostsToPL(plSheet, costData);
      
      SpreadsheetApp.getUi().alert('Profit & Loss data updated successfully!');
      
    } catch (error) {
      console.error('Error creating Profit & Loss:', error);
      SpreadsheetApp.getUi().alert('Error creating Profit & Loss: ' + error.message);
    }
  },
  
  /**
   * Collects revenue data from Main sheet, Column F
   * Uses logged data to identify which rows belong to each category
   * @param {Sheet} mainSheet - The Main sheet
   * @returns {Object} Revenue data organized by category
   */
  collectRevenueFromMain: function(mainSheet) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    const revenueData = {
      printing: 0,
      fabrication: 0,
      apparel: 0
    };
    
    // Get all rows from the Main sheet
    const lastRow = mainSheet.getLastRow();
    if (lastRow < 2) {
      return revenueData;
    }
    
    // Get data from columns B (ID) and F (Total Price)
    const data = mainSheet.getRange(2, 2, lastRow - 1, 5).getValues();
    
    // Sum revenues by ID prefix
    data.forEach(row => {
      const id = row[0]; // Column B - ID
      const totalPrice = parseFloat(row[4]) || 0; // Column F - Total Price (offset by 4 from column B)
      
      if (id && typeof id === 'string') {
        if (id.startsWith('PR')) {
          revenueData.printing += totalPrice;
        } else if (id.startsWith('F')) {
          revenueData.fabrication += totalPrice;
        } else if (id.startsWith('AP')) {
          revenueData.apparel += totalPrice;
        }
      }
    });
    
    return revenueData;
  },
  
  /**
   * Writes revenue data to the mapped cells in PL sheet
   * @param {Sheet} plSheet - The PL sheet
   * @param {Object} revenueData - Revenue data by category
   */
  writeRevenueToPL: function(plSheet, revenueData) {
    // Revenue Mapping:
    // Printing = Row 8, Column D
    // Fabrication = Row 9, Column D
    // Apparel = Row 10, Column D
    
    plSheet.getRange(8, 4).setValue(revenueData.printing);
    plSheet.getRange(9, 4).setValue(revenueData.fabrication);
    plSheet.getRange(10, 4).setValue(revenueData.apparel);
    
    // Format as currency
    plSheet.getRange(8, 4, 3, 1).setNumberFormat('$#,##0.00');
  },
  
  /**
   * Writes cost data to the mapped cells in PL sheet
   * @param {Sheet} plSheet - The PL sheet
   * @param {Object} costData - Cost data by category and line item
   */
  writeCostsToPL: function(plSheet, costData) {
    // Printing Costs - Column G
    // Material Costs - Row 3
    // Ink Costs - Row 4
    // Equipment Costs - Row 5
    // Operator Labor Costs - Row 6
    // Design Labor Costs - Row 7
    
    plSheet.getRange(3, 7).setValue(costData.printing.materials);
    plSheet.getRange(4, 7).setValue(costData.printing.ink);
    plSheet.getRange(5, 7).setValue(costData.printing.equipment);
    plSheet.getRange(6, 7).setValue(costData.printing.operator);
    plSheet.getRange(7, 7).setValue(costData.printing.design);
    
    // Fabrication Costs - Column G
    // Material Costs - Row 11
    // Personnel Costs - Row 12
    // Component Costs - Row 13
    
    plSheet.getRange(11, 7).setValue(costData.fabrication.materials);
    plSheet.getRange(12, 7).setValue(costData.fabrication.labor);
    plSheet.getRange(13, 7).setValue(costData.fabrication.components);
    
    // Apparel Costs - Column G
    // Garment Costs - Row 17
    // Printing Costs - Row 18
    // Labor Costs - Row 19
    
    plSheet.getRange(17, 7).setValue(costData.apparel.garments);
    plSheet.getRange(18, 7).setValue(costData.apparel.printing);
    plSheet.getRange(19, 7).setValue(costData.apparel.labor);
    
    // Format all cost cells as currency
    const costRanges = [
      plSheet.getRange(3, 7, 5, 1),   // Printing costs
      plSheet.getRange(11, 7, 3, 1),  // Fabrication costs
      plSheet.getRange(17, 7, 3, 1)   // Apparel costs
    ];
    
    costRanges.forEach(range => {
      range.setNumberFormat('$#,##0.00');
    });
  },
  
  /**
   * Collects cost data from all log sheets
   * @returns {Object} Cost data organized by category and line item
   */
  collectCostData: function() {
    const costData = {
      printing: {
        materials: 0,
        ink: 0,
        equipment: 0,
        design: 0,
        operator: 0
      },
      fabrication: {
        materials: 0,
        labor: 0,
        components: 0
      },
      apparel: {
        garments: 0,
        printing: 0,
        labor: 0
      }
    };
    
    // Collect printing costs
    const printingCosts = this.getPrintingCosts();
    costData.printing.materials = printingCosts.materialCost + printingCosts.laminationCost;
    costData.printing.ink = printingCosts.inkCost;
    costData.printing.equipment = printingCosts.cuttingCost + printingCosts.equipmentCost;
    costData.printing.design = printingCosts.designCost;
    costData.printing.operator = printingCosts.operatorCost;
    
    // Collect fabrication costs
    const fabricationCosts = this.getFabricationCosts();
    costData.fabrication.materials = fabricationCosts.materialTotal;
    costData.fabrication.labor = fabricationCosts.personnelTotal;
    costData.fabrication.components = fabricationCosts.componentTotal;
    
    // Collect apparel costs
    const apparelCosts = this.getApparelCosts();
    costData.apparel.garments = apparelCosts.garmentTotal;
    costData.apparel.printing = apparelCosts.totalPrintCosts + apparelCosts.screenSetupCosts;
    costData.apparel.labor = apparelCosts.additionalOptionsCosts;
    
    return costData;
  },
  
  /**
   * Extracts printing costs from PrintingLog
   * @returns {Object} Printing cost breakdown
   */
  getPrintingCosts: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('PrintingLog');
    
    const costs = {
      materialCost: 0,
      laminationCost: 0,
      inkCost: 0,
      cuttingCost: 0,
      equipmentCost: 0,
      designCost: 0,
      operatorCost: 0
    };
    
    if (!logSheet) {
      return costs;
    }
    
    const dataRange = logSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Skip header row
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      
      try {
        const formData = JSON.parse(formDataJson);
        
        // Calculate costs based on the printing algorithm
        const calculatedCosts = this.calculatePrintingCosts(formData);
        
        costs.materialCost += calculatedCosts.materialCost;
        costs.laminationCost += calculatedCosts.laminationCost;
        costs.inkCost += calculatedCosts.inkCost;
        costs.cuttingCost += calculatedCosts.cuttingCost;
        costs.equipmentCost += calculatedCosts.equipmentCost;
        costs.designCost += calculatedCosts.designCost;
        costs.operatorCost += calculatedCosts.operatorCost;
        
      } catch (error) {
        console.error('Error parsing PrintingLog data:', error);
      }
    }
    
    return costs;
  },
  
  /**
   * Calculates printing costs based on form data (mirrors PrintingIndex.html logic)
   * @param {Object} formData - The printing form data
   * @returns {Object} Calculated costs
   */
  /**
 * Calculates printing costs based on form data (mirrors PrintingIndex.html logic)
 * @param {Object} formData - The printing form data
 * @returns {Object} Calculated costs
 */
calculatePrintingCosts: function(formData) {
  const costs = {
    materialCost: 0,
    laminationCost: 0,
    inkCost: 0,
    cuttingCost: 0,
    equipmentCost: 0,
    designCost: 0,
    operatorCost: 0
  };
  
  // Extract form data with defaults
  const qty = Number(formData.quantity) || 0;
  const artWidth = Number(formData.width) || 0;
  const artHeight = Number(formData.height) || 0;
  const materialName = formData.materialName || '';
  
  if (!materialName || qty <= 0 || artWidth <= 0 || artHeight <= 0) {
    return costs;
  }
  
  const bleed = 0.25;
  const spacing = 0.25;
  const artWidthTotal = artWidth + bleed;
  const artHeightTotal = artHeight + bleed;
  
  // Look up material from Materials sheet
  const material = this.getMaterialByName(materialName);
  
  if (!material) {
    console.warn('Material not found:', materialName);
    return costs;
  }
  
  let totalLinearFeet = 0;
  
  // Calculate material cost based on type (ROLL vs SHEET)
  if (material.type === 'ROLL') {
    const rollWidth = material.width;
    
    // Check if artwork fits on roll
    if (artWidthTotal > rollWidth && artHeightTotal > rollWidth) {
      console.warn('Artwork is wider than the selected roll:', materialName);
      return costs;
    }
    
    const linFtCost = material.costLinFt;
    
    // Calculate how many columns fit across the roll width
    const colsPortrait = Math.floor((rollWidth + spacing) / (artWidthTotal + spacing));
    const colsLandscape = Math.floor((rollWidth + spacing) / (artHeightTotal + spacing));
    let numColumns = Math.max(colsPortrait, colsLandscape, 1);
    
    // Determine layout row height based on orientation
    let layoutRow = artHeightTotal;
    if (colsLandscape > colsPortrait) {
      layoutRow = artWidthTotal;
    }
    
    // Calculate rows needed
    const numRows = Math.ceil(qty / numColumns);
    
    // Calculate total linear inches and convert to feet
    const totalLinearInches = (numRows * layoutRow) + ((numRows - 1) * spacing);
    totalLinearFeet = totalLinearInches / 12;
    
    // Add 2.5 feet buffer and calculate cost
    costs.materialCost = (totalLinearFeet + 2.5) * linFtCost;
    
  } else if (material.type === 'SHEET') {
    const sheetWidth = material.width;
    const sheetHeight = material.height;
    const sheetCost = material.costSheet;
    
    // Check if artwork fits on sheet in either orientation
    const artFitsPortrait = (artWidthTotal <= sheetWidth && artHeightTotal <= sheetHeight);
    const artFitsLandscape = (artWidthTotal <= sheetHeight && artHeightTotal <= sheetWidth);
    
    if (!artFitsPortrait && !artFitsLandscape) {
      console.warn('Artwork is larger than the selected sheet:', materialName);
      return costs;
    }
    
    // Calculate pieces per sheet for both orientations
    const perSheet1 = artFitsPortrait 
      ? Math.floor((sheetWidth + spacing) / (artWidthTotal + spacing)) * 
        Math.floor((sheetHeight + spacing) / (artHeightTotal + spacing)) 
      : 0;
    const perSheet2 = artFitsLandscape 
      ? Math.floor((sheetWidth + spacing) / (artHeightTotal + spacing)) * 
        Math.floor((sheetHeight + spacing) / (artWidthTotal + spacing)) 
      : 0;
    
    const piecesPerSheet = Math.max(perSheet1, perSheet2, 1);
    const totalSheets = Math.ceil(qty / piecesPerSheet);
    
    // Calculate cost with half-sheet minimum
    let calculatedMaterialCost = totalSheets * sheetCost;
    const halfSheetCost = sheetCost * 0.5;
    costs.materialCost = Math.max(calculatedMaterialCost, halfSheetCost);
  }
  
  // Calculate total artwork square footage
  const singlePieceSqFt = (artWidthTotal * artHeightTotal) / 144;
  let totalArtworkSqFt = singlePieceSqFt * qty;
  if (formData.doubleSided) {
    totalArtworkSqFt *= 2;
  }
  
  // Calculate total artwork perimeter for cutting
  const totalArtworkPerimeter = (artWidth * 2 + artHeight * 2) * qty;
  
  // Calculate time-based values
  const printTimeHours = (totalArtworkSqFt / 0.83) / 60;
  let cutTimeHours = (totalArtworkPerimeter / 120) / 60;
  if (formData.complexShape) {
    cutTimeHours *= 1.5;
  }
  const ripTimeHours = (totalArtworkSqFt / 20.52) / 60;
  const printComputeTimeHours = (totalArtworkSqFt / 6.2) / 60;
  
  // Calculate labor times
  const designTimeInHours = this.getTimeInHours(formData.designTime, formData.designTimeUnit);
  const laborDecalsTimeInHours = this.getTimeInHours(formData.laborDecalsTime, formData.laborDecalsTimeUnit);
  const laborFinishingTimeInHours = this.getTimeInHours(formData.laborFinishingTime, formData.laborFinishingTimeUnit);
  const laborInstallingTimeInHours = this.getTimeInHours(formData.laborInstallingTime, formData.laborInstallingTimeUnit);
  
  const manualOperatorTimeInHours = laborDecalsTimeInHours + laborFinishingTimeInHours + laborInstallingTimeInHours;
  const totalProjectRunTimeHours = printTimeHours + cutTimeHours + ripTimeHours + printComputeTimeHours;
  
  // Calculate costs
  costs.inkCost = totalArtworkSqFt * 0.165;
  costs.cuttingCost = cutTimeHours * 25.00;
  
  // Design cost: base + manual
  const baseDesignCost = (totalArtworkSqFt / 25) * 0.0625 * 60.00;
  const manualDesignCost = designTimeInHours * 60.00;
  costs.designCost = baseDesignCost + manualDesignCost;
  
  // Lamination cost
  if (formData.lamination) {
    if (material.type === 'ROLL') {
      costs.laminationCost = totalLinearFeet * 1.02;
    } else {
      costs.laminationCost = totalArtworkSqFt * 0.2267;
    }
  }
  
  // Equipment and operator costs
  costs.equipmentCost = totalProjectRunTimeHours * 4.95;
  costs.operatorCost = (totalProjectRunTimeHours + manualOperatorTimeInHours) * 28.00;
  
  return costs;
},

/**
 * Looks up a material by name from the Materials sheet
 * @param {string} materialName - Name of the material to find
 * @returns {Object|null} Material object with type, width, height, costSheet, costLinFt
 */
getMaterialByName: function(materialName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Materials');
  
  if (!sheet) {
    console.error('Materials sheet not found');
    return null;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }
  
  const data = sheet.getRange('A2:P' + lastRow).getValues();
  
  for (let row of data) {
    const name = row[1]; // Column B
    const primaryCategory = row[4]; // Column E
    
    // Check if this is a PRINT material with matching name
    if (name && name.toString().trim() === materialName && 
        primaryCategory && primaryCategory.toString().toUpperCase().includes('PRINT')) {
      
      const type = row[6] ? row[6].toString().trim().toUpperCase() : 'SHEET'; // Column G
      const width = parseFloat(row[7]) || 0; // Column H (inches)
      const length = parseFloat(row[8]) || 0; // Column I (feet for ROLL, inches for SHEET)
      let unitCost = row[9]; // Column J
      
      // Parse unit cost
      if (unitCost && typeof unitCost === 'string') {
        unitCost = parseFloat(unitCost.replace(/[^0-9.-]+/g, '')) || 0;
      } else if (typeof unitCost !== 'number') {
        unitCost = 0;
      }
      
      // Calculate cost per linear foot for ROLL materials
      let costLinFt = 0;
      if (type === 'ROLL' && length > 0) {
        costLinFt = unitCost / length;
      }
      
      return {
        name: name.toString().trim(),
        type: type,
        width: width,
        height: length, // For SHEET this is height in inches, for ROLL this is length in feet
        costSheet: unitCost,
        costLinFt: costLinFt
      };
    }
  }
  
  return null;
},
  
  /**
   * Helper function to convert time to hours
   */
  getTimeInHours: function(time, unit) {
    const numTime = Number(time) || 0;
    return unit === 'Minutes' ? numTime / 60 : numTime;
  },
  
  /**
   * Extracts fabrication costs from FabricationLog
   * @returns {Object} Fabrication cost breakdown
   */
  getFabricationCosts: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('FabricationLog');
    
    const costs = {
      materialTotal: 0,
      personnelTotal: 0,
      componentTotal: 0
    };
    
    if (!logSheet) {
      return costs;
    }
    
    const dataRange = logSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Skip header row
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      
      try {
        const formData = JSON.parse(formDataJson);
        
        // Sum materials
        if (formData.materials && Array.isArray(formData.materials)) {
          formData.materials.forEach(material => {
            costs.materialTotal += material.total || 0;
          });
        }
        
        // Sum personnel
        if (formData.personnel && Array.isArray(formData.personnel)) {
          formData.personnel.forEach(person => {
            costs.personnelTotal += person.total || 0;
          });
        }
        
        // Sum components
        if (formData.components && Array.isArray(formData.components)) {
          formData.components.forEach(component => {
            costs.componentTotal += component.total || 0;
          });
        }
        
      } catch (error) {
        console.error('Error parsing FabricationLog data:', error);
      }
    }
    
    return costs;
  },
  
  /**
   * Extracts apparel costs from ApparelLog
   * @returns {Object} Apparel cost breakdown
   */
  getApparelCosts: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('ApparelLog');
    
    const costs = {
      garmentTotal: 0,
      totalPrintCosts: 0,
      screenSetupCosts: 0,
      additionalOptionsCosts: 0
    };
    
    if (!logSheet) {
      return costs;
    }
    
    const dataRange = logSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Skip header row
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      
      try {
        const formData = JSON.parse(formDataJson);
        
        // Calculate costs based on the apparel algorithm
        const calculatedCosts = this.calculateApparelCosts(formData);
        
        costs.garmentTotal += calculatedCosts.garmentTotal;
        costs.totalPrintCosts += calculatedCosts.totalPrintCosts;
        costs.screenSetupCosts += calculatedCosts.screenSetupCosts;
        costs.additionalOptionsCosts += calculatedCosts.additionalOptionsCosts;
        
      } catch (error) {
        console.error('Error parsing ApparelLog data:', error);
      }
    }
    
    return costs;
  },
  
  /**
   * Calculates apparel costs based on form data (mirrors ApparelIndex.html logic)
   * @param {Object} formData - The apparel form data
   * @returns {Object} Calculated costs
   */
  calculateApparelCosts: function(formData) {
    const costs = {
      garmentTotal: 0,
      totalPrintCosts: 0,
      screenSetupCosts: 0,
      additionalOptionsCosts: 0
    };
    
    const quantity = Number(formData.quantity) || 0;
    const garmentUnitCost = Number(formData.garmentUnitCost) || 0;
    
    // Garment total
    costs.garmentTotal = quantity * garmentUnitCost;
    
    // Print costs and screen setup (simplified pricing tier logic)
    const frontColors = Number(formData.frontColors) || 0;
    const backColors = Number(formData.backColors) || 0;
    let totalColors = frontColors + backColors;
    
    // Add additional location colors
    if (formData.additionalLocations && Array.isArray(formData.additionalLocations)) {
      formData.additionalLocations.forEach(loc => {
        totalColors += Number(loc.colors) || 0;
      });
    }
    
    // Screen setup cost
    costs.screenSetupCosts = totalColors * 13; // $13 per color
    
    // Print costs (simplified - would need full pricing tier logic)
    // For now, use a basic calculation
    if (totalColors > 0) {
      costs.totalPrintCosts = quantity * totalColors * 2; // Simplified estimate
    }
    
    // Additional options costs
    const constants = {
      oversize: 1,
      colorChange: 5,
      polyNylonSpandexMesh: 0.25,
      metallicShimmerInk: 0.50,
      glow: 1.25,
      fleece: 0.25,
      designLabor: 60
    };
    
    if (formData.additionalOptions) {
      const options = formData.additionalOptions;
      
      if (options.oversized) {
        costs.additionalOptionsCosts += constants.oversize * quantity;
      }
      
      if (options.colorChange && options.colorChange.enabled) {
        const colorChangeQty = Number(options.colorChange.quantity) || 1;
        costs.additionalOptionsCosts += constants.colorChange * colorChangeQty;
      }
      
      if (options.polyNylon) {
        costs.additionalOptionsCosts += constants.polyNylonSpandexMesh * quantity;
      }
      
      if (options.metallicShimmer) {
        costs.additionalOptionsCosts += constants.metallicShimmerInk * quantity;
      }
      
      if (options.glow) {
        costs.additionalOptionsCosts += constants.glow * quantity;
      }
      
      if (options.fleece) {
        costs.additionalOptionsCosts += constants.fleece * quantity;
      }
      
      if (options.designLabor && options.designLabor.enabled) {
        const laborQty = Number(options.designLabor.quantity) || 0;
        const unit = options.designLabor.unit || 'hours';
        let hours = laborQty;
        if (unit === 'minutes') {
          hours = laborQty / 60;
        }
        costs.additionalOptionsCosts += constants.designLabor * hours;
      }
    }
    
    return costs;
  },
  
  /**
   * Placeholder for Create Estimate functionality
   */
  createEstimate: function() {
    SpreadsheetApp.getUi().alert('Create Estimate functionality coming soon!');
  },
  
  /**
   * Placeholder for Create Invoice functionality
   */
  createInvoice: function() {
    SpreadsheetApp.getUi().alert('Create Invoice functionality coming soon!');
  },
  
  /**
   * Creates Bill of Materials from logged data
   * Collects material quantities and prepares data for Google Docs template
   */
  createBillOfMaterials: function() {
    try {
      // Collect BOM data from all log sheets
      const bomData = this.collectBOMData();
      
      // TODO: Replace with actual template document ID
      // For now, display the data structure that would be sent to the template
      this.displayBOMData(bomData);
      
    } catch (error) {
      console.error('Error creating Bill of Materials:', error);
      SpreadsheetApp.getUi().alert('Error creating Bill of Materials: ' + error.message);
    }
  },
  
  /**
   * Collects and aggregates BOM data from all log sheets
   * @returns {Object} BOM data organized by category with aggregated line items
   */
  collectBOMData: function() {
    const bomData = {
      printing: [],
      fabrication: [],
      apparel: []
    };
    
    // Collect printing materials
    const printingItems = this.collectPrintingBOM();
    if (printingItems.length > 0) {
      bomData.printing = printingItems;
    }
    
    // Collect fabrication materials
    const fabricationItems = this.collectFabricationBOM();
    if (fabricationItems.length > 0) {
      bomData.fabrication = fabricationItems;
    }
    
    // Collect apparel materials
    const apparelItems = this.collectApparelBOM();
    if (apparelItems.length > 0) {
      bomData.apparel = apparelItems;
    }
    
    return bomData;
  },
  
  /**
   * Collects printing BOM items from PrintingLog
   * @returns {Array} Array of printing line items
   */
  collectPrintingBOM: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('PrintingLog');
    
    if (!logSheet) {
      return [];
    }
    
    const materialMap = new Map(); // Key: material name, Value: {quantity, unitOfMeasure, vendor}
    
    const dataRange = logSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Skip header row
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      
      try {
        const formData = JSON.parse(formDataJson);
        
        // Extract material information from printing data
        const materialName = formData.materialName;
        if (!materialName) continue;
        
        // Calculate material quantity based on form data
        const materialInfo = this.calculatePrintingMaterialQuantity(formData);
        
        if (materialInfo && materialInfo.quantity > 0) {
          const key = materialName;
          
          if (materialMap.has(key)) {
            // Sum quantities for duplicate materials
            const existing = materialMap.get(key);
            existing.quantity += materialInfo.quantity;
          } else {
            materialMap.set(key, {
              description: materialName,
              quantity: materialInfo.quantity,
              unitOfMeasure: materialInfo.unitOfMeasure,
              vendor: materialInfo.vendor || '',
              status: ''
            });
          }
        }
        
      } catch (error) {
        console.error('Error parsing PrintingLog data:', error);
      }
    }
    
    // Convert map to array
    return Array.from(materialMap.values());
  },
  
  /**
   * Calculates material quantity for a printing item
   * @param {Object} formData - Printing form data
   * @returns {Object} Material quantity info
   */
  calculatePrintingMaterialQuantity: function(formData) {
    // This would need access to the materials data to determine if it's a sheet or roll
    // For now, we'll look up the material from the Materials sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const materialsSheet = spreadsheet.getSheetByName('Materials');
    
    if (!materialsSheet) {
      return null;
    }
    
    const materialName = formData.materialName;
    const qty = Number(formData.quantity) || 0;
    const artWidth = Number(formData.width) || 0;
    const artHeight = Number(formData.height) || 0;
    const bleed = 0.25;
    const spacing = 0.25;
    const artWidthTotal = artWidth + bleed;
    const artHeightTotal = artHeight + bleed;
    
    // Find material in Materials sheet
    const lastRow = materialsSheet.getLastRow();
    const materialsData = materialsSheet.getRange('A2:P' + lastRow).getValues();
    
    let materialType = 'SHEET';
    let materialWidth = 0;
    let materialHeight = 0;
    let vendor = '';
    
    for (let row of materialsData) {
      const name = row[1]; // Column B
      if (name && name.toString().trim() === materialName) {
        materialType = row[6] ? row[6].toString().trim().toUpperCase() : 'SHEET'; // Column G
        materialWidth = parseFloat(row[7]) || 0; // Column H
        materialHeight = parseFloat(row[8]) || 0; // Column I
        vendor = row[11] || ''; // Column L - Vendor
        break;
      }
    }
    
    let quantity = 0;
    let unitOfMeasure = '';
    
    if (materialType === 'ROLL') {
      // Calculate linear feet needed
      const rollWidth = materialWidth;
      const colsPortrait = Math.floor((rollWidth + spacing) / (artWidthTotal + spacing));
      const colsLandscape = Math.floor((rollWidth + spacing) / (artHeightTotal + spacing));
      let numColumns = Math.max(colsPortrait, colsLandscape, 1);
      let layoutRow = artHeightTotal;
      if (colsLandscape > colsPortrait) layoutRow = artWidthTotal;
      const numRows = Math.ceil(qty / numColumns);
      const totalLinearInches = (numRows * layoutRow) + ((numRows - 1) * spacing);
      quantity = (totalLinearInches / 12) + 2.5; // Add 2.5 feet buffer
      unitOfMeasure = 'Lin Feet';
    } else if (materialType === 'SHEET') {
      // Calculate number of sheets needed
      const sheetWidth = materialWidth;
      const sheetHeight = materialHeight;
      const artFitsPortrait = (artWidthTotal <= sheetWidth && artHeightTotal <= sheetHeight);
      const artFitsLandscape = (artWidthTotal <= sheetHeight && artHeightTotal <= sheetWidth);
      
      const perSheet1 = artFitsPortrait ? 
        Math.floor((sheetWidth + spacing) / (artWidthTotal + spacing)) * 
        Math.floor((sheetHeight + spacing) / (artHeightTotal + spacing)) : 0;
      const perSheet2 = artFitsLandscape ? 
        Math.floor((sheetWidth + spacing) / (artHeightTotal + spacing)) * 
        Math.floor((sheetHeight + spacing) / (artWidthTotal + spacing)) : 0;
      const piecesPerSheet = Math.max(perSheet1, perSheet2, 1);
      quantity = Math.ceil(qty / piecesPerSheet);
      unitOfMeasure = 'Sheets';
    }
    
    return {
      quantity: Math.ceil(quantity), // Round up to whole units
      unitOfMeasure: unitOfMeasure,
      vendor: vendor
    };
  },
  
  /**
   * Collects fabrication BOM items from FabricationLog
   * @returns {Array} Array of fabrication line items
   */
  collectFabricationBOM: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('FabricationLog');
    
    if (!logSheet) {
      return [];
    }
    
    const materialMap = new Map(); // For materials
    const personnelMap = new Map(); // For personnel (labor)
    const componentMap = new Map(); // For components
    
    const dataRange = logSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Skip header row
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      
      try {
        const formData = JSON.parse(formDataJson);
        
        // Process materials
        if (formData.materials && Array.isArray(formData.materials)) {
          formData.materials.forEach(material => {
            const key = material.name;
            const qty = Number(material.quantity) || 0;
            
            if (materialMap.has(key)) {
              const existing = materialMap.get(key);
              existing.quantity += qty;
            } else {
              // Look up vendor from Materials sheet
              const vendor = this.getMaterialVendor(material.name);
              materialMap.set(key, {
                description: material.name,
                quantity: qty,
                unitOfMeasure: 'Units', // Default, could be enhanced
                vendor: vendor,
                status: ''
              });
            }
          });
        }
        
        // Process personnel (labor)
        if (formData.personnel && Array.isArray(formData.personnel)) {
          formData.personnel.forEach(person => {
            const key = person.name;
            const days = Number(person.days) || 0;
            const hours = Number(person.hours) || 0;
            const totalHours = days * hours;
            
            if (personnelMap.has(key)) {
              const existing = personnelMap.get(key);
              existing.quantity += totalHours;
            } else {
              personnelMap.set(key, {
                description: person.name,
                quantity: totalHours,
                unitOfMeasure: 'Hours',
                vendor: '',
                status: ''
              });
            }
          });
        }
        
        // Process components
        if (formData.components && Array.isArray(formData.components)) {
          formData.components.forEach(component => {
            const key = component.description;
            const qty = Number(component.quantity) || 0;
            
            if (componentMap.has(key)) {
              const existing = componentMap.get(key);
              existing.quantity += qty;
            } else {
              componentMap.set(key, {
                description: component.description,
                quantity: qty,
                unitOfMeasure: 'Units',
                vendor: '',
                status: ''
              });
            }
          });
        }
        
      } catch (error) {
        console.error('Error parsing FabricationLog data:', error);
      }
    }
    
    // Combine all items: materials, personnel, then components
    const allItems = [
      ...Array.from(materialMap.values()),
      ...Array.from(personnelMap.values()),
      ...Array.from(componentMap.values())
    ];
    
    return allItems;
  },
  
  /**
   * Looks up vendor for a material from the Materials sheet
   * @param {string} materialName - Name of the material
   * @returns {string} Vendor name
   */
  getMaterialVendor: function(materialName) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const materialsSheet = spreadsheet.getSheetByName('Materials');
    
    if (!materialsSheet) {
      return '';
    }
    
    const lastRow = materialsSheet.getLastRow();
    const materialsData = materialsSheet.getRange('A2:P' + lastRow).getValues();
    
    for (let row of materialsData) {
      const name = row[1]; // Column B
      if (name && name.toString().trim() === materialName) {
        return row[11] || ''; // Column L - Vendor
      }
    }
    
    return '';
  },
  
  /**
   * Collects apparel BOM items from ApparelLog
   * @returns {Array} Array of apparel line items
   */
  collectApparelBOM: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('ApparelLog');
    
    if (!logSheet) {
      return [];
    }
    
    const garmentMap = new Map(); // For garments
    
    const dataRange = logSheet.getDataRange();
    const values = dataRange.getValues();
    
    // Skip header row
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      
      try {
        const formData = JSON.parse(formDataJson);
        
        // Extract garment information
        const garmentName = formData.garment;
        const quantity = Number(formData.quantity) || 0;
        
        if (garmentName && quantity > 0) {
          const key = garmentName;
          
          if (garmentMap.has(key)) {
            const existing = garmentMap.get(key);
            existing.quantity += quantity;
          } else {
            garmentMap.set(key, {
              description: garmentName,
              quantity: quantity,
              unitOfMeasure: 'Units',
              vendor: '', // Could be enhanced to look up vendor
              status: ''
            });
          }
        }
        
      } catch (error) {
        console.error('Error parsing ApparelLog data:', error);
      }
    }
    
    return Array.from(garmentMap.values());
  },
  
  /**
   * Creates Bill of Materials from logged data using hybrid template approach
   * Uses a Google Docs template for branding and programmatically builds tables
   */
  createBillOfMaterials: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      
      // Template Document ID
      const TEMPLATE_ID = '1mAbdWVwLGn8v146oqmNVxxQyYO1zX3CtAUN9wXj6rWw';
      
      // Collect BOM data from all log sheets
      const bomData = this.collectBOMData();
      
      // Check if there's any data to include
      if (bomData.printing.length === 0 && 
          bomData.fabrication.length === 0 && 
          bomData.apparel.length === 0) {
        ui.alert(
          'No Data Found',
          'No items found in any category. Please ensure you have logged items before creating a BOM.',
          ui.ButtonSet.OK
        );
        return;
      }
      
      // Get project info (TODO: Customize this based on your needs)
      const projectInfo = this.getProjectInfo();
      
      // Create document from template
      const docUrl = this.createBOMDocument(TEMPLATE_ID, bomData, projectInfo);
      
      // Show success message
      const response = ui.alert(
        'Bill of Materials Created',
        'Your Bill of Materials has been created successfully!\n\n' +
        'Would you like to open it now?',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        // Open the document in a new browser tab
        const htmlOutput = HtmlService.createHtmlOutput(
          '<script>window.open("' + docUrl + '", "_blank"); google.script.host.close();</script>'
        );
        ui.showModelessDialog(htmlOutput, 'Opening Document...');
      }
      
    } catch (error) {
      console.error('Error creating Bill of Materials:', error);
      SpreadsheetApp.getUi().alert('Error creating Bill of Materials: ' + error.message);
    }
  },
  
  /**
   * Gets project information for the BOM header
   * TODO: Customize this based on where you store project info
   * @returns {Object} Project information
   */
  getProjectInfo: function() {
    const ui = SpreadsheetApp.getUi();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // OPTION 1: Get from specific cells (uncomment and customize)
    /*
    const infoSheet = spreadsheet.getSheetByName('Project Info');
    return {
      projectNumber: infoSheet.getRange('B2').getValue(),
      clientName: infoSheet.getRange('B3').getValue(),
      date: new Date().toLocaleDateString()
    };
    */
    
    // OPTION 2: Prompt user for input (current default)
    const projectNumber = ui.prompt(
      'Project Number',
      'Enter the project number (e.g., 19171LV):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (projectNumber.getSelectedButton() !== ui.Button.OK) {
      throw new Error('BOM creation cancelled by user');
    }
    
    const clientName = ui.prompt(
      'Client Name',
      'Enter the client name:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (clientName.getSelectedButton() !== ui.Button.OK) {
      throw new Error('BOM creation cancelled by user');
    }
    
    return {
      projectNumber: projectNumber.getResponseText(),
      clientName: clientName.getResponseText(),
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy')
    };
  },
  
  /**
   * Creates the BOM document from template
   * @param {string} templateId - Google Docs template ID
   * @param {Object} bomData - BOM data structure
   * @param {Object} projectInfo - Project header information
   * @returns {string} URL of created document
   */
  createBOMDocument: function(templateId, bomData, projectInfo) {
    // Copy the template
    const templateFile = DriveApp.getFileById(templateId);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const newFileName = `BOM - ${projectInfo.projectNumber} - ${timestamp}`;
    const newFile = templateFile.makeCopy(newFileName);
    
    // Open the document for editing
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();
    
    // Replace simple placeholders
    body.replaceText('\\{\\{projectNumber\\}\\}', projectInfo.projectNumber);
    body.replaceText('\\{\\{clientName\\}\\}', projectInfo.clientName);
    body.replaceText('\\{\\{date\\}\\}', projectInfo.date);
    
    // Process each section
    this.insertBOMSection(body, 'PRINTING', bomData.printing);
    this.insertBOMSection(body, 'FABRICATION', bomData.fabrication);
    this.insertBOMSection(body, 'APPAREL', bomData.apparel);
    
    // Save and close
    doc.saveAndClose();
    
    return newFile.getUrl();
  },
  
  /**
   * Inserts or removes a BOM section based on data availability
   * @param {Body} body - Document body
   * @param {string} sectionName - Section name (PRINTING, FABRICATION, APPAREL)
   * @param {Array} items - Array of items for this section
   */
  insertBOMSection: function(body, sectionName, items) {
    const placeholder = `{{${sectionName}_TABLE}}`;
    
    if (items.length === 0) {
      // Remove entire section if no items
      this.removeBOMSection(body, sectionName, placeholder);
      return;
    }
    
    // Find the placeholder
    const placeholderSearch = body.findText(placeholder);
    
    if (!placeholderSearch) {
      console.warn(`Placeholder ${placeholder} not found in template`);
      return;
    }
    
    // Get the element containing the placeholder
    const element = placeholderSearch.getElement();
    const parent = element.getParent();
    const index = body.getChildIndex(parent);
    
    // Remove the placeholder paragraph
    parent.removeFromParent();
    
    // Create table data
    const tableData = [
      ['Description', 'Quantity', 'Unit of Measure', 'Vendor', 'Status']
    ];
    
    items.forEach(item => {
      tableData.push([
        item.description,
        item.quantity.toString(),
        item.unitOfMeasure,
        item.vendor,
        item.status
      ]);
    });
    
    // Insert table at the placeholder position
    const table = body.insertTable(index, tableData);
    
    // Style the table
    this.styleBOMTable(table);
  },
  
  /**
   * Removes a BOM section that has no items
   * @param {Body} body - Document body
   * @param {string} sectionName - Section name
   * @param {string} placeholder - Placeholder text
   */
  removeBOMSection: function(body, sectionName, placeholder) {
    // Remove the placeholder
    const placeholderSearch = body.findText(placeholder);
    if (placeholderSearch) {
      const element = placeholderSearch.getElement();
      const parent = element.getParent();
      parent.removeFromParent();
    }
    
    // Remove the section heading
    // Look for the capitalized version or title case version
    const headingVariations = [
      sectionName, // "PRINTING"
      sectionName.charAt(0) + sectionName.slice(1).toLowerCase(), // "Printing"
      sectionName.toLowerCase() // "printing"
    ];
    
    for (let heading of headingVariations) {
      const headingSearch = body.findText(heading);
      if (headingSearch) {
        const element = headingSearch.getElement();
        const parent = element.getParent();
        // Only remove if this is the entire content of the paragraph
        const text = parent.asText().getText().trim();
        if (text === heading) {
          parent.removeFromParent();
          break;
        }
      }
    }
  },
  
  /**
   * Applies styling to a BOM table per design specifications
   * @param {Table} table - The table to style
   */
  styleBOMTable: function(table) {
    const numRows = table.getNumRows();
    const headerRow = table.getRow(0);
    const numCols = headerRow.getNumCells();
    
    // Google Docs color constants (approximations)
    // Dark Gray 4: #434343
    // Dark Gray 2: #666666
    const DARK_GRAY_4 = '#434343';
    const DARK_GRAY_2 = '#666666';
    const BLACK = '#000000';
    
    // Set column widths (as percentage of table width)
    // Column widths: 50%, 15%, 15%, 15%, 5%
    const columnWidths = [50, 10, 15, 15, 10];
    
    // Calculate actual widths in points (assuming 468 points total width for standard page)
    // Google Docs standard page width with margins is approximately 468 points
    const totalWidth = 550;
    const columnWidthsInPoints = columnWidths.map(percent => (percent / 100) * totalWidth);
    
    // Set column widths by setting width on cells in first row
    for (let i = 0; i < numCols && i < columnWidthsInPoints.length; i++) {
      headerRow.getCell(i).setWidth(columnWidthsInPoints[i]);
    }
    
    // Style header row
    for (let i = 0; i < numCols; i++) {
      const cell = headerRow.getCell(i);
      
      // No background color for header
      cell.setBackgroundColor(null);
      
      // Header text styling
      const text = cell.getChild(0).asText();
      text.setFontFamily('Inter');
      text.setFontSize(9);
      text.setBold(true);
      text.setForegroundColor(DARK_GRAY_4);
      
      // Cell padding
      cell.setPaddingTop(6);
      cell.setPaddingBottom(6);
      cell.setPaddingLeft(8);
      cell.setPaddingRight(8);
    }
    
    // Style data rows
    for (let i = 1; i < numRows; i++) {
      const row = table.getRow(i);
      for (let j = 0; j < numCols; j++) {
        const cell = row.getCell(j);
        
        // Data text styling
        const text = cell.getChild(0).asText();
        text.setFontFamily('Inter');
        text.setFontSize(9);
        text.setBold(false);
        text.setForegroundColor(BLACK);
        
        // No background color
        cell.setBackgroundColor(null);
        
        // Cell padding
        cell.setPaddingTop(4);
        cell.setPaddingBottom(4);
        cell.setPaddingLeft(8);
        cell.setPaddingRight(8);
      }
    }
    
    // Set table borders: Perimeter only, no internal vertical lines
    // Border width: 1pt
    table.setBorderWidth(1);
    table.setBorderColor(DARK_GRAY_2);
    
    // Remove internal vertical borders by setting them to 0 width
    // We need to iterate through cells and remove their right borders
    // for (let i = 0; i < numRows; i++) {
    //   const row = table.getRow(i);
    //   for (let j = 0; j < numCols - 1; j++) { // All columns except last
    //     const cell = row.getCell(j);
    //     // Google Docs doesn't have direct cell border control, 
    //     // but we can style the table to minimize internal borders
    //     // by using editAsText and paragraph formatting
    //   }
    // }
    
    // Note: Google Docs API has limitations on selective border removal
    // The perimeter border is set, internal horizontal borders will remain
    // but vertical internal borders cannot be removed via Apps Script
    // This is a known limitation of the DocumentApp API
  },
};

// ============================================================================
// MENU CREATION FUNCTIONS
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Production Menu (existing)
  ui.createMenu('Production')
      .addItem('Fabrication', 'openFabricationApp')
      .addItem('Apparel', 'openApparelApp')
      .addItem('Printing', 'openPrintingApp')
      .addSeparator()
      .addItem('Edit Selected Item', 'editSelectedItem')
      .addToUi();
  
  // NICH Docs Menu (new)
  ui.createMenu('NICH Docs')
      .addItem('Create Estimate', 'createEstimate')
      .addItem('Create Invoice', 'createInvoice')
      .addItem('Create Profit & Loss', 'createProfitLoss')
      .addItem('Create Bill of Materials', 'createBillOfMaterials')
      .addToUi();
}

// ============================================================================
// GLOBAL FUNCTION WRAPPERS
// ============================================================================

// Production Menu Functions
function openFabricationApp() {
  fabricationApp.showDialog();
}

function openApparelApp() {
  apparelApp.showDialog();
}

function openPrintingApp() {
  printingApp.showDialog();
}

function getMaterials() {
  return fabricationApp.getMaterials();
}

function getPersonnel() {
  return fabricationApp.getPersonnel();
}

function getPrintingMaterials() {
  return printingApp.getMaterials();
}

function addFabricationToProject(fabricationData) {
  return fabricationApp.addToProject(fabricationData);
}

function addApparelToProject(apparelData) {
  return apparelApp.addToProject(apparelData);
}

function addPrintingToProject(printingData) {
  return printingApp.addToProject(printingData);
}

function openFabricationAppForEdit(logId) {
  return fabricationApp.openForEdit(logId);
}

function openApparelAppForEdit(logId) {
  return apparelApp.openForEdit(logId);
}

function openPrintingAppForEdit(logId) {
  return printingApp.openForEdit(logId);
}

function getLoggedFormData(logId, logSheetName) {
  return projectSheet.getLoggedFormData(logId, logSheetName);
}

function editSelectedItem() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeCell = sheet.getActiveCell();
    
    const column = activeCell.getColumn();
    if (column !== 7) {
      SpreadsheetApp.getUi().alert(
        'Edit Item', 
        'Please select an "Edit" cell first, then try again.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const cellNote = activeCell.getNote();
    if (!cellNote || !cellNote.includes('LogID:')) {
      SpreadsheetApp.getUi().alert(
        'Edit Item', 
        'No edit data found for this item.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const logIdMatch = cellNote.match(/LogID:\s*([^\n\r]+)/);
    if (!logIdMatch) {
      SpreadsheetApp.getUi().alert(
        'Edit Item', 
        'Could not find LogID in the selected cell.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const logId = logIdMatch[1].trim();
    
    if (logId.startsWith('FAB_')) {
      fabricationApp.openForEdit(logId);
    } else if (logId.startsWith('APP_')) {
      apparelApp.openForEdit(logId);
    } else if (logId.startsWith('PRT_')) {
      printingApp.openForEdit(logId);
    } else {
      SpreadsheetApp.getUi().alert(
        'Edit Item', 
        'Unknown item type. Cannot determine which editor to open.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('Error in editSelectedItem:', error);
    SpreadsheetApp.getUi().alert(
      'Error', 
      'An error occurred while trying to edit the item: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// NICH Docs Menu Functions
function createEstimate() {
  nichDocs.createEstimate();
}

function createInvoice() {
  nichDocs.createInvoice();
}

function createProfitLoss() {
  nichDocs.createProfitLoss();
}

function createBillOfMaterials() {
  nichDocs.createBillOfMaterials();
}
