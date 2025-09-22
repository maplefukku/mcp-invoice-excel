#!/usr/bin/env node

import ExcelJS from 'exceljs';

/**
 * Example of how to create a new invoice while preserving ALL formatting
 * from the Japanese invoice template
 */

class JapaneseInvoiceFormatter {
  constructor() {
    this.template = null;
    this.workbook = null;
    this.worksheet = null;
  }

  /**
   * Load template and create a working copy
   */
  async loadTemplate(templatePath) {
    this.workbook = new ExcelJS.Workbook();
    this.template = new ExcelJS.Workbook();

    // Load the original template
    await this.template.xlsx.readFile(templatePath);

    // Create a complete copy with all formatting preserved
    const templateSheet = this.template.worksheets[0];
    this.worksheet = this.workbook.addWorksheet(templateSheet.name);

    // Copy ALL worksheet properties
    this.copyWorksheetProperties(templateSheet, this.worksheet);

    // Copy ALL cells with complete formatting
    await this.copyAllCellsWithFormatting(templateSheet, this.worksheet);

    return this.worksheet;
  }

  /**
   * Copy worksheet-level properties
   */
  copyWorksheetProperties(source, target) {
    // Copy column definitions (widths, styles)
    source.columns.forEach((col, index) => {
      if (!target.columns[index]) {
        target.columns[index] = {};
      }
      Object.assign(target.columns[index], {
        width: col.width,
        style: col.style ? { ...col.style } : undefined,
        hidden: col.hidden,
        outlineLevel: col.outlineLevel
      });
    });

    // Copy page setup
    if (source.pageSetup) {
      target.pageSetup = { ...source.pageSetup };
    }

    // Copy header/footer
    if (source.headerFooter) {
      target.headerFooter = { ...source.headerFooter };
    }

    // Copy views (zoom levels, etc.)
    if (source.views) {
      target.views = source.views.map(view => ({ ...view }));
    }

    // Copy print settings
    if (source.printSettings) {
      target.printSettings = { ...source.printSettings };
    }
  }

  /**
   * Copy every cell with complete formatting preservation
   */
  async copyAllCellsWithFormatting(source, target) {
    // First, copy all merged cell ranges
    Object.keys(source._merges || {}).forEach(range => {
      target.mergeCells(range);
    });

    // Copy each row with height
    for (let rowNum = 1; rowNum <= source.actualRowCount; rowNum++) {
      const sourceRow = source.getRow(rowNum);
      const targetRow = target.getRow(rowNum);

      // Copy row height and properties
      if (sourceRow.height) {
        targetRow.height = sourceRow.height;
      }
      if (sourceRow.hidden) {
        targetRow.hidden = sourceRow.hidden;
      }
      if (sourceRow.outlineLevel) {
        targetRow.outlineLevel = sourceRow.outlineLevel;
      }

      // Copy each cell in the row
      sourceRow.eachCell({ includeEmpty: true }, (sourceCell, colNumber) => {
        const targetCell = targetRow.getCell(colNumber);
        this.copyCellCompletely(sourceCell, targetCell);
      });
    }
  }

  /**
   * Copy a single cell with ALL its properties
   */
  copyCellCompletely(sourceCell, targetCell) {
    // Copy value (but we'll override this for data cells)
    targetCell.value = sourceCell.value;

    // Copy all formatting properties
    if (sourceCell.font) {
      targetCell.font = { ...sourceCell.font };
    }

    if (sourceCell.alignment) {
      targetCell.alignment = { ...sourceCell.alignment };
    }

    if (sourceCell.border) {
      targetCell.border = this.deepCopyBorder(sourceCell.border);
    }

    if (sourceCell.fill) {
      targetCell.fill = { ...sourceCell.fill };
      // Deep copy fgColor and bgColor if they exist
      if (sourceCell.fill.fgColor) {
        targetCell.fill.fgColor = { ...sourceCell.fill.fgColor };
      }
      if (sourceCell.fill.bgColor) {
        targetCell.fill.bgColor = { ...sourceCell.fill.bgColor };
      }
    }

    if (sourceCell.numFmt) {
      targetCell.numFmt = sourceCell.numFmt;
    }

    if (sourceCell.protection) {
      targetCell.protection = { ...sourceCell.protection };
    }

    if (sourceCell.dataValidation) {
      targetCell.dataValidation = { ...sourceCell.dataValidation };
    }

    // Copy formula if it exists
    if (sourceCell.formula) {
      targetCell.formula = sourceCell.formula;
    }

    // Copy note/comment if it exists
    if (sourceCell.note) {
      targetCell.note = { ...sourceCell.note };
    }
  }

  /**
   * Deep copy border object
   */
  deepCopyBorder(border) {
    const newBorder = {};
    ['top', 'bottom', 'left', 'right', 'diagonal'].forEach(side => {
      if (border[side]) {
        newBorder[side] = { ...border[side] };
        if (border[side].color) {
          newBorder[side].color = { ...border[side].color };
        }
      }
    });
    return newBorder;
  }

  /**
   * Fill the template with data while preserving ALL formatting
   */
  fillInvoiceData(invoiceData) {
    // Issue date (G1) - preserve the Japanese date format
    if (invoiceData.issueDate) {
      const cell = this.worksheet.getCell('G1');
      cell.value = new Date(invoiceData.issueDate);
      // The numFmt is already copied: yyyy" 年 "m" 月 "d" 日"
    }

    // Client information (preserve all formatting)
    if (invoiceData.clientName) {
      const cell = this.worksheet.getCell('B4');
      cell.value = invoiceData.clientName;
      // All formatting (font, borders, alignment) already preserved
    }

    if (invoiceData.clientPostal) {
      this.worksheet.getCell('B6').value = invoiceData.clientPostal;
    }

    if (invoiceData.clientAddress) {
      const addressLines = invoiceData.clientAddress.split('\\n');
      this.worksheet.getCell('B7').value = addressLines[0] || '';
      this.worksheet.getCell('B8').value = addressLines[1] || '';
    }

    // Company information (sender)
    if (invoiceData.companyPostal) {
      this.worksheet.getCell('G7').value = invoiceData.companyPostal;
    }

    if (invoiceData.companyAddress) {
      const addressLines = invoiceData.companyAddress.split('\\n');
      this.worksheet.getCell('G8').value = addressLines[0] || '';
      this.worksheet.getCell('G9').value = addressLines[1] || '';
    }

    if (invoiceData.companyEmail) {
      this.worksheet.getCell('H11').value = invoiceData.companyEmail;
    }

    if (invoiceData.companyName) {
      this.worksheet.getCell('G12').value = invoiceData.companyName;
    }

    // Due date (C18) - preserve Japanese date format
    if (invoiceData.dueDate) {
      this.worksheet.getCell('C18').value = new Date(invoiceData.dueDate);
    }

    // Clear existing item data (preserve formatting)
    for (let row = 21; row <= 28; row++) {
      // Clear content but keep ALL formatting
      this.clearCellValueOnly(`A${row}:C${row}`); // Description
      this.clearCellValueOnly(`D${row}`);         // Quantity
      this.clearCellValueOnly(`E${row}:F${row}`); // Unit price
      // Don't clear G column - it has formulas that should be preserved
      this.clearCellValueOnly(`H${row}:I${row}`); // Remarks
    }

    // Fill items (preserve formatting)
    if (invoiceData.items && Array.isArray(invoiceData.items)) {
      invoiceData.items.forEach((item, index) => {
        if (index < 8) { // Max 8 items (rows 21-28)
          const rowNum = 21 + index;

          // Description (A:C merged) - formatting already preserved
          this.worksheet.getCell(`A${rowNum}`).value = item.description;

          // Quantity (D) - text format (@) already preserved
          this.worksheet.getCell(`D${rowNum}`).value = item.quantity.toString();

          // Unit price (E:F merged) - ¥#,##0 format already preserved
          this.worksheet.getCell(`E${rowNum}`).value = item.unitPrice;

          // Amount (G) - formula already preserved, will auto-calculate
          // Remarks if provided
          if (item.remarks) {
            this.worksheet.getCell(`H${rowNum}`).value = item.remarks;
          }
        }
      });
    }

    // Bank information
    if (invoiceData.bankAccount) {
      this.worksheet.getCell('C32').value = invoiceData.bankAccount;
    }

    if (invoiceData.bankName) {
      this.worksheet.getCell('C33').value = `名義：${invoiceData.bankName}`;
    }

    // Notes (if provided)
    if (invoiceData.notes) {
      this.worksheet.getCell('A37').value = invoiceData.notes;
    }
  }

  /**
   * Clear only the value of a cell/range, preserving all formatting
   */
  clearCellValueOnly(cellAddress) {
    if (cellAddress.includes(':')) {
      // It's a range
      const range = cellAddress.split(':');
      const startCell = range[0];
      const endCell = range[1];

      // Get the range and clear values
      const startCol = startCell.charCodeAt(0) - 64;
      const startRow = parseInt(startCell.slice(1));
      const endCol = endCell.charCodeAt(0) - 64;
      const endRow = parseInt(endCell.slice(1));

      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          const cell = this.worksheet.getCell(row, col);
          // Only clear value, keep all formatting
          if (!cell.formula) { // Don't clear if it has a formula
            cell.value = null;
          }
        }
      }
    } else {
      // Single cell
      const cell = this.worksheet.getCell(cellAddress);
      if (!cell.formula) {
        cell.value = null;
      }
    }
  }

  /**
   * Save the formatted invoice
   */
  async save(outputPath) {
    await this.workbook.xlsx.writeFile(outputPath);
  }
}

// Example usage
async function createFormattedInvoice() {
  const formatter = new JapaneseInvoiceFormatter();

  // Load template with all formatting preserved
  await formatter.loadTemplate('/Users/fukku_maple/Downloads/invoice-template.xlsx');

  // Example invoice data
  const invoiceData = {
    issueDate: '2025-09-22',
    dueDate: '2025-10-31',
    clientName: '株式会社テストクライアント',
    clientPostal: '〒100-0001',
    clientAddress: '東京都千代田区千代田1-1-1\\nテストビル10階',
    companyName: '田中商事',
    companyPostal: '〒150-0001',
    companyAddress: '東京都渋谷区神宮前1-1-1\\n神宮前ビル5階',
    companyEmail: 'contact@tanaka-corp.jp',
    items: [
      {
        description: 'ウェブサイト制作費',
        quantity: 1,
        unitPrice: 500000,
        remarks: '初期制作費'
      },
      {
        description: 'メンテナンス費用',
        quantity: 3,
        unitPrice: 50000,
        remarks: '月額'
      }
    ],
    bankAccount: '〇〇銀行 渋谷支店（123） 普通 1234567',
    bankName: 'タナカタロウ',
    notes: '請求書に関するお問い合わせは上記メールアドレスまでお願いします。'
  };

  // Fill data while preserving formatting
  formatter.fillInvoiceData(invoiceData);

  // Save the result
  await formatter.save('/Users/fukku_maple/Downloads/formatted-invoice-output.xlsx');

  console.log('Invoice created with complete formatting preservation!');
  console.log('Output: /Users/fukku_maple/Downloads/formatted-invoice-output.xlsx');
}

// Uncomment to run
// createFormattedInvoice().catch(console.error);

export { JapaneseInvoiceFormatter };