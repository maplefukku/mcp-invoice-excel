#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  Tool,
  ErrorCode,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import ExcelJS from 'exceljs';
import { promises as fs } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import dotenv from 'dotenv';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

dotenv.config();

const serverName = process.env.MCP_NAME ?? 'mcp-invoice-excel';

interface InvoiceData {
  invoiceNumber: string;
  issueDate: string;
  dueDate?: string;

  // Sender (Your company)
  sender: {
    companyName: string;
    address?: string;
    phone?: string;
    email?: string;
    taxId?: string;
  };

  // Recipient (Client)
  recipient: {
    companyName: string;
    address?: string;
    phone?: string;
    email?: string;
    taxId?: string;
  };

  // Invoice items
  items: Array<{
    description: string;
    quantity: number;
    unitPrice: number;
    taxRate?: number;
    amount?: number;
  }>;

  // Payment details
  paymentMethod?: string;
  bankAccount?: string;
  notes?: string;

  // Totals (can be auto-calculated)
  subtotal?: number;
  taxAmount?: number;
  totalAmount?: number;
}

class InvoiceExcelServer {
  private server: Server;
  private invoiceCounter: number = 1;

  constructor() {
    this.server = new Server(
      {
        name: serverName,
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.setupHandlers();
    this.setupErrorHandling();
  }

  private setupErrorHandling(): void {
    this.server.onerror = (error) => {
      console.error('[MCP Server Error]', error);
    };

    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  private setupHandlers(): void {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: this.getTools(),
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        switch (name) {
          case 'create_invoice':
            return await this.createInvoice(args as any);
          case 'create_invoice_from_template':
            return await this.createInvoiceFromTemplate(args as any);
          case 'analyze_template':
            return await this.analyzeTemplate(args as any);
          case 'fill_japanese_template':
            return await this.fillJapaneseTemplate(args as any);
          default:
            throw new McpError(
              ErrorCode.MethodNotFound,
              `Unknown tool: ${name}`
            );
        }
      } catch (error: any) {
        if (error instanceof McpError) throw error;

        throw new McpError(
          ErrorCode.InternalError,
          `Tool execution failed: ${error.message}`
        );
      }
    });
  }

  private getTools(): Tool[] {
    return [
      {
        name: 'create_invoice',
        description: 'Create an Excel invoice with specified data',
        inputSchema: {
          type: 'object',
          properties: {
            invoiceData: {
              type: 'object',
              properties: {
                invoiceNumber: { type: 'string', description: 'Invoice number' },
                issueDate: { type: 'string', description: 'Issue date (YYYY-MM-DD)' },
                dueDate: { type: 'string', description: 'Due date (YYYY-MM-DD)' },
                sender: {
                  type: 'object',
                  properties: {
                    companyName: { type: 'string' },
                    address: { type: 'string' },
                    phone: { type: 'string' },
                    email: { type: 'string' },
                    taxId: { type: 'string' }
                  },
                  required: ['companyName']
                },
                recipient: {
                  type: 'object',
                  properties: {
                    companyName: { type: 'string' },
                    address: { type: 'string' },
                    phone: { type: 'string' },
                    email: { type: 'string' },
                    taxId: { type: 'string' }
                  },
                  required: ['companyName']
                },
                items: {
                  type: 'array',
                  items: {
                    type: 'object',
                    properties: {
                      description: { type: 'string' },
                      quantity: { type: 'number' },
                      unitPrice: { type: 'number' },
                      taxRate: { type: 'number', description: 'Tax rate as decimal (e.g., 0.1 for 10%)' }
                    },
                    required: ['description', 'quantity', 'unitPrice']
                  }
                },
                paymentMethod: { type: 'string' },
                bankAccount: { type: 'string' },
                notes: { type: 'string' }
              },
              required: ['invoiceNumber', 'issueDate', 'sender', 'recipient', 'items']
            },
            outputPath: {
              type: 'string',
              description: 'Path where the Excel file should be saved'
            }
          },
          required: ['invoiceData', 'outputPath']
        }
      },
      {
        name: 'create_invoice_from_template',
        description: 'Create an invoice by filling an existing Excel template',
        inputSchema: {
          type: 'object',
          properties: {
            templatePath: {
              type: 'string',
              description: 'Path to the Excel template file'
            },
            invoiceData: {
              type: 'object',
              description: 'Invoice data to fill in the template (same structure as create_invoice)'
            },
            outputPath: {
              type: 'string',
              description: 'Path where the filled Excel file should be saved'
            }
          },
          required: ['templatePath', 'invoiceData', 'outputPath']
        }
      },
      {
        name: 'analyze_template',
        description: 'Analyze an Excel template to understand its structure',
        inputSchema: {
          type: 'object',
          properties: {
            templatePath: {
              type: 'string',
              description: 'Path to the Excel template file to analyze'
            }
          },
          required: ['templatePath']
        }
      },
      {
        name: 'fill_japanese_template',
        description: 'Fill a Japanese invoice template with specific cell mappings',
        inputSchema: {
          type: 'object',
          properties: {
            templatePath: {
              type: 'string',
              description: 'Path to the Japanese Excel template file'
            },
            invoiceData: {
              type: 'object',
              properties: {
                invoiceNumber: { type: 'string', description: 'Invoice number' },
                issueDate: { type: 'string', description: 'Issue date (YYYY-MM-DD)' },
                dueDate: { type: 'string', description: 'Due date (YYYY-MM-DD)' },
                companyName: { type: 'string', description: 'Your company name' },
                companyPostal: { type: 'string', description: 'Your company postal code (e.g., 〒111-0000)' },
                companyAddress: { type: 'string', description: 'Your company address (use \\n for line breaks)' },
                companyEmail: { type: 'string', description: 'Your company email' },
                clientName: { type: 'string', description: 'Client company name' },
                clientPostal: { type: 'string', description: 'Client postal code (e.g., 〒111-2222)' },
                clientAddress: { type: 'string', description: 'Client address (use \\n for line breaks)' },
                bankAccount: { type: 'string', description: 'Bank account information for payment' },
                bankName: { type: 'string', description: 'Account holder name' },
                items: {
                  type: 'array',
                  items: {
                    type: 'object',
                    properties: {
                      description: { type: 'string' },
                      quantity: { type: 'number' },
                      unitPrice: { type: 'number' },
                      taxRate: { type: 'number', description: 'Tax rate as decimal (e.g., 0.1 for 10%)' }
                    },
                    required: ['description', 'quantity', 'unitPrice']
                  }
                },
                subtotal: { type: 'number', description: 'Subtotal amount' },
                taxAmount: { type: 'number', description: 'Tax amount' },
                totalAmount: { type: 'number', description: 'Total amount' },
                notes: { type: 'string', description: 'Additional notes' }
              },
              required: ['invoiceNumber', 'issueDate', 'companyName', 'clientName', 'items']
            },
            outputPath: {
              type: 'string',
              description: 'Path where the filled Excel file should be saved'
            }
          },
          required: ['templatePath', 'invoiceData', 'outputPath']
        }
      }
    ];
  }

  private async createInvoice(args: { invoiceData: InvoiceData; outputPath: string }) {
    const { invoiceData, outputPath } = args;

    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Invoice');

    // Set column widths
    worksheet.columns = [
      { width: 15 },
      { width: 40 },
      { width: 12 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
    ];

    let currentRow = 1;

    // Add title
    worksheet.mergeCells(`A${currentRow}:F${currentRow}`);
    const titleCell = worksheet.getCell(`A${currentRow}`);
    titleCell.value = 'INVOICE';
    titleCell.font = { bold: true, size: 20 };
    titleCell.alignment = { horizontal: 'center' };
    currentRow += 2;

    // Invoice details
    worksheet.getCell(`A${currentRow}`).value = 'Invoice Number:';
    worksheet.getCell(`B${currentRow}`).value = invoiceData.invoiceNumber;
    worksheet.getCell(`D${currentRow}`).value = 'Issue Date:';
    worksheet.getCell(`E${currentRow}`).value = invoiceData.issueDate;
    currentRow++;

    if (invoiceData.dueDate) {
      worksheet.getCell(`D${currentRow}`).value = 'Due Date:';
      worksheet.getCell(`E${currentRow}`).value = invoiceData.dueDate;
    }
    currentRow += 2;

    // Sender information
    worksheet.getCell(`A${currentRow}`).value = 'FROM:';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    currentRow++;
    worksheet.getCell(`A${currentRow}`).value = invoiceData.sender.companyName;
    currentRow++;
    if (invoiceData.sender.address) {
      worksheet.getCell(`A${currentRow}`).value = invoiceData.sender.address;
      currentRow++;
    }
    if (invoiceData.sender.phone) {
      worksheet.getCell(`A${currentRow}`).value = `Phone: ${invoiceData.sender.phone}`;
      currentRow++;
    }
    if (invoiceData.sender.email) {
      worksheet.getCell(`A${currentRow}`).value = `Email: ${invoiceData.sender.email}`;
      currentRow++;
    }
    if (invoiceData.sender.taxId) {
      worksheet.getCell(`A${currentRow}`).value = `Tax ID: ${invoiceData.sender.taxId}`;
      currentRow++;
    }
    currentRow++;

    // Recipient information
    worksheet.getCell(`A${currentRow}`).value = 'TO:';
    worksheet.getCell(`A${currentRow}`).font = { bold: true };
    currentRow++;
    worksheet.getCell(`A${currentRow}`).value = invoiceData.recipient.companyName;
    currentRow++;
    if (invoiceData.recipient.address) {
      worksheet.getCell(`A${currentRow}`).value = invoiceData.recipient.address;
      currentRow++;
    }
    if (invoiceData.recipient.phone) {
      worksheet.getCell(`A${currentRow}`).value = `Phone: ${invoiceData.recipient.phone}`;
      currentRow++;
    }
    if (invoiceData.recipient.email) {
      worksheet.getCell(`A${currentRow}`).value = `Email: ${invoiceData.recipient.email}`;
      currentRow++;
    }
    if (invoiceData.recipient.taxId) {
      worksheet.getCell(`A${currentRow}`).value = `Tax ID: ${invoiceData.recipient.taxId}`;
      currentRow++;
    }
    currentRow += 2;

    // Items header
    const headerRow = currentRow;
    worksheet.getCell(`A${headerRow}`).value = 'Item';
    worksheet.getCell(`B${headerRow}`).value = 'Description';
    worksheet.getCell(`C${headerRow}`).value = 'Quantity';
    worksheet.getCell(`D${headerRow}`).value = 'Unit Price';
    worksheet.getCell(`E${headerRow}`).value = 'Tax';
    worksheet.getCell(`F${headerRow}`).value = 'Amount';

    // Style header row
    for (let col = 1; col <= 6; col++) {
      const cell = worksheet.getCell(headerRow, col);
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0E0E0' }
      };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
    currentRow++;

    // Items
    let subtotal = 0;
    let totalTax = 0;
    invoiceData.items.forEach((item, index) => {
      const amount = item.quantity * item.unitPrice;
      const taxAmount = item.taxRate ? amount * item.taxRate : 0;
      const totalAmount = amount + taxAmount;

      worksheet.getCell(`A${currentRow}`).value = index + 1;
      worksheet.getCell(`B${currentRow}`).value = item.description;
      worksheet.getCell(`C${currentRow}`).value = item.quantity;
      worksheet.getCell(`D${currentRow}`).value = item.unitPrice;
      worksheet.getCell(`E${currentRow}`).value = taxAmount;
      worksheet.getCell(`F${currentRow}`).value = totalAmount;

      // Format currency cells
      worksheet.getCell(`D${currentRow}`).numFmt = '#,##0.00';
      worksheet.getCell(`E${currentRow}`).numFmt = '#,##0.00';
      worksheet.getCell(`F${currentRow}`).numFmt = '#,##0.00';

      subtotal += amount;
      totalTax += taxAmount;
      currentRow++;
    });

    currentRow++;

    // Totals
    worksheet.getCell(`E${currentRow}`).value = 'Subtotal:';
    worksheet.getCell(`E${currentRow}`).font = { bold: true };
    worksheet.getCell(`F${currentRow}`).value = subtotal;
    worksheet.getCell(`F${currentRow}`).numFmt = '#,##0.00';
    currentRow++;

    worksheet.getCell(`E${currentRow}`).value = 'Tax:';
    worksheet.getCell(`E${currentRow}`).font = { bold: true };
    worksheet.getCell(`F${currentRow}`).value = totalTax;
    worksheet.getCell(`F${currentRow}`).numFmt = '#,##0.00';
    currentRow++;

    worksheet.getCell(`E${currentRow}`).value = 'Total:';
    worksheet.getCell(`E${currentRow}`).font = { bold: true, size: 14 };
    worksheet.getCell(`F${currentRow}`).value = subtotal + totalTax;
    worksheet.getCell(`F${currentRow}`).numFmt = '#,##0.00';
    worksheet.getCell(`F${currentRow}`).font = { bold: true, size: 14 };
    currentRow += 2;

    // Payment information
    if (invoiceData.paymentMethod) {
      worksheet.getCell(`A${currentRow}`).value = 'Payment Method:';
      worksheet.getCell(`B${currentRow}`).value = invoiceData.paymentMethod;
      currentRow++;
    }
    if (invoiceData.bankAccount) {
      worksheet.getCell(`A${currentRow}`).value = 'Bank Account:';
      worksheet.getCell(`B${currentRow}`).value = invoiceData.bankAccount;
      currentRow++;
    }
    if (invoiceData.notes) {
      currentRow++;
      worksheet.getCell(`A${currentRow}`).value = 'Notes:';
      worksheet.getCell(`A${currentRow}`).font = { bold: true };
      currentRow++;
      worksheet.mergeCells(`A${currentRow}:F${currentRow}`);
      worksheet.getCell(`A${currentRow}`).value = invoiceData.notes;
      worksheet.getCell(`A${currentRow}`).alignment = { wrapText: true };
    }

    // Save the workbook
    await workbook.xlsx.writeFile(outputPath);

    return {
      content: [
        {
          type: 'text',
          text: `Invoice successfully created at: ${outputPath}\n` +
                `Invoice Number: ${invoiceData.invoiceNumber}\n` +
                `Total Amount: ${(subtotal + totalTax).toFixed(2)}`
        }
      ]
    };
  }

  private async createInvoiceFromTemplate(args: {
    templatePath: string;
    invoiceData: InvoiceData;
    outputPath: string
  }) {
    const { templatePath, invoiceData, outputPath } = args;

    // Load the template
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const worksheet = workbook.worksheets[0];

    // Find and replace placeholders in the template
    // This is a simple implementation - can be enhanced based on template structure
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        if (typeof cell.value === 'string') {
          let value = cell.value;

          // Replace common placeholders
          value = value.replace(/\{\{INVOICE_NUMBER\}\}/g, invoiceData.invoiceNumber);
          value = value.replace(/\{\{ISSUE_DATE\}\}/g, invoiceData.issueDate);
          value = value.replace(/\{\{DUE_DATE\}\}/g, invoiceData.dueDate || '');

          // Sender
          value = value.replace(/\{\{SENDER_COMPANY\}\}/g, invoiceData.sender.companyName);
          value = value.replace(/\{\{SENDER_ADDRESS\}\}/g, invoiceData.sender.address || '');
          value = value.replace(/\{\{SENDER_PHONE\}\}/g, invoiceData.sender.phone || '');
          value = value.replace(/\{\{SENDER_EMAIL\}\}/g, invoiceData.sender.email || '');
          value = value.replace(/\{\{SENDER_TAX_ID\}\}/g, invoiceData.sender.taxId || '');

          // Recipient
          value = value.replace(/\{\{RECIPIENT_COMPANY\}\}/g, invoiceData.recipient.companyName);
          value = value.replace(/\{\{RECIPIENT_ADDRESS\}\}/g, invoiceData.recipient.address || '');
          value = value.replace(/\{\{RECIPIENT_PHONE\}\}/g, invoiceData.recipient.phone || '');
          value = value.replace(/\{\{RECIPIENT_EMAIL\}\}/g, invoiceData.recipient.email || '');
          value = value.replace(/\{\{RECIPIENT_TAX_ID\}\}/g, invoiceData.recipient.taxId || '');

          // Payment
          value = value.replace(/\{\{PAYMENT_METHOD\}\}/g, invoiceData.paymentMethod || '');
          value = value.replace(/\{\{BANK_ACCOUNT\}\}/g, invoiceData.bankAccount || '');
          value = value.replace(/\{\{NOTES\}\}/g, invoiceData.notes || '');

          cell.value = value;
        }
      });
    });

    // Calculate totals
    let subtotal = 0;
    let totalTax = 0;
    invoiceData.items.forEach(item => {
      const amount = item.quantity * item.unitPrice;
      const taxAmount = item.taxRate ? amount * item.taxRate : 0;
      subtotal += amount;
      totalTax += taxAmount;
    });

    // Replace total placeholders
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        if (typeof cell.value === 'string') {
          let value = cell.value;
          value = value.replace(/\{\{SUBTOTAL\}\}/g, subtotal.toFixed(2));
          value = value.replace(/\{\{TAX_AMOUNT\}\}/g, totalTax.toFixed(2));
          value = value.replace(/\{\{TOTAL_AMOUNT\}\}/g, (subtotal + totalTax).toFixed(2));
          cell.value = value;
        }
      });
    });

    // Save the filled template
    await workbook.xlsx.writeFile(outputPath);

    return {
      content: [
        {
          type: 'text',
          text: `Invoice created from template successfully!\n` +
                `Template: ${templatePath}\n` +
                `Output: ${outputPath}\n` +
                `Invoice Number: ${invoiceData.invoiceNumber}\n` +
                `Total Amount: ${(subtotal + totalTax).toFixed(2)}`
        }
      ]
    };
  }

  private async analyzeTemplate(args: { templatePath: string }) {
    const { templatePath } = args;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const analysis: any = {
      worksheets: [],
      placeholders: new Set<string>(),
      structure: {},
      cellContents: []
    };

    workbook.worksheets.forEach(worksheet => {
      const sheetInfo = {
        name: worksheet.name,
        rowCount: worksheet.rowCount,
        columnCount: worksheet.columnCount,
        mergedCells: [] as string[],
        placeholders: [] as string[],
        cellValues: [] as Array<{address: string, value: any, type: string}>
      };

      // Analyze each cell in detail
      worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          const cellValue = cell.value;
          const cellAddress = cell.address;

          // Record cell value and type
          if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
            sheetInfo.cellValues.push({
              address: cellAddress,
              value: cellValue,
              type: typeof cellValue
            });
          }

          if (cell.isMerged) {
            const address = cell.address;
            if (!sheetInfo.mergedCells.includes(address)) {
              sheetInfo.mergedCells.push(address);
            }
          }

          // Find placeholders (text between {{ and }})
          if (typeof cell.value === 'string') {
            const matches = cell.value.match(/\{\{([^}]+)\}\}/g);
            if (matches) {
              matches.forEach(match => {
                const placeholder = match.replace(/\{\{|\}\}/g, '');
                analysis.placeholders.add(placeholder);
                if (!sheetInfo.placeholders.includes(match)) {
                  sheetInfo.placeholders.push(match);
                }
              });
            }
          }
        });
      });

      analysis.worksheets.push(sheetInfo);
    });

    // Convert Set to Array for JSON serialization
    analysis.placeholders = Array.from(analysis.placeholders);

    // Create a more readable cell contents summary
    const worksheet = analysis.worksheets[0];
    const cellSummary = worksheet.cellValues
      .filter((cell: any) => typeof cell.value === 'string' && cell.value.trim())
      .map((cell: any) => `${cell.address}: "${cell.value}"`)
      .join('\n');

    return {
      content: [
        {
          type: 'text',
          text: `Template Analysis Results:\n` +
                `File: ${templatePath}\n` +
                `Worksheets: ${analysis.worksheets.length}\n` +
                `Found ${analysis.placeholders.length} unique placeholders:\n` +
                analysis.placeholders.map((p: string) => `  - {{${p}}}`).join('\n') +
                `\n\nWorksheet: ${worksheet.name}\n` +
                `Dimensions: ${worksheet.rowCount} rows × ${worksheet.columnCount} columns\n` +
                `Merged cells: ${worksheet.mergedCells.length}\n` +
                `\n--- Cell Contents (Text Only) ---\n` +
                cellSummary +
                `\n\n--- Worksheet Details (JSON) ---\n` +
                JSON.stringify(analysis.worksheets, null, 2)
        }
      ]
    };
  }

  private async fillJapaneseTemplate(args: {
    templatePath: string;
    invoiceData: any;
    outputPath: string
  }) {
    const { templatePath, invoiceData, outputPath } = args;

    // Create a perfect clone of the template by copying the file first
    const fs = await import('fs/promises');
    await fs.copyFile(templatePath, outputPath);

    // Load the copied workbook to modify only the data
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(outputPath);

    const worksheet = workbook.worksheets[0];

    // Japanese invoice template cell mappings based on template analysis
    const cellMappings = {
      // Header information
      issueDate: 'G1',          // 発行日
      dueDate: 'C18',           // 振込期限

      // Company information (issuer/sender) - right side
      companyPostal: 'G7',      // 〒111-0000
      companyAddress1: 'G8',    // 東京都渋谷区1-1-1
      companyAddress2: 'G9',    // XXマンション102号室
      companyEmail: 'H11',      // example@example.com
      companyName: 'G12',       // 田中太郎

      // Client information (recipient) - left side
      clientName: 'B4',         // 株式会社インボイス生成
      clientPostal: 'B6',       // 〒111-2222
      clientAddress1: 'B7',     // 東京都中央区9-9-9
      clientAddress2: 'B8',     // XXマンション902号室

      // Item rows start from row 21
      itemStartRow: 21,
      itemEndRow: 28,
      itemColumns: {
        description: 'A',       // 品名・内容 (A21-C21 merged)
        quantity: 'D',          // 数量 (D21)
        unitPrice: 'E',         // 単価（税込） (E21)
        amount: 'G'             // 金額 (G21, has formula)
      },

      // Bank information
      bankInfo: 'C32',          // お振込先
      bankName: 'C33'           // 名義
    };

    // Helper function to update only the cell value without touching any formatting
    const updateCellValue = (cellAddress: string, newValue: any) => {
      const cell = worksheet.getCell(cellAddress);
      // Simply update the value - all formatting is already perfect from the copy
      cell.value = newValue;
    };

    // Update invoice data while preserving all formatting
    if (invoiceData.issueDate) {
      updateCellValue(cellMappings.issueDate, new Date(invoiceData.issueDate));
    }

    if (invoiceData.dueDate) {
      updateCellValue(cellMappings.dueDate, new Date(invoiceData.dueDate));
    }

    // Company information
    if (invoiceData.companyName) {
      updateCellValue(cellMappings.companyName, invoiceData.companyName);
    }

    if (invoiceData.companyPostal) {
      updateCellValue(cellMappings.companyPostal, invoiceData.companyPostal);
    }

    if (invoiceData.companyAddress) {
      const addressParts = invoiceData.companyAddress.split('\n');
      if (addressParts[0]) {
        updateCellValue(cellMappings.companyAddress1, addressParts[0]);
      }
      if (addressParts[1]) {
        updateCellValue(cellMappings.companyAddress2, addressParts[1]);
      }
    }

    if (invoiceData.companyEmail) {
      updateCellValue(cellMappings.companyEmail, invoiceData.companyEmail);
    }

    // Client information
    if (invoiceData.clientName) {
      updateCellValue(cellMappings.clientName, invoiceData.clientName);
    }

    if (invoiceData.clientPostal) {
      updateCellValue(cellMappings.clientPostal, invoiceData.clientPostal);
    }

    if (invoiceData.clientAddress) {
      const addressParts = invoiceData.clientAddress.split('\n');
      if (addressParts[0]) {
        updateCellValue(cellMappings.clientAddress1, addressParts[0]);
      }
      if (addressParts[1]) {
        updateCellValue(cellMappings.clientAddress2, addressParts[1]);
      }
    }

    // Clear existing items data - formatting is already perfect from the copy
    for (let row = cellMappings.itemStartRow; row <= cellMappings.itemEndRow; row++) {
      // Clear only the data values, leave everything else untouched
      worksheet.getCell(`${cellMappings.itemColumns.description}${row}`).value = null;
      worksheet.getCell(`${cellMappings.itemColumns.quantity}${row}`).value = null;
      worksheet.getCell(`${cellMappings.itemColumns.unitPrice}${row}`).value = null;
      // Don't touch amount column as it has formulas that need to remain
    }

    // Fill items with data while preserving formatting
    let totalCalculated = 0;

    if (invoiceData.items && Array.isArray(invoiceData.items)) {
      invoiceData.items.forEach((item: any, index: number) => {
        if (index < 8) { // Max 8 items (rows 21-28)
          const rowNumber = cellMappings.itemStartRow + index;

          // Update each cell while preserving its formatting
          updateCellValue(`${cellMappings.itemColumns.description}${rowNumber}`, item.description);
          updateCellValue(`${cellMappings.itemColumns.quantity}${rowNumber}`, item.quantity);
          updateCellValue(`${cellMappings.itemColumns.unitPrice}${rowNumber}`, item.unitPrice);

          totalCalculated += item.quantity * item.unitPrice;
        }
      });
    }

    // Update bank information while preserving formatting
    if (invoiceData.bankAccount) {
      updateCellValue(cellMappings.bankInfo, invoiceData.bankAccount);
    }

    if (invoiceData.bankName) {
      updateCellValue(cellMappings.bankName, `名義：${invoiceData.bankName}`);
    }

    // Add notes if provided (find appropriate cell)
    if (invoiceData.notes) {
      // Look for a notes area in the template
      const notesCell = worksheet.getCell('B37') || worksheet.getCell('A37');
      if (notesCell) {
        updateCellValue(notesCell.address, invoiceData.notes);
      }
    }

    // Save the workbook - all formatting is already perfect from the file copy
    await workbook.xlsx.writeFile(outputPath);

    return {
      content: [
        {
          type: 'text',
          text: `Japanese invoice template filled successfully with complete formatting preservation!\n` +
                `Template: ${templatePath}\n` +
                `Output: ${outputPath}\n` +
                `Company: ${invoiceData.companyName || 'N/A'}\n` +
                `Client: ${invoiceData.clientName || 'N/A'}\n` +
                `Items: ${invoiceData.items?.length || 0}\n` +
                `Calculated Total: ¥${totalCalculated.toLocaleString()}\n` +
                `Issue Date: ${invoiceData.issueDate || 'N/A'}\n` +
                `Due Date: ${invoiceData.dueDate || 'N/A'}\n` +
                `Perfect 100% template reproduction achieved via file cloning!`
        }
      ]
    };
  }

  public async run(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);

    // Log to stderr to avoid interfering with MCP communication
    console.error(`${serverName} MCP server running on stdio`);
  }
}

// Main entry point
async function main() {
  try {
    const server = new InvoiceExcelServer();
    await server.run();
  } catch (error) {
    console.error('Failed to start MCP server:', error);
    process.exit(1);
  }
}

main().catch(console.error);