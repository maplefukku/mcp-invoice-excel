#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

async function testFillJapaneseTemplate() {
    console.log('Starting live debug test...');

    const templatePath = '/Users/fukku_maple/Downloads/invoice-template.xlsx';
    const outputPath = '/Users/fukku_maple/Downloads/debug-live-test.xlsx';

    const testData = {
        invoiceNumber: "DEBUG-001",
        issueDate: "2025-01-23",
        companyName: "デバッグテスト株式会社",
        clientName: "テストクライアント",
        items: [
            {
                description: "デバッグテスト項目",
                quantity: 1,
                unitPrice: 100000
            }
        ]
    };

    try {
        console.log('Loading template from:', templatePath);

        // Check if template exists
        if (!fs.existsSync(templatePath)) {
            throw new Error(`Template file not found: ${templatePath}`);
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);

        console.log('Template loaded successfully');
        console.log('Number of worksheets:', workbook.worksheets.length);

        const worksheet = workbook.worksheets[0];
        console.log('Working with worksheet:', worksheet.name);

        // Apply the data to the template
        console.log('Applying test data...');

        // Find and fill cells based on common patterns
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                if (cell.value && typeof cell.value === 'string') {
                    let value = cell.value;

                    // Replace invoice number patterns
                    if (value.includes('{{invoiceNumber}}') || value.includes('請求書番号')) {
                        console.log(`Found invoice number cell at ${cell.address}: ${value}`);
                        if (value.includes('{{invoiceNumber}}')) {
                            cell.value = value.replace('{{invoiceNumber}}', testData.invoiceNumber);
                        }
                    }

                    // Replace date patterns
                    if (value.includes('{{issueDate}}') || value.includes('発行日')) {
                        console.log(`Found date cell at ${cell.address}: ${value}`);
                        if (value.includes('{{issueDate}}')) {
                            cell.value = value.replace('{{issueDate}}', testData.issueDate);
                        }
                    }

                    // Replace company name patterns
                    if (value.includes('{{companyName}}')) {
                        console.log(`Found company name cell at ${cell.address}: ${value}`);
                        cell.value = value.replace('{{companyName}}', testData.companyName);
                    }

                    // Replace client name patterns
                    if (value.includes('{{clientName}}')) {
                        console.log(`Found client name cell at ${cell.address}: ${value}`);
                        cell.value = value.replace('{{clientName}}', testData.clientName);
                    }
                }
            });
        });

        // Add item data if we find item placeholders
        let itemRowStart = null;
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell) => {
                if (cell.value && typeof cell.value === 'string' && cell.value.includes('{{item.description}}')) {
                    itemRowStart = rowNumber;
                }
            });
        });

        if (itemRowStart) {
            console.log(`Found item row template at row ${itemRowStart}`);
            const item = testData.items[0];
            const row = worksheet.getRow(itemRowStart);

            row.eachCell((cell) => {
                if (cell.value && typeof cell.value === 'string') {
                    let value = cell.value;
                    value = value.replace('{{item.description}}', item.description);
                    value = value.replace('{{item.quantity}}', item.quantity.toString());
                    value = value.replace('{{item.unitPrice}}', item.unitPrice.toLocaleString('ja-JP'));
                    value = value.replace('{{item.total}}', (item.quantity * item.unitPrice).toLocaleString('ja-JP'));
                    cell.value = value;
                }
            });
        }

        console.log('Saving to:', outputPath);
        await workbook.xlsx.writeFile(outputPath);

        // Verify the file was created
        const stats = fs.statSync(outputPath);
        console.log(`✅ File created successfully: ${outputPath}`);
        console.log(`📁 File size: ${stats.size} bytes`);
        console.log(`🕒 Created at: ${stats.birthtime}`);

        return {
            success: true,
            message: `Invoice successfully generated at ${outputPath}`,
            size: stats.size,
            createdAt: stats.birthtime
        };

    } catch (error) {
        console.error('❌ Error during test:', error.message);
        return {
            success: false,
            error: error.message
        };
    }
}

// Run the test
testFillJapaneseTemplate().then(result => {
    console.log('\n=== Test Results ===');
    console.log(JSON.stringify(result, null, 2));
});