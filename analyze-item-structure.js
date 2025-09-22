#!/usr/bin/env node

import ExcelJS from 'exceljs';

async function analyzeItemStructure() {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/fukku_maple/Downloads/invoice-template.xlsx');
    const worksheet = workbook.getWorksheet(1);

    console.log('=== TEMPLATE ITEM STRUCTURE ANALYSIS ===');

    // Check row 20 (header row)
    console.log('\nRow 20 (Headers):');
    for (let col = 1; col <= 8; col++) {
      const cell = worksheet.getCell(20, col);
      const letter = String.fromCharCode(64 + col); // A=65, so 64+1=A
      console.log(`  ${letter}20: "${cell.value}"`);
    }

    // Check row 21 (first item row)
    console.log('\nRow 21 (First Item):');
    for (let col = 1; col <= 8; col++) {
      const cell = worksheet.getCell(21, col);
      const letter = String.fromCharCode(64 + col);
      console.log(`  ${letter}21: value="${cell.value}", formula="${cell.formula}", merged="${cell.isMerged}"`);
    }

    // Check specific merged ranges for row 21
    console.log('\nMerged cells in row 21:');
    const merges = worksheet.model.merges;
    Object.entries(merges).forEach(([range, merge]) => {
      if (range.includes('21')) {
        console.log(`  ${range}`);
      }
    });

    // Based on typical Japanese invoice structure, let me guess the correct mapping
    console.log('\n=== TYPICAL JAPANESE INVOICE STRUCTURE ===');
    console.log('Usually:');
    console.log('  A21-C21: Item description (merged)');
    console.log('  D21: Quantity');
    console.log('  E21-F21: Unit Price (merged) OR E21: Unit Price, F21: Quantity');
    console.log('  G21: Amount (calculated)');
    console.log('  H21: Tax or other');

  } catch (error) {
    console.error('Error:', error);
  }
}

analyzeItemStructure();