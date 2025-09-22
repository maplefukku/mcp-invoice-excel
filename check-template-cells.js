#!/usr/bin/env node

import ExcelJS from 'exceljs';

async function checkTemplateCells() {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/fukku_maple/Downloads/invoice-template.xlsx');
    const worksheet = workbook.getWorksheet(1);

    console.log('=== ORIGINAL TEMPLATE CELLS ===');
    console.log(`E21 value: ${worksheet.getCell('E21').value}`);
    console.log(`E21 formula: ${worksheet.getCell('E21').formula}`);
    console.log(`F21 value: ${worksheet.getCell('F21').value}`);
    console.log(`F21 formula: ${worksheet.getCell('F21').formula}`);
    console.log(`E22 value: ${worksheet.getCell('E22').value}`);
    console.log(`E22 formula: ${worksheet.getCell('E22').formula}`);
    console.log(`F22 value: ${worksheet.getCell('F22').value}`);
    console.log(`F22 formula: ${worksheet.getCell('F22').formula}`);

    // Also check merged cells around this area
    console.log('\n=== MERGED CELLS ANALYSIS ===');
    const merges = worksheet.model.merges;
    Object.entries(merges).forEach(([range, merge]) => {
      if (range.includes('E2') || range.includes('F2') || range.includes('E21') || range.includes('F21')) {
        console.log(`Merged range: ${range}`);
      }
    });
  } catch (error) {
    console.error('Error:', error);
  }
}

checkTemplateCells();