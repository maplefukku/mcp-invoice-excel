#!/usr/bin/env node

import ExcelJS from 'exceljs';
import { promises as fs } from 'fs';

async function analyzeTemplateDetailed(templatePath) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    console.log(`=== COMPREHENSIVE ANALYSIS OF ${templatePath} ===\n`);

    workbook.worksheets.forEach((worksheet, wsIndex) => {
      console.log(`WORKSHEET ${wsIndex + 1}: "${worksheet.name}"`);
      console.log(`Dimensions: ${worksheet.rowCount} rows × ${worksheet.columnCount} columns`);
      console.log(`Actual Row Count: ${worksheet.actualRowCount}`);
      console.log(`Actual Column Count: ${worksheet.actualColumnCount}`);
      console.log('');

      // Column widths
      console.log('=== COLUMN WIDTHS ===');
      worksheet.columns.forEach((col, index) => {
        if (col.width) {
          console.log(`Column ${String.fromCharCode(65 + index)} (${index + 1}): width=${col.width}`);
        }
      });
      console.log('');

      // Row heights
      console.log('=== ROW HEIGHTS ===');
      for (let i = 1; i <= worksheet.actualRowCount; i++) {
        const row = worksheet.getRow(i);
        if (row.height) {
          console.log(`Row ${i}: height=${row.height}`);
        }
      }
      console.log('');

      // Merged cells
      console.log('=== MERGED CELLS ===');
      Object.keys(worksheet._merges).forEach(range => {
        console.log(`Merged range: ${range}`);
      });
      console.log('');

      // Detailed cell analysis
      console.log('=== DETAILED CELL ANALYSIS ===');
      worksheet.eachRow((row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          const colLetter = String.fromCharCode(64 + colNumber);
          const address = `${colLetter}${rowNumber}`;

          console.log(`\n--- Cell ${address} ---`);
          console.log(`Value: ${JSON.stringify(cell.value)}`);
          console.log(`Type: ${typeof cell.value}`);

          // Font details
          if (cell.font) {
            console.log(`Font: ${JSON.stringify(cell.font)}`);
          }

          // Alignment
          if (cell.alignment) {
            console.log(`Alignment: ${JSON.stringify(cell.alignment)}`);
          }

          // Border
          if (cell.border) {
            console.log(`Border: ${JSON.stringify(cell.border)}`);
          }

          // Fill (background)
          if (cell.fill) {
            console.log(`Fill: ${JSON.stringify(cell.fill)}`);
          }

          // Number format
          if (cell.numFmt) {
            console.log(`Number Format: ${cell.numFmt}`);
          }

          // Style
          if (cell.style) {
            console.log(`Style: ${JSON.stringify(cell.style)}`);
          }

          // Data validation
          if (cell.dataValidation) {
            console.log(`Data Validation: ${JSON.stringify(cell.dataValidation)}`);
          }

          // Is merged
          if (cell.isMerged) {
            console.log(`Is Merged: true`);
            console.log(`Master: ${cell.master ? cell.master.address : 'N/A'}`);
          }

          // Note/comment
          if (cell.note) {
            console.log(`Note: ${JSON.stringify(cell.note)}`);
          }

          // Formula
          if (cell.formula) {
            console.log(`Formula: ${cell.formula}`);
          }

          // Cell type
          console.log(`Cell Type: ${cell.type}`);
        });
      });

      // Print page setup
      console.log('\n=== PAGE SETUP ===');
      if (worksheet.pageSetup) {
        console.log(`Page Setup: ${JSON.stringify(worksheet.pageSetup, null, 2)}`);
      }

      // Print settings
      console.log('\n=== PRINT SETTINGS ===');
      if (worksheet.headerFooter) {
        console.log(`Header/Footer: ${JSON.stringify(worksheet.headerFooter, null, 2)}`);
      }

      // Views (zoom, etc.)
      console.log('\n=== VIEWS ===');
      if (worksheet.views) {
        console.log(`Views: ${JSON.stringify(worksheet.views, null, 2)}`);
      }

      console.log('\n=== WORKBOOK PROPERTIES ===');
      if (workbook.creator) console.log(`Creator: ${workbook.creator}`);
      if (workbook.lastModifiedBy) console.log(`Last Modified By: ${workbook.lastModifiedBy}`);
      if (workbook.created) console.log(`Created: ${workbook.created}`);
      if (workbook.modified) console.log(`Modified: ${workbook.modified}`);
      if (workbook.subject) console.log(`Subject: ${workbook.subject}`);
      if (workbook.title) console.log(`Title: ${workbook.title}`);
      if (workbook.description) console.log(`Description: ${workbook.description}`);
      if (workbook.keywords) console.log(`Keywords: ${workbook.keywords}`);
      if (workbook.category) console.log(`Category: ${workbook.category}`);
      if (workbook.manager) console.log(`Manager: ${workbook.manager}`);
      if (workbook.company) console.log(`Company: ${workbook.company}`);

      console.log('\n' + '='.repeat(80));
    });

  } catch (error) {
    console.error('Error analyzing template:', error);
  }
}

// Get template path from command line arguments
const templatePath = process.argv[2];
if (!templatePath) {
  console.error('Usage: node analyze-template.js <path-to-template.xlsx>');
  process.exit(1);
}

analyzeTemplateDetailed(templatePath);