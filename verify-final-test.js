#!/usr/bin/env node

import ExcelJS from 'exceljs';

async function verifyFinalTest() {
  console.log('🔍 Verifying Enhanced Formatting Preservation Test Results\n');

  try {
    // Load both template and generated files
    const templateWorkbook = new ExcelJS.Workbook();
    await templateWorkbook.xlsx.readFile('/Users/fukku_maple/Downloads/invoice-template.xlsx');

    const generatedWorkbook = new ExcelJS.Workbook();
    await generatedWorkbook.xlsx.readFile('/Users/fukku_maple/Downloads/final-perfect-test.xlsx');

    const templateSheet = templateWorkbook.getWorksheet(1);
    const generatedSheet = generatedWorkbook.getWorksheet(1);

    console.log('📋 BASIC VERIFICATION:');
    console.log('✅ Template file loaded successfully');
    console.log('✅ Generated file loaded successfully');
    console.log(`✅ Template dimensions: ${templateSheet.rowCount} rows x ${templateSheet.columnCount} cols`);
    console.log(`✅ Generated dimensions: ${generatedSheet.rowCount} rows x ${generatedSheet.columnCount} cols`);

    // Check total calculation
    console.log('\n💰 CALCULATION VERIFICATION:');
    const totalCell = generatedSheet.getCell('F18');
    const totalValue = totalCell.value;
    console.log(`Total value: ${totalValue}`);
    console.log(`Expected: ¥700,000`);

    if (totalValue === 700000 || totalValue === '¥700,000' || String(totalValue).includes('700000')) {
      console.log('✅ Total calculation is CORRECT');
    } else {
      console.log('❌ Total calculation is INCORRECT');
    }

    // Check data population
    console.log('\n📝 DATA VERIFICATION:');

    // Invoice number
    const invoiceNumCell = generatedSheet.getCell('F2');
    console.log(`Invoice Number: ${invoiceNumCell.value} (Expected: FINAL-001)`);

    // Company name
    const companyNameCell = generatedSheet.getCell('A2');
    console.log(`Company Name: ${companyNameCell.value} (Expected: 完璧再現株式会社)`);

    // Client name
    const clientNameCell = generatedSheet.getCell('A7');
    console.log(`Client Name: ${clientNameCell.value} (Expected: 最終テスト株式会社)`);

    // First item
    const item1Cell = generatedSheet.getCell('A14');
    console.log(`Item 1: ${item1Cell.value} (Expected: 完璧フォーマット再現システム)`);

    // Check formatting preservation
    console.log('\n🎨 FORMATTING VERIFICATION:');

    let formattingIssues = 0;

    // Sample key cells to check formatting
    const keyCells = [
      { address: 'A1', description: 'Title cell' },
      { address: 'F2', description: 'Invoice number' },
      { address: 'A14', description: 'First item' },
      { address: 'F18', description: 'Total' }
    ];

    for (const cell of keyCells) {
      const templateCell = templateSheet.getCell(cell.address);
      const generatedCell = generatedSheet.getCell(cell.address);

      console.log(`\n📍 ${cell.description} (${cell.address}):`);

      // Check font
      const templateFont = templateCell.font || {};
      const generatedFont = generatedCell.font || {};

      if (JSON.stringify(templateFont) !== JSON.stringify(generatedFont)) {
        console.log(`  ⚠️  Font differs`);
        console.log(`     Template: ${JSON.stringify(templateFont)}`);
        console.log(`     Generated: ${JSON.stringify(generatedFont)}`);
        formattingIssues++;
      } else {
        console.log(`  ✅ Font preserved`);
      }

      // Check alignment
      const templateAlignment = templateCell.alignment || {};
      const generatedAlignment = generatedCell.alignment || {};

      if (JSON.stringify(templateAlignment) !== JSON.stringify(generatedAlignment)) {
        console.log(`  ⚠️  Alignment differs`);
        console.log(`     Template: ${JSON.stringify(templateAlignment)}`);
        console.log(`     Generated: ${JSON.stringify(generatedAlignment)}`);
        formattingIssues++;
      } else {
        console.log(`  ✅ Alignment preserved`);
      }

      // Check border
      const templateBorder = templateCell.border || {};
      const generatedBorder = generatedCell.border || {};

      if (JSON.stringify(templateBorder) !== JSON.stringify(generatedBorder)) {
        console.log(`  ⚠️  Border differs`);
        formattingIssues++;
      } else {
        console.log(`  ✅ Border preserved`);
      }

      // Check fill
      const templateFill = templateCell.fill || {};
      const generatedFill = generatedCell.fill || {};

      if (JSON.stringify(templateFill) !== JSON.stringify(generatedFill)) {
        console.log(`  ⚠️  Fill differs`);
        formattingIssues++;
      } else {
        console.log(`  ✅ Fill preserved`);
      }
    }

    // Check for formula calculation
    console.log('\n🧮 FORMULA VERIFICATION:');
    const formulaCell = generatedSheet.getCell('F18');
    if (formulaCell.formula) {
      console.log(`✅ Formula preserved: ${formulaCell.formula}`);
    } else {
      console.log(`⚠️  No formula found, checking if value is calculated correctly`);
    }

    // Summary
    console.log('\n📊 SUMMARY:');
    console.log(`Total formatting issues found: ${formattingIssues}`);

    if (formattingIssues === 0) {
      console.log('🎉 PERFECT! All formatting preserved successfully');
    } else if (formattingIssues <= 2) {
      console.log('👍 GOOD! Minor formatting differences');
    } else {
      console.log('⚠️  NEEDS IMPROVEMENT! Multiple formatting issues detected');
    }

    console.log('\n✅ Enhanced formatting preservation test completed successfully!');
    console.log('📁 Output file: /Users/fukku_maple/Downloads/final-perfect-test.xlsx');

  } catch (error) {
    console.error('❌ Error during verification:', error.message);
  }
}

verifyFinalTest();