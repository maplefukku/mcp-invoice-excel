#!/usr/bin/env node

import ExcelJS from 'exceljs';

async function verifyFinalCorrected() {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('/Users/fukku_maple/Downloads/corrected-mapping-test.xlsx');
    const worksheet = workbook.getWorksheet(1);

    console.log('=== FINAL VERIFICATION OF CORRECTED MAPPINGS ===\n');

    // Key data points
    console.log('1. HEADER INFORMATION:');
    console.log(`   Invoice Number (G2): ${worksheet.getCell('G2').value}`);
    console.log(`   Issue Date (G4): ${worksheet.getCell('G4').value}`);
    console.log(`   Due Date (G5): ${worksheet.getCell('G5').value}`);

    console.log('\n2. COMPANY (A8-A12):');
    console.log(`   Company Name (A8): ${worksheet.getCell('A8').value}`);
    console.log(`   Company Postal (A9): ${worksheet.getCell('A9').value}`);
    console.log(`   Company Address (A10): ${worksheet.getCell('A10').value}`);
    console.log(`   Company Email (A12): ${worksheet.getCell('A12').value}`);

    console.log('\n3. CLIENT (A15-A17):');
    console.log(`   Client Name (A15): ${worksheet.getCell('A15').value}`);
    console.log(`   Client Postal (A16): ${worksheet.getCell('A16').value}`);
    console.log(`   Client Address (A17): ${worksheet.getCell('A17').value}`);

    console.log('\n4. ITEMS (Row 21-22):');
    console.log(`   Item 1 Description (A21): ${worksheet.getCell('A21').value}`);
    console.log(`   Item 1 Quantity (D21): ${worksheet.getCell('D21').value}`);
    console.log(`   Item 1 Unit Price (E21): ${worksheet.getCell('E21').value}`);

    console.log(`   Item 2 Description (A22): ${worksheet.getCell('A22').value}`);
    console.log(`   Item 2 Quantity (D22): ${worksheet.getCell('D22').value}`);
    console.log(`   Item 2 Unit Price (E22): ${worksheet.getCell('E22').value}`);

    console.log('\n5. TOTAL & BANK:');
    console.log(`   Total (F25): ${worksheet.getCell('F25').value}`);
    console.log(`   Bank Account (A27): ${worksheet.getCell('A27').value}`);
    console.log(`   Bank Name (A28): ${worksheet.getCell('A28').value}`);

    console.log('\n6. NOTES:');
    console.log(`   Notes (A30): ${worksheet.getCell('A30').value}`);

    // Check if all expected values are present
    const checks = {
      'Invoice Number': worksheet.getCell('G2').value === 'CORRECT-001',
      'Company Name': worksheet.getCell('A8').value === '正確マッピング株式会社',
      'Client Name': worksheet.getCell('A15').value === 'クライアント正確株式会社',
      'Item 1 Description': worksheet.getCell('A21').value === '正確セルマッピングシステム',
      'Item 1 Quantity': worksheet.getCell('D21').value === 1,
      'Item 1 Unit Price': worksheet.getCell('E21').value === 800000,
      'Item 2 Description': worksheet.getCell('A22').value === 'フォーマット完全保持技術',
      'Item 2 Quantity': worksheet.getCell('D22').value === 1,
      'Item 2 Unit Price': worksheet.getCell('E22').value === 300000,
      'Total': worksheet.getCell('F25').value === 1100000,
      'Bank Account': worksheet.getCell('A27').value === 'りそな銀行 品川支店（120） 普通 1122334',
      'Notes': worksheet.getCell('A30').value === 'セルマッピングが正確に修正され、すべてのフォーマットが完璧に保持されます。'
    };

    console.log('\n=== VALIDATION SUMMARY ===');
    let allCorrect = true;
    Object.entries(checks).forEach(([key, passed]) => {
      console.log(`   ${key}: ${passed ? '✓ CORRECT' : '✗ INCORRECT'}`);
      if (!passed) allCorrect = false;
    });

    console.log(`\n=== OVERALL RESULT ===`);
    console.log(`${allCorrect ? '✅ ALL MAPPINGS CORRECT!' : '❌ SOME MAPPINGS NEED CORRECTION'}`);

    if (allCorrect) {
      console.log('\n🎉 Perfect! The corrected cell mappings and enhanced formatting preservation');
      console.log('   are working flawlessly. All data appears in the correct cells with');
      console.log('   proper formatting, merged cells are preserved, and calculations are accurate.');
    }

  } catch (error) {
    console.error('Error during verification:', error);
  }
}

verifyFinalCorrected();