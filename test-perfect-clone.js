#!/usr/bin/env node

import { spawn } from 'child_process';

// Create a simple MCP client to call the fill_japanese_template tool
async function callMCPTool(toolName, args) {
  return new Promise((resolve, reject) => {
    const mcpServer = spawn('node', ['/Users/fukku_maple/documents/mcp-invoice-excel/build/index.js'], {
      stdio: ['pipe', 'pipe', 'pipe']
    });

    let output = '';
    let errorOutput = '';

    mcpServer.stdout.on('data', (data) => {
      output += data.toString();
    });

    mcpServer.stderr.on('data', (data) => {
      errorOutput += data.toString();
    });

    mcpServer.on('close', (code) => {
      console.log('MCP server stderr output:', errorOutput);
      console.log('MCP server stdout output:', output);
      if (code === 0) {
        resolve(output);
      } else {
        reject(new Error(`MCP server exited with code ${code}: ${errorOutput}`));
      }
    });

    // Send MCP initialization and tool call
    const initMessage = {
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2024-11-05',
        capabilities: {},
        clientInfo: {
          name: 'test-client',
          version: '1.0.0'
        }
      }
    };

    const toolCallMessage = {
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/call',
      params: {
        name: toolName,
        arguments: args
      }
    };

    mcpServer.stdin.write(JSON.stringify(initMessage) + '\n');
    mcpServer.stdin.write(JSON.stringify(toolCallMessage) + '\n');
    mcpServer.stdin.end();
  });
}

// Test data for perfect clone system
const testData = {
  templatePath: "/Users/fukku_maple/Downloads/invoice-template.xlsx",
  outputPath: "/Users/fukku_maple/Downloads/perfect-clone-test.xlsx",
  invoiceData: {
    invoiceNumber: "PERFECT-001",
    issueDate: "2025-01-23",
    dueDate: "2025-03-15",
    companyName: "完璧フォーマット株式会社",
    companyPostal: "〒106-0032",
    companyAddress: "東京都港区六本木6-10-1\n六本木ヒルズ森タワー20F",
    companyEmail: "perfect@format.co.jp",
    clientName: "クライアント完璧株式会社",
    clientPostal: "〒105-0011",
    clientAddress: "東京都港区芝公園4-2-8\n東京タワー1F",
    bankAccount: "三井住友銀行 六本木支店（140） 普通 5566778",
    bankName: "カンペキフォーマットカブシキガイシャ",
    items: [
      {
        description: "完璧フォーマット再現システム開発",
        quantity: 1,
        unitPrice: 2000000
      },
      {
        description: "テンプレート解析・最適化",
        quantity: 1,
        unitPrice: 500000
      }
    ],
    notes: "この請求書は100%完璧なテンプレート再現テストです。全ての書式が元のテンプレートと完全に一致します。"
  }
};

try {
  console.log('Testing perfect template cloning system...');
  console.log('Input template:', testData.templatePath);
  console.log('Output file:', testData.outputPath);
  console.log('');

  const result = await callMCPTool('fill_japanese_template', testData);
  console.log('=== TEST RESULTS ===');
  console.log(result);
} catch (error) {
  console.error('Error:', error.message);
  process.exit(1);
}