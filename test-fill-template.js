#!/usr/bin/env node

import { spawn } from 'child_process';

// Create a simple MCP client to call the fill_japanese_template tool
async function callMCPTool(toolName, args) {
  return new Promise((resolve, reject) => {
    const mcpServer = spawn('node', ['/Users/fukku_maple/Documents/mcp-invoice-excel/build/index.js'], {
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

// Test data for the enhanced formatting preservation
const invoiceData = {
  "invoiceNumber": "FINAL-001",
  "issueDate": "2025-01-24",
  "dueDate": "2025-02-24",
  "companyName": "完璧再現株式会社",
  "companyPostal": "〒107-0052",
  "companyAddress": "東京都港区赤坂9-7-1\nミッドタウン・タワー25F",
  "companyEmail": "perfect@repro.co.jp",
  "clientName": "最終テスト株式会社",
  "clientPostal": "〒100-0006",
  "clientAddress": "東京都千代田区有楽町1-1-1\n有楽町マリオン10F",
  "bankAccount": "三菱UFJ銀行 赤坂支店（051） 普通 7788990",
  "bankName": "カンペキサイゲンカブシキガイシャ",
  "items": [
    {
      "description": "完璧フォーマット再現システム",
      "quantity": 1,
      "unitPrice": 500000
    },
    {
      "description": "品質保証・テスト",
      "quantity": 1,
      "unitPrice": 200000
    }
  ],
  "notes": "この請求書は完全なフォーマット再現テストです。すべての書式が元のテンプレートと完璧に一致するはずです。"
};

try {
  console.log('Testing enhanced formatting preservation fix...');
  const result = await callMCPTool('fill_japanese_template', {
    templatePath: "/Users/fukku_maple/Downloads/invoice-template.xlsx",
    outputPath: "/Users/fukku_maple/Downloads/final-perfect-test.xlsx",
    invoiceData: invoiceData
  });
  console.log('Result:', result);
} catch (error) {
  console.error('Error:', error.message);
}