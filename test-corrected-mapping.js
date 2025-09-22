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

// Test data for corrected cell mappings and enhanced formatting preservation
const invoiceData = {
  "invoiceNumber": "CORRECT-001",
  "issueDate": "2025-01-25",
  "dueDate": "2025-03-01",
  "companyName": "正確マッピング株式会社",
  "companyPostal": "〒108-0075",
  "companyAddress": "東京都港区港南2-16-3\n品川グランドセントラルタワー30F",
  "companyEmail": "mapping@correct.co.jp",
  "clientName": "クライアント正確株式会社",
  "clientPostal": "〒100-0013",
  "clientAddress": "東京都千代田区霞が関1-3-1\n経済産業省別館5F",
  "bankAccount": "りそな銀行 品川支店（120） 普通 1122334",
  "bankName": "セイカクマッピングカブシキガイシャ",
  "items": [
    {
      "description": "正確セルマッピングシステム",
      "quantity": 1,
      "unitPrice": 800000
    },
    {
      "description": "フォーマット完全保持技術",
      "quantity": 1,
      "unitPrice": 300000
    }
  ],
  "notes": "セルマッピングが正確に修正され、すべてのフォーマットが完璧に保持されます。"
};

try {
  console.log('Testing corrected cell mappings and enhanced formatting preservation...');
  const result = await callMCPTool('fill_japanese_template', {
    templatePath: "/Users/fukku_maple/Downloads/invoice-template.xlsx",
    outputPath: "/Users/fukku_maple/Downloads/corrected-mapping-test.xlsx",
    invoiceData: invoiceData
  });
  console.log('Result:', result);
} catch (error) {
  console.error('Error:', error.message);
}