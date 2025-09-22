#!/usr/bin/env node

import { spawn } from 'child_process';

// Create a simple MCP client to call the analyze_template tool
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

// Call the analyze_template tool
const templatePath = process.argv[2];
if (!templatePath) {
  console.error('Usage: node call-mcp-tool.js <template-path>');
  process.exit(1);
}

try {
  const result = await callMCPTool('analyze_template', { templatePath });
  console.log(result);
} catch (error) {
  console.error('Error:', error.message);
}