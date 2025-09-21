# MCP Invoice Excel

The Universal MCP Server for generating Excel invoices from structured data input.

## Installation

### Prerequisites
- Node.js 18+
- Set `MCP_INVOICE_EXCEL_API_KEY` in your environment (if using external services)

### Get an API key
- This server operates locally and doesn't require external API keys
- Skip this step for basic invoice generation

### Install from npm (Recommended)
```bash
npm install -g mcp-invoice-excel
```

### Build locally (Development)
```bash
cd /path/to/mcp-invoice-excel
npm i
npm run build
```

## Setup: Claude Code (CLI)

Use this one-liner (replace with your real values):

```bash
claude mcp add "mcp-invoice-excel" -s user -- npx mcp-invoice-excel
```

To remove:

```bash
claude mcp remove "mcp-invoice-excel"
```

## Setup: Cursor

Create `.cursor/mcp.json` in your client (do not commit it here):

```json
{
  "mcpServers": {
    "mcp-invoice-excel": {
      "command": "npx",
      "args": ["mcp-invoice-excel"],
      "env": {},
      "autoStart": true
    }
  }
}
```

## Other Clients and Agents

<details>
<summary>VS Code</summary>

Install via URI or CLI:

```bash
code --add-mcp '{"name":"mcp-invoice-excel","command":"npx","args":["mcp-invoice-excel"],"env":{}}'
```

</details>

<details>
<summary>Claude Desktop</summary>

Follow the MCP install guide and reuse the standard config above.

</details>

<details>
<summary>LM Studio</summary>

- Command: npx
- Args: ["mcp-invoice-excel"]
- Env: {}

</details>

<details>
<summary>Goose</summary>

- Type: STDIO
- Command: npx
- Args: mcp-invoice-excel
- Enabled: true

</details>

<details>
<summary>opencode</summary>

Example `~/.config/opencode/opencode.json`:

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "mcp-invoice-excel": {
      "type": "local",
      "command": ["npx", "mcp-invoice-excel"],
      "enabled": true
    }
  }
}
```

</details>

<details>
<summary>Qodo Gen</summary>

Add a new MCP and paste the standard JSON config.

</details>

<details>
<summary>Windsurf</summary>

See docs and reuse the standard config above.

</details>

## Setup: Codex (TOML)

Example (Serena):

```toml
[mcp_servers.serena]
command = "uvx"
args = ["--from", "git+https://github.com/oraios/serena", "serena", "start-mcp-server", "--context", "codex"]
```

This server (minimal):

```toml
[mcp_servers.mcp-invoice-excel]
command = "npx"
args = ["mcp-invoice-excel"]
# Optional:
# MCP_NAME = "mcp-invoice-excel"
```

## Configuration (Env)
- MCP_NAME: Server name override (default: mcp-invoice-excel)
- No API keys required for local invoice generation

## Available Tools

### create_invoice
- **Description**: Create an Excel invoice with specified data
- **Inputs**:
  - invoiceData (object, required):
    - invoiceNumber (string, required)
    - issueDate (string, required) - Format: YYYY-MM-DD
    - dueDate (string, optional) - Format: YYYY-MM-DD
    - sender (object, required):
      - companyName (string, required)
      - address (string, optional)
      - phone (string, optional)
      - email (string, optional)
      - taxId (string, optional)
    - recipient (object, required):
      - companyName (string, required)
      - address (string, optional)
      - phone (string, optional)
      - email (string, optional)
      - taxId (string, optional)
    - items (array, required):
      - description (string, required)
      - quantity (number, required)
      - unitPrice (number, required)
      - taxRate (number, optional) - Tax rate as decimal (e.g., 0.1 for 10%)
    - paymentMethod (string, optional)
    - bankAccount (string, optional)
    - notes (string, optional)
  - outputPath (string, required) - Path where the Excel file should be saved
- **Output**: Success message with file path and invoice details

### create_invoice_from_template
- **Description**: Create an invoice by filling an existing Excel template
- **Inputs**:
  - templatePath (string, required) - Path to the Excel template file
  - invoiceData (object, required) - Same structure as create_invoice
  - outputPath (string, required) - Path where the filled Excel file should be saved
- **Output**: Success message with template and output paths

### analyze_template
- **Description**: Analyze an Excel template to understand its structure
- **Inputs**:
  - templatePath (string, required) - Path to the Excel template file to analyze
- **Output**: Analysis results including worksheets, placeholders, and structure

## Example invocation (MCP tool call)

```json
{
  "tool": "create_invoice",
  "arguments": {
    "invoiceData": {
      "invoiceNumber": "INV-2025-001",
      "issueDate": "2025-01-21",
      "dueDate": "2025-02-21",
      "sender": {
        "companyName": "Your Company Inc.",
        "address": "123 Business St, City, State 12345",
        "phone": "+1-234-567-8900",
        "email": "billing@yourcompany.com",
        "taxId": "12-3456789"
      },
      "recipient": {
        "companyName": "Client Corp.",
        "address": "456 Client Ave, City, State 54321",
        "email": "accounts@clientcorp.com"
      },
      "items": [
        {
          "description": "Consulting Services - January 2025",
          "quantity": 40,
          "unitPrice": 150,
          "taxRate": 0.1
        },
        {
          "description": "Software License",
          "quantity": 1,
          "unitPrice": 500,
          "taxRate": 0.1
        }
      ],
      "paymentMethod": "Bank Transfer",
      "bankAccount": "Account: 1234567890, Routing: 987654321",
      "notes": "Thank you for your business!"
    },
    "outputPath": "/path/to/invoice_2025_001.xlsx"
  }
}
```

## Troubleshooting
- Ensure Node 18+ is installed
- Local runs: `npx mcp-invoice-excel` after `npm run build`
- Inspect publish artifacts: `npm pack --dry-run`
- For template analysis, provide the path to your Excel template file

## References
- MCP SDK: https://modelcontextprotocol.io/docs/sdks
- Architecture: https://modelcontextprotocol.io/docs/learn/architecture
- Server Concepts: https://modelcontextprotocol.io/docs/learn/server-concepts
- Specification: https://modelcontextprotocol.io/specification/2025-06-18/server/index

## Name Consistency & Troubleshooting
- Always use CANONICAL_ID (mcp-invoice-excel) for identifiers and keys.
- Use CANONICAL_DISPLAY (MCP Invoice Excel) only for UI labels.
- Do not mix legacy keys after registration.

### Consistency Matrix:
- npm package name → mcp-invoice-excel
- Binary name → mcp-invoice-excel
- MCP server name (SDK metadata) → mcp-invoice-excel
- Env default MCP_NAME → mcp-invoice-excel
- Client registry key → mcp-invoice-excel
- UI label → MCP Invoice Excel

### Conflict Cleanup:
- Remove any stale keys (e.g., old display names) and re-add with mcp-invoice-excel only.
- Cursor: configure in the UI; this project intentionally omits .cursor/mcp.json.