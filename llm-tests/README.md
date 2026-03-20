# PptMcp LLM Integration Tests

LLM-powered integration tests for both PptMcp MCP Server and PowerPoint CLI using pytest-aitest.

## Prerequisites

- Windows desktop with Microsoft PowerPoint installed
- .NET 10 SDK
- Azure OpenAI endpoint configured
- PptMcp MCP Server and CLI built/installed

### Azure OpenAI

Set the endpoint for Entra ID auth:

```powershell
$env:AZURE_OPENAI_ENDPOINT = "https://<your-resource>.openai.azure.com/"
```

## Setup (uv + local pytest-aitest)

From this directory:

```powershell
uv sync
```

This uses a local editable dependency via:

```toml
[tool.uv.sources]
pytest-aitest = { path = "../../../pytest-aitest", editable = true }
```

## Build MCP Server (Required)

```powershell
dotnet build ..\..\src\PptMcp.McpServer\PptMcp.McpServer.csproj -c Release
```

## Run Tests (Manual Only)

### MCP Server tests

```powershell
uv run pytest -m mcp -v
```

### CLI tests

```powershell
uv run pytest -m cli -v
```

### All LLM tests

```powershell
uv run pytest -m aitest -v
```

### Canonical regression gate

Run the standard manual gate with the helper script from the repository root:

```powershell
.\scripts\Test-LlmRegressionGate.ps1
```

This runs the canonical six scenarios:

- `cli/test_cli_table.py::test_cli_table_create_query`
- `cli/test_cli_chart.py::test_cli_chart_workflows`
- `cli/test_cli_styling.py::test_cli_styling_header_fill`
- `mcp_tests/test_mcp_table.py::test_mcp_table_create_query`
- `mcp_tests/test_mcp_chart.py::test_mcp_chart_workflows`
- `mcp_tests/test_mcp_styling.py::test_mcp_styling_header_fill`

Use this gate after changing skill content, MCP tool descriptions, CLI help text, or other LLM-facing workflow guidance.

## Configuration Overrides

- `ppt_mcp_SERVER_COMMAND` — override MCP server command (full command line)
- `PPT_CLI_COMMAND` — override CLI command (default: `pptcli`)

Example:

```powershell
$env:ppt_mcp_SERVER_COMMAND = "d:\\source\\mcp-server-ppt\\src\\PptMcp.McpServer\\bin\\Release\\net9.0-windows\\PptMcp.McpServer.exe"
$env:PPT_CLI_COMMAND = "d:\\source\\mcp-server-ppt\\src\\PptMcp.CLI\\bin\\Release\\net9.0-windows\\pptcli.exe"
```

## Test Structure

- `test_mcp_*.py` — MCP Server workflows
- `test_cli_*.py` — CLI workflows
- `test_*calculation_mode*.py` — new calculation mode scenarios
- `Fixtures/` — shared test inputs (CSV/JSON/M files)
- `TestResults/` — HTML reports and artifacts
