# MCP Server for Interaction with Office Apps

This is a [Model Context Protocol](https://github.com/modelcontextprotocol) (MCP) server project.

**Warning**: Be VERY careful with open Word or PowerPoint apps. Your content may be unexpectedly modified or deleted.

## ✨ Enhanced Features

This MCP server provides enhanced Word document manipulation capabilities with:

### 🔧 **Enhanced `word.batch_format` Tool**
- **Atomic Operations**: All-or-nothing execution with automatic rollback on failure
- **Performance Limits**: Maximum 50 operations per call, 10,000 matches processed
- **Comprehensive Validation**: 
  - Font sizes: 1-1638 points
  - Paragraph spacing: 0-1584 points
  - Alignment: left, center, right, justify
  - List types: bullet, numbered, none
- **Enhanced Error Handling**: Specific validation errors with clear messages
- **Superscript/Subscript Exclusion**: Prevents conflicting text formatting

### 🎨 **Enhanced `word.highlight` Tool**  
- **Unified API**: Single tool for both first and all highlighting
- **Performance Limits**: Maximum 10,000 matches for optimal performance
- **Enhanced Color Validation**: Supports CSS colors, hex codes, and Word constants
- **Comprehensive Error Messages**: Clear guidance on supported formats

### 🛡️ **Robust Error Handling**
- **Validation Errors**: Parameter validation with specific ranges
- **Performance Errors**: Clear limits to prevent server crashes  
- **Execution Errors**: Document access and style existence validation
- **Atomic Rollback**: Document restoration on operation failure

### 📊 **Supported Operations**

#### Font Formatting
```json
{
  "font": {
    "name": "Arial",
    "size": 12,
    "color": "#FF0000",
    "bold": true,
    "italic": false,
    "underline": true,
    "strikethrough": false,
    "superscript": false,
    "subscript": false
  }
}
```

#### Paragraph Formatting  
```json
{
  "paragraph": {
    "alignment": "center",
    "line_spacing_rule": "single",
    "space_before": 6,
    "space_after": 6
  }
}
```

#### List Formatting
```json
{
  "list": {
    "type": "bullet"
  }
}
```

#### Style Application
```json
{
  "style": {
    "name": "Heading 1"
  }
}
```

### 🎯 **Scope Options**
- **`document`**: Apply to entire document
- **`selection`**: Apply to current selection (errors if no selection)
- **`matches`**: Apply to text matches using find parameters

### 🎨 **Color Formats**
- **CSS Colors**: `"red"`, `"blue"`, `"yellow"`
- **Hex Codes**: `"#FF0000"`, `"#0000FF"`
- **Word Constants**: `16711680` (red), `255` (blue)

## Setup and Installation

Simply run:

```bash
make
```

To create the virtual environment and install dependencies.

### Building the Standalone Executable (Windows Only)

To build the standalone executable for this project, you must:

1. Ensure you are on a **Windows system**.
2. Install the development dependencies (including PyInstaller):
   ```bash
   make
   ```
3. Run the build command to generate the executable:
   ```bash
   make package
   ```

This will create a `mcp-server-office.exe` file inside the `dist/` folder.

### Running the Standalone Executable

Once built, the executable can be run by simply double-clicking it, or from the command prompt:

```bash
./dist/mcp-server-office.exe
```

The server will start in SSE mode and run on port `25566`. To expose the server publicly, use the provided batch file (`run_with_devtunnel.bat`) to set up a Dev Tunnel and start the server.

---

### Running the Server

Use the VSCode launch configuration, or run manually:

#### Defaults to stdio transport:

```bash
uv run -m mcp_server.start
```

#### For SSE transport:

```bash
uv run -m mcp_server.start --transport sse --port 25566
```

To use this MCP server with a hosted Semantic Workbench assistant, go to [libraries:mcp-tunnel](../../libraries/python/mcp-tunnel) and run the following command and copy its output into your assistant configuration:

```bash
uv run mcp-tunnel --servers "office:25566"
```

#### If you need a public-facing server, use the `--use-ngrok-tunnel` option:

```bash
uv run -m mcp_server.start --use-ngrok-tunnel
```

or for .exe:

```bash
mcp-server-office.exe --use-ngrok-tunnel
```

The SSE URL is:

```bash
http://127.0.0.1:25566/sse
```

## Client Configuration

To use this MCP server in your setup, consider the following configuration:

### Stdio

```json
{
  "mcpServers": {
    "mcp-server-word": {
      "command": "uv",
      "args": ["run", "-m", "mcp_server.start"]
    }
  }
}
```

### SSE

```json
{
  "mcpServers": {
    "mcp-server-word": {
      "command": "http://127.0.0.1:25566/sse",
      "args": []
    }
  }
}
```
