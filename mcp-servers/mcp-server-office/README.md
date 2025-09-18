# MCP Server for Interaction with Office Apps

This is a [Model Context Protocol](https://github.com/modelcontextprotocol) (MCP) server project.

**Warning**: Be VERY careful with open Word or PowerPoint apps. Your content may be unexpectedly modified or deleted.

## ✨ Enhanced Features

This MCP server provides comprehensive Office application manipulation with:

### 📊 **PowerPoint Tools - NEW!**
Complete PowerPoint automation with precision coordinate placement for OCR reconstruction workflows:

#### **🎯 Presentation Management**
- **`ppt.presentation.create(a4_portrait=True, close_existing=False)`**: Create new presentation (auto slide 1)
- **`ppt.presentation.content()`**: Extract full structured content
- **`ppt.presentation.list()`**: List open presentations
- **`ppt.presentation.activate(presentation_index)`**: Focus a specific presentation
- **`ppt.presentation.save(presentation_index?)`** / **`ppt.presentation.save_as(...)`**: Persist changes
- **`ppt.presentation.close(presentation_index?, save=False)`**: Close one or all
- **`ppt.presentation.reveal(presentation_index?)`**: Bring window front

#### **📝 Precise Text Box Placement**  
- **`ppt.shape.textbox.add(slide_index, left, top, width, height, text, font?, paragraph?, dpi=96)`**: Add text at exact coordinates
- **`ppt.shape.textbox.update(slide_index, shape_id, text?, font?, paragraph?)`**: Update existing text boxes
- **Full Word Formatting Parity**: name, size, color, bold, italic, underline, strikethrough, superscript, subscript
- **Advanced Paragraph Controls**: alignment, line_spacing, space_before, space_after, bullets, indentation

#### **🖼️ Enhanced Image Handling**
- **`ppt.shape.image.add(slide_index, path, left, top, width?, height?, preserve_aspect=True, dpi=96)`**: Precise image placement
- **Bounds validation**: Direct width+height validated pre-insert; aspect-derived validated post intrinsic read
- **DPI Support**: Convert pixel coordinates from OCR (300 DPI typical)
- **Format Support**: PNG, JPG/JPEG, GIF, BMP, WMF, EMF, TIFF

#### **🎛️ Shape Management**
- **`ppt.shape.list(slide_index)`**: Get all shapes with precise geometry
- **`ppt.shape.delete(slide_index, shape_id)`**: Remove shapes
- **`ppt.shape.zorder(slide_index, shape_id, action)`**: Layer control (bring_to_front, send_to_back, step_forward, step_backward)

#### **🎯 Coordinate System & Units**
- **Origin**: Top-left of slide (0,0)

```
PowerPoint Coordinate System (A4 Portrait):
┌─────────────────────────────────────────────────────────────┐ (0,0)
│ Origin: Top-left corner                                     │
│                                                             │
│  (72,144)                                                   │
│     ┌─────────────┐                                         │
│     │ Text Box    │ ← 200pt width                           │
│     │ 50pt height │                                         │
│     └─────────────┘                                         │
│                                                             │
│                                                             │
│                                                             │
│  Units: Points (pt), Millimeters (mm), Pixels (px)         │
│  A4 Dimensions: 210mm × 297mm = 595.276 × 841.890 points   │
│                                                             │
│                                            (595.276,841.890)│
└─────────────────────────────────────────────────────────────┘
                      Bottom-right corner
```

- **Units Supported**: Points (default), millimeters (`"30mm"`), pixels (`"120px"` with DPI) (inches intentionally not supported)
- **A4 Portrait**: 210mm × 297mm (595.276 × 841.890 points)
- **Precision**: 3 decimal places, drift-tested for 80+ shape workflows

#### **🛡️ Validation & Error Handling**
```json
{
  "ok": true,
  "shape_id": 123,
  "left_pt": 72.000,
  "top_pt": 144.000,
  "width_pt": 200.000,
  "height_pt": 50.000
}
```
**Error Codes** (authoritative subset): `not-found`, `ambiguous`, `unsaved`, `overwrite-denied`, `operation-failed`, `io-error`, `invalid-unit`, `invalid-dimensions`, `invalid-action`, `creation-failed`, `validation-failed`, `out-of-range`, `unsupported-format`, `close-failed`

#### **📋 OCR Reconstruction Example**
```python
# 1. Create A4 presentation (slide 1 exists automatically)
await ppt.presentation.create(a4_portrait=True)

# 2. Add text blocks from OCR coordinates
await ppt.shape.textbox.add(
  slide_index=1,
  left="25mm", top="30mm",
  width="160mm", height="12mm",
  text="Document Title",
  font={"name": "Arial", "size": 16, "bold": True, "color": "#000000"}
)

# 3. Add image
await ppt.shape.image.add(
  slide_index=1,
  path="/path/to/logo.png",
  left="170mm", top="10mm",
  width="30mm", height="15mm",
  dpi=300
)

# 4. Save
await ppt.presentation.save_as("~/Documents/output.pptx")
```

### 🔧 **Enhanced Word Tools**
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

## File Operations (Word)

Two new tools provide essential file lifecycle actions for Word:

- `create_word_document(path, content?, content_format='markdown', overwrite=false, make_active=true, visible=false)`
- `save_word_document_as(target_path, source_path?, format='docx', overwrite=false, include_comments=true, close_after=false)`

Notes:
- Supports formats: `docx`, `pdf`, and `md` (markdown export is written by the server).
- Set `MCP_OFFICE_SANDBOX_ROOT` to restrict file I/O to a directory tree.
- Per-document locking and atomic save patterns are used for safety.

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
