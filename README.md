# PowerPoint Automation MCP Server

A Model Context Protocol (MCP) server that provides comprehensive PowerPoint automation capabilities using python-pptx. This server enables AI assistants to create, modify, and manage PowerPoint presentations programmatically.

## Demo

https://github.com/user-attachments/assets/3daf3bef-4d75-4639-a891-0e64b80b4807

## Features

- **Presentation Management**: Create, open, save, and close presentations
- **Slide Operations**: Add, delete, and manage slides with various layouts
- **Content Creation**: Add text boxes, images, tables, and charts
- **Text Formatting**: Update text with font styling (name, size, bold, italic)
- **Shape Manipulation**: Move, resize, and modify shape properties
- **Table Operations**: Create tables and update cell content
- **Chart Creation**: Add various chart types (column, line, pie, bar) with data
- **Group Management**: Ungroup shapes for easier manipulation
- **Layout Support**: Work with different slide layouts and placeholders

## Installation

```bash
git clone https://github.com/buckylee2019/powerpoint-automation-mcp.git
cd powerpoint-automation-mcp
uv sync
```

## Running the Server

```bash
uv run --directory /path/to/powerpoint-automation-mcp mcp-ppt-server
```

## API Reference

The server provides 25+ tools for comprehensive PowerPoint automation. All operations work with a single active presentation to simplify the API.

### Presentation Management

#### `initialize_powerpoint() -> bool`
Initialize the PowerPoint automation system.

#### `create_presentation(template: str = None) -> Dict`
Create a new presentation, optionally from a template.
- `template`: Optional path to template file

#### `open_presentation(file_path: str) -> Dict`
Open an existing PowerPoint file.
- `file_path`: Full path to .pptx file

#### `get_presentation() -> Dict`
Get information about the currently active presentation.

#### `save_presentation(path: str = None) -> Dict`
Save the active presentation.
- `path`: Optional save path (uses current path if omitted)

#### `close_presentation() -> Dict`
Close the active presentation.

### Slide Management

#### `get_slides() -> List[Dict]`
Get list of all slides with metadata.

#### `add_slide(layout_index: int = 1) -> Dict`
Add a new slide with specified layout.
- `layout_index`: Layout type (0=Title, 1=Title+Content, etc.)

#### `delete_slide(slide_index: int) -> Dict`
Delete a slide from the presentation.
- `slide_index`: 0-based slide index

#### `get_slide_layouts() -> Dict`
Get all available slide layouts with their properties.

#### `set_slide_title(slide_index: int, title: str) -> Dict`
Set the title of a slide.

### Content Inspection

#### `get_slide_shapes(slide_index: int) -> Dict`
**ALWAYS RUN FIRST!** Get all shapes in a slide with properties and IDs.
- Returns shape types, positions, sizes, and text content
- Identifies grouped shapes that may need ungrouping

#### `get_slide_text(slide_index: int) -> Dict`
Get all text content in a slide organized by shape.
- Detects grouped shapes that may need ungrouping first

### Text and Shape Manipulation

#### `update_text(slide_index: int, shape_index: int, text: str, font_name: str = None, font_size: int = None, bold: bool = None, italic: bool = None, preserve_existing: bool = True) -> Dict`
Update text content with optional formatting.
- Supports font styling while preserving existing formatting
- `preserve_existing`: Keep original formatting for unspecified properties

#### `update_shape_by_id(slide_index: int, shape_id: str, text: str = None, left: float = None, top: float = None, width: float = None, height: float = None) -> Dict`
Update shape properties by ID.
- Positions and sizes in inches

### Content Creation

#### `add_textbox(slide_index: int, text: str, left: float = 1, top: float = 1, width: float = 4, height: float = 2) -> Dict`
Add a text box to a slide.

#### `add_image(slide_index: int, image_path: str, left: float = 1, top: float = 1, width: float = None, height: float = None) -> Dict`
Add an image to a slide.
- Maintains aspect ratio if only width or height specified

#### `add_table(slide_index: int, rows: int, cols: int, left: float = 1, top: float = 1, width: float = 8, height: float = 4) -> Dict`
Add a table to a slide.

#### `add_chart(slide_index: int, chart_type: str, categories: List[str], series_names: List[str], series_values: List[List[float]], left: float = 1, top: float = 1, width: float = 8, height: float = 4, has_legend: bool = True) -> Dict`
Add a chart with multiple data series.
- `chart_type`: 'COLUMN', 'LINE', 'PIE', 'BAR'
- `categories`: X-axis labels
- `series_names`: Legend labels
- `series_values`: Data values for each series

### Table Operations

#### `update_table_cell(slide_index: int, shape_index: int, row: int, col: int, text: str) -> Dict`
Update text in a specific table cell.

#### `get_table_content(slide_index: int, shape_index: int) -> Dict`
Retrieve all content from a table.

### Advanced Operations

#### `ungroup_shapes(slide_index: int) -> Dict`
Ungroup all text-containing groups in a slide.
- Only processes simple groups (no nested groups)
- Preserves shape positions and properties
- Skips groups without text content

## Usage Workflow

1. **Initialize**: Call `initialize_powerpoint()`
2. **Open/Create**: Use `open_presentation()` or `create_presentation()`
3. **Inspect**: Always run `get_slide_shapes()` first to understand slide structure
4. **Ungroup if needed**: Use `ungroup_shapes()` if grouped shapes are detected
5. **Modify**: Use appropriate tools to add/modify content
6. **Save**: Call `save_presentation()` to persist changes

## Error Handling

All functions return dictionaries with either success data or error information:

```json
// Success response
{
  "success": true,
  "message": "Operation completed",
  "data": {...}
}

// Error response
{
  "error": "Description of what went wrong"
}
```

## MCP Configuration

Add to your MCP settings:

```json
{
  "mcpServers": {
    "powerpoint-automation": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "/path/to/powerpoint-automation-mcp",
        "mcp-ppt-server"
      ]
    }
  }
}
```

## Dependencies

- `python-pptx`: PowerPoint file manipulation
- `mcp`: Model Context Protocol framework
- `fastmcp`: Simplified MCP server creation

## License

MIT License - see LICENSE file for details.
