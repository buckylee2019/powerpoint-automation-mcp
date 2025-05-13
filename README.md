# MCP Server Demo - PowerPoint Automation

This project demonstrates how to use MCP (Model Completion Protocol) to create a server that provides PowerPoint automation capabilities for Amazon Q.

## Demo

https://github.com/user-attachments/assets/3daf3bef-4d75-4639-a891-0e64b80b4807

## Features

- Create, open, and save PowerPoint presentations
- Add and modify slides
- Add text boxes, images, tables, and charts
- Update text content and shape properties
- Retrieve table content

## Running the Server

To run the server:

```
uv run --directory  /path/to/mcp-ppt-server mcp-ppt-server 
```

## Usage

The server exposes various tools for PowerPoint automation that can be called via MCP. The server is designed to work with a single active presentation at a time, which simplifies the API by eliminating the need to track and specify presentation IDs.

### Presentation Management
- `initialize_powerpoint()`: Initialize PowerPoint automation
- `create_presentation(template=None)`: Create a new presentation, optionally using a template file
- `open_presentation(file_path)`: Open an existing presentation from the specified path
- `get_presentation()`: Get information about the currently active presentation
- `save_presentation(path=None)`: Save the active presentation to disk (path is optional if presentation was previously saved)
- `close_presentation()`: Close the active presentation

### Slide Management
- `get_slides()`: Get a list of all slides in the active presentation
- `add_slide(layout_index=1)`: Add a new slide with specified layout (default is Title and Content)
  - Layout options: 0=Title slide, 1=Title and content, 2=Section header, 3=Two content, etc.
- `set_slide_title(slide_index, title)`: Set the title text of a slide
- `get_slide_layouts()`: Get all available slide layouts in the active presentation
- `delete_slide(slide_index)`: Delete a slide from the presentation

### Content Management
- `get_slide_text(slide_index)`: Get all text content in a slide
- `get_slide_shapes(slide_index)`: Get all shapes in a slide with their IDs and properties
- `update_text(slide_index, shape_index, text)`: Update the text content of a shape
- `update_shape_by_id(slide_index, shape_id, text=None, left=None, top=None, width=None, height=None)`: Update a shape by its ID with new properties
- `update_table_cell(slide_index, shape_index, row, col, text)`: Update the text in a table cell

### Adding Content
- `add_textbox(slide_index, text, left=1, top=1, width=4, height=2)`: Add a text box to a slide
- `add_image(slide_index, image_path, left=1, top=1, width=None, height=None)`: Add an image to a slide
- `add_table(slide_index, rows, cols, left=1, top=1, width=8, height=4)`: Add a table to a slide
- `get_table_content(slide_index, shape_index)`: Retrieve the content of a table in a slide
- `add_chart(slide_index, chart_type, categories, series_names, series_values, left=1, top=1, width=8, height=4, has_legend=True)`: Add a chart to a slide with multiple data series
  - Chart types: 'COLUMN', 'LINE', 'PIE', 'BAR'
  - `categories`: List of category names (x-axis labels)
  - `series_names`: List of series names for the legend
  - `series_values`: List of lists containing values for each series

## Error Handling

Most functions return a dictionary with either success information or error details. Always check for the presence of an "error" key in the returned dictionary to handle errors appropriately.

Example error response:
```json
{
  "error": "No active presentation. Please open or create a presentation first."
}
```

## MCP Config

```json
{
  "mcpServers": {
   "ppt-mcp-server": {
         "command": "uv",
         "args": [
         "run",
         "--directory",
         "/path/to/mcp-ppt-server",
         "mcp-ppt-server"
         ]
   }
  }
}
```
