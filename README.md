# MCP Server Demo - PowerPoint Automation

This project demonstrates how to use MCP (Model Completion Protocol) to create a server that provides PowerPoint automation capabilities.

## Demo

https://github.com/user-attachments/assets/3daf3bef-4d75-4639-a891-0e64b80b4807

## Features

- Create, open, and save PowerPoint presentations
- Add and modify slides
- Add text boxes, images, tables, and charts
- Update text content and shape properties


## Running the Server

To run the server:

```
uv run --directory  /path/to/mcp-ppt-server mcp-ppt-server 
```


## Usage

The server exposes various tools for PowerPoint automation that can be called via MCP. These include:

### Presentation Management
- `initialize_powerpoint()`: Initialize PowerPoint automation
- `create_presentation()`: Create a new presentation
- `open_presentation(file_path)`: Open an existing presentation from the specified path
- `get_presentations()`: Get a list of all tracked PowerPoint presentations with their metadata
- `save_presentation(presentation_id, path=None)`: Save a presentation to disk (path is optional if presentation was previously saved)
- `close_presentation(presentation_id)`: Close a presentation

### Slide Management
- `get_slides(presentation_id)`: Get a list of all slides in a presentation
- `add_slide(presentation_id, layout_index=1)`: Add a new slide with specified layout (default is Title and Content)
  - Layout options: 0=Title slide, 1=Title and content, 2=Section header, 3=Two content, etc.
- `set_slide_title(presentation_id, slide_index, title)`: Set the title text of a slide

### Content Management
- `get_slide_text(presentation_id, slide_index)`: Get all text content in a slide
- `get_slide_shapes(presentation_id, slide_index)`: Get all shapes in a slide with their IDs and properties
- `update_text(presentation_id, slide_index, shape_index, text)`: Update the text content of a shape
- `update_shape_by_id(presentation_id, slide_index, shape_id, text=None, left=None, top=None, width=None, height=None)`: Update a shape by its ID with new properties

### Adding Content
- `add_textbox(presentation_id, slide_index, text, left=1, top=1, width=4, height=2)`: Add a text box to a slide
- `add_image(presentation_id, slide_index, image_path, left=1, top=1, width=None, height=None)`: Add an image to a slide
- `add_table(presentation_id, slide_index, rows, cols, left=1, top=1, width=8, height=4)`: Add a table to a slide
- `update_table_cell(presentation_id, slide_index, shape_index, row, col, text)`: Update the text in a table cell
- `add_chart(presentation_id, slide_index, chart_type, categories, series_names, series_values, left=1, top=1, width=8, height=4, has_legend=True)`: Add a chart to a slide
  - Chart types: 'COLUMN', 'LINE', 'PIE', 'BAR'

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
