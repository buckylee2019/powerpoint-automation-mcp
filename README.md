# MCP Server Demo - PowerPoint Automation

This project demonstrates how to use MCP (Model Completion Protocol) to create a server that provides PowerPoint automation capabilities.

## Demo

https://user-images.githubusercontent.com/YOUR_USER_ID/mcp_demo.mp4

![Demo Video](./mcp_demo.mp4)

## Features

- Create, open, and save PowerPoint presentations
- Add and modify slides
- Add text boxes, images, tables, and charts
- Update text content and shape properties

## Setup

1. Create a virtual environment:
   ```
   python -m venv .venv
   ```

2. Activate the virtual environment:
   - On Windows: `.venv\Scripts\activate`
   - On macOS/Linux: `source .venv/bin/activate`

3. Install dependencies:
   ```
   pip install -e .
   ```
   
   Or if using uv:
   ```
   uv pip install -e .
   ```

4. Install python-pptx:
   ```
   pip install python-pptx
   ```

## Running the Server

To run the server:

```
uv run server
```


## Usage

The server exposes various tools for PowerPoint automation that can be called via MCP. These include:

- `initialize_powerpoint()`: Initialize PowerPoint automation
- `create_presentation()`: Create a new presentation
- `open_presentation(file_path)`: Open an existing presentation
- `add_slide(presentation_id, layout_index)`: Add a new slide
- `add_textbox(presentation_id, slide_index, text, ...)`: Add a text box
- `add_image(presentation_id, slide_index, image_path, ...)`: Add an image
- `add_table(presentation_id, slide_index, rows, cols, ...)`: Add a table
- `add_chart(presentation_id, slide_index, chart_type, ...)`: Add a chart
- `save_presentation(presentation_id, path)`: Save a presentation
- `close_presentation(presentation_id)`: Close a presentation

## MCP Config

```json
{
  "mcpServers": {
   "ppt-mcp-server": {
         "command": "uv",
         "args": [
         "run",
         "/path/to/mcp-ppt-server/server",
         ]
   }
  }
}
```
