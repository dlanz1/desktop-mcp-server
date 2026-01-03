# Desktop MCP Server

A Windows desktop automation MCP (Model Context Protocol) server that uses UI Automation instead of screenshots for efficient token usage.

## Features

- **UI Automation-based** - Reads actual UI element text instead of screenshots (~50-100x more token efficient)
- **Smart element finding** - Find and click elements by text content
- **Window management** - List, focus, and interact with windows
- **Full input control** - Mouse, keyboard, scrolling, hotkeys

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### With Gemini CLI

Add to `~/.gemini/settings.json`:

```json
{
  "mcpServers": {
    "desktop-controller": {
      "command": "python",
      "args": ["C:\\Users\\Derek\\Documents\\github\\desktop-mcp-server\\desktop_server.py"]
    }
  }
}
```

### Standalone

```bash
python desktop_server.py
```

## Available Tools

### UI Reading (Use these first!)
| Tool | Description |
|------|-------------|
| `get_window_text_content()` | Gets all readable text from active window with clickable coordinates |
| `get_active_window()` | Get info about the focused window |
| `list_all_windows()` | List all open windows |
| `find_element(text)` | Search for UI elements by text |

### Smart Interaction
| Tool | Description |
|------|-------------|
| `click_element(text)` | Click an element by its text (recommended over coordinates) |
| `focus_window(title)` | Bring a window to foreground by title |

### Mouse Control
| Tool | Description |
|------|-------------|
| `click_mouse(x, y)` | Click at coordinates |
| `double_click(x, y)` | Double-click |
| `move_mouse(x, y)` | Move cursor |
| `drag_mouse(x, y)` | Drag to position |
| `scroll(clicks)` | Scroll mouse wheel |

### Keyboard Control
| Tool | Description |
|------|-------------|
| `type_text(text)` | Type a string |
| `press_key(key)` | Press a single key |
| `hotkey(keys)` | Press key combination |

### Fallback
| Tool | Description |
|------|-------------|
| `take_screenshot_region(x, y, w, h)` | Screenshot a small region (for images/games only) |

## Why UI Automation?

| Screenshot-based | UI Automation |
|------------------|---------------|
| ~1-5MB base64 per call | ~2-10KB structured data |
| ~50,000+ tokens | ~500-2000 tokens |
| Requires vision AI | Just text/JSON |
| Coordinates via AI guessing | Exact click coordinates provided |

## Requirements

- Windows 10/11
- Python 3.10+
