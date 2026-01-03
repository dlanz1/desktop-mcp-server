from fastmcp import FastMCP
import pyautogui
import uiautomation as auto
import sys
import json

# Create the MCP server
mcp = FastMCP("desktop-controller")

# --- Efficient UI Context (No Screenshots!) ---

@mcp.tool()
def get_active_window() -> dict:
    """
    Gets information about the currently active/focused window.
    Returns window title, app name, position, and size.
    Much more efficient than screenshots!
    """
    try:
        window = auto.GetForegroundControl()
        if window:
            rect = window.BoundingRectangle
            return {
                "name": window.Name or "",
                "class_name": window.ClassName or "",
                "control_type": window.ControlTypeName or "",
                "position": {"x": rect.left, "y": rect.top},
                "size": {"width": rect.width(), "height": rect.height()}
            }
        return {"error": "No active window found"}
    except Exception as e:
        return {"error": str(e)}

@mcp.tool()
def get_window_text_content(max_depth: int = 3) -> dict:
    """
    Gets all readable text content from the active window.
    This is the PRIMARY way to understand what's on screen - use this instead of screenshots!
    Returns structured text content from UI elements.
    
    Args:
        max_depth: How deep to traverse the UI tree (default 3, max 5)
    """
    max_depth = min(max_depth, 5)
    
    def get_element_info(element, depth=0):
        if depth > max_depth:
            return None
        
        try:
            info = {}
            name = element.Name
            control_type = element.ControlTypeName
            
            # Only include elements with meaningful content
            if name or control_type in ['Edit', 'Text', 'Button', 'ListItem', 'MenuItem', 'TreeItem', 'TabItem', 'Document']:
                info['type'] = control_type
                if name:
                    info['text'] = name[:500]  # Limit text length
                
                # Get value for editable controls
                if control_type == 'Edit':
                    try:
                        value_pattern = element.GetValuePattern()
                        if value_pattern:
                            info['value'] = value_pattern.Value[:500]
                    except:
                        pass
                
                # Get position for interactive elements
                if control_type in ['Button', 'Edit', 'ListItem', 'MenuItem', 'TabItem', 'Link', 'CheckBox', 'RadioButton']:
                    try:
                        rect = element.BoundingRectangle
                        info['clickable_at'] = {
                            "x": rect.left + rect.width() // 2,
                            "y": rect.top + rect.height() // 2
                        }
                    except:
                        pass
                
                # Get children
                if depth < max_depth:
                    children = []
                    try:
                        for child in element.GetChildren():
                            child_info = get_element_info(child, depth + 1)
                            if child_info:
                                children.append(child_info)
                    except:
                        pass
                    if children:
                        info['children'] = children
                
                return info if (name or children if 'children' in info else False) else None
            
            # Even if this element has no name, check children
            if depth < max_depth:
                children = []
                try:
                    for child in element.GetChildren():
                        child_info = get_element_info(child, depth + 1)
                        if child_info:
                            children.append(child_info)
                except:
                    pass
                if children:
                    return {'children': children}
            
            return None
        except:
            return None
    
    try:
        window = auto.GetForegroundControl()
        if not window:
            return {"error": "No active window"}
        
        result = {
            "window_title": window.Name or "",
            "window_class": window.ClassName or "",
            "content": get_element_info(window)
        }
        return result
    except Exception as e:
        return {"error": str(e)}

@mcp.tool()
def list_all_windows() -> list:
    """
    Lists all open windows with their titles and positions.
    Useful for finding and switching to specific applications.
    """
    windows = []
    try:
        for window in auto.GetRootControl().GetChildren():
            if window.ControlTypeName == 'Window' and window.Name:
                try:
                    rect = window.BoundingRectangle
                    windows.append({
                        "title": window.Name,
                        "class": window.ClassName,
                        "position": {"x": rect.left, "y": rect.top},
                        "size": {"width": rect.width(), "height": rect.height()}
                    })
                except:
                    windows.append({"title": window.Name, "class": window.ClassName})
    except Exception as e:
        return [{"error": str(e)}]
    return windows

@mcp.tool()
def find_element(text: str, control_type: str = None) -> list:
    """
    Finds UI elements containing the specified text in the active window.
    Returns clickable coordinates for each match.
    
    Args:
        text: Text to search for (case-insensitive partial match)
        control_type: Optional filter by type (Button, Edit, Text, ListItem, MenuItem, etc.)
    """
    results = []
    text_lower = text.lower()
    
    def search_element(element, depth=0):
        if depth > 6:
            return
        
        try:
            name = element.Name or ""
            elem_type = element.ControlTypeName
            
            if text_lower in name.lower():
                if control_type is None or elem_type == control_type:
                    try:
                        rect = element.BoundingRectangle
                        results.append({
                            "text": name[:200],
                            "type": elem_type,
                            "click_x": rect.left + rect.width() // 2,
                            "click_y": rect.top + rect.height() // 2
                        })
                    except:
                        results.append({"text": name[:200], "type": elem_type})
            
            for child in element.GetChildren():
                search_element(child, depth + 1)
        except:
            pass
    
    try:
        window = auto.GetForegroundControl()
        if window:
            search_element(window)
        return results if results else [{"message": f"No elements found containing '{text}'"}]
    except Exception as e:
        return [{"error": str(e)}]

@mcp.tool()
def click_element(text: str, control_type: str = None) -> str:
    """
    Finds and clicks an element by its text content.
    Much more reliable than coordinate-based clicking!
    
    Args:
        text: Text of the element to click (case-insensitive partial match)
        control_type: Optional filter (Button, MenuItem, ListItem, Link, etc.)
    """
    elements = find_element(text, control_type)
    
    if elements and 'click_x' in elements[0]:
        x, y = elements[0]['click_x'], elements[0]['click_y']
        pyautogui.click(x, y)
        return f"Clicked '{elements[0]['text']}' at ({x}, {y})"
    elif elements and 'error' in elements[0]:
        return f"Error: {elements[0]['error']}"
    else:
        return f"Could not find clickable element containing '{text}'"

@mcp.tool()
def focus_window(title: str) -> str:
    """
    Brings a window to the foreground by its title (partial match).
    
    Args:
        title: Part of the window title to match
    """
    title_lower = title.lower()
    try:
        for window in auto.GetRootControl().GetChildren():
            if window.ControlTypeName == 'Window' and window.Name:
                if title_lower in window.Name.lower():
                    window.SetFocus()
                    return f"Focused window: {window.Name}"
        return f"No window found matching '{title}'"
    except Exception as e:
        return f"Error: {str(e)}"

# --- Screen Info ---

@mcp.tool()
def get_screen_size() -> dict:
    """Returns the width and height of the primary monitor."""
    width, height = pyautogui.size()
    return {"width": width, "height": height}

@mcp.tool()
def get_mouse_position() -> dict:
    """Returns the current mouse cursor position."""
    x, y = pyautogui.position()
    return {"x": x, "y": y}

# --- Mouse Control ---

@mcp.tool()
def move_mouse(x: int, y: int) -> str:
    """Moves the mouse cursor to the specified (x, y) coordinates."""
    pyautogui.moveTo(x, y)
    return f"Moved mouse to ({x}, {y})"

@mcp.tool()
def click_mouse(x: int = None, y: int = None, button: str = 'left') -> str:
    """
    Clicks the mouse at coordinates or current position.
    Prefer using click_element() when you know the text of what to click!
    
    Args:
        x: X coordinate (optional)
        y: Y coordinate (optional)  
        button: 'left', 'right', or 'middle'
    """
    if x is not None and y is not None:
        pyautogui.click(x=x, y=y, button=button)
        return f"Clicked {button} button at ({x}, {y})"
    else:
        pyautogui.click(button=button)
        return f"Clicked {button} button at current position"

@mcp.tool()
def double_click(x: int = None, y: int = None) -> str:
    """Double-clicks the left mouse button."""
    if x is not None and y is not None:
        pyautogui.doubleClick(x=x, y=y)
        return f"Double-clicked at ({x}, {y})"
    else:
        pyautogui.doubleClick()
        return "Double-clicked at current position"

@mcp.tool()
def drag_mouse(x: int, y: int, duration: float = 0.5) -> str:
    """Drags the mouse from current position to (x, y)."""
    pyautogui.dragTo(x, y, duration=duration)
    return f"Dragged mouse to ({x}, {y})"

@mcp.tool()
def scroll(clicks: int, x: int = None, y: int = None) -> str:
    """
    Scrolls the mouse wheel.
    
    Args:
        clicks: Number of scroll clicks (positive = up, negative = down)
        x: Optional x position to scroll at
        y: Optional y position to scroll at
    """
    if x is not None and y is not None:
        pyautogui.scroll(clicks, x=x, y=y)
        return f"Scrolled {clicks} clicks at ({x}, {y})"
    else:
        pyautogui.scroll(clicks)
        return f"Scrolled {clicks} clicks"

# --- Keyboard Control ---

@mcp.tool()
def type_text(text: str, interval: float = 0.0) -> str:
    """Types the given text string."""
    pyautogui.write(text, interval=interval)
    return f"Typed: {text}"

@mcp.tool()
def press_key(key: str) -> str:
    """
    Presses a specific key.
    Examples: 'enter', 'esc', 'tab', 'space', 'backspace', 'delete',
              'up', 'down', 'left', 'right', 'home', 'end',
              'f1'-'f12', 'win', 'alt', 'ctrl', 'shift'
    """
    pyautogui.press(key)
    return f"Pressed key: {key}"

@mcp.tool()
def hotkey(keys: list[str]) -> str:
    """
    Presses a combination of keys simultaneously.
    Examples: ['ctrl', 'c'], ['alt', 'tab'], ['ctrl', 'shift', 's']
    """
    pyautogui.hotkey(*keys)
    return f"Pressed hotkey: {'+'.join(keys)}"

# --- Fallback Screenshot (use sparingly!) ---

@mcp.tool()
def take_screenshot_region(x: int, y: int, width: int, height: int) -> str:
    """
    Takes a screenshot of a specific region only.
    Use this ONLY when UI automation cannot read the content (e.g., images, canvas, games).
    For normal UI, use get_window_text_content() instead!
    
    Args:
        x: Left position
        y: Top position
        width: Width of region
        height: Height of region
    """
    import base64
    import io
    
    screenshot = pyautogui.screenshot(region=(x, y, width, height))
    buffered = io.BytesIO()
    screenshot.save(buffered, format="PNG", optimize=True)
    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
    return img_str


if __name__ == "__main__":
    print("Starting Desktop MCP Server (UI Automation Mode)...", file=sys.stderr)
    print("Use get_window_text_content() for efficient screen reading!", file=sys.stderr)
    mcp.run()
