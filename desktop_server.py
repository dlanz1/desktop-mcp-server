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
    Gets text content from the active window.
    """
    max_depth = min(max_depth, 5)
    
    def get_element_info(element, depth=0):
        if depth > max_depth:
            return None
        
        try:
            # Skip invisible elements
            try:
                rect = element.BoundingRectangle
                if rect.width() == 0 or rect.height() == 0:
                    return None
            except:
                pass

            info = {}
            name = element.Name
            control_type = element.ControlTypeName
            
            if name:
                info['text'] = name[:200]
            
            if control_type == 'Edit':
                try:
                    value_pattern = element.GetValuePattern()
                    if value_pattern:
                        val = value_pattern.Value
                        if val: info['value'] = val[:200]
                except:
                    pass
            
            if control_type in ['Button', 'Edit', 'ListItem', 'MenuItem', 'TabItem', 'Link', 'CheckBox', 'RadioButton']:
                try:
                    rect = element.BoundingRectangle
                    info['clickable_at'] = {
                        "x": rect.left + rect.width() // 2,
                        "y": rect.top + rect.height() // 2
                    }
                except:
                    pass
            
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
            
            # Only return if meaningful content
            if 'text' in info or 'value' in info or 'children' in info:
                info['type'] = control_type
                return info
            
            return None
        except:
            return None
    
    try:
        window = auto.GetForegroundControl()
        if not window:
            return {"error": "No active window"}
        
        return {
            "title": window.Name or "",
            "content": get_element_info(window)
        }
    except Exception as e:
        return {"error": str(e)}

@mcp.tool()
def list_all_windows() -> list:
    """
    Lists all visible open windows (width/height > 0) with their titles and positions.
    Includes detected dialogs/child windows.
    """
    windows = []
    try:
        root = auto.GetRootControl()
        for window in root.GetChildren():
            # Check top-level window
            if (window.ControlTypeName == 'Window' or window.ControlTypeName == 'WindowControl') and window.Name:
                try:
                    rect = window.BoundingRectangle
                    if rect.width() > 0 and rect.height() > 0:
                        windows.append({
                            "title": window.Name,
                            "class": window.ClassName,
                            "position": {"x": rect.left, "y": rect.top},
                            "size": {"width": rect.width(), "height": rect.height()}
                        })
                except:
                    pass
                
                # Check for immediate child windows (dialogs)
                try:
                    for child in window.GetChildren():
                        if child.ControlTypeName == 'WindowControl' and child.Name:
                            c_rect = child.BoundingRectangle
                            if c_rect.width() > 0 and c_rect.height() > 0:
                                windows.append({
                                    "title": child.Name,
                                    "class": child.ClassName,
                                    "position": {"x": c_rect.left, "y": c_rect.top},
                                    "size": {"width": c_rect.width(), "height": c_rect.height()},
                                    "parent": window.Name
                                })
                except:
                    pass
    except Exception as e:
        return [{"error": str(e)}]
    return windows

def _find_element(text: str, control_type: str = None) -> list:
    """Helper function for finding elements."""
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
        # First try active window
        window = auto.GetForegroundControl()
        if window:
            search_element(window)
        
        # If nothing found, try from root (depth-limited)
        if not results:
            search_element(auto.GetRootControl(), depth=2) # Start deeper to avoid infinite root recursion
            
        return results if results else [{"message": f"No elements found containing '{text}'"}]
    except Exception as e:
        return [{"error": str(e)}]

@mcp.tool()
def find_element(text: str, control_type: str = None) -> list:
    """
    Finds UI elements containing the specified text in the active window.
    Returns clickable coordinates for each match.
    
    Args:
        text: Text to search for (case-insensitive partial match)
        control_type: Optional filter by type (Button, Edit, Text, ListItem, MenuItem, etc.)
    """
    return _find_element(text, control_type)

@mcp.tool()
def click_element(text: str, control_type: str = None) -> str:
    """
    Finds and clicks an element by its text content.
    Much more reliable than coordinate-based clicking!
    
    Args:
        text: Text of the element to click (case-insensitive partial match)
        control_type: Optional filter (Button, MenuItem, ListItem, Link, etc.)
    """
    elements = _find_element(text, control_type)
    
    if elements and 'click_x' in elements[0]:
        x, y = elements[0]['click_x'], elements[0]['click_y']
        try:
            pyautogui.click(x, y)
            return f"Clicked '{elements[0]['text']}' at ({x}, {y})"
        except Exception as e:
            return f"Error clicking: {str(e)}"
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
        root = auto.GetRootControl()
        for window in root.GetChildren():
            if (window.ControlTypeName == 'Window' or window.ControlTypeName == 'WindowControl') and window.Name:
                if title_lower in window.Name.lower():
                    window.SetFocus()
                    return f"Focused window: {window.Name}"
                
                # Check children (dialogs)
                for child in window.GetChildren():
                    if child.ControlTypeName == 'WindowControl' and child.Name:
                        if title_lower in child.Name.lower():
                            child.SetFocus()
                            return f"Focused window: {child.Name} (child of {window.Name})"
                            
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
def mouse_action(action: str, x: int = None, y: int = None, button: str = 'left', clicks: int = 1, duration: float = 0.5) -> str:
    """
    Performs mouse actions.
    Actions: 'move', 'click', 'double_click', 'drag', 'scroll'.
    Args:
        action: 'move', 'click', 'double_click', 'drag', or 'scroll'.
        x, y: Coordinates (optional for click/scroll, required for move/drag).
        button: 'left', 'right', 'middle' (for click).
        clicks: Number of clicks (for scroll).
        duration: Drag duration.
    """
    try:
        if action == 'move':
            if x is None or y is None: return "Error: x and y required for move"
            pyautogui.moveTo(x, y)
            return f"Moved to ({x}, {y})"
        elif action == 'click':
            if x is not None and y is not None:
                pyautogui.click(x=x, y=y, button=button)
                return f"Clicked {button} at ({x}, {y})"
            else:
                pyautogui.click(button=button)
                return f"Clicked {button} at current pos"
        elif action == 'double_click':
            if x is not None and y is not None:
                pyautogui.doubleClick(x=x, y=y)
                return f"Double-clicked at ({x}, {y})"
            else:
                pyautogui.doubleClick()
                return "Double-clicked at current pos"
        elif action == 'drag':
            if x is None or y is None: return "Error: x and y required for drag"
            pyautogui.dragTo(x, y, duration=duration)
            return f"Dragged to ({x}, {y})"
        elif action == 'scroll':
            if x is not None and y is not None:
                pyautogui.scroll(clicks, x=x, y=y)
            else:
                pyautogui.scroll(clicks)
            return f"Scrolled {clicks} clicks"
        else:
            return f"Unknown action: {action}"
    except Exception as e:
        return f"Error: {str(e)}"

# --- Keyboard Control ---

@mcp.tool()
def keyboard_action(action: str, text: str = None, key: str = None, keys: list[str] = None, interval: float = 0.0) -> str:
    """
    Performs keyboard actions.
    Actions: 'type', 'press', 'hotkey'.
    Args:
        action: 'type', 'press', or 'hotkey'.
        text: Text to type.
        key: Key to press.
        keys: List of keys for hotkey.
        interval: Typing interval.
    """
    try:
        if action == 'type':
            if not text: return "Error: text required"
            pyautogui.write(text, interval=interval)
            return f"Typed: {text}"
        elif action == 'press':
            if not key: return "Error: key required"
            pyautogui.press(key)
            return f"Pressed: {key}"
        elif action == 'hotkey':
            if not keys: return "Error: keys required"
            pyautogui.hotkey(*keys)
            return f"Hotkey: {'+'.join(keys)}"
        else:
            return f"Unknown action: {action}"
    except Exception as e:
        return f"Error: {str(e)}"

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
