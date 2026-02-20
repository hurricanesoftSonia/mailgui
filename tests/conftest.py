"""Shared test fixtures — mock tkinter before mailgui is imported."""
import sys
from unittest.mock import MagicMock

# Mock tkinter and related modules so mailgui can be imported in headless CI
for mod in ("tkinter", "tkinter.ttk", "tkinter.messagebox", "tkinter.filedialog",
            "tkinter.scrolledtext", "tkinter.font"):
    if mod not in sys.modules:
        sys.modules[mod] = MagicMock()
