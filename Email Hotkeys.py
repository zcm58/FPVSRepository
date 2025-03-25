import tkinter as tk
from tkinter import ttk
import keyboard
import pyperclip
import time

# Dictionary to keep track of keybinds (optional for further use)
keybinds = {}


def paste_text(text):
    """Copies the text to the clipboard and simulates a paste."""
    pyperclip.copy(text)
    time.sleep(0.1)  # Brief pause to ensure clipboard update
    keyboard.press_and_release("ctrl+v")


def add_keybind():
    """Reads the hotkey and text from the GUI, registers the hotkey,
    and displays the keybind in the list."""
    hotkey = hotkey_entry.get().strip()
    text = text_entry.get("1.0", tk.END).strip()
    if not hotkey or not text:
        return  # Both fields must be non-empty

    # If the hotkey already exists, remove the previous binding.
    if hotkey in keybinds:
        keyboard.remove_hotkey(keybinds[hotkey]['id'])

    # Register the hotkey and store the reference (ID) to the binding
    hotkey_id = keyboard.add_hotkey(hotkey, lambda text=text: paste_text(text))

    # Store keybind info (could be expanded later for removal/editing)
    keybinds[hotkey] = {'text': text, 'id': hotkey_id}

    # Insert into the Treeview for display
    tree.insert("", "end", values=(hotkey, text))

    # Clear the entry fields after adding the keybind
    hotkey_entry.delete(0, tk.END)
    text_entry.delete("1.0", tk.END)


# Create main window
root = tk.Tk()
root.title("Hotkey Text Paster")

# Main frame using ttk for modern styling
frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))

# Hotkey label and entry field
hotkey_label = ttk.Label(frame, text="Hotkey (e.g., ctrl+alt+e):")
hotkey_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
hotkey_entry = ttk.Entry(frame, width=30)
hotkey_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5))

# Text label and Text widget (for multi-line input)
text_label = ttk.Label(frame, text="Text to paste:")
text_label.grid(row=1, column=0, sticky=tk.NW, pady=(0, 5))
text_entry = tk.Text(frame, width=40, height=10)
text_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 5))

# Button to add the keybind
add_button = ttk.Button(frame, text="Add Keybind", command=add_keybind)
add_button.grid(row=2, column=1, sticky=tk.E, pady=(0, 10))

# Treeview to display active keybinds and associated text
tree = ttk.Treeview(frame, columns=("Hotkey", "Text"), show="headings", height=5)
tree.heading("Hotkey", text="Hotkey")
tree.heading("Text", text="Text")
tree.column("Hotkey", width=150)
tree.column("Text", width=300)
tree.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E))

# Configure grid weights for resizing
frame.columnconfigure(1, weight=1)
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# Start the GUI event loop
root.mainloop()
