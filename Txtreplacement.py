#!/usr/bin/env python3
"""
Windows Text Replacement Program
A system-wide text expander similar to macOS Text Replacement
Author: Claude Assistant
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json
import os
import sys
import threading
import time
from pynput import keyboard
from pynput.keyboard import Key, Listener
import pyperclip
import win32clipboard
import win32con
import win32gui
import win32api
from datetime import datetime
import pystray
from PIL import Image, ImageDraw

class TextReplacer:
    def __init__(self):
        self.replacements = {}
        # Use user's AppData folder to avoid permission issues
        self.config_file = os.path.join(os.path.expanduser("~"), "AppData", "Local", "text_replacements.json")
        self.buffer = ""
        self.max_buffer_length = 50
        self.listener = None
        self.running = False
        self.load_replacements()
        
    def load_replacements(self):
        """Load replacements from JSON file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.replacements = json.load(f)
            else:
                # Default replacements
                self.replacements = {
                    "@@": "steve@workperfect.io",
                    "myname": "Steve Derviniotis",
                    "wp": "Work Perfect",
                    "tgif": "Thank God It's Friday!",
                    "addr": "Adelaide, South Australia",
                    "sig": "Best regards,\nSteve Derviniotis\nHead of Customer Success\nWork Perfect",
                    "today": datetime.now().strftime("%Y-%m-%d"),
                    "time": datetime.now().strftime("%H:%M:%S")
                }
                self.save_replacements()
        except Exception as e:
            print(f"Error loading replacements: {e}")
            self.replacements = {}
    
    def save_replacements(self):
        """Save replacements to JSON file"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.replacements, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving replacements: {e}")
    
    def on_press(self, key):
        """Handle key press events"""
        try:
            if hasattr(key, 'char') and key.char:
                # Add character to buffer
                self.buffer += key.char
                
                # Keep buffer size manageable
                if len(self.buffer) > self.max_buffer_length:
                    self.buffer = self.buffer[-self.max_buffer_length:]
                
                # Check for replacements
                self.check_replacements()
                
        except AttributeError:
            # Special keys (ctrl, alt, etc.)
            if key == Key.space or key == Key.enter or key == Key.tab:
                self.check_replacements()
                self.buffer = ""
            elif key == Key.backspace:
                if self.buffer:
                    self.buffer = self.buffer[:-1]
    
    def check_replacements(self):
        """Check if any text in buffer should be replaced"""
        for shortcut, replacement in self.replacements.items():
            if self.buffer.endswith(shortcut):
                # Perform replacement
                self.perform_replacement(shortcut, replacement)
                break
    
    def perform_replacement(self, shortcut, replacement):
        """Perform the actual text replacement"""
        try:
            # Handle dynamic replacements
            if shortcut == "today":
                replacement = datetime.now().strftime("%Y-%m-%d")
            elif shortcut == "time":
                replacement = datetime.now().strftime("%H:%M:%S")
            
            # Clear the buffer
            self.buffer = ""
            
            # Send backspaces to delete the shortcut
            for _ in range(len(shortcut)):
                win32api.keybd_event(0x08, 0, 0, 0)  # Backspace down
                win32api.keybd_event(0x08, 0, win32con.KEYEVENTF_KEYUP, 0)  # Backspace up
                time.sleep(0.01)
            
            # Type the replacement text
            self.type_text(replacement)
            
        except Exception as e:
            print(f"Error performing replacement: {e}")
    
    def type_text(self, text):
        """Type text using clipboard method (more reliable for complex text)"""
        try:
            # Save current clipboard content
            old_clipboard = ""
            try:
                old_clipboard = pyperclip.paste()
            except:
                pass
            
            # Put replacement text in clipboard
            pyperclip.copy(text)
            time.sleep(0.1)
            
            # Paste the text (Ctrl+V)
            win32api.keybd_event(0x11, 0, 0, 0)  # Ctrl down
            win32api.keybd_event(0x56, 0, 0, 0)  # V down
            win32api.keybd_event(0x56, 0, win32con.KEYEVENTF_KEYUP, 0)  # V up
            win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)  # Ctrl up
            
            # Restore old clipboard content after a delay
            def restore_clipboard():
                time.sleep(1)
                try:
                    pyperclip.copy(old_clipboard)
                except:
                    pass
            
            threading.Thread(target=restore_clipboard, daemon=True).start()
            
        except Exception as e:
            print(f"Error typing text: {e}")
    
    def start_listener(self):
        """Start the keyboard listener"""
        self.running = True
        self.listener = Listener(on_press=self.on_press)
        self.listener.start()
    
    def stop_listener(self):
        """Stop the keyboard listener"""
        self.running = False
        if self.listener:
            self.listener.stop()


class TextReplacerGUI:
    def __init__(self):
        self.replacer = TextReplacer()
        self.root = tk.Tk()
        self.root.title("Windows Text Replacement")
        self.root.geometry("800x600")
        self.setup_gui()
        self.setup_system_tray()
        
    def setup_gui(self):
        """Setup the main GUI"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Text Replacement Manager", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Status
        self.status_var = tk.StringVar(value="Stopped")
        status_label = ttk.Label(main_frame, text="Status:")
        status_label.grid(row=1, column=0, sticky=tk.W)
        self.status_display = ttk.Label(main_frame, textvariable=self.status_var, 
                                       foreground="red")
        self.status_display.grid(row=1, column=1, sticky=tk.W)
        
        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))
        
        self.start_btn = ttk.Button(button_frame, text="Start", command=self.start_service)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_btn = ttk.Button(button_frame, text="Stop", command=self.stop_service, 
                                  state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Replacements list
        list_frame = ttk.LabelFrame(main_frame, text="Text Replacements", padding="10")
        list_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Treeview for replacements
        columns = ("Shortcut", "Replacement")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        
        # Define headings
        self.tree.heading("Shortcut", text="Shortcut")
        self.tree.heading("Replacement", text="Replacement Text")
        
        # Define column widths
        self.tree.column("Shortcut", width=150)
        self.tree.column("Replacement", width=400)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Grid the treeview and scrollbar
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Buttons for managing replacements
        btn_frame = ttk.Frame(list_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0), sticky=(tk.W, tk.E))
        
        ttk.Button(btn_frame, text="Add", command=self.add_replacement).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Edit", command=self.edit_replacement).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Delete", command=self.delete_replacement).pack(side=tk.LEFT, padx=(0, 10))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Load initial data
        self.refresh_list()
        
        # Bind double-click to edit
        self.tree.bind("<Double-1>", lambda e: self.edit_replacement())
    
    def setup_system_tray(self):
        """Setup system tray functionality"""
        # Create system tray icon
        def create_image():
            # Create a simple icon
            image = Image.new('RGB', (64, 64), color=(0, 100, 200))
            draw = ImageDraw.Draw(image)
            draw.text((20, 20), "TR", fill=(255, 255, 255))
            return image
        
        # System tray menu
        def show_window(icon, item):
            self.root.deiconify()
            self.root.lift()
            self.root.attributes('-topmost', True)
            self.root.after_idle(self.root.attributes, '-topmost', False)
        
        def quit_app(icon, item):
            self.replacer.stop_listener()
            icon.stop()
            self.root.quit()
        
        def toggle_service(icon, item):
            if self.replacer.running:
                self.stop_service()
            else:
                self.start_service()
        
        menu = pystray.Menu(
            pystray.MenuItem("Show Window", show_window, default=True),
            pystray.MenuItem("Toggle Service", toggle_service),
            pystray.MenuItem("Quit", quit_app)
        )
        
        try:
            self.tray_icon = pystray.Icon("TextReplacer", create_image(), "Text Replacer", menu)
        except:
            # Fallback if PIL/pystray not available
            self.tray_icon = None
        
        # Handle window close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def refresh_list(self):
        """Refresh the replacements list"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add replacements
        for shortcut, replacement in self.replacer.replacements.items():
            # Truncate long replacements for display
            display_replacement = replacement[:100] + "..." if len(replacement) > 100 else replacement
            display_replacement = display_replacement.replace('\n', '\\n')
            self.tree.insert("", tk.END, values=(shortcut, display_replacement))
    
    def start_service(self):
        """Start the text replacement service"""
        try:
            self.replacer.start_listener()
            self.status_var.set("Running")
            self.status_display.config(foreground="green")
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            messagebox.showinfo("Started", "Text replacement service started!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start service: {e}")
    
    def stop_service(self):
        """Stop the text replacement service"""
        try:
            self.replacer.stop_listener()
            self.status_var.set("Stopped")
            self.status_display.config(foreground="red")
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            messagebox.showinfo("Stopped", "Text replacement service stopped!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to stop service: {e}")
    
    def add_replacement(self):
        """Add a new text replacement"""
        dialog = ReplacementDialog(self.root, "Add Replacement")
        if dialog.result:
            shortcut, replacement = dialog.result
            if shortcut in self.replacer.replacements:
                if not messagebox.askyesno("Confirm", f"Shortcut '{shortcut}' already exists. Replace it?"):
                    return
            
            self.replacer.replacements[shortcut] = replacement
            self.replacer.save_replacements()
            self.refresh_list()
    
    def edit_replacement(self):
        """Edit selected replacement"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a replacement to edit.")
            return
        
        item = selection[0]
        shortcut = self.tree.item(item)["values"][0]
        current_replacement = self.replacer.replacements[shortcut]
        
        dialog = ReplacementDialog(self.root, "Edit Replacement", shortcut, current_replacement)
        if dialog.result:
            new_shortcut, new_replacement = dialog.result
            
            # Remove old entry if shortcut changed
            if new_shortcut != shortcut:
                del self.replacer.replacements[shortcut]
            
            self.replacer.replacements[new_shortcut] = new_replacement
            self.replacer.save_replacements()
            self.refresh_list()
    
    def delete_replacement(self):
        """Delete selected replacement"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a replacement to delete.")
            return
        
        item = selection[0]
        shortcut = self.tree.item(item)["values"][0]
        
        if messagebox.askyesno("Confirm Delete", f"Delete replacement '{shortcut}'?"):
            del self.replacer.replacements[shortcut]
            self.replacer.save_replacements()
            self.refresh_list()
    
    def on_closing(self):
        """Handle window closing - minimize to tray instead of closing"""
        if self.tray_icon:
            self.root.withdraw()  # Hide window
            if not hasattr(self, 'tray_started'):
                # Start system tray in background thread
                def run_tray():
                    self.tray_icon.run()
                
                threading.Thread(target=run_tray, daemon=True).start()
                self.tray_started = True
        else:
            # Fallback: minimize to taskbar
            self.root.iconify()
    
    def quit_application(self):
        """Actually quit the application"""
        self.replacer.stop_listener()
        if hasattr(self, 'tray_icon') and self.tray_icon:
            self.tray_icon.stop()
        self.root.destroy()
    
    def run(self):
        """Run the GUI"""
        self.root.mainloop()


class ReplacementDialog:
    def __init__(self, parent, title, shortcut="", replacement=""):
        self.result = None
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("500x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.geometry("+%d+%d" % (parent.winfo_rootx() + 50, parent.winfo_rooty() + 50))
        
        # Create form
        frame = ttk.Frame(self.dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Shortcut field
        ttk.Label(frame, text="Shortcut:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.shortcut_var = tk.StringVar(value=shortcut)
        shortcut_entry = ttk.Entry(frame, textvariable=self.shortcut_var, width=30)
        shortcut_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Replacement field
        ttk.Label(frame, text="Replacement Text:").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        self.replacement_text = tk.Text(frame, height=10, width=50)
        self.replacement_text.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        self.replacement_text.insert("1.0", replacement)
        
        # Scrollbar for text area
        text_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.replacement_text.yview)
        self.replacement_text.configure(yscrollcommand=text_scroll.set)
        text_scroll.grid(row=3, column=1, sticky=(tk.N, tk.S), pady=(0, 15))
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=4, column=0, sticky=(tk.W, tk.E))
        
        ttk.Button(btn_frame, text="OK", command=self.ok_clicked).pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(btn_frame, text="Cancel", command=self.cancel_clicked).pack(side=tk.RIGHT)
        
        # Configure grid
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)
        
        # Focus on shortcut entry
        shortcut_entry.focus()
        
        # Bind Enter key to OK (when not in text area)
        shortcut_entry.bind("<Return>", lambda e: self.ok_clicked())
        
        # Wait for dialog to close
        self.dialog.wait_window()
    
    def ok_clicked(self):
        shortcut = self.shortcut_var.get().strip()
        replacement = self.replacement_text.get("1.0", tk.END).strip()
        
        if not shortcut:
            messagebox.showwarning("Invalid Input", "Shortcut cannot be empty.")
            return
        
        if not replacement:
            messagebox.showwarning("Invalid Input", "Replacement text cannot be empty.")
            return
        
        self.result = (shortcut, replacement)
        self.dialog.destroy()
    
    def cancel_clicked(self):
        self.dialog.destroy()


def main():
    """Main function to run the application"""
    try:
        app = TextReplacerGUI()
        app.run()
    except Exception as e:
        print(f"Error running application: {e}")
        messagebox.showerror("Error", f"Failed to start application: {e}")


if __name__ == "__main__":
    main()
