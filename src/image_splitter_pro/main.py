import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from ttkbootstrap import Style, Button as TBButton
import os
import csv
from pathlib import Path
from PIL import Image, ImageTk  # 449 Youll need pip install pillow
import sys  # Needed for sys.executable
import subprocess  # Needed to launch external scripts
import json  # To load profile JSON data
import re  # To sanitize profile names
import shutil  # To move files
from datetime import datetime  # To generate timestamp folders
import time
import webbrowser  # Open support link in default browser
from send2trash import send2trash

# HEIC support
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIC_SUPPORTED = True
except ImportError:
    HEIC_SUPPORTED = False

# Drag-and-drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    # Fallback if tkinterdnd2 isn't installed
    TkinterDnD = tk.Tk
    DND_FILES = None


#Image Splitter Pro
#Author: Abel Aramburo (@AbelXL) (https://github.com/AbelXL) (https://www.abelxl.com/)
#Created: 2026-01-19
#Copyright (c) 2026 Abel Aramburo
#This project was developed with the assistance of AI logic-modeling to ensure high-performance image handling and a modern user experience.
#This project is licensed under the **MIT License**. This means you are free to use, modify, and distribute the software, provided that the original copyright notice and this permission notice are included in all copies or substantial portions of the software.


# Removed file-based debug logging per user request (no image_wizard_debug.log will be created)


# Prefer a local `config` folder next to the script or executable on all OSes.
# This makes profiles and settings local to the application folder (or the
# exe's folder when frozen) instead of using per-user roaming AppData.
try:
    # When frozen by PyInstaller the executable lives at sys.executable; put
    # the config next to that exe so it remains local to the distribution.
    if getattr(sys, 'frozen', False):
        _base_dir = os.path.dirname(sys.executable)
    else:
        # Running from source: place config next to this module file.
        _base_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    # Fallback to the user's home directory if anything unexpected occurs.
    _base_dir = os.path.expanduser('~')

CONFIG_FOLDER = os.path.join(_base_dir, 'config')
CONFIG_FILE = os.path.join(CONFIG_FOLDER, 'config.csv')


# Lightweight fallback label object with a no-op config method to avoid AttributeError
class _NullLabel:
    def config(self, *args, **kwargs):
        return None


# Helper: centralized saving with desired defaults
def _save_image_preset(img: Image.Image, out_path: str, quality: int = 95):
    """Save an Image with sane defaults per format.
    - JPEG/JPG: save as JPEG, convert to RGB if needed, quality and no subsampling.
    - WEBP: save with quality.
    - PNG: optimized save.
    - HEIC/HEIF: convert to JPEG.
    - Otherwise: fallback to Image.save.
    This ensures we use quality=95 by default instead of Pillow's implicit defaults.
    """
    ext = os.path.splitext(out_path)[1].lower()
    fmt = None
    if ext in ('.jpg', '.jpeg'):
        fmt = 'JPEG'
    elif ext == '.webp':
        fmt = 'WEBP'
    elif ext == '.png':
        fmt = 'PNG'
    elif ext in ('.heic', '.heif'):
        # Convert HEIC/HEIF to JPEG
        fmt = 'JPEG'
        out_path = os.path.splitext(out_path)[0] + '.jpg'

    try:
        img_to_save = img
        if fmt == 'JPEG':
            # JPEG requires RGB
            if getattr(img, 'mode', None) != 'RGB':
                try:
                    img_to_save = img.convert('RGB')
                except Exception:
                    img_to_save = img
            img_to_save.save(out_path, format='JPEG', quality=quality, subsampling=0, optimize=True)
        elif fmt == 'WEBP':
            img_to_save.save(out_path, format='WEBP', quality=quality)
        elif fmt == 'PNG':
            img_to_save.save(out_path, format='PNG', optimize=True)
        else:
            img_to_save.save(out_path)
    except Exception:
        # best-effort fallback
        try:
            img.save(out_path)
        except Exception:
            pass


# ====================== CONFIG SYSTEM (outside the class) ======================


def ensure_config_exists():
    # If the app was run from a PyInstaller onefile bundle, resources may have been
    # extracted to sys._MEIPASS in a temp folder. If there is a config folder there
    # (from earlier runs), migrate its contents to the persistent config folder
    # located next to the final exe (CONFIG_FOLDER).
    try:
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            meipass_cfg = os.path.join(sys._MEIPASS, 'config')
            if os.path.exists(meipass_cfg) and not os.path.exists(CONFIG_FOLDER):
                try:
                    os.makedirs(CONFIG_FOLDER, exist_ok=True)
                    for name in os.listdir(meipass_cfg):
                        s = os.path.join(meipass_cfg, name)
                        d = os.path.join(CONFIG_FOLDER, name)
                        try:
                            if os.path.isdir(s):
                                import shutil as _sh
                                _sh.copytree(s, d)
                            else:
                                import shutil as _sh
                                _sh.copy2(s, d)
                        except Exception:
                            # ignore per-file copy errors; continue best-effort
                            pass
                except Exception:
                    pass
    except Exception:
        pass

    if not os.path.exists(CONFIG_FOLDER):
        os.makedirs(CONFIG_FOLDER)
        print("Created folder: config")

    # If the config file doesn't exist, create with sane defaults
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            f.write("setting,value\n")
            f.write("source_folder,\n")
            f.write("destination_folder,\n")
            # Persist user's preference whether to delete originals after cropping
            f.write("delete_original_after_cropping,False\n")
            # Persist user's preference whether to delete originals after moving
            f.write("delete_original_after_moving,False\n")
            # Persist whether to show the confirmation dialog before deleting originals after cropping.
            # Default: True => show the confirmation dialog (user can opt-out with the checkbox).
            f.write("confirm_delete_after_cropping,True\n")
            # Persist whether to show the confirmation dialog before deleting originals after moving.
            # Default: True => show the confirmation dialog (user can opt-out with the checkbox).
            f.write("confirm_delete_after_moving,True\n")
            # Persist whether to show the onboarding / first-run wizard. Default True shows it on first run.
            f.write("show_onboarding,True\n")
        print("Created: config/config.csv")
    else:
        # Migration: if an older key 'confirm_delete_originals' exists, rename it to the new key
        try:
            rows = []
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                rows = list(csv.reader(f))

            # Convert to dict for easy lookup
            cfg = {r[0]: r[1] if len(r) > 1 else '' for r in rows if r}
            migrated = False
            if 'confirm_delete_originals' in cfg and 'confirm_delete_after_cropping' not in cfg:
                cfg['confirm_delete_after_cropping'] = cfg.get('confirm_delete_originals', 'True')
                try:
                    del cfg['confirm_delete_originals']
                except Exception:
                    pass
                migrated = True

            # If migration happened, write back CSV preserving other keys/order loosely
            if migrated:
                out_rows = [["setting", "value"]]
                for k, v in cfg.items():
                    out_rows.append([k, v])
                with open(CONFIG_FILE, 'w', encoding='utf-8', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerows(out_rows)
                print("Migrated config: confirm_delete_originals -> confirm_delete_after_cropping")
        except Exception:
            # best-effort: ignore migration errors to avoid breaking startup
            pass


def save_config(setting: str, value: str):
    rows = []
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            rows = list(csv.reader(f))
    except:
        rows = [["setting", "value"]]

    found = False
    for i, row in enumerate(rows):
        if len(row) >= 2 and row[0] == setting:
            rows[i] = [setting, value]
            found = True
    if not found:
        rows.append([setting, value])

    with open(CONFIG_FILE, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(rows)


def load_config(setting: str) -> str:
    if not os.path.exists(CONFIG_FILE):
        return ""
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) >= 2 and row[0] == setting:
                return row[1]
    return ""


def load_profiles() -> list[str]:
    """Scans the CONFIG_FOLDER for files ending in .profile and returns a list of their base names."""
    profiles = []
    if os.path.exists(CONFIG_FOLDER):
        for item in os.listdir(CONFIG_FOLDER):
            if item.lower().endswith(".profile") and os.path.isfile(os.path.join(CONFIG_FOLDER, item)):
                profiles.append(item[:-8])
    return profiles


def load_profile_rules(profile_name: str) -> dict[int, list[dict]] | None:
    """Loads the cropping rules from a specified profile file."""
    if profile_name == "— No Profile Selected —":
        return None

    safe_name = re.sub(r'[\\/:*?"<>|]', '_', profile_name)
    file_path = os.path.join(CONFIG_FOLDER, f"{safe_name}.profile")

    if not os.path.exists(file_path):
        return None

    # Group rules by position into a list (preserve file order so multiple rules
    # targeting the same position are applied in file order)
    rules_map: dict[int, list[dict]] = {}
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            profile_data = json.load(f)
            rules_list = profile_data.get('rules', [])

            def _position_of(rule_obj):
                # Support both 'position' (current) and older 'position_number' keys
                pos_raw = rule_obj.get('position', rule_obj.get('position_number', None))
                try:
                    return int(pos_raw)
                except Exception:
                    return 0

            # Annotate each rule with its file-order index so we can later
            # sort applications in rule order. Keep grouping by position.
            for idx, rule in enumerate(rules_list, start=1):
                # store original rule index for ordering
                rule['_rule_index'] = idx
                position = _position_of(rule)
                if position > 0:
                    rules_map.setdefault(position, []).append(rule)
        return rules_map
    except Exception as e:
        print(f"ERROR loading profile '{profile_name}': {e}")
        return None


ensure_config_exists()

# Fixed left panel width (pixels). Change this value to manually adjust the left column width.
LEFT_PANEL_FIXED_WIDTH = 200


# ========================= MAIN APPLICATION CLASS =========================
class ImageCroppingApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        print("[image_wizard] __init__ start")
        try:
            # When the window is closed, just destroy it (no debug logging)
            self.protocol('WM_DELETE_WINDOW', self.destroy)
            # Bind destroy event without logging
            self.bind('<Destroy>', lambda e: None)
            # Do not override Tk's exception handler — avoid file logging
        except Exception:
            pass

        # Initialize ttkbootstrap style (Cosmo theme)
        try:
            self.style = Style('cosmo')
            print("[image_wizard] style initialized")
        except Exception:
            # fallback: continue without crashing
            self.style = None
            print("[image_wizard] style failed, continuing")

        # Set window title and geometry FIRST to prevent white box flash
        self.title("Image Splitter Pro v1.0.0 Beta")
        self.geometry("900x600")

        # Set Windows AppUserModelID for proper taskbar grouping and icon display
        try:
            if sys.platform == 'win32':
                import ctypes
                # Set a unique AppUserModelID so Windows shows our custom icon in the taskbar
                myappid = 'ImageSplitterPro.CropEditor.1.0'
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            # Non-fatal if this fails (older Windows or other issues)
            pass

        # --- set window icon (supports running from source and PyInstaller --onefile) ---
        try:
            icon_filename = os.path.join("_internal", "icon.ico")

            # Try multiple paths in order of preference
            icon_path = None

            # 1. If running from a PyInstaller bundle, check _MEIPASS first
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                bundled_icon = os.path.join(sys._MEIPASS, icon_filename)
                if os.path.exists(bundled_icon):
                    icon_path = bundled_icon

            # 2. Check next to the script/exe
            if icon_path is None:
                if getattr(sys, 'frozen', False):
                    # Running as compiled exe
                    local_icon = os.path.join(os.path.dirname(sys.executable), icon_filename)
                else:
                    # Running from source
                    local_icon = os.path.join(os.path.dirname(os.path.abspath(__file__)), icon_filename)
                if os.path.exists(local_icon):
                    icon_path = local_icon

            # 3. Try setting the icon if we found it
            if icon_path:
                # On Linux/Unix, wm_iconphoto is the standard method
                # On Windows, iconbitmap works better
                if sys.platform.startswith('linux') or sys.platform == 'darwin':
                    try:
                        # Use PIL.Image -> ImageTk.PhotoImage for Linux/Mac
                        img = Image.open(icon_path)
                        photo = ImageTk.PhotoImage(img)
                        self.wm_iconphoto(True, photo)
                        # retain reference so the image isn't garbage-collected
                        self._app_icon = photo
                    except Exception:
                        pass
                else:
                    # Windows: try iconbitmap first, fallback to wm_iconphoto
                    try:
                        self.iconbitmap(icon_path)
                    except Exception:
                        try:
                            img = Image.open(icon_path)
                            photo = ImageTk.PhotoImage(img)
                            self.wm_iconphoto(True, photo)
                            self._app_icon = photo
                        except Exception:
                            pass
        except Exception:
            # Non-fatal: don't prevent the app from launching if icon setup fails
            pass


        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(0, weight=1)

        self.file_paths = {
            "source_folder": "",
            "destination_folder": ""
        }

        # --- STATE FOR DYNAMIC POLLING ---
        self.last_seen_files = []
        self._last_width = 0
        self.profile_process = None

        # --- STATE FOR THUMBNAIL SELECTION ---
        self.selected_thumbnail_index = None  # Index of currently selected thumbnail
        self.thumbnail_widgets = []  # List of (image_label, name_label, filename) tuples

        saved_source = load_config("source_folder")
        saved_dest = load_config("destination_folder")
        print(f"[image_wizard] saved_source={saved_source!r}, saved_dest={saved_dest!r}")
        if saved_source and os.path.isdir(saved_source):
            self.file_paths["source_folder"] = saved_source
        if saved_dest and os.path.isdir(saved_dest):
            # use the correctly named local variable
            self.file_paths["destination_folder"] = saved_dest
        print(f"[image_wizard] file_paths after load: {self.file_paths}")

        self.create_widgets()
        print("[image_wizard] create_widgets done")

        # --- BIND RESIZE EVENT ---
        self.bind("<Configure>", self.on_window_resize)

        # --- START THE POLLING LOOP ---
        self.start_auto_refresh()
        print("[image_wizard] start_auto_refresh scheduled")
        try:
            self.after(300, self._maybe_show_onboarding)
        except Exception:
            pass
        print("[image_wizard] __init__ end")

    def create_widgets(self):
        # use ttk.Frame for modern look
        self.left_column = ttk.Frame(self, padding=(10, 10))
        self.left_column.grid(row=0, column=0, sticky="nsew")
        # Apply a fixed pixel width to the left column. Edit LEFT_PANEL_FIXED_WIDTH above to change it.
        try:
            self.left_column.configure(width=LEFT_PANEL_FIXED_WIDTH)
            # Make left column fixed and let right column expand
            self.grid_columnconfigure(0, weight=0)
            self.grid_columnconfigure(1, weight=1)
            # Prevent children from forcing the frame to resize
            self.left_column.grid_propagate(False)
        except Exception:
            pass
        self.setup_left_column(self.left_column)

        self.main_viewer_area = ttk.Frame(self, padding=(10, 10))
        self.main_viewer_area.grid(row=0, column=1, sticky="nsew")
        self.main_viewer_area.grid_columnconfigure(0, weight=1)
        self.main_viewer_area.grid_rowconfigure(1, weight=1)
        self.setup_viewer_area(self.main_viewer_area)

    def setup_left_column(self, parent_frame):
        # Profile label + combobox
        ttk.Label(parent_frame, text="Select a Profile:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.selected_profile = tk.StringVar(parent_frame)
        # When the selected profile changes, update enabled/disabled state of menu/buttons
        def _on_profile_var_changed(*args):
            try:
                self.on_profile_change()
            except Exception:
                pass
        try:
            self.selected_profile.trace_add('write', _on_profile_var_changed)
        except Exception:
            try:
                self.selected_profile.trace('w', _on_profile_var_changed)
            except Exception:
                pass
        # create combobox placeholder; refresh_profile_dropdown will populate values
        self.profile_combobox = ttk.Combobox(parent_frame, textvariable=self.selected_profile, state='readonly')
        self.profile_combobox.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        # when user selects an item from the combobox, ensure we update UI state
        self.profile_combobox.bind('<<ComboboxSelected>>', lambda e: self.on_profile_change())
        self.refresh_profile_dropdown()

        # Use ttkbootstrap 'secondary' bootstyle for gray buttons so color follows theme reliably
        # The Create/Edit Profiles buttons were removed from the left column per UI cleanup.
        # Their functionality remains available from the Edit menu (Create Profile... / Edit Profiles...).
        #
        # create_profile_btn = TBButton(parent_frame, text="Create Profiles", command=self.create_profile_window, bootstyle='secondary')
        # create_profile_btn.grid(row=2, column=0, sticky="ew", pady=5)
        #
        # edit_profile_btn = TBButton(parent_frame, text="Edit Profiles", command=self.edit_profile_window, bootstyle='secondary')
        # edit_profile_btn.grid(row=3, column=0, sticky="ew", pady=5)

        ttk.Separator(parent_frame, orient='horizontal').grid(row=4, column=0, sticky="ew", pady=10)

        # Use an explicit blue tk.Button so the Crop button appears blue across themes
        # Expose these as attributes so the menu can enable/disable them together with menu entries
        # On macOS, tk.Button doesn't respect bg/fg colors well, so use TBButton with primary bootstyle
        if sys.platform == 'darwin':
            self.crop_btn = TBButton(parent_frame, text="Crop", command=lambda: self.run_cropping(save_after=False), bootstyle='primary')
        else:
            self.crop_btn = tk.Button(parent_frame, text="Crop", command=lambda: self.run_cropping(save_after=False), bg='#0d6efd', fg='white', activebackground='#0b5bd0', activeforeground='white', bd=1, relief='raised')
        self.crop_btn.grid(row=5, column=0, sticky="ew", pady=5, ipady=8)
        self.move_btn = TBButton(parent_frame, text="Move", command=self.run_move_only, bootstyle='warning')
        self.move_btn.grid(row=6, column=0, sticky="ew", pady=5, ipady=8)
        # Use explicit blue tk.Button for Crop & Move to match the Crop button
        if sys.platform == 'darwin':
            self.crop_move_btn = TBButton(parent_frame, text="Crop & Move", command=lambda: self.run_cropping(save_after=True), bootstyle='primary')
        else:
            self.crop_move_btn = tk.Button(parent_frame, text="Crop & Move", command=lambda: self.run_cropping(save_after=True), bg='#0d6efd', fg='white', activebackground='#0b5bd0', activeforeground='white', bd=1, relief='raised')
        self.crop_move_btn.grid(row=7, column=0, sticky="ew", pady=5, ipady=12)

        # New: checkbox to allow deleting original images after cropping
        # Load saved preference (default False) and persist changes to the config file.
        try:
            saved_pref = load_config("delete_original_after_cropping")
            saved_val = str(saved_pref).lower() in ("1", "true", "yes", "on")
        except Exception:
            saved_val = False
        self.delete_original_var = tk.BooleanVar(value=saved_val)
        # New: delete-after-moving preference (mutually exclusive with delete_original_var)
        try:
            saved_move_pref = load_config("delete_original_after_moving")
            saved_move_val = str(saved_move_pref).lower() in ("1", "true", "yes", "on")
        except Exception:
            saved_move_val = False
        self.delete_move_var = tk.BooleanVar(value=saved_move_val)

        # Mutual-exclusion: when one is enabled, disable the other
        def _on_delete_crop_changed(*args):
            try:
                if self.delete_original_var.get():
                    try:
                        self.delete_move_var.set(False)
                    except Exception:
                        pass
                save_config("delete_original_after_cropping", "True" if self.delete_original_var.get() else "False")
            except Exception:
                pass

        def _on_delete_move_changed(*args):
            try:
                if self.delete_move_var.get():
                    try:
                        self.delete_original_var.set(False)
                    except Exception:
                        pass
                save_config("delete_original_after_moving", "True" if self.delete_move_var.get() else "False")
            except Exception:
                pass
        # Save any changes to the user's preference immediately to config.csv
        try:
            # trace_add is available in newer tkinter; fallback to trace for older versions
            self.delete_original_var.trace_add('write', _on_delete_crop_changed)
        except Exception:
            try:
                self.delete_original_var.trace('w', _on_delete_crop_changed)
            except Exception:
                pass

        try:
            self.delete_move_var.trace_add('write', _on_delete_move_changed)
        except Exception:
            try:
                self.delete_move_var.trace('w', _on_delete_move_changed)
            except Exception:
                pass
        # Left-column UI cleanup: remove the inline checkbox and Select/Open folder buttons.
        # The 'Delete Original Image after Cropping' option and the Select/Open folder actions
        # remain available in the Edit and File menus respectively.

        # Horizontal divider before folder controls (kept for layout)
        ttk.Separator(parent_frame, orient='horizontal').grid(row=9, column=0, columnspan=2, sticky="ew", pady=10)

    def setup_viewer_area(self, parent_frame):
        # Top area retained for potential future controls; folder buttons moved to left column
        top_frame = tk.Frame(parent_frame)
        # Expose top_frame so create_menubar can reliably place the visible toolbar there
        self.top_frame = top_frame
        top_frame.grid(row=0, column=0, sticky="new")
        top_frame.grid_columnconfigure(0, weight=1)
        # reserve a second column for the support link (right side)
        top_frame.grid_columnconfigure(1, weight=0)

        # Status label (left side) — shows short runtime messages (initialized here to avoid missing attribute errors)
        # try:
        #     self.status_label = ttk.Label(top_frame, text="Ready", foreground="green", font=('Helvetica', 9, 'italic'))
        #     self.status_label.grid(row=0, column=0, sticky='w', padx=(0,6))
        # except Exception:
        #     # Fallback: create a simple tk.Label if ttk fails for some reason
        #     try:
        #         self.status_label = tk.Label(top_frame, text="Ready", fg="green")
        #         self.status_label.grid(row=0, column=0, sticky='w', padx=(0,6))
        #     except Exception:
        #         self.status_label = _NullLabel()
        # Status messages removed per user request. Keep a no-op placeholder so
        # existing code that calls self.status_label.config(...) continues to work.
        self.status_label = _NullLabel()

        # Support link (top-right corner) that opens the external site
        # support_url = "https://www.abelxl.com/"
        # try:
        #     def open_support(event=None):
        #         webbrowser.open(support_url)
        #
        #     support_lbl = ttk.Label(top_frame, text="Support Us", foreground="#0d6efd", cursor="hand2")
        #     # Reduced vertical padding so the link sits closer to the top/bottom
        #     support_lbl.grid(row=0, column=1, sticky="ne", padx=(0, 6), pady=(0, 0))
        #     support_lbl.bind("<Button-1>", open_support)
        # except Exception:
        #     pass

        # Stronger visual treatment: dark, thick raised outer border with a
        # bright inner panel so the viewer area stands out clearly regardless
        # of the current ttk theme. This avoids relying on theme colors.
        canvas_bg = '#fbfcfe'         # very light inner panel so thumbnails pop
        canvas_outline = '#22262b'    # very dark outline for strong contrast
        outer_border = 6              # make outline thick so it's visible
        inner_padding = 8             # more breathing room inside the border

        # Outer frame provides a thick raised border that reads as a panel edge
        outer_frame = tk.Frame(parent_frame, bg=canvas_outline, bd=outer_border, relief='raised')
        outer_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        outer_frame.grid_rowconfigure(0, weight=1)
        outer_frame.grid_columnconfigure(0, weight=1)

        # inner_frame creates an inset effect (padding) so the outline reads as
        # a visible border even on light themes
        inner_frame = tk.Frame(outer_frame, bg=canvas_bg, bd=0, relief='flat')
        inner_frame.grid(row=0, column=0, sticky="nsew", padx=inner_padding, pady=inner_padding)
        inner_frame.grid_rowconfigure(0, weight=1)
        inner_frame.grid_columnconfigure(0, weight=1)

        # Canvas uses the inner_frame background so it reads as a panel; don't
        # add extra highlights — outer_frame's bd gives the needed contrast.
        self.canvas = tk.Canvas(inner_frame, bg=canvas_bg, bd=0, highlightthickness=0, relief='flat')
        v_scroll = tk.Scrollbar(inner_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=v_scroll.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")

        # Keep the thumb_frame white so thumbnail 'cards' still pop against the panel
        self.thumb_frame = tk.Frame(self.canvas, bg="white")
        self.canvas.create_window((0, 0), window=self.thumb_frame, anchor="nw")
        self.thumb_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Enable mouse wheel scrolling on the canvas
        def on_mousewheel(event):
            scroll_amount = int(-1 * (event.delta / 120))

            # Get current scroll position (0.0 to 1.0)
            current_pos = self.canvas.yview()

            # If scrolling up (negative scroll_amount means scrolling down the page, positive means up)
            # and already at the top, don't scroll
            if scroll_amount < 0 and current_pos[0] <= 0.0:
                return "break"

            self.canvas.yview_scroll(scroll_amount, "units")
            return "break"

        # Store the mousewheel handler so it can be bound to dynamically created widgets
        self._on_mousewheel = on_mousewheel

        self.canvas.bind("<MouseWheel>", on_mousewheel)
        self.thumb_frame.bind("<MouseWheel>", on_mousewheel)

        # Enable drag-and-drop on the canvas
        if DND_FILES:
            try:
                self.canvas.drop_target_register(DND_FILES)
                self.canvas.dnd_bind('<<Drop>>', self._on_drop)
                self.canvas.dnd_bind('<<DragEnter>>', self._on_drag_enter)
                self.canvas.dnd_bind('<<DragLeave>>', self._on_drag_leave)
            except Exception:
                pass  # Silently fail if drag-and-drop isn't available

        # Bind arrow keys for thumbnail navigation
        self.canvas.bind("<Left>", lambda e: self.navigate_thumbnail('left'))
        self.canvas.bind("<Right>", lambda e: self.navigate_thumbnail('right'))
        self.canvas.bind("<Up>", lambda e: self.navigate_thumbnail('up'))
        self.canvas.bind("<Down>", lambda e: self.navigate_thumbnail('down'))
        # Create the application menubar after the UI is built so variables exist
        try:
            self.create_menubar()
        except Exception:
            # If menubar creation fails, don't crash the UI; just log traceback to console
            import traceback
            traceback.print_exc()

    # ========================= DYNAMIC POLLING & SORTING =========================
    def on_window_resize(self, event):
        if abs(self.canvas.winfo_width() - self._last_width) > 20:
            self.refresh_thumbnails(is_polling=True)

    def start_auto_refresh(self):
        self.refresh_thumbnails(is_polling=True)
        self.after(2000, self.start_auto_refresh)

    # ========================= DRAG-AND-DROP HANDLERS =========================
    def _on_drag_enter(self, event):
        """Visual feedback when dragging over the canvas."""
        try:
            self.canvas.config(bg='#e6f3ff')  # Light blue background
        except Exception:
            pass

    def _on_drag_leave(self, event):
        """Reset visual feedback when drag leaves the canvas."""
        try:
            self.canvas.config(bg='#fbfcfe')  # Original background
        except Exception:
            pass

    def _on_drop(self, event):
        """Handle dropped files - copy them to the source folder."""
        try:
            # Reset canvas background
            self.canvas.config(bg='#fbfcfe')

            # Get source folder
            source_folder = self.file_paths.get("source_folder", "")
            if not source_folder or not os.path.isdir(source_folder):
                messagebox.showwarning(
                    "No Source Folder",
                    "Please select a Source Folder before dropping images.",
                    parent=self
                )
                return

            # Parse dropped files (tkinterdnd2 returns a string with curly braces)
            files = self._parse_drop_files(event.data)

            # Filter for image files only
            valid_extensions = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp", ".heic", ".heif"}

            # Separate supported and unsupported files
            image_files = []
            unsupported_files = []
            for f in files:
                if os.path.splitext(f)[1].lower() in valid_extensions:
                    image_files.append(f)
                else:
                    unsupported_files.append(os.path.basename(f))

            # If there are unsupported files, show detailed message
            if unsupported_files:
                total_files = len(files)
                supported_count = len(image_files)
                unsupported_count = len(unsupported_files)

                msg = f"Dropped {total_files} file(s):\n\n"
                msg += f"✓ Supported: {supported_count}\n"
                msg += f"✗ Unsupported: {unsupported_count}\n\n"

                if unsupported_count > 0:
                    msg += "Unsupported files:\n"
                    # Limit display to first 10 unsupported files to avoid huge message box
                    display_files = unsupported_files[:10]
                    for filename in display_files:
                        msg += f"  • {filename}\n"
                    if unsupported_count > 10:
                        msg += f"  ... and {unsupported_count - 10} more\n"

                    msg += "\nSupported formats: JPG, JPEG, PNG, BMP, GIF, TIF, TIFF, WEBP, HEIC, HEIF"

                if supported_count == 0:
                    messagebox.showwarning("No Supported Images", msg, parent=self)
                    return
                else:
                    # Show info and continue processing supported files
                    messagebox.showinfo("Mixed File Types", msg, parent=self)

            if not image_files:
                return

            # Copy files to source folder
            copied_count = 0
            skipped_count = 0

            for file_path in image_files:
                try:
                    filename = os.path.basename(file_path)
                    destination = os.path.join(source_folder, filename)

                    # Check if file already exists
                    if os.path.exists(destination):
                        skipped_count += 1
                        continue

                    # Copy the file
                    shutil.copy2(file_path, destination)
                    copied_count += 1
                except Exception:
                    skipped_count += 1

            # Show result message (only if no unsupported files warning was shown, or as a second message)
            if not unsupported_files:
                msg = f"Copied {copied_count} image(s) to Source Folder."
                if skipped_count:
                    msg += f"\n{skipped_count} file(s) skipped (already exist or failed)."
                messagebox.showinfo("Files Copied", msg, parent=self)
            elif copied_count > 0 or skipped_count > 0:
                # Show a brief success message after the unsupported files warning
                msg = f"Result: {copied_count} copied"
                if skipped_count:
                    msg += f", {skipped_count} skipped"
                messagebox.showinfo("Copy Complete", msg, parent=self)

            # Refresh thumbnails to show new images
            self.refresh_thumbnails()

        except Exception as e:
            messagebox.showerror("Drop Error", f"Failed to process dropped files: {e}", parent=self)

    def _parse_drop_files(self, data):
        """Parse the file paths from tkinterdnd2 drop event data.
        The data format can be '{file1} {file2}' or just 'file1 file2'.
        """
        files = []
        try:
            # Remove curly braces and split by spaces (respecting quoted paths)
            # Match either {path} or "path" or unquoted paths
            pattern = r'\{([^}]+)\}|"([^"]+)"|(\S+)'
            matches = re.findall(pattern, data)
            for match in matches:
                # match is a tuple (group1, group2, group3) - use whichever matched
                path = match[0] or match[1] or match[2]
                if path and os.path.isfile(path):
                    files.append(path)
        except Exception:
            # Fallback: simple split
            files = [f.strip('{}\"') for f in data.split() if os.path.isfile(f.strip('{}\"'))]

        return files

    def refresh_thumbnails(self, is_polling=False):
        """Refreshes images, sorted by time, wrapping based on window width."""
        source = self.file_paths["source_folder"]
        if not source or not os.path.isdir(source):
            if not is_polling:
                for widget in self.thumb_frame.winfo_children(): widget.destroy()
                no_source_label = tk.Label(self.thumb_frame, text="No source folder selected", bg="white")
                no_source_label.pack(pady=50)
                no_source_label.bind("<MouseWheel>", self._on_mousewheel)
            return

        extensions = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp", ".heic", ".heif"}

        try:
            # --- UPDATED SORTING LOGIC: SORT BY MODIFIED TIME ---
            # This ensures preview order matches the internal cropping sequence
            all_files = [p for p in Path(source).iterdir() if p.suffix.lower() in extensions and p.is_file()]
            all_files.sort(key=lambda p: p.stat().st_mtime)
            current_files = [f.name for f in all_files]
        except:
            return

        current_width = self.canvas.winfo_width()

        if current_files == self.last_seen_files and current_width == self._last_width:
            return

        self.last_seen_files = current_files
        self._last_width = current_width

        for widget in self.thumb_frame.winfo_children():
            widget.destroy()

        # Clear thumbnail tracking
        self.thumbnail_widgets = []
        self.selected_thumbnail_index = None

        if not current_files:
            # Get the canvas dimensions to properly center the message
            canvas_width = self.canvas.winfo_width() if self.canvas.winfo_width() > 1 else 800
            canvas_height = self.canvas.winfo_height() if self.canvas.winfo_height() > 1 else 600

            # Create a frame to hold the multi-line message with explicit dimensions
            message_frame = tk.Frame(self.thumb_frame, bg="white", width=canvas_width, height=canvas_height)
            message_frame.pack(fill="both", expand=True)
            message_frame.pack_propagate(False)  # Prevent the frame from shrinking
            message_frame.bind("<MouseWheel>", self._on_mousewheel)

            # Create an inner frame for vertical positioning
            inner_frame = tk.Frame(message_frame, bg="white")
            inner_frame.place(relx=0.5, rely=0.3, anchor="center")
            inner_frame.bind("<MouseWheel>", self._on_mousewheel)

            # First line: Drag and drop instruction
            line1_label = tk.Label(inner_frame, text="Drag and Drop images here", fg="gray", bg="white", font=("Arial", 11, "bold"))
            line1_label.pack()
            line1_label.bind("<MouseWheel>", self._on_mousewheel)

            # Second line: Alternative method
            line2_label = tk.Label(inner_frame, text="Or go to File → Open Source Folder and paste images there.", fg="gray", bg="white", font=("Arial", 9))
            line2_label.pack(pady=(5, 0))
            line2_label.bind("<MouseWheel>", self._on_mousewheel)
            return

        row, col = 0, 0
        thumb_size_px = 150
        padding_px = 30
        slot_width = thumb_size_px + padding_px
        num_cols = max(1, current_width // slot_width) if current_width > 1 else 2

        for idx, fname in enumerate(current_files):
            try:
                img_path = os.path.join(source, fname)
                # Open the image before creating a thumbnail
                with Image.open(img_path) as img:
                    img.thumbnail((thumb_size_px, thumb_size_px), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img.copy())

                label = tk.Label(self.thumb_frame, image=photo, bg="white", bd=1, relief="solid",
                                highlightthickness=2, highlightbackground="#E0E0E0", highlightcolor="#E0E0E0")
                label.image = photo
                label.grid(row=row, column=col, padx=15, pady=(15, 0))
                label.bind("<MouseWheel>", self._on_mousewheel)

                name_label = tk.Label(self.thumb_frame, text=fname[:20], bg="white", fg="#333", font=("Arial", 8))
                name_label.grid(row=row + 1, column=col, sticky="we", padx=15, pady=(0, 15))
                name_label.bind("<MouseWheel>", self._on_mousewheel)

                # Store thumbnail references
                self.thumbnail_widgets.append((label, name_label, fname))

                # Bind click events for selection
                label.bind("<Button-1>", lambda e, i=idx: self.select_thumbnail(i))
                name_label.bind("<Button-1>", lambda e, i=idx: self.select_thumbnail(i))

                col += 1
                if col >= num_cols:
                    col = 0
                    row += 2
            except Exception:
                continue

    def select_thumbnail(self, index):
        """Select a thumbnail by index and update visual feedback."""
        if not self.thumbnail_widgets or index < 0 or index >= len(self.thumbnail_widgets):
            return

        # Deselect previous thumbnail - only change colors, keep same border/highlight thickness
        if self.selected_thumbnail_index is not None and self.selected_thumbnail_index < len(self.thumbnail_widgets):
            prev_label, prev_name_label, _ = self.thumbnail_widgets[self.selected_thumbnail_index]
            try:
                prev_label.config(highlightbackground="#E0E0E0", highlightcolor="#E0E0E0")
                prev_name_label.config(bg="white", fg="#333")
            except Exception:
                pass

        # Select new thumbnail - only change colors, keep same border/highlight thickness
        self.selected_thumbnail_index = index
        label, name_label, filename = self.thumbnail_widgets[index]

        # Highlight the selected thumbnail with blue color
        label.config(highlightbackground="#007ACC", highlightcolor="#007ACC")
        name_label.config(bg="#007ACC", fg="white")

        # Set focus to canvas so arrow keys work
        self.canvas.focus_set()

        # Scroll to make selected thumbnail visible
        self._scroll_to_thumbnail(index)

    def _scroll_to_thumbnail(self, index):
        """Scroll the canvas to ensure the selected thumbnail is visible."""
        if not self.thumbnail_widgets or index < 0 or index >= len(self.thumbnail_widgets):
            return

        try:
            label, _, _ = self.thumbnail_widgets[index]
            # Get the label's position in the canvas
            self.thumb_frame.update_idletasks()

            # Get the bbox of the label relative to thumb_frame
            y = label.winfo_y()
            height = label.winfo_height() + 50  # Add some padding

            # Get canvas dimensions
            canvas_height = self.canvas.winfo_height()

            # Get current scroll region
            scroll_region = self.canvas.cget("scrollregion").split()
            if len(scroll_region) == 4:
                total_height = float(scroll_region[3])

                # Calculate the fraction to scroll to
                if total_height > 0:
                    # Scroll so the thumbnail is centered in view if possible
                    target_y = max(0, y - canvas_height // 2)
                    fraction = target_y / total_height
                    self.canvas.yview_moveto(fraction)
        except Exception:
            pass

    def navigate_thumbnail(self, direction):
        """Navigate thumbnails using arrow keys. Direction: 'up', 'down', 'left', 'right'."""
        if not self.thumbnail_widgets:
            return

        # If nothing selected, select the first thumbnail
        if self.selected_thumbnail_index is None:
            self.select_thumbnail(0)
            return

        current_index = self.selected_thumbnail_index
        new_index = current_index

        # Calculate grid dimensions
        current_width = self.canvas.winfo_width()
        thumb_size_px = 150
        padding_px = 30
        slot_width = thumb_size_px + padding_px
        num_cols = max(1, current_width // slot_width) if current_width > 1 else 2

        if direction == 'left':
            new_index = max(0, current_index - 1)
        elif direction == 'right':
            new_index = min(len(self.thumbnail_widgets) - 1, current_index + 1)
        elif direction == 'up':
            new_index = max(0, current_index - num_cols)
        elif direction == 'down':
            new_index = min(len(self.thumbnail_widgets) - 1, current_index + num_cols)

        if new_index != current_index:
            self.select_thumbnail(new_index)

    # ========================= CORE LOGIC =========================
    def run_move_only(self):
        # Require a profile to be selected before moving
        try:
            profile_name = self.selected_profile.get()
        except Exception:
            profile_name = None
        if not profile_name or profile_name == "— No Profile Selected —":
            try:
                messagebox.showwarning("Profile Required", "Please select a profile before proceeding.")
            except Exception:
                pass
            return

        source_folder = self.file_paths["source_folder"]
        destination_folder = self.file_paths["destination_folder"]
        if not source_folder or not os.path.isdir(source_folder): return
        if not destination_folder or not os.path.isdir(destination_folder): return

        # If delete-after-moving is enabled, maybe prompt for confirmation first.
        try:
            do_delete_after_move = bool(getattr(self, 'delete_move_var', None) and self.delete_move_var.get())
        except Exception:
            do_delete_after_move = False
        if do_delete_after_move:
            try:
                pref = load_config('confirm_delete_after_moving')
                show_dialog = True
                if isinstance(pref, str) and pref != '':
                    show_dialog = str(pref).lower() in ('1', 'true', 'yes', 'on')
            except Exception:
                show_dialog = True

            if show_dialog:
                confirmed, do_not_show_again = self._show_confirm_delete_after_move_dialog()
                if not confirmed:
                    try:
                        self.status_label.config(text="Move cancelled.", foreground="red")
                    except Exception:
                        pass
                    return
                if do_not_show_again:
                    try:
                        save_config('confirm_delete_after_moving', 'False')
                    except Exception:
                        pass

        timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        new_archive_folder = os.path.join(destination_folder, timestamp)

        try:
            os.makedirs(new_archive_folder, exist_ok=True)

            # Determine which files are "cropped outputs" for this profile.
            # This mirrors the naming used by run_cropping: <base_name>_suffix.ext
            base_name = re.sub(r'[^a-zA-Z0-9_ -]', '', profile_name).replace(' ', '_').strip('_')
            output_prefix = f"{base_name}_"
            image_exts = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp", ".heic", ".heif"}

            all_files = [f for f in os.listdir(source_folder) if os.path.isfile(os.path.join(source_folder, f))]
            # Files that look like cropped outputs (to be moved)
            cropped_files = [f for f in all_files if f.startswith(output_prefix) and os.path.splitext(f)[1].lower() in image_exts]
            # Originals are image files that are NOT cropped outputs
            original_files = [f for f in all_files if os.path.splitext(f)[1].lower() in image_exts and f not in cropped_files]

            moved_count = 0
            deleted_count = 0

            do_delete_after_move = bool(getattr(self, 'delete_move_var', None) and self.delete_move_var.get())

            # If delete-after-moving is enabled we will copy outputs then delete originals.
            # Otherwise we simply move the outputs (previous behavior changed to only move outputs).
            # Use the already-imported send2trash function from the top of the file
            send2trash_fn = send2trash if do_delete_after_move else None

            created_out_paths = set()

            # Move or copy files. If delete-after-moving is enabled we copy
            # outputs then delete originals; otherwise move everything (both
            # cropped outputs and the original image files) so the Source
            # folder is emptied by the Move action.
            if do_delete_after_move:
                # Copy outputs to archive; originals will be deleted below.
                for fname in cropped_files:
                    src = os.path.join(source_folder, fname)
                    dst = os.path.join(new_archive_folder, fname)
                    try:
                        shutil.copy2(src, dst)
                        moved_count += 1
                        created_out_paths.add(os.path.abspath(dst))
                    except Exception:
                        # skip failures per-file
                        continue
            else:
                # Move both cropped outputs and any original images so nothing is left behind.
                all_to_move = []
                # preserve order: move cropped outputs first then originals
                all_to_move.extend(cropped_files)
                # add originals that are not already in cropped_files
                for f in original_files:
                    if f not in cropped_files:
                        all_to_move.append(f)

                for fname in all_to_move:
                    src = os.path.join(source_folder, fname)
                    dst = os.path.join(new_archive_folder, fname)
                    try:
                        shutil.move(src, dst)
                        moved_count += 1
                        created_out_paths.add(os.path.abspath(dst))
                    except Exception:
                        # skip failures per-file
                        continue

            # If delete-after-moving is enabled: delete source copies of the cropped outputs first,
            # then delete original image files (not the cropped outputs)
            if do_delete_after_move:
                # First delete the source copies of the cropped outputs (we copied them above)
                for fname in cropped_files:
                    try:
                        src_path = os.path.abspath(os.path.join(source_folder, fname))
                        if src_path in created_out_paths:
                            # very unlikely, but don't delete if equal to created output
                            continue
                        if os.path.isfile(src_path):
                            try:
                                if send2trash_fn:
                                    try:
                                        send2trash_fn(src_path)
                                        deleted_count += 1
                                    except Exception:
                                        try:
                                            os.remove(src_path)
                                            deleted_count += 1
                                        except Exception:
                                            pass
                                else:
                                    try:
                                        os.remove(src_path)
                                        deleted_count += 1
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                    except Exception:
                        continue

                # Then delete the original image files (which are distinct from cropped outputs)
                if original_files:
                    for fname in original_files:
                        try:
                            orig_path = os.path.abspath(os.path.join(source_folder, fname))
                            # Skip if somehow orig_path equals a created output
                            if orig_path in created_out_paths:
                                continue
                            if os.path.isfile(orig_path):
                                try:
                                    if send2trash_fn:
                                        try:
                                            send2trash_fn(orig_path)
                                            deleted_count += 1
                                        except Exception:
                                            try:
                                                os.remove(orig_path)
                                                deleted_count += 1
                                            except Exception:
                                                pass
                                    else:
                                        try:
                                            os.remove(orig_path)
                                            deleted_count += 1
                                        except Exception:
                                            pass
                                except Exception:
                                    pass
                        except Exception:
                            continue

            # Update status
            try:
                if deleted_count:
                    self.status_label.config(text=f"Moved {moved_count} files to {timestamp}. Deleted {deleted_count} original file(s).", foreground="green")
                else:
                    self.status_label.config(text=f"Moved {moved_count} files to {timestamp}", foreground=("green" if moved_count else "red"))
            except Exception:
                pass

        except Exception as e:
            self.status_label.config(text=f"Move failed: {e}", foreground="red")

    def run_cropping(self, save_after=False):
        # Before starting, if the user has enabled delete-originals, maybe prompt for confirmation.
        try:
            delete_enabled = bool(getattr(self, 'delete_original_var', None) and self.delete_original_var.get())
        except Exception:
            delete_enabled = False
        if delete_enabled:
            # Load suppression preference (default True meaning show dialog). If confirm_delete_originals is True, show dialog.
            try:
                suppress_pref = load_config('confirm_delete_after_cropping')
                # interpret values: True => show dialog, False => do NOT show
                show_dialog = True
                if isinstance(suppress_pref, str) and suppress_pref != '':
                    show_dialog = str(suppress_pref).lower() in ('1', 'true', 'yes', 'on')
                # If config value is missing/empty, default to True (show dialog)
            except Exception:
                show_dialog = True

            # The config column 'confirm_delete_originals' uses True to indicate we should show the confirmation.
            if show_dialog:
                confirmed, do_not_show_again = self._show_confirm_delete_dialog()
                if not confirmed:
                    # user declined -> abort cropping
                    try:
                        self.status_label.config(text="Cropping cancelled.", foreground="red")
                    except Exception:
                        pass
                    return
                # If the user checked 'do not show again', save the inverse flag: False means do not show
                if do_not_show_again:
                    try:
                        save_config('confirm_delete_after_cropping', 'False')
                    except Exception:
                        pass

        profile_name = self.selected_profile.get()
        source_folder = self.file_paths["source_folder"]

        # If no profile is selected, alert the user and return early
        if not profile_name or profile_name == "— No Profile Selected —":
            try:
                messagebox.showwarning("Profile Required", "Please select a profile before proceeding.")
            except Exception:
                pass
            return

        if not source_folder:
            # Keep existing behavior for missing source folder
            self.status_label.config(text="No source folder selected.", foreground="red")
            return

        rules_map = load_profile_rules(profile_name)
        if not rules_map: return

        extensions = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp", ".heic", ".heif"}
        image_paths = sorted(
            [p for p in Path(source_folder).iterdir() if p.suffix.lower() in extensions and p.is_file()],
            key=lambda p: p.stat().st_mtime)

        # Snapshot the initial list of originals (absolute paths) so we have
        # a stable index to refer to during deletion. This prevents newly
        # created cropped files in the same folder from shifting positions.
        initial_originals = [str(p.resolve()) for p in image_paths]

        if not image_paths:
            self.status_label.config(text="No images found in source folder.", foreground="red")
            return

        base_name = re.sub(r'[^a-zA-Z0-9_ -]', '', profile_name).replace(' ', '_').strip('_')

        processed_count = 0
        skipped_count = 0
        # Track originals that were actually processed and outputs we created.
        # We'll delete originals after processing if the user enabled the checkbox.
        processed_originals: set[str] = set()
        created_out_paths: set[str] = set()
        # Precompute sorted rule positions for fallback lookups
        sorted_rule_positions = sorted([p for p in rules_map.keys()])

        # Informative diagnostics for debugging
        try:
            self.status_label.config(text=f"Loaded profile '{profile_name}' with {len(sorted_rule_positions)} rule(s). Found {len(image_paths)} image(s).", foreground="black")
        except Exception:
            pass

        # Build the complete list of (rule_index, position, rule) applications
        # by iterating rules in file order. Rules that target an explicit position
        # are applied to that position. Rules with apply_to_all_remaining=True
        # are applied to any positions that don't already have an explicit rule
        # (and that haven't been filled yet by an earlier apply_all).
        applications: list[tuple[int, int, dict]] = []  # (rule_index, position, rule)

        # Flatten all rules in file order based on their stored _rule_index
        all_rules: list[dict] = []
        for p in sorted(rules_map.keys()):
            for r in rules_map.get(p, []):
                all_rules.append(r)
        all_rules.sort(key=lambda rr: int(rr.get('_rule_index', 999999)))

        image_count = len(image_paths)
        # positions that have explicit rules in the profile (do not override these)
        explicit_positions = set(rules_map.keys())
        # positions already assigned by an earlier apply_all
        filled_by_apply_all: set[int] = set()

        for rule in all_rules:
            rule_index = int(rule.get('_rule_index', 999999))

            # If this rule targets an explicit position, apply it there
            try:
                rule_pos = int(rule.get('position', rule.get('position_number', 0)))
            except Exception:
                rule_pos = 0

            if 1 <= rule_pos <= image_count:
                applications.append((rule_index, rule_pos, rule))

            # If this rule requests apply-to-all, apply it to any positions that
            # don't already have an explicit rule and haven't already been filled
            # (and that haven't been filled yet by an earlier apply_all).
            if rule.get('apply_to_all_remaining'):
                for p in range(1, image_count + 1):
                    if p in explicit_positions:
                        # explicit rule exists for this position; skip
                        continue
                    if p in filled_by_apply_all:
                        # already filled by an earlier apply_all
                        continue
                    # Assign this apply_all rule to position p
                    applications.append((rule_index, p, rule))
                    filled_by_apply_all.add(p)

        # If no applications were created, nothing applies
        if not applications:
            self.status_label.config(text="No rules applicable to images.", foreground="red")
            return

        # Sort applications by rule file-order so outputs follow the rule listing order.
        applications.sort(key=lambda t: t[0])

        # Now perform cropping in the order of applications; name files sequentially a,b,c...
        def _suffix_for_index(i: int) -> str:
            # 1 -> a, 2 -> b, ... 26 -> z, 27 -> 27, 28 -> 28, etc.
            if 1 <= i <= 26:
                return chr(96 + i)
            return str(i)

        for app_idx, (rule_index, position, rule) in enumerate(applications, start=1):
            # Map position to image path (guard bounds)
            if position <= 0 or position > len(image_paths):
                # invalid position — skip
                continue
            img_path = image_paths[position - 1]

            try:
                with Image.open(img_path) as source_img:
                    # Determine whether this application came from an apply_all rule
                    # that originates from another position. If so, and if that
                    # originating rule's crop equals the originating image's full
                    # dimensions and the aspect_ratio is 'none', then the user's
                    # intent is likely to keep the target image at its original
                    # size — so do not force the originating image dimensions.
                    use_original_size_for_target = False
                    try:
                        rule_origin_pos = int(rule.get('position', rule.get('position_number', 0)))
                    except Exception:
                        rule_origin_pos = 0

                    if 1 <= rule_origin_pos <= len(image_paths) and rule_origin_pos != position and rule.get('aspect_ratio', 'none') == 'none':
                        try:
                            # get the originating image path and inspect its crop
                            origin_path = image_paths[rule_origin_pos - 1]
                            with Image.open(origin_path) as origin_img:
                                c_rule = rule.get('crop', {})
                                try:
                                    rx1 = int(c_rule.get('x1', 0))
                                    ry1 = int(c_rule.get('y1', 0))
                                    rx2 = int(c_rule.get('x2', origin_img.width))
                                    ry2 = int(c_rule.get('y2', origin_img.height))
                                except Exception:
                                    rx1, ry1, rx2, ry2 = 0, 0, origin_img.width, origin_img.height

                                # if the rule's crop exactly matches the origin's full size
                                if rx1 == 0 and ry1 == 0 and rx2 == origin_img.width and ry2 == origin_img.height:
                                    use_original_size_for_target = True
                        except Exception:
                            # if anything fails, fall back to normal cropping
                            use_original_size_for_target = False

                    if use_original_size_for_target:
                        # create an un-cropped copy of the target image
                        cropped_img = source_img.copy()
                        # mark coordinates as full-target so later logic treats it as full image
                        x1, y1, x2, y2 = 0, 0, source_img.width, source_img.height
                    else:
                        c = rule.get('crop', {})
                        try:
                            x1 = int(c.get('x1', 0))
                            y1 = int(c.get('y1', 0))
                            x2 = int(c.get('x2', source_img.width))
                            y2 = int(c.get('y2', source_img.height))
                        except Exception:
                            continue

                        x1 = max(0, min(x1, source_img.width - 1))
                        y1 = max(0, min(y1, source_img.height - 1))
                        x2 = max(0, min(x2, source_img.width))
                        y2 = max(0, min(y2, source_img.height))
                        if x2 <= x1 or y2 <= y1:
                            continue

                        cropped_img = source_img.crop((x1, y1, x2, y2))

                    suf = _suffix_for_index(app_idx)
                    # Convert HEIC/HEIF images to JPEG
                    original_ext = img_path.suffix.lower()
                    if original_ext in ('.heic', '.heif'):
                        output_ext = '.jpg'
                    else:
                        output_ext = img_path.suffix
                    output_file_name = f"{base_name}_{suf}{output_ext}"
                    out_path = os.path.join(source_folder, output_file_name)

                    compression_percent = int(rule.get('compression', rule.get('compression_percent', 0)))

                    # If no compression requested and the crop is the full image, prefer
                    # to re-save via Pillow at high quality to strip EXIF, otherwise fall back
                    # to copying bytes. For cropped images we must save the cropped image.
                    is_full_image = (x1 == 0 and y1 == 0 and x2 == source_img.width and y2 == source_img.height)

                    try:
                        if compression_percent <= 0 and is_full_image:
                            # Try re-saving at high quality (95) which removes metadata.
                            try:
                                _save_image_preset(source_img, out_path, quality=95)
                                try:
                                    now = time.time()
                                    os.utime(out_path, (now, now))
                                except Exception:
                                    pass
                            except Exception:
                                # Fallback: copy original bytes and update mtime.
                                try:
                                    src_path = os.path.abspath(img_path)
                                    dst_path = os.path.abspath(out_path)
                                    if src_path != dst_path:
                                        shutil.copy2(src_path, dst_path)
                                        try:
                                            now = time.time()
                                            os.utime(dst_path, (now, now))
                                        except Exception:
                                            pass
                                except Exception:
                                    pass
                        else:
                            # Save cropped image; map compression percent to Pillow quality.
                            if compression_percent <= 0:
                                _save_image_preset(cropped_img, out_path, quality=95)
                            else:
                                pillow_quality = max(1, min(95, int(round(95 * (100 - compression_percent) / 100))))
                                _save_image_preset(cropped_img, out_path, quality=pillow_quality)
                    except Exception:
                        # Best-effort fallback: try saving with defaults
                        try:
                            _save_image_preset(cropped_img, out_path, quality=95)
                        except Exception:
                            pass

                    # record that this original was processed and note the new output path
                    try:
                        processed_originals.add(str(img_path))
                        created_out_paths.add(os.path.abspath(out_path))
                    except Exception:
                        pass

                    processed_count += 1
            except Exception:
                continue

        # If the user enabled deletion of original images, remove the originals that were processed.
        deleted_count = 0
        if getattr(self, 'delete_original_var', None) and self.delete_original_var.get() and processed_originals:
            # Prefer moving files to the OS Trash using send2trash for safety.
            # Use the already-imported send2trash function from the top of the file
            send2trash_fn = send2trash

            # Iterate the initial snapshot so deletions target the original
            # files that were present at the start of the crop operation.
            for orig in list(initial_originals):
                try:
                    # We now delete ALL originals that existed at the start,
                    # regardless of whether a rule processed them. This matches
                    # the requested behavior: remove every original image.
                    orig_abspath = os.path.abspath(orig)
                    # Never delete an output we just created by this run
                    if orig_abspath in created_out_paths:
                        continue
                    if os.path.isfile(orig_abspath):
                        try:
                            if send2trash_fn:
                                try:
                                    send2trash_fn(orig_abspath)
                                    deleted_count += 1
                                except Exception:
                                    try:
                                        os.remove(orig_abspath)
                                        deleted_count += 1
                                    except Exception:
                                        pass
                            else:
                                try:
                                    os.remove(orig_abspath)
                                    deleted_count += 1
                                except Exception:
                                    pass
                        except Exception:
                            pass
                except Exception:
                    continue

        # Final status message
        if save_after:
            self.run_move_only()
        else:
            msg = f"Finished cropping {processed_count} image(s)."
            if skipped_count:
                msg += f" Skipped {skipped_count} image(s) with no matching rule."
            if deleted_count:
                msg += f" Deleted {deleted_count} original file(s)."
            self.status_label.config(text=msg, foreground=("green" if processed_count else "red"))

    # ========================= HELPERS =========================
    def refresh_profile_dropdown(self):
        self.available_profiles = load_profiles()
        options = ["— No Profile Selected —"] + self.available_profiles
        try:
            # set combobox values and choose default
            self.profile_combobox['values'] = options
            if not self.selected_profile.get():
                self.selected_profile.set(options[0])
            else:
                # ensure current value is valid
                if self.selected_profile.get() not in options:
                    self.selected_profile.set(options[0])
            # set current selection index
            try:
                idx = options.index(self.selected_profile.get())
                self.profile_combobox.current(idx)
            except Exception:
                self.profile_combobox.current(0)
        except Exception:
            pass
        # If the File->Edit Profiles submenu exists, rebuild it to reflect current profiles
        try:
            if hasattr(self, 'edit_profiles_menu'):
                try:
                    # Remove any existing entries
                    self.edit_profiles_menu.delete(0, 'end')
                except Exception:
                    pass
                if self.available_profiles:
                    for _p in self.available_profiles:
                        # bind each profile name to open in the external editor
                        try:
                            self.edit_profiles_menu.add_command(label=_p, command=(lambda name=_p: self.open_profile_in_editor(name)))
                        except Exception:
                            pass
                else:
                    try:
                        self.edit_profiles_menu.add_command(label='(no profiles)', state='disabled')
                    except Exception:
                        pass
        except Exception:
            # best-effort: if rebuilding submenu fails, ignore and continue
            pass
        # update menu/button enabled state after profiles refresh
        try:
            self.update_menu_state()
        except Exception:
            pass

    def select_source_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.file_paths["source_folder"] = folder
            save_config("source_folder", folder)
            self.last_seen_files = []
            self.refresh_thumbnails()
            try:
                if os.path.isdir(folder):
                    self.view_src_btn.config(state='normal')
            except Exception:
                pass
            try:
                self.update_menu_state()
            except Exception:
                pass

    def select_destination_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.file_paths["destination_folder"] = folder
            save_config("destination_folder", folder)
            try:
                if os.path.isdir(folder):
                    self.view_dst_btn.config(state='normal')
            except Exception:
                pass
            try:
                self.update_menu_state()
            except Exception:
                pass

    def view_source_folder(self):
        """Open the source folder in the OS file manager."""
        path = self.file_paths.get('source_folder', '')
        if not path or not os.path.isdir(path):
            try:
                messagebox.showwarning("Folder not available", "No valid source folder is selected.")
            except Exception:
                pass
            return
        try:
            self._open_folder(path)
        except Exception:
            try:
                messagebox.showerror("Open failed", f"Could not open folder: {path}")
            except Exception:
                pass

    def view_destination_folder(self):
        """Open the destination folder in the OS file manager."""
        path = self.file_paths.get('destination_folder', '')
        if not path or not os.path.isdir(path):
            try:
                messagebox.showwarning("Folder not available", "No valid destination folder is selected.")
            except Exception:
                pass
            return
        try:
            self._open_folder(path)
        except Exception:
            try:
                messagebox.showerror("Open failed", f"Could not open folder: {path}")
            except Exception:
                pass

    def _open_folder(self, path: str):
        """Cross-platform open folder helper.
        - Windows: os.startfile
        - macOS: open
        - Linux: xdg-open (fallback to webbrowser)
        """
        if sys.platform.startswith('win'):
            # type: ignore
            os.startfile(path)
            return
        if sys.platform == 'darwin':
            subprocess.Popen(['open', path])
            return
        # Assume Linux/Unix
        try:
            subprocess.Popen(['xdg-open', path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return
        except Exception:
            # Fallback: try webbrowser to open a file:// URL
            try:
                import webbrowser
                webbrowser.open('file://' + os.path.abspath(path))
            except Exception:
                # last resort: raise
                raise

    def _show_edit_profiles_menu(self):
        """Show the Edit Profiles submenu via keyboard shortcut (Ctrl+E)."""
        try:
            if not hasattr(self, 'edit_profiles_menu'):
                return

            # Try to programmatically open the Edit menu and show the profiles submenu
            # Unfortunately, Tkinter doesn't support true programmatic menu cascading
            # So we show the profiles menu directly at a logical position

            # Calculate position right below the menubar to avoid overlap with Select a Profile dropdown
            x = self.winfo_rootx() + 150
            y = self.winfo_rooty() + 0  # Positioned higher, just below menubar

            # tk_popup is better than post() - it properly grabs focus and stays visible
            self.edit_profiles_menu.tk_popup(x, y, 0)

            # Release the grab after the menu is dismissed (tk_popup handles this automatically)
        except Exception:
            pass

    def create_profile_window(self):
        # Prefer the delegated flag when frozen so the bootloader runs our
        # delegated handler which imports the embedded profile_editor from sys._MEIPASS.
        try:
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                # Invoke the same exe with a flag the __main__ handles
                self.profile_process = subprocess.Popen([sys.executable, '--run-cropping-gui'])
                self.check_profile_process()
                return
        except Exception:
            pass

        # Non-frozen: run the local script next to the module
        target_script = os.path.join(os.path.dirname(__file__), "profile_editor.py")
        if os.path.exists(target_script):
            self.profile_process = subprocess.Popen([sys.executable, target_script])
            self.check_profile_process()

    def edit_profile_window(self):
        profile_path = filedialog.askopenfilename(initialdir=CONFIG_FOLDER, filetypes=(("Profile Files", "*.profile"),))
        if not profile_path:
            return
        try:
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                # delegate to the frozen exe which will import the embedded module
                self.profile_process = subprocess.Popen([sys.executable, '--run-cropping-gui', profile_path])
                self.check_profile_process()
                return
        except Exception:
            pass

        target_script = os.path.join(os.path.dirname(__file__), "profile_editor.py")
        if os.path.exists(target_script):
            self.profile_process = subprocess.Popen([sys.executable, target_script, profile_path])
            self.check_profile_process()

    def check_profile_process(self):
        if self.profile_process and self.profile_process.poll() is None:
            self.after(500, self.check_profile_process)
        else:
            self.refresh_profile_dropdown()
            self.profile_process = None

    def create_menubar(self):
        # Build a traditional top menubar (File, Edit, Select, Actions)
        import traceback
        try:
            self.menubar = tk.Menu(self)

            # Determine modifier key for keyboard shortcuts (Command on macOS, Ctrl elsewhere)
            mod_key = 'Cmd' if sys.platform == 'darwin' else 'Ctrl'

            # --- File menu ---
            self.file_menu = tk.Menu(self.menubar, tearoff=0)
            # Select entries (no accelerator) — user can still pick a folder via dialog
            self.file_menu.add_command(label='Select Source Folder...', command=lambda: self.select_source_folder())
            self.file_menu.add_command(label='Select Destination Folder...', command=lambda: self.select_destination_folder())
            self.file_menu.add_separator()
            # Open entries with accelerators: Ctrl+O -> open source, Ctrl+Shift+O -> open destination
            self.file_menu.add_command(label='Open Source Folder', command=lambda: self.view_source_folder(), accelerator=f'{mod_key}+O')
            self.file_menu.add_command(label='Open Destination Folder', command=lambda: self.view_destination_folder(), accelerator=f'{mod_key}+Shift+O')
            self.file_menu.add_separator()
            self.file_menu.add_command(label='Exit', command=lambda: self.quit(), accelerator=f'{mod_key}+Q')
            self.menubar.add_cascade(label='File', menu=self.file_menu)

            # --- Edit menu ---
            self.edit_menu = tk.Menu(self.menubar, tearoff=0)
            # Create Profile menu item now lives under Edit
            shortcut_n = '⌘N' if sys.platform == 'darwin' else 'Ctrl+N'
            self.edit_menu.add_command(label=f'Create Profile... ({shortcut_n})', command=lambda: self.create_profile_window())
            # Edit Profiles submenu under Edit menu (populated now and refreshed later)
            self.edit_profiles_menu = tk.Menu(self.edit_menu, tearoff=0)
            profiles_initial = load_profiles()
            if profiles_initial:
                for _p in profiles_initial:
                    # bind with default arg to avoid late-binding
                    self.edit_profiles_menu.add_command(label=_p, command=(lambda name=_p: self.open_profile_in_editor(name)))
            else:
                self.edit_profiles_menu.add_command(label='(no profiles)', state='disabled')
            # Note: Tkinter doesn't display accelerators for cascade menus, so we show the shortcut in the label
            shortcut_display = '⌘E' if sys.platform == 'darwin' else 'Ctrl+E'
            self.edit_menu.add_cascade(label=f'Edit Profiles... ({shortcut_display})', menu=self.edit_profiles_menu)
            # visual divider after Edit Profiles submenu
            self.edit_menu.add_separator()
            # bind checkbutton to the existing BooleanVar so it persists
            self.edit_menu.add_checkbutton(label='Delete Original Images after Cropping', variable=self.delete_original_var)
            # New: Delete original images after Move action (mutually exclusive with cropping option)
            try:
                self.edit_menu.add_checkbutton(label='Delete Original Images after Moving', variable=self.delete_move_var)
            except Exception:
                # fallback: still try to add without crashing
                try:
                    self.edit_menu.add_checkbutton(label='Delete Original Images after Moving', variable=self.delete_move_var)
                except Exception:
                    pass
            # Removed divider and 'Preferences...' per user request
            self.menubar.add_cascade(label='Edit', menu=self.edit_menu)

            # --- Actions menu ---
            self.actions_menu = tk.Menu(self.menubar, tearoff=0)
            self.actions_menu.add_command(label='Crop', command=lambda: self.run_cropping(save_after=False), accelerator=f'{mod_key}+R')
            self.actions_menu.add_command(label='Crop & Move', command=lambda: self.run_cropping(save_after=True), accelerator=f'{mod_key}+Shift+R')
            self.actions_menu.add_command(label='Move', command=lambda: self.run_move_only(), accelerator=f'{mod_key}+M')
            self.menubar.add_cascade(label='Actions', menu=self.actions_menu)

            # --- Help menu (previously 'Info') ---
            self.info_menu = tk.Menu(self.menubar, tearoff=0)
            # Support link updated to the Image Splitter Pro release/post URL
            self.info_menu.add_command(label='Support Us', command=lambda: webbrowser.open('https://www.abelxl.com/2026/01/image-splitter-pro.html'))
            # Open the project repository so users can check releases/updates
            self.info_menu.add_command(label='Check for Updates', command=lambda: webbrowser.open('https://github.com/AbelXL/Image-Splitter-Pro'))
            # New: quick links to file issues on GitHub
            self.info_menu.add_separator()
            self.info_menu.add_command(
                label='Report a Bug',
                command=lambda: webbrowser.open('https://github.com/AbelXL/Image-Splitter-Pro/issues/new?template=submit-bug-report.md')
            )
            self.info_menu.add_command(
                label='Request a Feature',
                command=lambda: webbrowser.open('https://github.com/AbelXL/Image-Splitter-Pro/issues/new?template=request-a-feature.md')
            )
            # Version displayed as a disabled menu item
            self.info_menu.add_command(label='Version 1.0.0', state='disabled')
            # Show as 'Help' in the menubar (renamed from 'Info')
            self.menubar.add_cascade(label='Help', menu=self.info_menu)

            # Attach the menubar to the root window
            try:
                self.config(menu=self.menubar)
            except Exception:
                pass

            # Keyboard accelerators (Control on Windows/Linux, Command on macOS)
            mod = 'Command' if sys.platform == 'darwin' else 'Control'
            try:
                self.bind_all(f'<{mod}-n>', lambda e: self.create_profile_window())
                # Ctrl/Cmd+O opens the source folder (view) instead of Select dialog
                self.bind_all(f'<{mod}-o>', lambda e: self.view_source_folder())
                try:
                    self.bind_all(f'<{mod}-O>', lambda e: self.view_source_folder())
                except Exception:
                    pass
                # Shift+O opens the destination folder (view)
                if sys.platform == 'darwin':
                    self.bind_all('<Command-Shift-O>', lambda e: self.view_destination_folder())
                else:
                    # Bind both lowercase and uppercase Shift variants
                    self.bind_all('<Control-Shift-o>', lambda e: self.view_destination_folder())
                    try:
                        self.bind_all('<Control-Shift-O>', lambda e: self.view_destination_folder())
                    except Exception:
                        pass
                self.bind_all(f'<{mod}-q>', lambda e: self.quit())
                self.bind_all(f'<{mod}-r>', lambda e: self.run_cropping(save_after=False))
                if sys.platform == 'darwin':
                    self.bind_all('<Command-Shift-R>', lambda e: self.run_cropping(save_after=True))
                else:
                    self.bind_all('<Control-Shift-r>', lambda e: self.run_cropping(save_after=True))
                self.bind_all(f'<{mod}-m>', lambda e: self.run_move_only())
                self.bind_all('<F5>', lambda e: self.refresh_thumbnails(is_polling=False))
                # Ctrl/Cmd+E opens the Edit Profiles menu
                self.bind_all(f'<{mod}-e>', lambda e: self._show_edit_profiles_menu())
            except Exception:
                pass

            # initialize menu/button state
            try:
                self.update_menu_state()
            except Exception:
                traceback.print_exc()
        except Exception:
            # If menu creation fails, print the traceback so the user can see why
            print("[ERROR] create_menubar: failed to create menubar")
            traceback.print_exc()
            return

    def update_menu_state(self):
        # Determine current app state
        try:
            profile_selected = bool(self.selected_profile.get()) and self.selected_profile.get() != '— No Profile Selected —'
        except Exception:
            profile_selected = False

        source_exists = bool(self.file_paths.get('source_folder')) and os.path.isdir(self.file_paths.get('source_folder', ''))
        dest_exists = bool(self.file_paths.get('destination_folder')) and os.path.isdir(self.file_paths.get('destination_folder', ''))

        # Actions: Crop/Crop & Move require profile + source; Move requires profile + source + dest
        try:
            # Fallback: prefer combobox displayed value if StringVar isn't reflecting selection
            try:
                if not profile_selected and hasattr(self, 'profile_combobox'):
                    cb_val = self.profile_combobox.get()
                    if cb_val:
                        profile_selected = (cb_val != '— No Profile Selected —')
            except Exception:
                pass
            if hasattr(self, 'actions_menu'):
                self.actions_menu.entryconfig('Crop', state=('normal' if (profile_selected and source_exists) else 'disabled'))
                self.actions_menu.entryconfig('Crop & Move', state=('normal' if (profile_selected and source_exists and dest_exists) else 'disabled'))
                self.actions_menu.entryconfig('Move', state=('normal' if (profile_selected and source_exists and dest_exists) else 'disabled'))
        except Exception:
            pass

        # File menu view items
        try:
            if hasattr(self, 'file_menu'):
                # Find indices for view entries by label search (safer across versions).
                try:
                    self.file_menu.entryconfig('Open Source Folder', state=('normal' if source_exists else 'disabled'))
                    self.file_menu.entryconfig('Open Destination Folder', state=('normal' if dest_exists else 'disabled'))
                except Exception:
                    pass
        except Exception:
            pass

        # Also enable/disable the left-column buttons to match
        try:
            btn_state = tk.NORMAL if (profile_selected and source_exists) else tk.DISABLED
            self.crop_btn.config(state=btn_state)
            self.crop_move_btn.config(state=(tk.NORMAL if (profile_selected and source_exists and dest_exists) else tk.DISABLED))
            # note: move_btn is a TBButton (ttkbootstrap Button) which uses configure
            try:
                self.move_btn.configure(state=(tk.NORMAL if (profile_selected and source_exists and dest_exists) else tk.DISABLED))
            except Exception:
                try:
                    self.move_btn.config(state=(tk.NORMAL if (profile_selected and source_exists and dest_exists) else tk.DISABLED))
                except Exception:
                    pass
        except Exception:
            pass

    def on_profile_change(self):
        # Called whenever the selected profile changes (combobox selection or programmatic change)
        try:
            # Update menu/button enabled state to reflect new profile selection
            self.update_menu_state()
        except Exception:
            pass

    def open_profile_in_editor(self, profile_name: str):
        """Open the specified profile in the external editor."""
        profile_arg = os.path.join(CONFIG_FOLDER, f"{profile_name}.profile")
        try:
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                self.profile_process = subprocess.Popen([sys.executable, '--run-cropping-gui', profile_arg])
                self.check_profile_process()
                return
        except Exception:
            pass

        target_script = os.path.join(os.path.dirname(__file__), "profile_editor.py")
        if os.path.exists(target_script):
            # Pass the selected profile directly as an argument
            self.profile_process = subprocess.Popen([sys.executable, target_script, profile_arg])
            self.check_profile_process()

    def _show_confirm_delete_dialog(self):
        """Show the confirmation dialog for delete-originals.
        Returns a tuple (confirmed, do_not_show_again).
        - confirmed: True if the user confirmed, False if canceled.
        - do_not_show_again: True if the user chose 'Do not show again', else False.
        """
        dialog = None
        try:
            # Create a transient dialog window
            dialog = tk.Toplevel(self)
            dialog.title("Confirm Delete Originals")
            dialog.geometry("400x200")
            dialog.grab_set()  # Make the dialog modal

            # Center the dialog on the parent window
            x = self.winfo_x() + (self.winfo_width() // 2) - 200
            y = self.winfo_y() + (self.winfo_height() // 2) - 100
            dialog.geometry(f"+{x}+{y}")

            label = tk.Label(dialog, text="This action will delete the original images after cropping.\n\nAre you sure you want to continue?", wraplength=350, justify="center")
            label.pack(pady=20)

            button_frame = tk.Frame(dialog)
            button_frame.pack(pady=10)

            confirmed = [False]  # Use a list to allow modification in the callback
            def on_confirm():
                confirmed[0] = True
                dialog.destroy()

            def on_cancel():
                dialog.destroy()

            yes_button = tk.Button(button_frame, text="Yes", command=on_confirm, width=10, bg="#4CAF50", fg="white")
            yes_button.pack(side="left", padx=5)

            no_button = tk.Button(button_frame, text="No", command=on_cancel, width=10, bg="#f44336", fg="white")
            no_button.pack(side="left", padx=5)

            # New: "Do not show again" checkbox
            do_not_show_var = tk.BooleanVar()
            checkbox = tk.Checkbutton(dialog, text="Do not show again", variable=do_not_show_var, font=("Arial", 10))
            checkbox.pack(pady=(10, 0))

            # Wait for the dialog to close
            self.wait_window(dialog)

            # Return the result
            return confirmed[0], do_not_show_var.get()
        except Exception as e:
            print(f"Error showing confirmation dialog: {e}")
            if dialog:
                dialog.destroy()
        return False, False

    def _show_confirm_delete_after_move_dialog(self):
        """Show the confirmation dialog for delete-after-move.
        Returns a tuple (confirmed, do_not_show_again).
        - confirmed: True if the user confirmed, False if canceled.
        - do_not_show_again: True if the user chose 'Do not show again', else False.
        """
        dialog = None
        try:
            # Create a transient dialog window
            dialog = tk.Toplevel(self)
            dialog.title("Confirm Delete Originals After Move")
            dialog.geometry("400x200")
            dialog.grab_set()  # Make the dialog modal

            # Center the dialog on the parent window
            x = self.winfo_x() + (self.winfo_width() // 2) - 200
            y = self.winfo_y() + (self.winfo_height() // 2) - 100
            dialog.geometry(f"+{x}+{y}")

            label = tk.Label(dialog, text="This action will delete the original images after they are moved.\n\nAre you sure you want to continue?", wraplength=350, justify="center")
            label.pack(pady=20)

            button_frame = tk.Frame(dialog)
            button_frame.pack(pady=10)

            confirmed = [False]  # Use a list to allow modification in the callback
            def on_confirm():
                confirmed[0] = True
                dialog.destroy()

            def on_cancel():
                dialog.destroy()

            yes_button = tk.Button(button_frame, text="Yes", command=on_confirm, width=10, bg="#4CAF50", fg="white")
            yes_button.pack(side="left", padx=5)

            no_button = tk.Button(button_frame, text="No", command=on_cancel, width=10, bg="#f44336", fg="white")
            no_button.pack(side="left", padx=5)

            # New: "Do not show again" checkbox
            do_not_show_var = tk.BooleanVar()
            checkbox = tk.Checkbutton(dialog, text="Do not show again", variable=do_not_show_var, font=("Arial", 10))
            checkbox.pack(pady=(10, 0))

            # Wait for the dialog to close
            self.wait_window(dialog)

            # Return the result
            return confirmed[0], do_not_show_var.get()
        except Exception as e:
            print(f"Error showing confirmation dialog: {e}")
            if dialog:
                dialog.destroy()
        return False, False

    def _maybe_show_onboarding(self):
        """Check config and source_folder and show onboarding dialog on first run.
        Recommended behavior: if the user skips or successfully selects a folder, mark onboarding complete so it doesn't reappear.
        """
        try:
            pref = load_config('show_onboarding')
            show_dialog = True
            if isinstance(pref, str) and pref != '':
                show_dialog = str(pref).lower() in ('1', 'true', 'yes', 'on')
        except Exception:
            show_dialog = True

        # If onboarding is suppressed, don't show anything
        if not show_dialog:
            return

        # Now check folder states
        try:
            sf = self.file_paths.get('source_folder', '')
            df = self.file_paths.get('destination_folder', '')
            # If both are set and valid, nothing to do.
            if sf and os.path.isdir(sf) and df and os.path.isdir(df):
                return

            # If the user has a source but no destination, only prompt for destination.
            if sf and os.path.isdir(sf) and (not df or not os.path.isdir(df)):
                try:
                    dest_selected, dest_do_not_show = self._show_destination_dialog()
                    if dest_selected or dest_do_not_show:
                        try:
                            save_config('show_onboarding', 'False')
                        except Exception:
                            pass
                except Exception:
                    pass
                return
        except Exception:
            # if anything failed while inspecting folders, fall back to showing full onboarding
            pass

        # Otherwise, only show the full onboarding if the user hasn't suppressed it.

        try:
            self._show_onboarding_dialog()
        except Exception:
            # don't crash if dialogs fail (e.g., headless)
            pass

    def _show_onboarding_dialog(self):
        """Simpler, safe onboarding flow using messagebox + filedialog.
        Uses plain messageboxes instead of creating a custom Toplevel to avoid
        startup race conditions that can cause the root window to be destroyed.
        """
        try:
            text = (
                "Please select a Source Folder.\n\n"
                "1. The Source Folder is the location the app monitors for images.\n\n"
                "2. The Source Folder should contain ONLY the images you want the app to edit, move, or delete.\n\n"
                "3. Any other images placed here may be lost.\n\n"
                "Would you like to select a Source Folder now?"
            )
            res = messagebox.askyesno("Welcome — Select Source Folder", text, parent=self)
            if res:
                folder = filedialog.askdirectory(parent=self, title="Select Source Folder")
                if folder:
                    self.file_paths['source_folder'] = folder
                    save_config('source_folder', folder)
                    self.last_seen_files = []
                    try:
                        self.refresh_thumbnails()
                    except Exception:
                        pass
                    # Ask for destination next using the shared helper so behavior is consistent
                    try:
                        dest_selected, dest_do_not_show = self._show_destination_dialog()
                        # If the user completed destination selection or explicitly chose to suppress onboarding,
                        # mark onboarding complete. Otherwise keep it enabled so the prompt returns on next run.
                        if dest_selected or dest_do_not_show:
                            try:
                                save_config('show_onboarding', 'False')
                            except Exception:
                                pass
                        else:
                            try:
                                save_config('show_onboarding', 'True')
                            except Exception:
                                pass
                    except Exception:
                        pass
                else:
                    # cancelled source selection -> ask whether to suppress onboarding
                    s = messagebox.askyesno("Do not show again?", "Do not show this message again?", parent=self)
                    if s:
                        save_config('show_onboarding', 'False')
            else:
                # user skipped -> ask whether to suppress onboarding
                s = messagebox.askyesno("Do not show again?", "Do not show this message again?", parent=self)
                if s:
                    save_config('show_onboarding', 'False')
                else:
                    # keep onboarding for next run
                    save_config('show_onboarding', 'True')
        except Exception:
            # Fail silently if dialogs cannot be shown
            pass

    def _show_destination_dialog(self):
        """Simpler destination prompt using messagebox + filedialog.
        Returns (selected: bool, do_not_show_again: bool).
        """
        try:
            text = (
                "Please select a Destination Folder.\n\n"
                "The Destination Folder is where edited images will be moved to.\n\n"
                "Would you like to select a Destination Folder now?"
            )
            res = messagebox.askyesno("Destination Folder", text, parent=self)
            if res:
                folder = filedialog.askdirectory(parent=self, title="Select Destination Folder")
                if folder:
                    self.file_paths['destination_folder'] = folder
                    save_config('destination_folder', folder)
                    try:
                        self.update_menu_state()
                    except Exception:
                        pass
                    # After user selects destination via this prompt, show the final onboarding info dialog
                    try:
                        messagebox.showinfo(
                            "Next step",
                            "Please add an image(s) to the Source Folder.\n\nThen create a Profile.\n\nProfiles are a set of rules on how to split your images.",
                            parent=self,
                        )
                    except Exception:
                        pass
                    return True, False
                else:
                    # cancelled selection
                    return False, False
            else:
                # ask if user wants to suppress future prompts
                s = messagebox.askyesno("Do not show again?", "Do not show this message again?", parent=self)
                if s:
                    return False, True
                return False, False
        except Exception:
            return False, False

if __name__ == "__main__":
    try:
        # Diagnostic flag: print config locations and exit
        if len(sys.argv) >= 2 and sys.argv[1] == '--print-config':
            print(f"CONFIG_FOLDER={CONFIG_FOLDER}")
            print(f"CONFIG_FILE={CONFIG_FILE}")
            print(f"sys.executable={sys.executable}")
            print(f"sys._MEIPASS={'<not set>' if not hasattr(sys, '_MEIPASS') else sys._MEIPASS}")
            sys.exit(0)
        # Delegated runner: when the onefile exe is invoked as
        #   image_wizard.exe --run-cropping-gui [optional_profile_path]
        # we import and run the embedded profile_editor module instead of
        # launching a second copy of the full app.
        if len(sys.argv) >= 2 and sys.argv[1] == '--run-cropping-gui':
            profile_arg = sys.argv[2] if len(sys.argv) > 2 else None
            # Locate profile_editor.py: prefer sys._MEIPASS when frozen
            try:
                if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                    target = os.path.join(sys._MEIPASS, 'profile_editor.py')
                else:
                    target = os.path.join(os.path.dirname(__file__), 'profile_editor.py')
                if os.path.exists(target):
                    import importlib.util
                    spec = importlib.util.spec_from_file_location('embedded_profile_editor', target)
                    if spec and spec.loader:
                        mod = importlib.util.module_from_spec(spec)
                        spec.loader.exec_module(mod)
                        if hasattr(mod, 'open_cropping_window'):
                            try:
                                mod.open_cropping_window(parent=None, profile_file_path=profile_arg)
                            except TypeError:
                                mod.open_cropping_window(parent=None)
                        else:
                            # Fallback: if module has no entrypoint, try running as script
                            if hasattr(mod, '__file__'):
                                run_target = mod.__file__
                                subprocess.call([sys.executable, run_target] + ([profile_arg] if profile_arg else []))
                        # After cropping GUI exits, terminate this process
                        sys.exit(0)
            except Exception:
                # If anything fails, fall through to normal startup so user still gets the main app.
                import traceback
                traceback.print_exc()

        print("[image_wizard] launching ImageCroppingApp...")
        app = ImageCroppingApp()
        print("[image_wizard] created app, entering mainloop")
        app.mainloop()
        print("[image_wizard] mainloop exited")
    except Exception:
        # Global exception handler: log any uncaught exceptions to the console
        import traceback
        traceback.print_exc()
