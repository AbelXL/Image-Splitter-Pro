"""Minimal cropping GUI replacement used for testing bundling.
Provides open_cropping_window(parent=None, profile_file_path=None) so it can be
called either embedded (imported) or launched as a standalone script.
"""
import tkinter as tk
from tkinter import ttk
import sys
import os
import json
from tkinter import messagebox
from tkinter import simpledialog
import csv
import re



#Image Splitter Pro
#Author: Abel Aramburo (@AbelXL) (https://github.com/AbelXL) (https://www.abelxl.com/)
#Created: 2026-01-19
#Copyright (c) 2026 Abel Aramburo
#This project was developed with the assistance of AI logic-modeling to ensure high-performance image handling and a modern user experience.
#This project is licensed under the **MIT License**. This means you are free to use, modify, and distribute the software, provided that the original copyright notice and this permission notice are included in all copies or substantial portions of the software.

# Try to import ttkbootstrap if available. We keep a minimal, safe fallback so
# the module can run whether ttkbootstrap is installed (useful for bundling).
try:
    import ttkbootstrap as tb
    TTB_AVAILABLE = True
except Exception:
    tb = None
    TTB_AVAILABLE = False


try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except Exception:
    Image = None
    ImageTk = None
    PIL_AVAILABLE = False

# HEIC support
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIC_SUPPORTED = True
except ImportError:
    HEIC_SUPPORTED = False

# Choose a resampling filter in a way that doesn't raise a static-analysis
# error when PIL is not available. If PIL is present it will pick LANCZOS if
# available, otherwise fall back to ANTIALIAS or a numeric sentinel.
_RESAMPLE_FILTER = None
if PIL_AVAILABLE:
    _RESAMPLE_FILTER = getattr(Image, 'LANCZOS', getattr(Image, 'ANTIALIAS', None))
else:
    _RESAMPLE_FILTER = None


def open_cropping_window(parent=None, profile_file_path=None):
    """Open the full MinimalProfileEditor UI.

    If parent is None we create a new Tk root and run the mainloop. When used
    as an imported module the caller can provide a parent and embed the editor
    into an existing event loop.
    """
    print("[cropping_gui2] open_cropping_window called, parent=", parent)
    created_root = False
    if parent is None:
        root = tk.Tk()
        created_root = True
        # ensure a sensible default size so left controls are visible
        try:
            root.geometry('1000x600')
        except Exception:
            pass
    else:
        root = tk.Toplevel(parent)
        try:
            root.transient(parent)
        except Exception:
            pass

    try:
        root.title('Profile Editor')
    except Exception:
        pass

    # Set Windows AppUserModelID FIRST for proper taskbar grouping and icon display
    # This must be set before setting the icon
    try:
        if sys.platform == 'win32':
            import ctypes
            # Set a unique AppUserModelID so Windows shows our custom icon in the taskbar
            myappid = 'ImageSplitterPro.CropEditor.1.0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        # Non-fatal if this fails (older Windows or other issues)
        pass

    # Set window icon (same as main app for consistent branding)
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
                # Running as compiled exe - icon is in _internal subfolder next to exe
                local_icon = os.path.join(os.path.dirname(sys.executable), icon_filename)
            else:
                # Running from source - check _internal subfolder relative to script directory
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
                    if PIL_AVAILABLE:
                        img = Image.open(icon_path)
                        photo = ImageTk.PhotoImage(img)
                        root.wm_iconphoto(True, photo)
                        # Keep reference to avoid garbage collection
                        root._app_icon = photo
                except Exception:
                    pass
            else:
                # Windows: try iconbitmap first, fallback to wm_iconphoto
                try:
                    root.iconbitmap(icon_path)
                except Exception:
                    try:
                        if PIL_AVAILABLE:
                            img = Image.open(icon_path)
                            photo = ImageTk.PhotoImage(img)
                            root.wm_iconphoto(True, photo)
                            root._app_icon = photo
                    except Exception:
                        pass
    except Exception:
        pass

    # Apply ttkbootstrap 'cosmo' theme if available. This is done after
    # creating the root so the theme affects the widgets created below. We
    # keep this guarded so the module still works when ttkbootstrap isn't
    # installed (or not bundled) — the standard ttk theme will be used instead.
    if TTB_AVAILABLE:
        try:
            # Create a Style instance which will apply the theme to ttk widgets.
            # Using 'cosmo' by default to match image_wizard's look.
            try:
                tb.Style(theme='cosmo')
            except TypeError:
                # older/newer versions of ttkbootstrap may accept the theme name
                # as the first positional argument.
                tb.Style('cosmo')
        except Exception:
            # If anything goes wrong, silently fall back to default ttk.
            pass

    # Instantiate the editor frame and pack it to fill the window
    editor = None
    try:
        editor = MinimalProfileEditor(root, initial_profile_path=profile_file_path)
        editor.pack(fill='both', expand=True)
    except Exception:
        # Fallback: display a simple label if editor creation fails
        try:
            lbl = ttk.Label(root, text='Cropping GUI (failed to initialize editor)')
            lbl.pack(expand=True)
        except Exception:
            pass

    # If this function created the root, run the mainloop and optionally
    # support a quick --test auto-close mode for automated checks.
    if created_root:
        try:
            print("[cropping_gui2] running mainloop (created_root=True)")
            if '--test' in sys.argv:
                try:
                    # schedule the window to close after a short delay for automated tests
                    # Use a small helper so static analysis tools are happy
                    def _auto_close(arg=None, *args):
                        try:
                            root.destroy()
                        except Exception:
                            pass

                    # pass a concrete argument (None) so static analyzers see args provided
                    root.after(1000, _auto_close, None)
                except Exception:
                    pass
            # Force geometry/layout pass so minsize and widths apply before showing
            try:
                root.update_idletasks()
            except Exception:
                pass
            # Diagnostics: print environment and root state so we can debug
            # unexpected early mainloop exits when run from the user's setup.
            try:
                print('[cropping_gui2] sys.argv=', sys.argv)
            except Exception:
                pass
            try:
                print('[cropping_gui2] root.winfo_exists()=', int(bool(root.winfo_exists())))
            except Exception:
                pass
            try:
                print('[cropping_gui2] root.winfo_ismapped()=', int(bool(root.winfo_ismapped())))
            except Exception:
                pass
            try:
                print('[cropping_gui2] root.state()=', root.state())
            except Exception:
                pass
            try:
                print('[cropping_gui2] root children=', root.winfo_children())
            except Exception:
                pass
            # Enter the Tk main loop normally
            root.mainloop()
            try:
                print('[cropping_gui2] after mainloop root.winfo_exists()=', int(bool(root.winfo_exists())))
            except Exception:
                pass
            try:
                print('[cropping_gui2] after mainloop root.winfo_children()=', root.winfo_children())
            except Exception:
                pass
            print("[cropping_gui2] mainloop exited")
        except Exception:
            pass
        return editor
    return editor


# Determine config folder local to the script/executable (do NOT use APPDATA)
# Profiles and .profile files will be stored in <script_or_exe_dir>/config
# This keeps saved profiles local to the application's config folder.
if getattr(sys, 'frozen', False):
    _module_dir = os.path.dirname(sys.executable)
else:
    _module_dir = os.path.dirname(os.path.abspath(__file__))

CONFIG_FOLDER = os.path.join(_module_dir, 'config')
if not os.path.exists(CONFIG_FOLDER):
    try:
        os.makedirs(CONFIG_FOLDER, exist_ok=True)
    except Exception:
        pass


def load_profile_by_path(path: str) -> dict | None:
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def save_profile_to_path(profile: dict, path: str) -> bool:
    try:
        os.makedirs(os.path.dirname(path) or CONFIG_FOLDER, exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(profile, f, indent=2)
        return True
    except Exception:
        return False


def load_config(setting: str) -> str:
    """Load a setting from config.csv located next to the running script/exe.

    Behavior:
    - When running normally (development) the config file is looked up at
      <script_dir>/config/config.csv (i.e. next to `cropping_gui2.py`).
    - When running as a frozen exe, the config file is looked up next to the
      executable: <exe_dir>/config/config.csv. This ensures the exe and the
      script behave identically and do not look in APPDATA.
    Returns the stored value for `setting` or an empty string if not found.
    """
    # Determine the directory to consider "local". When frozen use the
    # executable location so the bundled exe looks for config next to itself.
    if getattr(sys, 'frozen', False):
        module_dir = os.path.dirname(sys.executable)
    else:
        module_dir = os.path.dirname(os.path.abspath(__file__))

    cfg_path = os.path.join(module_dir, 'config', 'config.csv')

    if not os.path.exists(cfg_path):
        return ""
    try:
        with open(cfg_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 2 and row[0] == setting:
                    return row[1]
    except Exception:
        pass
    return ""


class MinimalProfileEditor(tk.Frame):
    def _parse_aspect(self, aspect: str) -> tuple | None:
        """Parse aspect string like '1:1' or '1.91:1' -> (width, height). Returns None for 'none' or invalid."""
        try:
            if not aspect or aspect == 'none' or aspect == 'custom':
                return None
            parts = aspect.split(':')
            if len(parts) != 2:
                return None
            w = float(parts[0])
            h = float(parts[1])
            if w <= 0 or h <= 0:
                return None
            return (w, h)
        except Exception:
            return None

    def _apply_aspect_to_crop(self, aspect: str):
        """Adjust current crop rectangle to match the given aspect while keeping it inside preview bounds.
        If aspect is 'none' nothing is done.
        """
        try:
            asp = self._parse_aspect(aspect)
            if asp is None:
                return
            aw, ah = asp
            # ratio as float
            r = aw / ah
            # ensure we have an existing crop rect; otherwise initialize
            if not self._crop_rect:
                ix, iy = self._preview_image_pos
                iw, ih = self._preview_image_size
                if ix is None:
                    ix, iy = 0, 0
                self._crop_rect = [ix, iy, ix + iw, iy + ih]
            x1, y1, x2, y2 = map(int, self._crop_rect)
            iw_img = self._preview_image_size[0]
            ih_img = self._preview_image_size[1]
            ix, iy = self._preview_image_pos
            if ix is None:
                ix, iy = 0, 0
            # center of current rect
            cx = (x1 + x2) / 2.0
            cy = (y1 + y2) / 2.0
            # start with current size but adjust to aspect
            cur_w = max(1, x2 - x1)
            cur_h = max(1, y2 - y1)
            # prefer to keep size but change one dimension to match ratio
            if (cur_w / cur_h) > r:
                # too wide -> reduce width
                target_h = cur_h
                target_w = max(1, int(round(target_h * r)))
            else:
                target_w = cur_w
                target_h = max(1, int(round(target_w / r)))
            # Ensure target fits within image; if not, scale down
            max_w = iw_img
            max_h = ih_img
            if target_w > max_w:
                target_w = max_w
                target_h = max(1, int(round(target_w / r)))
            if target_h > max_h:
                target_h = max_h
                target_w = max(1, int(round(target_h * r)))
            # build new rect centered on cx,cy but clamped to image bounds
            nx1 = int(round(cx - target_w / 2.0))
            ny1 = int(round(cy - target_h / 2.0))
            nx2 = nx1 + target_w
            ny2 = ny1 + target_h
            # clamp into image
            if nx1 < ix:
                shift = ix - nx1
                nx1 += shift
                nx2 += shift
            if nx2 > ix + iw_img:
                shift = nx2 - (ix + iw_img)
                nx1 -= shift
                nx2 -= shift
            if ny1 < iy:
                shift = iy - ny1
                ny1 += shift
                ny2 += shift
            if ny2 > iy + ih_img:
                shift = ny2 - (iy + ih_img)
                ny1 -= shift
                ny2 -= shift
            # final sanity
            if nx2 - nx1 < 1 or ny2 - ny1 < 1:
                nx1, ny1, nx2, ny2 = ix, iy, ix + iw_img, iy + ih_img
            self._crop_rect = [nx1, ny1, nx2, ny2]
            try:
                self._clamp_crop_rect_to_preview()
            except Exception:
                pass
            # update linked entry vars so UI fields reflect the crop
            try:
                vals = [self._crop_rect[i] if self._crop_rect[i] is not None else 0 for i in range(4)]
                # Use central helper so we update either the shared central vars or
                # the per-rule vars depending on which rule is active.
                try:
                    active_outer = getattr(self, '_active_rule_outer', None)
                except Exception:
                    active_outer = None
                try:
                    self._sync_crop_values(int(vals[0]), int(vals[1]), int(vals[2]), int(vals[3]), active_outer=active_outer)
                except Exception:
                    # Fallback to directly updating central vars if helper missing
                    try:
                        self.x1_var.set(str(int(vals[0])))
                        self.y1_var.set(str(int(vals[1])))
                        self.x2_var.set(str(int(vals[2])))
                        self.y2_var.set(str(int(vals[3])))
                    except Exception:
                        pass
            except Exception:
                pass
        except Exception:
            pass

    def _sync_crop_values(self, x1: int, y1: int, x2: int, y2: int, active_outer=None):
        """Centralized helper to update crop entry StringVars.

        - Updates the central (initial rule) vars only when there is no active
          rule or when the active rule is the first rule (backwards-compat).
        - Also updates the currently-active rule's per-rule vars (if present).
        - Converts from preview space to original image space for display.
        """
        try:
            # Convert from preview coordinates to original image coordinates for display
            x1_orig, y1_orig = self._preview_to_original_coords(x1, y1)
            x2_orig, y2_orig = self._preview_to_original_coords(x2, y2)

            # Determine whether central shared vars should be updated
            update_central = False
            try:
                if not self.rule_boxes:
                    update_central = True
                else:
                    first_outer = self.rule_boxes[0].get('outer')
                    if active_outer is None or active_outer is first_outer:
                        update_central = True
            except Exception:
                update_central = False

            if update_central:
                try:
                    self.x1_var.set(str(int(x1_orig)))
                    self.y1_var.set(str(int(y1_orig)))
                    self.x2_var.set(str(int(x2_orig)))
                    self.y2_var.set(str(int(y2_orig)))
                except Exception:
                    pass

            # Update per-rule vars for the active rule (if any)
            try:
                if active_outer is not None:
                    for rb in self.rule_boxes:
                        try:
                            if rb.get('outer') is active_outer:
                                rvars = rb.get('vars', {}) or {}
                                try:
                                    if rvars.get('x1') is not None:
                                        rvars.get('x1').set(str(int(x1_orig)))
                                except Exception:
                                    pass
                                try:
                                    if rvars.get('y1') is not None:
                                        rvars.get('y1').set(str(int(y1_orig)))
                                except Exception:
                                    pass
                                try:
                                    if rvars.get('x2') is not None:
                                        rvars.get('x2').set(str(int(x2_orig)))
                                except Exception:
                                    pass
                                try:
                                    if rvars.get('y2') is not None:
                                        rvars.get('y2').set(str(int(y2_orig)))
                                except Exception:
                                    pass
                                break
                        except Exception:
                            pass
            except Exception:
                pass
        except Exception:
            pass

    def _resize_with_aspect(self, handle: str, orig: list, mouse_x: int, mouse_y: int, aspect: str) -> list:
        """Return new [x1,y1,x2,y2] for a resize action that preserves aspect (if aspect != 'none').
        handle: 'nw','ne','sw','se'
        orig: original rect [x1,y1,x2,y2]
        mouse_x, mouse_y: current mouse canvas coords
        """
        try:
            asp = self._parse_aspect(aspect)
            if asp is None:
                return orig
            aw, ah = asp
            r = aw / ah
            x1, y1, x2, y2 = orig
            ix, iy = self._preview_image_pos
            iw, ih = self._preview_image_size
            if ix is None:
                ix, iy = 0, 0
            minx = ix
            miny = iy
            maxx = ix + iw
            maxy = iy + ih
            # clamp mouse to image bounds
            mx = max(minx, min(mouse_x, maxx))
            my = max(miny, min(mouse_y, maxy))
            # depending on handle, the opposite corner is fixed
            if handle == 'se':
                fx, fy = x1, y1
                # desired width/height from fixed corner to mouse
                desired_w = max(1, int(mx - fx))
                desired_h = max(1, int(my - fy))
                # adjust to aspect: compute width from height then check bounds
                w_from_h = int(round(desired_h * r))
                h_from_w = int(round(desired_w / r))
                nw = desired_w
                nh = h_from_w
                if fx + nw > maxx:
                    nw = maxx - fx
                    nh = int(round(nw / r))
                if fy + nh > maxy:
                    nh = maxy - fy
                    nw = int(round(nh * r))
                # final coords
                nx1, ny1, nx2, ny2 = fx, fy, fx + nw, fy + nh
                return [nx1, ny1, nx2, ny2]
            if handle == 'nw':
                fx, fy = x2, y2
                desired_w = max(1, int(fx - mx))
                desired_h = max(1, int(fy - my))
                w_from_h = int(round(desired_h * r))
                h_from_w = int(round(desired_w / r))
                nw = desired_w
                nh = h_from_w
                if fx - nw < minx:
                    nw = fx - minx
                    nh = int(round(nw / r))
                if fy - nh < miny:
                    nh = fy - miny
                    nw = int(round(nh * r))
                nx1, ny1, nx2, ny2 = fx - nw, fy - nh, fx, fy
                return [nx1, ny1, nx2, ny2]
            if handle == 'ne':
                fx, fy = x1, y2
                # top-right: fixed x1,y2 ; changing x2 (mx) and y1 (my)
                desired_w = max(1, int(mx - fx))
                desired_h = max(1, int(fy - my))
                w_from_h = int(round(desired_h * r))
                h_from_w = int(round(desired_w / r))
                nw = desired_w
                nh = h_from_w
                if fx + nw > maxx:
                    nw = maxx - fx
                    nh = int(round(nw / r))
                if fy - nh < miny:
                    nh = fy - miny
                    nw = int(round(nh * r))
                nx1, ny1, nx2, ny2 = fx, fy - nh, fx + nw, fy
                return [nx1, ny1, nx2, ny2]
            if handle == 'sw':
                fx, fy = x2, y1
                desired_w = max(1, int(fx - mx))
                desired_h = max(1, int(my - fy))
                w_from_h = int(round(desired_h * r))
                h_from_w = int(round(desired_w / r))
                nw = desired_w
                nh = h_from_w
                if fx - nw < minx:
                    nw = fx - minx
                    nh = int(round(nw / r))
                if fy + nh > maxy:
                    nh = maxy - fy
                    nw = int(round(nh * r))
                nx1, ny1, nx2, ny2 = fx - nw, fy, fx, fy + nh
                return [nx1, ny1, nx2, ny2]
            return orig
        except Exception:
            return orig

    def __init__(self, master, initial_profile_path: str | None = None):
        super().__init__(master)
        self.master = master
        # Track which rule outer frame is currently active (clicked/selected)
        self._active_rule_outer = None
        # Track updating state for crop inputs to prevent infinite loops
        self._updating_rule_crop_inputs = {}
        self.current_profile_path = None
        self.current_profile = {'profile_name': '', 'rules': []}
        # keep track of created rule boxes (each is a dict with 'outer' frame and 'label')
        self.rule_boxes: list[dict] = []
        # central selected position variable (kept for backward compatibility)
        self.selected_position_var = tk.StringVar(value='')

        # Left column UI - TOP: Profile name entry
        left = ttk.Frame(self)
        # Pin the left frame to the west (left) edge of its grid cell so
        # expanding the column produces whitespace on the right instead
        # of shifting the content inward.
        left.grid(row=0, column=0, sticky='nsw')
        try:
            # Reserve a larger fixed width for the left panel so the canvas
            # content stays the intended size even when a vertical scrollbar
            # is present. Mouse/OS themes vary; 340 leaves room for a scrollbar.
            left.configure(width=240)
            # prevent the left column from collapsing to its children
            try:
                left.grid_propagate(False)
            except Exception:
                pass
            try:
                # ensure the overall frame reserves space for the left column
                # and keep the column fixed (weight=0) so the right column takes remaining space
                self.grid_columnconfigure(0, minsize=240)
                self.grid_columnconfigure(0, weight=0)
            except Exception:
                pass
        except Exception:
            pass

        # Profile name at top
        top_name_row = ttk.Frame(left)
        top_name_row.pack(fill='x', padx=8, pady=(8, 4))
        ttk.Label(top_name_row, text='Profile Name:', font=('Segoe UI', 10, 'bold')).pack(side='left')
        # Default name for new profiles so the Entry shows a helpful placeholder
        self.name_var = tk.StringVar(value='New Cropping Rule')
        ttk.Entry(top_name_row, textvariable=self.name_var, width=28).pack(side='left', padx=(8,0))

        # Container that will hold one or more rule boxes. New rules are added here.
        # Scrollable container that will hold one or more rule boxes. New rules
        # are added into an inner frame hosted by a Canvas so the user can
        # vertically scroll when many rules exist.
        self.rules_scroll_container = ttk.Frame(left)
        self.rules_scroll_container.pack(fill='both', expand=True)

        # Canvas + vertical scrollbar to host the rules
        # Give the Canvas an explicit width slightly smaller than the left
        # panel width so padding and the vertical scrollbar do not reduce
        # the available content space unexpectedly.
        self.rules_canvas = tk.Canvas(self.rules_scroll_container, bd=0, highlightthickness=0, width=220)
        self.rules_vscroll = ttk.Scrollbar(self.rules_scroll_container, orient='vertical', command=self.rules_canvas.yview)
        self.rules_canvas.configure(yscrollcommand=self.rules_vscroll.set)
        self.rules_vscroll.pack(side='right', fill='y')
        self.rules_canvas.pack(side='left', fill='both', expand=True)

        # Inner frame placed inside the canvas - we will add rule frames to this
        # inner frame (kept as self.rules_container so later code is unchanged).
        self.rules_container = ttk.Frame(self.rules_canvas)
        self.rules_window = self.rules_canvas.create_window((0, 0), window=self.rules_container, anchor='nw')

        # Keep the canvas scrollregion in sync with the inner content
        def _on_rules_config(e):
            try:
                self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox('all'))
            except Exception:
                pass

        def _on_rules_canvas_config(e):
            try:
                self.rules_canvas.itemconfig(self.rules_window, width=e.width)
            except Exception:
                pass

        self.rules_container.bind('<Configure>', _on_rules_config)
        self.rules_canvas.bind('<Configure>', _on_rules_canvas_config)

        # Mouse wheel support on Windows (works when mouse is over the canvas)
        def _on_mousewheel(e):
            try:
                # Only scroll if the inner content is larger than the visible canvas
                try:
                    bbox = self.rules_canvas.bbox('all')
                    if not bbox:
                        return
                    content_h = bbox[3] - bbox[1]
                    vis_h = self.rules_canvas.winfo_height()
                    # if content fits in the visible area, don't scroll
                    if content_h <= vis_h:
                        return
                except Exception:
                    # if any error occurs querying sizes, fall back to attempting to scroll
                    pass
                # e.delta is 120 units per wheel 'click' on Windows
                self.rules_canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
            except Exception:
                pass

        try:
            self.rules_canvas.bind_all('<MouseWheel>', _on_mousewheel)
        except Exception:
            pass

        # Blueish box containing the cropping rule(s)
        # Use a thinner, subtler border: remove frame bd/relief and use a 1-pixel
        # highlight border which appears less heavy than the default 'solid' relief.
        # Adjust `highlightbackground` to a slightly darker blue to keep a soft
        # separation while making the border visually thinner.
        # Keep the same blue when the frame is focused by also setting highlightcolor
        rule_box_outer = tk.Frame(self.rules_container, bg="#e8f2ff", bd=0, highlightthickness=1, highlightbackground="#cfe6ff", highlightcolor="#cfe6ff", takefocus=1)
        rule_box_outer.pack(fill='both', expand=False, padx=8, pady=(4,8))

        # Bind click and focus so the user sees a subtle active background tint
        try:
            rule_box_outer.bind('<Button-1>', lambda e, o=rule_box_outer: self.set_active_rule(o))
            rule_box_outer.bind('<FocusIn>', lambda e, o=rule_box_outer: self.set_active_rule(o))
        except Exception:
            pass

        # Inner padding/frame for nicer layout
        rule_box = ttk.Frame(rule_box_outer, padding=10)
        rule_box.pack(fill='both', expand=True)

        # Bind inner area and header area too so clicks on labels/buttons set the active rule
        try:
            rule_box.bind('<Button-1>', lambda e, o=rule_box_outer: self.set_active_rule(o))
        except Exception:
            pass

        # Rule header with title and Delete button (red) - center the Delete button
        header = ttk.Frame(rule_box)
        header.grid(row=0, column=0, columnspan=3, sticky='ew')
        # Make left and right columns flexible so the middle column stays centered
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(2, weight=1)
        # keep a reference to the label so we can renumber rules when needed
        rule_label = ttk.Label(header, text='Rule 1:', font=('Segoe UI', 12, 'bold'))
        rule_label.grid(row=0, column=0, sticky='w')

        # Make the Delete button a bit larger by adding internal padding and
        # extra horizontal grid padding so it visually has more space left/right.
        try:
            if TTB_AVAILABLE:
                try:
                    # Use ttkbootstrap's themed Button (bootstyle 'danger')
                    del_btn = tb.Button(header, text='Delete', bootstyle='danger', command=lambda o=rule_box_outer: self._remove_rule(o))
                except Exception:
                    # Fall back to a simple tk Button if tb.Button isn't available or fails
                    del_btn = tk.Button(header, text='Delete', bg='#ff4d6d', fg='white', bd=0, padx=8, pady=4, command=lambda o=rule_box_outer: self._remove_rule(o))
            else:
                del_btn = tk.Button(header, text='Delete', bg='#ff4d6d', fg='white', bd=0, padx=8, pady=4, command=lambda o=rule_box_outer: self._remove_rule(o))
            # place the Delete button in the center column with extra horizontal gap
            del_btn.grid(row=0, column=1, padx=12, pady=2)
            # Add hover behavior: lighten the button on mouse enter and restore on leave.
            try:
                # Record original visuals to restore later
                orig_bootstyle = None
                try:
                    orig_bootstyle = del_btn.cget('bootstyle')
                except Exception:
                    orig_bootstyle = None
                orig_bg = None
                try:
                    orig_bg = del_btn.cget('bg')
                except Exception:
                    orig_bg = None
                orig_fg = None
                try:
                    # tk uses 'fg', ttk may use 'foreground'
                    orig_fg = del_btn.cget('fg')
                except Exception:
                    try:
                        orig_fg = del_btn.cget('foreground')
                    except Exception:
                        orig_fg = None
                orig_relief = None
                try:
                    orig_relief = del_btn.cget('relief')
                except Exception:
                    orig_relief = None
                orig_bd = None
                try:
                    orig_bd = del_btn.cget('bd')
                except Exception:
                    orig_bd = None
                orig_cursor = None
                try:
                    orig_cursor = del_btn.cget('cursor')
                except Exception:
                    orig_cursor = None

                def _on_enter(e):
                    # Prefer changing bootstyle for themed button; fall back to bg change.
                    try:
                        if orig_bootstyle is not None:
                            # use a stronger danger style if available for a clearer hover
                            try:
                                del_btn.configure(bootstyle='danger')
                            except Exception:
                                try:
                                    del_btn.configure(bootstyle='danger-outline')
                                except Exception:
                                    pass
                            # amplify the visual cue with a raised relief and thicker border
                            try:
                                del_btn.configure(relief='raised', bd=3)
                            except Exception:
                                pass
                            try:
                                del_btn.configure(cursor='hand2')
                            except Exception:
                                pass
                            return
                    except Exception:
                        pass
                    try:
                        if orig_bg is not None:
                            # stronger lighter red for a more noticeable hover
                            hover_bg = '#ff6666'
                            try:
                                del_btn.configure(bg=hover_bg, activebackground=hover_bg)
                            except Exception:
                                del_btn.configure(bg=hover_bg)
                            try:
                                # adjust foreground if it improves contrast
                                if orig_fg is not None:
                                    del_btn.configure(fg=orig_fg)
                            except Exception:
                                pass
                            try:
                                del_btn.configure(relief='raised', bd=3)
                            except Exception:
                                pass
                            try:
                                del_btn.configure(cursor='hand2')
                            except Exception:
                                pass
                    except Exception:
                        pass

                def _on_leave(e):
                    try:
                        if orig_bootstyle is not None:
                            del_btn.configure(bootstyle=orig_bootstyle)
                            return
                    except Exception:
                        pass
                    try:
                        if orig_bg is not None:
                            # restore original colors and visuals
                            try:
                                del_btn.configure(bg=orig_bg, activebackground=orig_bg)
                            except Exception:
                                del_btn.configure(bg=orig_bg)
                        try:
                            # restore foreground
                            if orig_fg is not None:
                                try:
                                    del_btn.configure(fg=orig_fg)
                                except Exception:
                                    try:
                                        del_btn.configure(foreground=orig_fg)
                                    except Exception:
                                        pass
                        except Exception:
                            pass
                        try:
                            # restore relief/border
                            if orig_relief is not None:
                                del_btn.configure(relief=orig_relief)
                        except Exception:
                            pass
                        try:
                            if orig_bd is not None:
                                del_btn.configure(bd=orig_bd)
                        except Exception:
                            pass
                        try:
                            # restore cursor
                            if orig_cursor is not None:
                                del_btn.configure(cursor=orig_cursor)
                            else:
                                del_btn.configure(cursor='')
                        except Exception:
                            pass
                    except Exception:
                        pass

                # bind events (works for both tk and ttk widgets)
                try:
                    del_btn.bind('<Enter>', _on_enter)
                    del_btn.bind('<Leave>', _on_leave)
                except Exception:
                    pass
            except Exception:
                # non-fatal: ignore hover wiring errors
                pass
        except Exception:
            # If anything unexpected fails, silently continue (UI still usable)
            pass

        # Select Image button + placeholder label (per-rule pos var)
        sel_frame = ttk.Frame(rule_box)
        sel_frame.grid(row=1, column=0, columnspan=3, sticky='w', pady=(8,6))
        # per-rule StringVar used to show the currently-selected image position to the right of the button
        pos_var = tk.StringVar(value='')

        # Create a gray Select Image button that opens a dialog showing images
        try:
            if TTB_AVAILABLE:
                try:
                    select_btn = tb.Button(sel_frame, text='Select Image', bootstyle='secondary', command=lambda: None)
                except Exception:
                    select_btn = tk.Button(sel_frame, text='Select Image', bg='#6c757d', fg='white', bd=0, padx=6, pady=4, command=lambda: None)
            else:
                select_btn = tk.Button(sel_frame, text='Select Image', bg='#6c757d', fg='white', bd=0, padx=6, pady=4, command=lambda: None)
            select_btn.pack(side='left')
            # position label placed to the right of the button
            try:
                pos_lbl = ttk.Label(sel_frame, textvariable=pos_var, font=('Segoe UI', 10))
                pos_lbl.pack(side='left', padx=(8,0))
            except Exception:
                pass
        except Exception:
            try:
                select_btn = ttk.Button(sel_frame, text='Select Image', command=lambda: None)
                select_btn.pack(side='left')
                try:
                    pos_lbl = ttk.Label(sel_frame, textvariable=pos_var, font=('Segoe UI', 10))
                    pos_lbl.pack(side='left', padx=(8,0))
                except Exception:
                    pass
            except Exception:
                select_btn = None

        # wire the behavior: when an image is chosen call the per-rule handler
        try:
            if select_btn is not None:
                try:
                    select_btn.configure(command=lambda pv=pos_var, o=rule_box_outer: self.open_source_images_dialog(on_select=lambda p, pos, pv=pv, o=o: self._on_rule_image_selected(o, p, pos, pv)))
                except Exception:
                    try:
                        select_btn.config(command=lambda pv=pos_var, o=rule_box_outer: self.open_source_images_dialog(on_select=lambda p, pos, pv=pv, o=o: self._on_rule_image_selected(o, p, pos, pv)))
                    except Exception:
                        pass
        except Exception:
            pass

        # Aspect Ratio radios
        ar_label = ttk.Label(rule_box, text='Aspect Ratio:', font=('Segoe UI', 9, 'underline'))
        ar_label.grid(row=2, column=0, columnspan=3, sticky='w', pady=(6,2))
        self.aspect_var = tk.StringVar(value='none')
        # Arrange aspect ratio options into three rows
        # Row 1: 1:1, 3:4, 4:3, 4:5
        # Row 2: 16:9, 9:16, 1.91:1, 3:2
        # Row 3: custom, none
        ar_values = [
            ('1:1','1:1'), ('3:4','3:4'), ('4:3','4:3'), ('4:5','4:5'),
            ('16:9','16:9'), ('9:16','9:16'), ('1.91:1','1.91:1'), ('3:2','3:2'),
            ('Custom','custom'), ('None','none')
        ]
        ar_frame = ttk.Frame(rule_box)
        ar_frame.grid(row=3, column=0, columnspan=3, sticky='w')
        # Ensure four aligned columns so radios stack vertically in each column
        for col in range(4):
            ar_frame.grid_columnconfigure(col, weight=1, minsize=65)
        for i, (label, val) in enumerate(ar_values):
            # place first four in row 0, next four in row 1, remaining in row 2
            if i < 4:
                r = 0
                c = i
            elif i < 8:
                r = 1
                c = i - 4
            else:
                r = 2
                c = i - 8
            rb = ttk.Radiobutton(ar_frame, text=label, value=val, variable=self.aspect_var)
            # left-align each radio and use consistent padding so columns line up
            rb.grid(row=r, column=c, sticky='w', padx=2, pady=1)

        # Custom aspect ratio input field
        custom_ar_frame = ttk.Frame(rule_box)
        custom_ar_frame.grid(row=4, column=0, columnspan=3, sticky='w', pady=(4,0))
        ttk.Label(custom_ar_frame, text='').grid(row=0, column=0, sticky='w')
        self.custom_aspect_width_var = tk.StringVar(value='')
        custom_width_entry = ttk.Entry(custom_ar_frame, textvariable=self.custom_aspect_width_var, width=6)
        custom_width_entry.grid(row=0, column=1, sticky='w', padx=(4,0))
        ttk.Label(custom_ar_frame, text=':').grid(row=0, column=2, sticky='w', padx=(2,2))
        self.custom_aspect_height_var = tk.StringVar(value='')
        custom_height_entry = ttk.Entry(custom_ar_frame, textvariable=self.custom_aspect_height_var, width=6)
        custom_height_entry.grid(row=0, column=3, sticky='w', padx=(0,4))
        ttk.Label(custom_ar_frame, text='(e.g., 16:9)', font=('Segoe UI', 8), foreground='#666').grid(row=0, column=4, sticky='w', padx=(4,0))
        # Initially hide custom input
        custom_ar_frame.grid_remove()

        # Wire up aspect_var trace to snap crop to aspect when changed
        def _aspect_trace(*args):
            try:
                val = self.aspect_var.get()
                if val == 'custom':
                    # Show custom input field
                    custom_ar_frame.grid()
                    # Apply custom aspect if there's a value
                    custom_width = self.custom_aspect_width_var.get().strip()
                    custom_height = self.custom_aspect_height_var.get().strip()
                    if custom_width and custom_height:
                        try:
                            # Validate that both are numbers
                            float(custom_width)
                            float(custom_height)
                            custom_val = f"{custom_width}:{custom_height}"
                            self._apply_aspect_to_crop(custom_val)
                            self._draw_crop_rectangle()
                        except ValueError:
                            pass
                else:
                    # Hide custom input field
                    custom_ar_frame.grid_remove()
                    if val and val != 'none':
                        self._apply_aspect_to_crop(val)
                        self._draw_crop_rectangle()
            except Exception:
                pass

        # Also trace custom aspect input
        def _custom_aspect_trace(*args):
            try:
                if self.aspect_var.get() == 'custom':
                    custom_width = self.custom_aspect_width_var.get().strip()
                    custom_height = self.custom_aspect_height_var.get().strip()
                    if custom_width and custom_height:
                        try:
                            # Validate that both are numbers
                            float(custom_width)
                            float(custom_height)
                            custom_val = f"{custom_width}:{custom_height}"
                            self._apply_aspect_to_crop(custom_val)
                            self._draw_crop_rectangle()
                        except ValueError:
                            pass
            except Exception:
                pass

        try:
            self.aspect_var.trace_add('write', _aspect_trace)
            self.custom_aspect_width_var.trace_add('write', _custom_aspect_trace)
            self.custom_aspect_height_var.trace_add('write', _custom_aspect_trace)
        except Exception:
            try:
                self.aspect_var.trace('w', _aspect_trace)
                self.custom_aspect_width_var.trace('w', _custom_aspect_trace)
                self.custom_aspect_height_var.trace('w', _custom_aspect_trace)
            except Exception:
                pass

        # Crop Area inputs (X1, Y1, X2, Y2)
        ca_label = ttk.Label(rule_box, text='Crop Area (in Pixels):', font=('Segoe UI', 9, 'underline'))
        ca_label.grid(row=5, column=0, columnspan=3, sticky='w', pady=(8,4))

        ca_frame = ttk.Frame(rule_box)
        ca_frame.grid(row=6, column=0, columnspan=3, sticky='w')
        ttk.Label(ca_frame, text='X1').grid(row=0, column=0)
        self.x1_var = tk.StringVar(value='0')
        ttk.Entry(ca_frame, textvariable=self.x1_var, width=6).grid(row=0, column=1, padx=(4,12))
        ttk.Label(ca_frame, text='Y1').grid(row=0, column=2)
        self.y1_var = tk.StringVar(value='0')
        ttk.Entry(ca_frame, textvariable=self.y1_var, width=6).grid(row=0, column=3, padx=(4,12))

        ttk.Label(ca_frame, text='X2').grid(row=1, column=0, pady=(6,0))
        self.x2_var = tk.StringVar(value='0')
        ttk.Entry(ca_frame, textvariable=self.x2_var, width=6).grid(row=1, column=1, padx=(4,12), pady=(6,0))
        ttk.Label(ca_frame, text='Y2').grid(row=1, column=2, pady=(6,0))
        self.y2_var = tk.StringVar(value='0')
        ttk.Entry(ca_frame, textvariable=self.y2_var, width=6).grid(row=1, column=3, padx=(4,12), pady=(6,0))

        # Add trace callbacks to crop area inputs to update the crop box in real-time
        self._updating_crop_inputs = False  # Flag to prevent infinite loops

        def _on_crop_input_changed(*args):
            if self._updating_crop_inputs:
                return
            try:
                self._updating_crop_inputs = True

                # Get current aspect ratio
                aspect_val = self.aspect_var.get()

                # If custom aspect, construct the aspect string from width/height inputs
                if aspect_val == 'custom':
                    width_str = self.custom_aspect_width_var.get().strip()
                    height_str = self.custom_aspect_height_var.get().strip()
                    if width_str and height_str:
                        aspect_val = f"{width_str}:{height_str}"
                    else:
                        aspect_val = 'none'

                # Get the original image dimensions for boundary checking
                orig_w, orig_h = getattr(self, '_original_image_size', (1, 1))

                # Get the values from input boxes (in original coordinates)
                try:
                    x1_orig = int(float(self.x1_var.get()))
                    y1_orig = int(float(self.y1_var.get()))
                    x2_orig = int(float(self.x2_var.get()))
                    y2_orig = int(float(self.y2_var.get()))
                except (ValueError, TypeError):
                    self._updating_crop_inputs = False
                    return

                # ALWAYS clamp coordinates to image boundaries FIRST
                x1_orig = max(0, min(x1_orig, orig_w - 1))
                y1_orig = max(0, min(y1_orig, orig_h - 1))
                x2_orig = max(1, min(x2_orig, orig_w))
                y2_orig = max(1, min(y2_orig, orig_h))

                # Ensure x1 < x2 and y1 < y2
                if x1_orig >= x2_orig:
                    x2_orig = min(x1_orig + 1, orig_w)
                if y1_orig >= y2_orig:
                    y2_orig = min(y1_orig + 1, orig_h)

                # If aspect ratio is set and not 'none', adjust coordinates to maintain aspect
                if aspect_val and aspect_val != 'none':
                    asp = self._parse_aspect(aspect_val)
                    if asp is not None:
                        aw, ah = asp
                        r = aw / ah

                        # Calculate current dimensions
                        cur_w = x2_orig - x1_orig
                        cur_h = y2_orig - y1_orig

                        # Adjust to maintain aspect ratio - prefer width
                        target_h = int(round(cur_w / r))
                        if target_h != cur_h:
                            y2_orig = y1_orig + target_h
                            # Clamp y2 to image boundary
                            if y2_orig > orig_h:
                                y2_orig = orig_h
                                # Recalculate width to maintain aspect
                                cur_h = y2_orig - y1_orig
                                target_w = int(round(cur_h * r))
                                x2_orig = x1_orig + target_w
                                # Clamp x2 to image boundary
                                if x2_orig > orig_w:
                                    x2_orig = orig_w

                # Final boundary check after aspect ratio adjustment (ALWAYS applied)
                x1_orig = max(0, min(x1_orig, orig_w - 1))
                y1_orig = max(0, min(y1_orig, orig_h - 1))
                x2_orig = max(x1_orig + 1, min(x2_orig, orig_w))
                y2_orig = max(y1_orig + 1, min(y2_orig, orig_h))

                # Update input boxes with clamped values to show user the corrected values
                self.x1_var.set(str(x1_orig))
                self.y1_var.set(str(y1_orig))
                self.x2_var.set(str(x2_orig))
                self.y2_var.set(str(y2_orig))

                # Convert from original coordinates to preview coordinates
                x1_prev, y1_prev = self._original_to_preview_coords(x1_orig, y1_orig)
                x2_prev, y2_prev = self._original_to_preview_coords(x2_orig, y2_orig)

                # Update the crop rectangle in preview space
                if x1_prev is not None and y1_prev is not None and x2_prev is not None and y2_prev is not None:
                    self._crop_rect = [int(x1_prev), int(y1_prev), int(x2_prev), int(y2_prev)]
                    self._draw_crop_rectangle()
            except Exception:
                pass
            finally:
                self._updating_crop_inputs = False

        # Register trace callbacks
        try:
            self.x1_var.trace_add('write', _on_crop_input_changed)
            self.y1_var.trace_add('write', _on_crop_input_changed)
            self.x2_var.trace_add('write', _on_crop_input_changed)
            self.y2_var.trace_add('write', _on_crop_input_changed)
        except Exception:
            try:
                self.x1_var.trace('w', _on_crop_input_changed)
                self.y1_var.trace('w', _on_crop_input_changed)
                self.x2_var.trace('w', _on_crop_input_changed)
                self.y2_var.trace('w', _on_crop_input_changed)
            except Exception:
                pass

        # Compression level
        comp_label = ttk.Label(rule_box, text='Compression Level:', font=('Segoe UI', 9, 'underline'))
        comp_label.grid(row=7, column=0, columnspan=3, sticky='w', pady=(12,2))
        self.comp_var = tk.StringVar(value='0')
        # Place the entry and helper into a compact horizontal frame so the helper sits
        # immediately to the right of the entry with minimal gap.
        comp_frame = ttk.Frame(rule_box)
        comp_frame.grid(row=8, column=0, columnspan=3, sticky='w', pady=(0,2))
        # use grid to eliminate any pack-introduced gaps and keep widgets tightly aligned
        comp_entry = ttk.Entry(comp_frame, textvariable=self.comp_var, width=8)
        comp_entry.grid(row=0, column=0, sticky='w')
        helper_lbl = ttk.Label(comp_frame, text='% (0 = no compression)', font=('Segoe UI', 9), foreground='#444')
        helper_lbl.grid(row=0, column=1, sticky='w', padx=(2,0))
        comp_frame.grid_columnconfigure(0, minsize=0)

        # Checkbox: Apply this rule to the remaining images
        self.apply_to_remaining_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(rule_box, text='Apply this Rule to the Remaining Images', variable=self.apply_to_remaining_var).grid(row=9, column=0, columnspan=3, sticky='w', pady=(6,0))

        # Create New Cropping Rule button moved to the bottom action toolbar
        # (widget created in the toolbar below so it aligns with global actions)
        # The left panel now only contains the scrollable rules container above.

        # Note: we keep an invisible listbox widget to preserve existing helper methods
        # that expect self.listbox to exist; it is not shown in the UI.
        self.listbox = tk.Listbox(left, height=1)

        # register the initial rule box so add/remove and renumbering work
        try:
            # store variables for the initial rule so it's consistent with dynamically added rules
            self.rule_boxes.append({'outer': rule_box_outer, 'label': rule_label, 'vars': {
                'aspect': getattr(self, 'aspect_var', None),
                'x1': getattr(self, 'x1_var', None),
                'y1': getattr(self, 'y1_var', None),
                'x2': getattr(self, 'x2_var', None),
                'y2': getattr(self, 'y2_var', None),
                'comp': getattr(self, 'comp_var', None),
                'apply': getattr(self, 'apply_to_remaining_var', None),
                'pos': pos_var,
                'custom_aspect_width': getattr(self, 'custom_aspect_width_var', None),
                'custom_aspect_height': getattr(self, 'custom_aspect_height_var', None),
            }, 'img': None})
            # mark the initial rule active so the user sees the tint immediately
            try:
                self.set_active_rule(rule_box_outer)
            except Exception:
                pass
        except Exception:
            pass

        # Right: canvas placeholder
        right = ttk.Frame(self)
        right.grid(row=0, column=1, sticky='nsew')
        # Reserve row 0 for the header label and make row 1 expandable for the canvas
        right.grid_rowconfigure(0, weight=0)
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        # Header for the preview area
        try:
            header_lbl = ttk.Label(right, text='Image Preview', font=('Segoe UI', 16, 'bold'), anchor='center', justify='center')
            # make the label expand horizontally and center its text
            header_lbl.grid(row=0, column=0, sticky='ew', padx=8, pady=(8, 4))
        except Exception:
            pass

        self.canvas = tk.Canvas(right, bg='white', bd=0, highlightthickness=0)
        self.canvas.grid(row=1, column=0, sticky='nsew', padx=8, pady=8)
        # Small status backing var for preview diagnostics. The visible label
        # was removed per user request (don't show filenames like 'Showing: 1.jpg'),
        # but we keep the StringVar so existing set() calls elsewhere do not fail.
        try:
            self.preview_status_var = tk.StringVar(value='')
            # NOTE: intentionally do not create or grid a Label for preview_status_var
            # to avoid displaying the filename text in the UI.
        except Exception:
            self.preview_status_var = tk.StringVar(value='')
        # Crop rectangle state and mouse bindings
        # _crop_rect: [x1,y1,x2,y2] in canvas/preview coordinates (initially None until an image is shown)
        self._crop_rect = None
        # map of handle canvas id -> handle name ('nw','ne','sw','se')
        self._crop_handles = {}
        # drag data used while the user is interacting with the rectangle
        self._crop_drag = {'action': None, 'handle': None, 'start_x': 0, 'start_y': 0, 'orig_rect': None}
        # preview image bookkeeping to avoid GC and provide bounds before image loaded
        self._preview_image_ref = None
        self._preview_image_pos = (0, 0)
        self._preview_image_size = (0, 0)
        # Store original image dimensions (before scaling for preview) for accurate crop coordinate display
        self._original_image_size = (1, 1)

        # Bind mouse events for interacting with the crop rectangle
        try:
            self.canvas.bind('<Button-1>', self._on_canvas_mouse_down)
            self.canvas.bind('<B1-Motion>', self._on_canvas_mouse_move)
            self.canvas.bind('<ButtonRelease-1>', self._on_canvas_mouse_up)
            # when the canvas is resized we keep the rectangle as-is (image will be redrawn on new load)
            self.canvas.bind('<Configure>', lambda e: None)
        except Exception:
            pass

        # Layout weights for the overall frame
        self.grid_rowconfigure(0, weight=1)
        # reserve a row for bottom action buttons (Save & Close)
        try:
            self.grid_rowconfigure(1, weight=0)
        except Exception:
            pass
        self.grid_columnconfigure(1, weight=1)

        # Bottom action row: toolbar spanning the window. Create button on the
        # left, Save & Close on the right so both sit on the same horizontal line.
        try:
            action_row = ttk.Frame(self)
            action_row.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(6,8), padx=8)
            # Create button placed inside a fixed-width left container so it visually
            # matches the left panel width (left column minsize=240). We keep the
            # button in the toolbar per your request but make it span the left-panel width.
            try:
                # slightly wider than the left panel's reserved width so the
                # button visually aligns with the rule box (tweak as needed)
                left_width_container = ttk.Frame(action_row, width=268)
                # prevent the frame from shrinking to its child so the width is honored
                try:
                    left_width_container.pack_propagate(False)
                except Exception:
                    pass
                left_width_container.pack(side='left', fill='y')

                # create the button and pack it to fill the container horizontally
                # On macOS, tk.Button doesn't respect bg/fg colors well, so use ttkbootstrap if available
                if sys.platform == 'darwin' and TTB_AVAILABLE:
                    try:
                        create_btn = tb.Button(left_width_container, text='+ Create New Cropping Rule', bootstyle='primary', command=self.add_rule)
                    except Exception:
                        create_btn = tk.Button(left_width_container, text='+ Create New Cropping Rule', bg='#007bff', fg='white', bd=0, padx=10, pady=6, command=self.add_rule)
                else:
                    create_btn = tk.Button(left_width_container, text='+ Create New Cropping Rule', bg='#007bff', fg='white', bd=0, padx=10, pady=6, command=self.add_rule)
                # add a small inset so the button lines up with the rule box content
                create_btn.pack(fill='x', expand=True, padx=8, pady=(0,0))
                # keep hover effect
                try:
                    orig_bg = create_btn.cget('bg')
                except Exception:
                    orig_bg = None

                def _create_on_enter(e):
                    try:
                        hover_blue = '#339bff'
                        try:
                            create_btn.configure(bg=hover_blue, activebackground=hover_blue)
                        except Exception:
                            create_btn.configure(bg=hover_blue)
                    except Exception:
                        pass

                def _create_on_leave(e):
                    try:
                        if orig_bg is not None:
                            try:
                                create_btn.configure(bg=orig_bg, activebackground=orig_bg)
                            except Exception:
                                create_btn.configure(bg=orig_bg)
                    except Exception:
                        pass

                try:
                    create_btn.bind('<Enter>', _create_on_enter)
                    create_btn.bind('<Leave>', _create_on_leave)
                except Exception:
                    pass
            except Exception:
                pass

            # Save & Close on the right
            try:
                # Pack Save & Close first (rightmost), then pack Duplicate Profile to its left
                save_close_btn = ttk.Button(action_row, text='Save & Close', command=self._on_save_and_close)
                save_close_btn.pack(side='right')
                try:
                    # Duplicate Profile button placed immediately left of Save & Close with a small gap
                    dup_btn = ttk.Button(action_row, text='Duplicate Profile', command=self._on_duplicate_profile)
                    dup_btn.pack(side='right', padx=(0,8))
                    try:
                        # Delete Profile button placed immediately left of Duplicate Profile
                        # Prefer ttkbootstrap styled button (red/danger). Fall back to a plain tk.Button
                        if TTB_AVAILABLE:
                            try:
                                del_btn = tb.Button(action_row, text='Delete Profile', bootstyle='danger', command=self.on_delete)
                            except Exception:
                                del_btn = tk.Button(action_row, text='Delete Profile', bg='#ff4d6d', fg='white', bd=0, padx=6, pady=4, command=self.on_delete)
                        else:
                            del_btn = tk.Button(action_row, text='Delete Profile', bg='#ff4d6d', fg='white', bd=0, padx=6, pady=4, command=self.on_delete)
                        del_btn.pack(side='right', padx=(0,8))
                        # Add hover behavior for the bottom Delete Profile button as well
                        try:
                            del_orig_bootstyle = None
                            try:
                                del_orig_bootstyle = del_btn.cget('bootstyle')
                            except Exception:
                                del_orig_bootstyle = None
                            del_orig_bg = None
                            try:
                                del_orig_bg = del_btn.cget('bg')
                            except Exception:
                                del_orig_bg = None

                            def _del_on_enter(e):
                                try:
                                    if del_orig_bootstyle is not None:
                                        try:
                                            del_btn.configure(bootstyle='danger')
                                        except Exception:
                                            try:
                                                del_btn.configure(bootstyle='danger-outline')
                                            except Exception:
                                                pass
                                        try:
                                            del_btn.configure(relief='raised', bd=3)
                                        except Exception:
                                            pass
                                        try:
                                            del_btn.configure(cursor='hand2')
                                        except Exception:
                                            pass
                                        return
                                except Exception:
                                    pass
                                try:
                                    if del_orig_bg is not None:
                                        hover_bg = '#ff6666'
                                        try:
                                            del_btn.configure(bg=hover_bg, activebackground=hover_bg)
                                        except Exception:
                                            try:
                                                del_btn.configure(bg=hover_bg)
                                            except Exception:
                                                pass
                                        try:
                                            del_btn.configure(cursor='hand2')
                                        except Exception:
                                            pass
                                except Exception:
                                    pass

                            def _del_on_leave(e):
                                try:
                                    if del_orig_bootstyle is not None:
                                        try:
                                            del_btn.configure(bootstyle=del_orig_bootstyle)
                                        except Exception:
                                            pass
                                        return
                                except Exception:
                                    pass
                                try:
                                    if del_orig_bg is not None:
                                        try:
                                            del_btn.configure(bg=del_orig_bg, activebackground=del_orig_bg)
                                        except Exception:
                                            try:
                                                del_btn.configure(bg=del_orig_bg)
                                            except Exception:
                                                pass
                                        try:
                                            del_btn.configure(cursor='')
                                        except Exception:
                                            pass
                                except Exception:
                                    pass

                            try:
                                del_btn.bind('<Enter>', _del_on_enter)
                                del_btn.bind('<Leave>', _del_on_leave)
                            except Exception:
                                pass
                        except Exception:
                            pass
                    except Exception:
                        pass
                except Exception:
                    pass
            except Exception:
                pass
        except Exception:
            pass

        self.refresh_profiles()

        if initial_profile_path:
            # If a profile path is provided, attempt to load it and select in list
            if os.path.exists(initial_profile_path):
                p = load_profile_by_path(initial_profile_path)
                if p:
                    self.current_profile = p
                    self.current_profile_path = initial_profile_path
                    # Populate the name field from the explicitly provided profile
                    self.name_var.set(p.get('profile_name', os.path.splitext(os.path.basename(initial_profile_path))[0]))
                    # populate UI rule boxes from loaded profile
                    try:
                        self._populate_rules_from_profile(p)
                    except Exception:
                        pass
                    self.refresh_profiles(select=os.path.splitext(os.path.basename(initial_profile_path))[0])

        # NOTE: Do NOT auto-select an existing profile when the editor is opened
        # without an explicit `initial_profile_path`. Previously we auto-selected
        # the first profile (index 0) and called `on_profile_select()`, which
        # loaded that profile and overwrote the entry holding the default name
        # ('New Cropping Rule'). The desired behavior is to keep the default
        # name unless the user explicitly loads or duplicates a profile.
        # Therefore we intentionally avoid selecting any profile here.

    def on_delete(self):
        """Delete the profile selected in the Edit Profiles list.

        Behavior:
        - Prefer the currently-selected name in self.listbox (Edit Profiles selection).
        - If nothing is selected, fall back to the currently-loaded profile file (self.current_profile_path).
        - Confirm with the user before deleting. Attempt to remove <name>.profile from CONFIG_FOLDER.
        - On success, clear current_profile if it referred to the deleted file and refresh the list.
        """
        try:
            # Attempt to get selected profile name from the (possibly hidden) listbox
            sel = None
            try:
                sel = self.listbox.curselection()
            except Exception:
                sel = None

            name = None
            if sel:
                try:
                    name = self.listbox.get(sel[0])
                except Exception:
                    name = None

            # Fallback: use currently-loaded profile path's basename (without extension)
            if not name:
                try:
                    if self.current_profile_path:
                        name = os.path.splitext(os.path.basename(self.current_profile_path))[0]
                except Exception:
                    name = None

            if not name:
                try:
                    messagebox.showerror('Delete Failed', 'No profile selected to delete.', parent=self.master)
                except Exception:
                    pass
                return

            filename = f"{name}.profile"
            path = os.path.join(CONFIG_FOLDER, filename)

            if not os.path.exists(path):
                try:
                    messagebox.showerror('Delete Failed', f'Profile file not found: {filename}', parent=self.master)
                except Exception:
                    pass
                return

            try:
                confirm = messagebox.askyesno('Confirm Delete', f"Are you sure you want to delete the profile '{filename}'? This cannot be undone.", parent=self.master)
            except Exception:
                # If confirmation dialog fails for some reason, abort to be safe
                return

            if not confirm:
                return

            try:
                os.remove(path)
            except Exception:
                try:
                    messagebox.showerror('Delete Failed', f'Could not delete profile: {name}', parent=self.master)
                except Exception:
                    pass
                return

            # If the deleted file was the currently-loaded profile, clear it
            try:
                if self.current_profile_path and os.path.normcase(os.path.abspath(self.current_profile_path)) == os.path.normcase(os.path.abspath(path)):
                    self.current_profile = {'profile_name': '', 'rules': []}
                    self.current_profile_path = None
                    try:
                        self.name_var.set('')
                    except Exception:
                        pass
            except Exception:
                pass

            # Refresh the profiles list and UI
            try:
                self.refresh_profiles()
            except Exception:
                pass

            # Close the window after successful deletion
            try:
                self.master.destroy()
            except Exception:
                try:
                    self.master.quit()
                except Exception:
                    pass
        except Exception:
            pass

    def _on_save_and_close(self):
        """Gather the UI data for all rules, save as JSON file named <ProfileName>.profile
        into the local CONFIG_FOLDER, refresh list, and close the window."""
        try:
            name = ''
            try:
                name = (self.name_var.get() or '').strip()
            except Exception:
                name = ''
            if not name:
                try:
                    messagebox.showerror('Save Failed', 'Profile name cannot be empty.', parent=self.master)
                except Exception:
                    pass
                return

            # sanitize filename
            safe = re.sub(r'[^A-Za-z0-9_. \-]', '_', name).strip()
            if not safe:
                safe = 'profile'
            filename = f"{safe}.profile"
            path = os.path.join(CONFIG_FOLDER, filename)

            # if the file already exists, confirm overwrite
            try:
                if os.path.exists(path):
                    try:
                        # askyesno returns True to overwrite
                        if not messagebox.askyesno('Overwrite Profile', f"A profile named '{name}' already exists. Overwrite?", parent=self.master):
                            return
                    except Exception:
                        # If dialog fails for some reason, proceed with caution (skip saving)
                        return
            except Exception:
                pass

            # build rules list from self.rule_boxes in the requested shape
            # Each saved rule should be an object like:
            # { "position": 1, "apply_to_all_remaining": false, "crop": {"x1":0,"y1":0,"x2":800,"y2":600}, "compression": 10, "aspect_ratio": "none" }
            rules = []
            for rb in self.rule_boxes:
                try:
                    vars_map = rb.get('vars', {}) or {}

                    def _get_var(v):
                        try:
                            if v is None:
                                return None
                            return v.get()
                        except Exception:
                            return v

                    # helper to coerce numeric-ish values to int
                    def _intish(v):
                        try:
                            return int(float(v))
                        except Exception:
                            return 0

                    asp = _get_var(vars_map.get('aspect'))
                    # If aspect is 'custom', combine width and height into aspect ratio
                    if asp == 'custom':
                        custom_width = _get_var(vars_map.get('custom_aspect_width'))
                        custom_height = _get_var(vars_map.get('custom_aspect_height'))
                        if custom_width and custom_height:
                            width_str = str(custom_width).strip()
                            height_str = str(custom_height).strip()
                            if width_str and height_str:
                                asp = f"{width_str}:{height_str}"
                            else:
                                asp = 'none'
                        else:
                            asp = 'none'
                    x1 = _intish(_get_var(vars_map.get('x1')))
                    y1 = _intish(_get_var(vars_map.get('y1')))
                    x2 = _intish(_get_var(vars_map.get('x2')))
                    y2 = _intish(_get_var(vars_map.get('y2')))
                    comp = _intish(_get_var(vars_map.get('comp')))

                    # position (optional) - some rule boxes keep a 'pos' variable
                    pos = None
                    try:
                        pv = vars_map.get('pos')
                        if pv is not None:
                            try:
                                pv_val = _get_var(pv)
                                if pv_val is not None:
                                    s = str(pv_val).strip()
                                    if s:
                                        # common UI format is 'Position N' -> extract digits
                                        m = re.search(r"(\d+)", s)
                                        if m:
                                            pos = int(m.group(1))
                                        else:
                                            # fallback: try numeric conversion
                                            try:
                                                pos = int(float(s))
                                            except Exception:
                                                pos = None
                            except Exception:
                                pos = None
                    except Exception:
                        pos = None

                    # apply-to-all flag (map existing 'apply' var to 'apply_to_all_remaining')
                    apply_to_all = False
                    try:
                        av = vars_map.get('apply')
                        if av is not None:
                            try:
                                apply_to_all = bool(av.get())
                            except Exception:
                                apply_to_all = bool(av)
                    except Exception:
                        apply_to_all = False

                    rule_obj = {
                        'crop': {'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2},
                        'compression': comp,
                        'aspect_ratio': asp if asp is not None else 'none',
                        'apply_to_all_remaining': bool(apply_to_all)
                    }

                    if pos is not None and isinstance(pos, int) and pos > 0:
                        rule_obj['position'] = pos

                    rules.append(rule_obj)
                except Exception:
                    # skip problematic rule but continue
                    pass

            # keep profile_name for convenience, but image_wizard only requires top-level 'rules'
            profile = {'profile_name': name, 'rules': rules}

            ok = save_profile_to_path(profile, path)
            if not ok:
                try:
                    messagebox.showerror('Save Failed', f'Could not save profile: {name}', parent=self.master)
                except Exception:
                    pass
                return

            # update current profile and refresh list
            try:
                self.current_profile = profile
                self.current_profile_path = path
                # refresh and select the saved profile
                self.refresh_profiles(select=os.path.splitext(filename)[0])
            except Exception:
                pass

            # close the window (if it's a Toplevel or root)
            try:
                self.master.destroy()
            except Exception:
                try:
                    self.master.quit()
                except Exception:
                    pass
            return
        except Exception:
            try:
                messagebox.showerror('Save Failed', 'An unexpected error occurred while saving the profile.', parent=self.master)
            except Exception:
                pass
            return

    def _on_duplicate_profile(self):
        """Duplicate the currently active or selected profile into a new profile file
        in the local CONFIG_FOLDER. Prompt the user for the duplicated profile name,
        sanitize it, handle collisions (ask to overwrite or re-enter), then save."""
        try:
            src = None
            # Prefer in-memory current_profile if it has a name and rules
            try:
                if self.current_profile and self.current_profile.get('profile_name'):
                    src = self.current_profile
            except Exception:
                src = None

            # If no in-memory profile, try listbox selection
            if not src:
                try:
                    sel = None
                    try:
                        sel = self.listbox.curselection()
                    except Exception:
                        sel = None
                    if sel:
                        name = self.listbox.get(sel[0])
                        ppath = os.path.join(CONFIG_FOLDER, f"{name}.profile")
                        sp = load_profile_by_path(ppath)
                        if sp:
                            src = sp
                except Exception:
                    src = None

            if not src:
                try:
                    messagebox.showerror('Duplicate Failed', 'No profile selected to duplicate.', parent=self.master)
                except Exception:
                    pass
                return

            orig_name = src.get('profile_name') or 'profile'
            # propose default duplicate name
            default = f"{orig_name} Copy"

            # Prompt user for new name (allow cancel)
            try:
                user_input = simpledialog.askstring('Duplicate Profile', 'Enter name for duplicated profile:', initialvalue=default, parent=self.master)
            except Exception:
                user_input = None

            if user_input is None:
                # user cancelled
                return

            user_input = (user_input or '').strip()
            if not user_input:
                try:
                    messagebox.showerror('Invalid Name', 'Profile name cannot be empty.', parent=self.master)
                except Exception:
                    pass
                return

            # sanitize filename
            safe = re.sub(r'[^A-Za-z0-9_. \-]', '_', user_input).strip()
            if not safe:
                safe = 'profile'

            candidate = safe
            # If collision occurs, offer overwrite or let user re-enter name
            while True:
                target_path = os.path.join(CONFIG_FOLDER, f"{candidate}.profile")
                if not os.path.exists(target_path):
                    break
                # file exists: ask to overwrite or provide a new name
                try:
                    resp = messagebox.askyesno('Overwrite Profile', f"A profile named '{candidate}' already exists. Overwrite?", parent=self.master)
                except Exception:
                    resp = False
                if resp:
                    # will overwrite
                    break
                # ask for a new name
                try:
                    user_input = simpledialog.askstring('Duplicate Profile', 'Choose a different name for the duplicated profile:', initialvalue=f"{safe} (copy)", parent=self.master)
                except Exception:
                    user_input = None
                if user_input is None:
                    # user cancelled
                    return
                user_input = (user_input or '').strip()
                if not user_input:
                    try:
                        messagebox.showerror('Invalid Name', 'Profile name cannot be empty.', parent=self.master)
                    except Exception:
                        pass
                    return
                safe = re.sub(r'[^A-Za-z0-9_. \-]', '_', user_input).strip() or 'profile'
                candidate = safe

            new_name = candidate

            # deep copy rules safely
            try:
                new_rules = json.loads(json.dumps(src.get('rules', [])))
            except Exception:
                new_rules = src.get('rules', []) if src.get('rules') is not None else []

            new_profile = {'profile_name': new_name, 'rules': new_rules}
            new_path = os.path.join(CONFIG_FOLDER, f"{new_name}.profile")
            ok = save_profile_to_path(new_profile, new_path)
            if not ok:
                try:
                    messagebox.showerror('Duplicate Failed', f'Could not save duplicated profile: {new_name}', parent=self.master)
                except Exception:
                    pass
                return

            try:
                self.current_profile = new_profile
                self.current_profile_path = new_path
                try:
                    self.name_var.set(new_name)
                except Exception:
                    pass
                self.refresh_profiles(select=os.path.splitext(f"{new_name}.profile")[0])
                try:
                    messagebox.showinfo('Profile Duplicated', f"Profile duplicated to '{new_name}'.", parent=self.master)
                except Exception:
                    pass
            except Exception:
                pass
            return
        except Exception:
            try:
                messagebox.showerror('Duplicate Failed', 'An unexpected error occurred while duplicating the profile.', parent=self.master)
            except Exception:
                pass
            return

    # Profile list helpers implemented as regular methods
    def refresh_profiles(self, select: str | None = None):
        try:
            names = []
            for f in os.listdir(CONFIG_FOLDER):
                if f.lower().endswith('.profile') and os.path.isfile(os.path.join(CONFIG_FOLDER, f)):
                    names.append(os.path.splitext(f)[0])
            names.sort()
            # update the visible listbox only if it exists
            try:
                self.listbox.delete(0, 'end')
                for n in names:
                    self.listbox.insert('end', n)
                if select and select in names:
                    idx = names.index(select)
                    self.listbox.select_clear(0, 'end')
                    self.listbox.select_set(idx)
                    self.listbox.see(idx)
            except Exception:
                # If listbox is not present (some layouts keep it hidden), ignore
                pass
        except Exception:
            pass

    def on_profile_select(self):
        try:
            sel = None
            try:
                sel = self.listbox.curselection()
            except Exception:
                sel = None
            if not sel:
                return
            name = self.listbox.get(sel[0])
            path = os.path.join(CONFIG_FOLDER, f"{name}.profile")
            p = load_profile_by_path(path)
            if p:
                self.current_profile = p
                self.current_profile_path = path
                try:
                    self.name_var.set(p.get('profile_name', name))
                except Exception:
                    pass
        except Exception:
            pass

    def _open_image_preview(self, path: str):
        """Load the image at `path` and display it centered and scaled in self.canvas.
        Keeps a reference to the PhotoImage in self._preview_image_ref to avoid GC.
        """
        try:
            # Diagnostics: if path doesn't exist, draw a helpful placeholder
            if not path or not os.path.exists(path):
                try:
                    print(f'[cropping_gui2] _open_image_preview: file not found: {path}')
                except Exception:
                    pass
                try:
                    self.canvas.delete('all')
                    w = max(1, self.canvas.winfo_width())
                    h = max(1, self.canvas.winfo_height())
                    txt = f'File not found:\n{os.path.basename(path) if path else "(no path)"}\n\nCheck configured source folder or select a different image.'
                    self.canvas.create_text(w//2, h//2, text=txt, fill='#666', font=('Segoe UI', 12), anchor='center', width=int(w*0.9))
                except Exception:
                    pass
                return

            if not PIL_AVAILABLE:
                try:
                    print('[cropping_gui2] _open_image_preview: Pillow not available')
                except Exception:
                    pass
                try:
                    self.canvas.delete('all')
                    w = max(1, self.canvas.winfo_width())
                    h = max(1, self.canvas.winfo_height())
                    txt = 'Image preview requires the Pillow library (PIL). Please install Pillow to enable previews.'
                    self.canvas.create_text(w//2, h//2, text=txt, fill='#666', font=('Segoe UI', 12), anchor='center', width=int(w*0.9))
                except Exception:
                    pass
                return

            # ensure canvas has up-to-date size
            try:
                self.canvas.update_idletasks()
                cw = max(1, self.canvas.winfo_width())
                ch = max(1, self.canvas.winfo_height())
            except Exception:
                cw, ch = 800, 600

            try:
                img = Image.open(path)
            except Exception as e:
                # failed to open image file despite existence
                try:
                    print(f'[cropping_gui2] _open_image_preview: Pillow failed to open {path}: {e}')
                except Exception:
                    pass
                try:
                    self.canvas.delete('all')
                    w = max(1, self.canvas.winfo_width())
                    h = max(1, self.canvas.winfo_height())
                    txt = f'Could not open image:\n{os.path.basename(path)}\n\n{str(e)}'
                    self.canvas.create_text(w//2, h//2, text=txt, fill='#666', font=('Segoe UI', 12), anchor='center', width=int(w*0.9))
                    try:
                        self.preview_status_var.set('Could not open image: ' + (os.path.basename(path) or '(no path)'))
                    except Exception:
                        pass
                except Exception:
                    pass
                return

            # compute resized size preserving aspect ratio to fit into canvas
            # Store original dimensions BEFORE any resizing
            original_iw, original_ih = img.size

            iw, ih = img.size
            if iw > cw or ih > ch:
                img_ratio = iw / ih
                canvas_ratio = cw / ch
                # Guard: canvas may be very small during initial layout. Ensure
                # computed new_w/new_h are at least 1 so Pillow.resize() never
                # receives non-positive dimensions which raise an exception.
                if img_ratio > canvas_ratio:
                    new_w = max(1, cw - 8)
                    new_h = max(1, round(new_w / img_ratio))
                else:
                    new_h = max(1, ch - 8)
                    new_w = max(1, round(new_h * img_ratio))
                if _RESAMPLE_FILTER is not None:
                    img = img.resize((new_w, new_h), _RESAMPLE_FILTER)
                else:
                    img = img.resize((new_w, new_h))

            photo = ImageTk.PhotoImage(img)
            # clear canvas and draw centered
            try:
                self.canvas.delete('all')
                # place image at top-left of the canvas per request
                self.canvas.create_image(0, 0, image=photo, anchor='nw')
                # remember the displayed image bounds so crop interactions can be clamped
                try:
                    # new_w/new_h exist when resizing was applied above; otherwise use original
                    iw, ih = photo.width(), photo.height()
                except Exception:
                    iw, ih = img.size if hasattr(img, 'size') else (0, 0)
                self._preview_image_pos = (0, 0)
                self._preview_image_size = (iw, ih)
                # Store original image dimensions for accurate crop coordinate display
                self._original_image_size = (original_iw, original_ih)
                # keep reference to avoid GC
                self._preview_image_ref = photo
                # initialize crop rectangle: prefer saved coords for active rule, otherwise use full preview
                # Note: _crop_rect is kept in PREVIEW space. Saved values are in ORIGINAL space and need conversion.
                try:
                    # Check if there are saved crop coordinates for the active rule (in original space)
                    saved_coords_original = None
                    try:
                        active_outer = getattr(self, '_active_rule_outer', None)
                        if active_outer is not None:
                            for rb in self.rule_boxes:
                                try:
                                    if rb.get('outer') is active_outer:
                                        rvars = rb.get('vars', {}) or {}
                                        # Try to read saved crop coords from the rule's vars (in original space)
                                        try:
                                            x1_val = rvars.get('x1')
                                            y1_val = rvars.get('y1')
                                            x2_val = rvars.get('x2')
                                            y2_val = rvars.get('y2')
                                            if x1_val and y1_val and x2_val and y2_val:
                                                x1_saved = int(x1_val.get() or 0)
                                                y1_saved = int(y1_val.get() or 0)
                                                x2_saved = int(x2_val.get() or 0)
                                                y2_saved = int(y2_val.get() or 0)
                                                # Only use saved coords if they appear valid (non-zero area)
                                                if x2_saved > x1_saved and y2_saved > y1_saved:
                                                    # Also verify they fit within the ORIGINAL image bounds
                                                    if x1_saved >= 0 and y1_saved >= 0 and x2_saved <= original_iw and y2_saved <= original_ih:
                                                        saved_coords_original = [x1_saved, y1_saved, x2_saved, y2_saved]
                                        except Exception:
                                            pass
                                        break
                                except Exception:
                                    pass
                    except Exception:
                        pass

                    # Convert saved coords from original space to preview space, or use full original dimensions
                    if saved_coords_original:
                        x1_prev, y1_prev = self._original_to_preview_coords(saved_coords_original[0], saved_coords_original[1])
                        x2_prev, y2_prev = self._original_to_preview_coords(saved_coords_original[2], saved_coords_original[3])
                        self._crop_rect = [x1_prev, y1_prev, x2_prev, y2_prev]
                    else:
                        # Default to full original image dimensions (converted to preview space)
                        # This ensures we always start with the correct full image crop area
                        x2_prev, y2_prev = self._original_to_preview_coords(original_iw, original_ih)
                        self._crop_rect = [0, 0, max(0, x2_prev), max(0, y2_prev)]

                    # draw the crop rectangle overlay
                    try:
                        self._draw_crop_rectangle()
                    except Exception:
                        pass
                    # update linked entry vars so UI fields reflect the crop
                    try:
                        # If we loaded from saved coordinates (in original space), use those directly
                        # to avoid precision loss from round-trip conversion (original→preview→original)
                        if saved_coords_original:
                            # Update UI vars directly with the saved original coordinates
                            try:
                                active_outer = getattr(self, '_active_rule_outer', None)
                            except Exception:
                                active_outer = None
                            # Determine whether central shared vars should be updated
                            update_central = False
                            try:
                                if not self.rule_boxes:
                                    update_central = True
                                else:
                                    first_outer = self.rule_boxes[0].get('outer')
                                    if active_outer is None or active_outer is first_outer:
                                        update_central = True
                            except Exception:
                                update_central = False
                            if update_central:
                                try:
                                    self.x1_var.set(str(int(saved_coords_original[0])))
                                    self.y1_var.set(str(int(saved_coords_original[1])))
                                    self.x2_var.set(str(int(saved_coords_original[2])))
                                    self.y2_var.set(str(int(saved_coords_original[3])))
                                except Exception:
                                    pass
                            # Update per-rule vars for the active rule
                            try:
                                if active_outer is not None:
                                    for rb in self.rule_boxes:
                                        try:
                                            if rb.get('outer') is active_outer:
                                                rvars = rb.get('vars', {}) or {}
                                                try:
                                                    if rvars.get('x1') is not None:
                                                        rvars.get('x1').set(str(int(saved_coords_original[0])))
                                                except Exception:
                                                    pass
                                                try:
                                                    if rvars.get('y1') is not None:
                                                        rvars.get('y1').set(str(int(saved_coords_original[1])))
                                                except Exception:
                                                    pass
                                                try:
                                                    if rvars.get('x2') is not None:
                                                        rvars.get('x2').set(str(int(saved_coords_original[2])))
                                                except Exception:
                                                    pass
                                                try:
                                                    if rvars.get('y2') is not None:
                                                        rvars.get('y2').set(str(int(saved_coords_original[3])))
                                                except Exception:
                                                    pass
                                                break
                                        except Exception:
                                            pass
                            except Exception:
                                pass
                        else:
                            # No saved coords, so we're using default full image - convert from preview to original
                            vals = [self._crop_rect[i] if self._crop_rect[i] is not None else 0 for i in range(4)]
                            # Use central helper so we update either the shared central vars or
                            # the per-rule vars depending on which rule is active.
                            try:
                                active_outer = getattr(self, '_active_rule_outer', None)
                            except Exception:
                                active_outer = None
                            try:
                                self._sync_crop_values(int(vals[0]), int(vals[1]), int(vals[2]), int(vals[3]), active_outer=active_outer)
                            except Exception:
                                # Fallback to directly updating central vars if helper missing
                                try:
                                    self.x1_var.set(str(int(vals[0])))
                                    self.y1_var.set(str(int(vals[1])))
                                    self.x2_var.set(str(int(vals[2])))
                                    self.y2_var.set(str(int(vals[3])))
                                except Exception:
                                    pass
                    except Exception:
                        pass
                except Exception:
                    pass
            except Exception:
                pass
        except Exception as e:
            try:
                print(f'[cropping_gui2] _open_image_preview: unexpected error while previewing {path}: {e}')
            except Exception:
                pass
            # non-fatal: draw a non-blocking placeholder with diagnostic text
            try:
                self.canvas.delete('all')
                w = max(1, self.canvas.winfo_width())
                h = max(1, self.canvas.winfo_height())
                txt = 'Preview not available\nAn unexpected error occurred while loading the image.'
                self.canvas.create_text(w//2, h//2, text=txt, fill='#666', font=('Segoe UI', 12), anchor='center', width=int(w*0.9))
                try:
                    self.preview_status_var.set('Preview not available (error)')
                except Exception:
                    pass
            except Exception:
                pass

    # --- Coordinate conversion helpers ---
    def _original_to_preview_coords(self, x, y):
        """Convert coordinates from original image space to preview canvas space."""
        try:
            orig_w, orig_h = getattr(self, '_original_image_size', (1, 1))
            prev_w, prev_h = self._preview_image_size
            if orig_w <= 0 or orig_h <= 0:
                return (x, y)
            scale_x = prev_w / orig_w
            scale_y = prev_h / orig_h
            return (int(x * scale_x), int(y * scale_y))
        except Exception:
            return (x, y)

    def _preview_to_original_coords(self, x, y):
        """Convert coordinates from preview canvas space to original image space."""
        try:
            orig_w, orig_h = getattr(self, '_original_image_size', (1, 1))
            prev_w, prev_h = self._preview_image_size
            if prev_w <= 0 or prev_h <= 0:
                return (x, y)
            scale_x = orig_w / prev_w
            scale_y = orig_h / prev_h
            return (int(x * scale_x), int(y * scale_y))
        except Exception:
            return (x, y)

    # --- Crop rectangle interaction methods ---
    def _draw_crop_rectangle(self):
        """Draw the red crop rectangle and corner handles on the canvas.

        Existing crop items are removed and recreated. Handles use the tag 'crop_handle'
        so hit testing can find them.

        Note: _crop_rect is in PREVIEW space for easy mouse interaction. Values are
        converted to ORIGINAL space when syncing with UI vars.
        """
        try:
            if self._crop_rect is None:
                return
            # _crop_rect is in preview coordinates
            x1, y1, x2, y2 = map(int, self._crop_rect)

            c = self.canvas
            # remove previous crop items
            try:
                c.delete('crop')
            except Exception:
                pass
            # main rectangle
            c.create_rectangle(x1, y1, x2, y2, outline='red', width=2, tags=('crop',))
            # compute adaptive handle size based on rect size so handles remain usable on small images
            rect_w = max(1, x2 - x1)
            rect_h = max(1, y2 - y1)
            # target handle size (pixels) clamped to a sensible range
            # Original size calculation (kept for reference):
            # orig_size = max(6, min(14, min(rect_w, rect_h) // 10))
            # Use half the original size so handles are smaller; ensure a
            # reasonable minimum and force an even value for consistent placement.
            try:
                orig_size = max(6, min(14, min(rect_w, rect_h) // 10))
                # half it (integer division) but enforce a minimum of 4 to keep handles usable
                size = max(4, orig_size // 2)
                # Ensure size is even so placement math is consistent
                if size % 2 != 0:
                    size += 1
            except Exception:
                # fallback to a small even size if anything goes wrong
                size = 4
            # place handles fully inside the crop rectangle (so they aren't drawn partially outside canvas)
            # for each corner compute handle bbox that sits inside the rect
            handles = {
                'nw': (x1, y1, x1 + size, y1 + size),
                'ne': (x2 - size, y1, x2, y1 + size),
                'sw': (x1, y2 - size, x1 + size, y2),
                'se': (x2 - size, y2 - size, x2, y2),
            }
            self._crop_handles = {}
            for name, (hx1, hy1, hx2, hy2) in handles.items():
                # Clamp handle bbox to canvas/image bounds to avoid negative coords
                try:
                    # canvas coords origin (0,0) is top-left; also clamp within preview image if available
                    ix, iy = self._preview_image_pos
                    iw, ih = self._preview_image_size
                    minx = ix if ix is not None else 0
                    miny = iy if iy is not None else 0
                    maxx = (ix + iw) if (ix is not None and iw is not None) else None
                    maxy = (iy + ih) if (iy is not None and ih is not None) else None
                    if minx is not None:
                        hx1 = max(hx1, minx)
                        hy1 = max(hy1, miny)
                    if maxx is not None:
                        hx2 = min(hx2, maxx)
                        hy2 = min(hy2, maxy)
                except Exception:
                    pass
                hid = c.create_rectangle(hx1, hy1, hx2, hy2,
                                         fill='red', outline='black', tags=('crop', 'crop_handle', f'handle_{name}'))
                self._crop_handles[hid] = name

            # Cursor feedback for handles (helps users discover they can drag corners)
            try:
                # bind tag-level enter/leave to change cursor when over any handle
                # prefer directional sizing cursors where available
                c.tag_bind('handle_nw', '<Enter>', lambda e, canv=c: canv.configure(cursor='size_nw_se'))
                c.tag_bind('handle_nw', '<Leave>', lambda e, canv=c: canv.configure(cursor=''))
                c.tag_bind('handle_se', '<Enter>', lambda e, canv=c: canv.configure(cursor='size_nw_se'))
                c.tag_bind('handle_se', '<Leave>', lambda e, canv=c: canv.configure(cursor=''))
                c.tag_bind('handle_ne', '<Enter>', lambda e, canv=c: canv.configure(cursor='size_ne_sw'))
                c.tag_bind('handle_ne', '<Leave>', lambda e, canv=c: canv.configure(cursor=''))
                c.tag_bind('handle_sw', '<Enter>', lambda e, canv=c: canv.configure(cursor='size_ne_sw'))
                c.tag_bind('handle_sw', '<Leave>', lambda e, canv=c: canv.configure(cursor=''))
            except Exception:
                pass
        except Exception:
            pass

    def _clamp_crop_rect_to_preview(self):
        """Ensure self._crop_rect lies fully within the preview image bounds.
        Adjusts the rectangle in-place if it would extend outside the preview image.
        """
        try:
            if not self._crop_rect:
                return
            ix, iy = self._preview_image_pos
            iw, ih = self._preview_image_size
            if ix is None:
                ix, iy = 0, 0
            minx = ix
            miny = iy
            maxx = ix + (iw or 0)
            maxy = iy + (ih or 0)
            x1, y1, x2, y2 = map(int, self._crop_rect)
            # clamp coordinates
            if x1 < minx:
                # shift right
                shift = minx - x1
                x1 += shift
                x2 += shift
            if x2 > maxx:
                # shift left
                shift = x2 - maxx
                x1 -= shift
                x2 -= shift
            if y1 < miny:
                shift = miny - y1
                y1 += shift
                y2 += shift
            if y2 > maxy:
                shift = y2 - maxy
                y1 -= shift
                y2 -= shift
            # Ensure rect still valid; if image smaller than rect, snap to image
            if x2 - x1 < 1:
                x1, x2 = minx, maxx
            if y2 - y1 < 1:
                y1, y2 = miny, maxy
            self._crop_rect = [int(max(minx, x1)), int(max(miny, y1)), int(min(maxx, x2)), int(min(maxy, y2))]
        except Exception:
            pass

    def _on_canvas_mouse_down(self, event):
        try:
            c = self.canvas
            x = c.canvasx(event.x)
            y = c.canvasy(event.y)
            # find overlapping items at point (topmost first)
            items = c.find_overlapping(x, y, x, y)
            action = None
            handle_name = None
            # check for handle hit
            for iid in reversed(items):
                if iid in self._crop_handles:
                    handle_name = self._crop_handles.get(iid)
                    action = 'resize'
                    break
            # if no handle, check if inside rect -> move
            if action is None and self._crop_rect is not None:
                x1, y1, x2, y2 = self._crop_rect
                if x1 <= x <= x2 and y1 <= y <= y2:
                    action = 'move'
            # store drag state
            self._crop_drag['action'] = action
            self._crop_drag['handle'] = handle_name
            self._crop_drag['start_x'] = x
            self._crop_drag['start_y'] = y
            self._crop_drag['orig_rect'] = None if self._crop_rect is None else list(self._crop_rect)
        except Exception:
            pass

    def _on_canvas_mouse_move(self, event):
        try:
            if self._crop_drag.get('action') is None:
                return
            c = self.canvas
            x = c.canvasx(event.x)
            y = c.canvasy(event.y)
            dx = x - self._crop_drag['start_x']
            dy = y - self._crop_drag['start_y']
            orig = self._crop_drag.get('orig_rect') or [0,0,0,0]
            if self._crop_drag['action'] == 'move':
                x1, y1, x2, y2 = orig
                # Instead of applying dx/dy directly then clamping (which can
                # allow the rectangle to be partially moved out of bounds),
                # compute the allowable translation range and clamp dx/dy so
                # the whole rectangle remains inside the preview image.
                ix, iy = self._preview_image_pos
                iw, ih = self._preview_image_size
                if ix is None:
                    ix, iy = 0, 0
                # allowable translation so rectangle stays inside [ix, ix+iw] / [iy, iy+ih]
                max_left = ix - x1          # most negative dx allowed
                max_right = (ix + iw) - x2  # most positive dx allowed
                max_up = iy - y1            # most negative dy allowed
                max_down = (iy + ih) - y2   # most positive dy allowed

                # clamp dx/dy into the allowable range
                try:
                    dx_clamped = int(min(max(dx, max_left), max_right))
                except Exception:
                    dx_clamped = 0
                try:
                    dy_clamped = int(min(max(dy, max_up), max_down))
                except Exception:
                    dy_clamped = 0

                nx1 = int(x1 + dx_clamped)
                ny1 = int(y1 + dy_clamped)
                nx2 = int(x2 + dx_clamped)
                ny2 = int(y2 + dy_clamped)

                # ensure dimensions remain valid
                if nx2 - nx1 < 1 or ny2 - ny1 < 1:
                    return
                self._crop_rect = [nx1, ny1, nx2, ny2]
                try:
                    self._clamp_crop_rect_to_preview()
                except Exception:
                    pass
            elif self._crop_drag['action'] == 'resize':
                handle = self._crop_drag.get('handle')
                x1, y1, x2, y2 = orig
                # Determine which aspect variable to use (active rule)
                aspect_val = None
                try:
                    # Try to find the aspect var for the active rule
                    for rb in self.rule_boxes:
                        if rb.get('outer') is getattr(self, '_active_rule_outer', None):
                            aspect_var = rb.get('vars', {}).get('aspect')
                            if aspect_var is not None:
                                aspect_val = aspect_var.get()
                            # If custom aspect, get the width and height values
                            if aspect_val == 'custom':
                                custom_width_var = rb.get('vars', {}).get('custom_aspect_width')
                                custom_height_var = rb.get('vars', {}).get('custom_aspect_height')
                                if custom_width_var and custom_height_var:
                                    width_str = custom_width_var.get().strip()
                                    height_str = custom_height_var.get().strip()
                                    if width_str and height_str:
                                        aspect_val = f"{width_str}:{height_str}"
                                    else:
                                        aspect_val = 'none'
                                else:
                                    aspect_val = 'none'
                            break
                except Exception:
                    pass
                if not aspect_val:
                    # fallback to self.aspect_var if present
                    aspect_val = getattr(self, 'aspect_var', None)
                    if aspect_val is not None:
                        aspect_val = aspect_val.get()
                    # If custom aspect, get the width and height values from self
                    if aspect_val == 'custom':
                        try:
                            custom_width_var = getattr(self, 'custom_aspect_width_var', None)
                            custom_height_var = getattr(self, 'custom_aspect_height_var', None)
                            if custom_width_var and custom_height_var:
                                width_str = custom_width_var.get().strip()
                                height_str = custom_height_var.get().strip()
                                if width_str and height_str:
                                    aspect_val = f"{width_str}:{height_str}"
                                else:
                                    aspect_val = 'none'
                            else:
                                aspect_val = 'none'
                        except Exception:
                            aspect_val = 'none'
                if aspect_val and aspect_val != 'none':
                    # Ensure we have a string aspect value
                    try:
                        aspect_val = str(aspect_val) if aspect_val is not None else 'none'
                    except Exception:
                        aspect_val = 'none'
                    # Coerce handle to str for safety
                    try:
                        handle_str = str(handle) if handle is not None else ''
                    except Exception:
                        handle_str = ''
                    # Use aspect-ratio preserving resize
                    self._crop_rect = self._resize_with_aspect(handle_str, orig, x, y, aspect_val)
                    try:
                        self._clamp_crop_rect_to_preview()
                    except Exception:
                        pass
                else:
                    ix, iy = self._preview_image_pos
                    iw, ih = self._preview_image_size
                    minx = ix
                    miny = iy
                    maxx = ix + iw
                    maxy = iy + ih
                    nx1, ny1, nx2, ny2 = x1, y1, x2, y2
                    if handle == 'nw':
                        nx1 = int(min(max(minx, x1 + dx), x2 - 1))
                        ny1 = int(min(max(miny, y1 + dy), y2 - 1))
                    elif handle == 'ne':
                        nx2 = int(max(min(maxx, x2 + dx), x1 + 1))
                        ny1 = int(min(max(miny, y1 + dy), y2 - 1))
                    elif handle == 'sw':
                        nx1 = int(min(max(minx, x1 + dx), x2 - 1))
                        ny2 = int(max(min(maxy, y2 + dy), y1 + 1))
                    elif handle == 'se':
                        nx2 = int(max(min(maxx, x2 + dx), x1 + 1))
                        ny2 = int(max(min(maxy, y2 + dy), y1 + 1))
                    # ensure within bounds and valid
                    nx1 = max(minx, min(nx1, maxx - 1))
                    nx2 = max(minx + 1, min(nx2, maxx))
                    ny1 = max(miny, min(ny1, maxy - 1))
                    ny2 = max(miny + 1, min(ny2, maxy))
                    self._crop_rect = [nx1, ny1, nx2, ny2]
                    try:
                        self._clamp_crop_rect_to_preview()
                    except Exception:
                        pass
            # redraw
            self._draw_crop_rectangle()
            # update the entry vars - use _sync_crop_values to convert from preview to original coords
            try:
                cx1 = int(self._crop_rect[0])
                cy1 = int(self._crop_rect[1])
                cx2 = int(self._crop_rect[2])
                cy2 = int(self._crop_rect[3])
                # Use _sync_crop_values which handles conversion from preview to original coordinates
                try:
                    active_outer = getattr(self, '_active_rule_outer', None)
                except Exception:
                    active_outer = None
                try:
                    self._sync_crop_values(cx1, cy1, cx2, cy2, active_outer=active_outer)
                except Exception:
                    pass
            except Exception:
                pass
        except Exception:
            pass

    def _on_canvas_mouse_up(self, event):
        try:
            # clear drag state
            self._crop_drag['action'] = None
            self._crop_drag['handle'] = None
            self._crop_drag['orig_rect'] = None
        except Exception:
            pass

    def _on_thumbnail_double(self, path: str, dlg: tk.Toplevel, position: int, on_select=None):
        """Handler for double-clicking a thumbnail: close dialog, show image in preview,
        and call optional on_select(path, position).
        """
        try:
            if dlg is not None:
                try:
                    dlg.destroy()
                except Exception:
                    pass
        except Exception:
            pass
        try:
            # call optional callback provided by the caller (e.g., to update which rule got selected)
            if callable(on_select):
                try:
                    on_select(path, position)
                except Exception:
                    pass
            # also update the central selected_position_var for backward compatibility
            try:
                self.selected_position_var.set(f'Position {position}')
            except Exception:
                pass
            # Open the preview by position so the preview always resolves the live
            # image currently assigned to that position (name-independent).
            try:
                self._open_image_preview_by_position(position)
            except Exception:
                # fallback to opening by explicit path if position lookup fails
                try:
                    self._open_image_preview(path)
                except Exception:
                    pass
        except Exception:
            pass

    def _on_rule_image_selected(self, outer, path: str, position: int, pos_var: tk.StringVar | None):
        """Handle an image chosen from the Select Image dialog for a specific rule.

        - outer: the outer Frame for the rule that initiated the selection
        - path: filesystem path to the chosen image
        - position: integer position (e.g., Position N) from the dialog
        - pos_var: the per-rule StringVar to update with the position text
        """
        try:
            # store the image path on the matching rule entry
            for rb in self.rule_boxes:
                try:
                    if rb.get('outer') is outer:
                        # Prefer storing the position rather than a persistent filename
                        # so rules bind to a source slot (Position N) instead of a name.
                        rb['position'] = position
                        # keep the originally-selected path as a transient example (optional)
                        rb['img_temp'] = path
                        # update per-rule position label
                        try:
                            if pos_var is not None:
                                pos_var.set(f'Position {position}')
                        except Exception:
                            pass
                        break
                except Exception:
                    pass

            # update central backward-compat var
            try:
                self.selected_position_var.set(f'Position {position}')
            except Exception:
                pass

            # make the rule active which will update the preview
            try:
                self.set_active_rule(outer)
            except Exception:
                # as a fallback, directly open the preview
                try:
                    self._open_image_preview_by_position(position)
                except Exception:
                    pass
        except Exception:
            pass

    def get_source_images(self) -> list:
        """Return the ordered list of source image paths as used by the source dialog.

        The ordering matches the dialog (oldest first). Returns an empty list on error.
        """
        try:
            src = load_config('source_folder')
        except Exception:
            src = ''
        if not src or not os.path.isdir(src):
            return []
        exts = ('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif', '.tif', '.tiff', '.heic', '.heif')
        files = []
        try:
            for fn in os.listdir(src):
                path = os.path.join(src, fn)
                if os.path.isfile(path) and fn.lower().endswith(exts):
                    try:
                        mtime = os.path.getmtime(path)
                    except Exception:
                        mtime = 0
                    files.append((mtime, fn, path))
            files.sort(key=lambda t: t[0])
            return [p for (_m, _n, p) in files]
        except Exception:
            return []

    def _open_image_preview_by_position(self, position: int):
        """Resolve the live source image for `position` (1-based) and show it in the preview.

        If no image exists for the given position, the canvas is cleared and a placeholder is shown.
        """
        try:
            imgs = self.get_source_images()
        except Exception as e:
            try:
                print(f'[cropping_gui2] _open_image_preview_by_position: error listing source images: {e}')
            except Exception:
                pass
            imgs = []

        # if no configured source folder or no images found, draw a helpful placeholder
        try:
            src_cfg = load_config('source_folder')
        except Exception:
            src_cfg = ''
        if not src_cfg or not os.path.isdir(src_cfg):
            try:
                self.canvas.delete('all')
                w = max(1, self.canvas.winfo_width())
                h = max(1, self.canvas.winfo_height())
                txt = 'Source folder not configured or does not exist.\nPlease set the source folder in the main application.'
                self.canvas.create_text(w//2, h//2, text=txt, fill='#666', font=('Segoe UI', 12), anchor='center', width=int(w*0.9))
                try:
                    self.preview_status_var.set('Source folder not configured')
                except Exception:
                    pass
            except Exception:
                pass
            return

        idx = position - 1
        if not imgs or idx < 0 or idx >= len(imgs):
            try:
                print(f'[cropping_gui2] _open_image_preview_by_position: no image for position {position} (found {len(imgs)} images)')
            except Exception:
                pass
            try:
                self.canvas.delete('all')
                w = max(1, self.canvas.winfo_width())
                h = max(1, self.canvas.winfo_height())
                txt = f'No image available for Position {position}\n\nFound {len(imgs)} image(s) in source folder.'
                self.canvas.create_text(w//2, h//2, text=txt, fill='#666', font=('Segoe UI', 12), anchor='center', width=int(w*0.9))
                try:
                    self.preview_status_var.set(f'No image for Position {position} (found {len(imgs)})')
                except Exception:
                    pass
            except Exception:
                pass
            return

        path = imgs[idx]
        if not os.path.exists(path):
            try:
                print(f'[cropping_gui2] _open_image_preview_by_position: resolved path missing for position {position}: {path}')
            except Exception:
                pass
            try:
                self.canvas.delete('all')
                w = max(1, self.canvas.winfo_width())
                h = max(1, self.canvas.winfo_height())
                txt = f'File for Position {position} not found:\n{os.path.basename(path)}'
                self.canvas.create_text(w//2, h//2, text=txt, fill='#666', font=('Segoe UI', 12), anchor='center', width=int(w*0.9))
                try:
                    self.preview_status_var.set('Resolved file missing: ' + os.path.basename(path))
                except Exception:
                    pass
            except Exception:
                pass
            return

        # Finally attempt to open the resolved path
        try:
            self._open_image_preview(path)
        except Exception as e:
            try:
                print(f'[cropping_gui2] _open_image_preview_by_position: failed to preview {path}: {e}')
            except Exception:
                pass
            try:
                self.canvas.delete('all')
                w = max(1, self.canvas.winfo_width())
                h = max(1, self.canvas.winfo_height())
                txt = f'Could not preview file for Position {position}:\n{os.path.basename(path)}'
                self.canvas.create_text(w//2, h//2, text=txt, fill='#666', font=('Segoe UI', 12), anchor='center', width=int(w*0.9))
                try:
                    self.preview_status_var.set('Could not preview: ' + os.path.basename(path))
                except Exception:
                    pass
            except Exception:
                pass
        try:
            print(f'[cropping_gui2] _open_image_preview_by_position: success for {path}')
        except Exception:
            pass

    def open_source_images_dialog(self, on_select=None):
        """Open a modal dialog listing image files found in the configured source folder.
        If `on_select` is provided it will be called as on_select(path, position) when a
        thumbnail is double-clicked. The dialog still previews the image locally.
        """
        # read configured source folder
        try:
            src = load_config('source_folder')
        except Exception:
            src = ''
        if not src or not os.path.isdir(src):
            try:
                messagebox.showerror('Source folder not found', 'Source folder is not configured or does not exist. Please set it in the main application.', parent=self.master)
            except Exception:
                pass
            return

        exts = ('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif', '.tif', '.tiff', '.heic', '.heif')

        # gather files with their modification time; sort oldest first
        files = []
        try:
            for fn in os.listdir(src):
                path = os.path.join(src, fn)
                if os.path.isfile(path) and fn.lower().endswith(exts):
                    try:
                        mtime = os.path.getmtime(path)
                    except Exception:
                        mtime = 0
                    files.append((mtime, fn, path))
            files.sort(key=lambda t: t[0])
        except Exception:
            files = []

        dlg = tk.Toplevel(self.master)
        try:
            dlg.title('Source Images')
            dlg.transient(self.master)
            dlg.geometry('900x520')
            dlg.grab_set()
            # track which thumbnail label is single-click selected (store on dialog)
            try:
                dlg._selected_thumb = None
                dlg._selected_thumb_path = None
                dlg._selected_thumb_index = None
            except Exception:
                pass
        except Exception:
            pass

        container = ttk.Frame(dlg)
        container.pack(fill='both', expand=True)

        # Canvas + vertical scrollbar to host a grid of thumbnails
        canvas = tk.Canvas(container, bd=0)
        vscroll = ttk.Scrollbar(container, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        inner = ttk.Frame(canvas)
        window_item = canvas.create_window((0, 0), window=inner, anchor='nw')

        # Thumbnail layout parameters
        cols = 4
        thumb_size = (200, 140)  # max width, max height
        thumb_refs = []
        thumb_data = []  # Store (lbl, path, idx, r, c) for keyboard navigation

        # populate grid
        r = 0
        c = 0
        # enumerate files so we can display position labels (oldest -> Position 1)
        for idx, (mtime, name, path) in enumerate(files, start=1):
            cell = ttk.Frame(inner, width=thumb_size[0], padding=6)
            cell.grid(row=r, column=c, padx=4, pady=4, sticky='n')

            if PIL_AVAILABLE:
                try:
                    img = Image.open(path)
                    img.thumbnail(thumb_size, _RESAMPLE_FILTER)
                    photo = ImageTk.PhotoImage(img)
                    # use tk.Label (not ttk) so we can control highlightthickness/bg easily
                    # Always set highlightthickness=2 to reserve space and prevent layout shifts
                    lbl = tk.Label(cell, image=photo, bd=0, relief='flat', highlightthickness=2, highlightbackground='#f0f0f0')
                    lbl.image = photo
                    lbl.pack()
                    thumb_refs.append(photo)
                except Exception:
                    lbl = tk.Label(cell, text='[Unreadable]', bd=0, relief='flat', highlightthickness=2, highlightbackground='#f0f0f0')
                    lbl.pack()
            else:
                # Pillow not available: show filename-only placeholder
                lbl = tk.Label(cell, text='(Pillow not installed)', fg='#666', bd=0, relief='flat', highlightthickness=2, highlightbackground='#f0f0f0')
                lbl.pack()

            # Add position label under the image
            pos_label = tk.Label(cell, text=f'Position {idx}', fg='#666', font=('Arial', 9))
            pos_label.pack(pady=(2, 0))

            # single-click: highlight this thumbnail (and clear previous)
            try:
                def _on_thumb_click(e, l=lbl, p=path, i=idx, d=dlg):
                    try:
                        prev = getattr(d, '_selected_thumb', None)
                        if prev is not None and prev is not l:
                            try:
                                # Reset previous highlight to match background (invisible)
                                prev.configure(highlightbackground='#f0f0f0')
                            except Exception:
                                pass
                        try:
                            # visual highlight: blue outline (only change color, not thickness)
                            l.configure(highlightbackground='#1e90ff')
                        except Exception:
                            try:
                                l.configure(bg='#cfe6ff')
                            except Exception:
                                pass
                        d._selected_thumb = l
                        d._selected_thumb_path = p
                        d._selected_thumb_index = i
                        # enable the Open button when a thumbnail is selected
                        try:
                            if hasattr(d, '_open_btn'):
                                d._open_btn.configure(state='normal')
                        except Exception:
                            pass
                    except Exception:
                        pass

                lbl.bind('<Button-1>', _on_thumb_click)
            except Exception:
                pass

            # bind double-click to open the preview, close the dialog, and call on_select if provided
            try:
                lbl.bind('<Double-Button-1>', lambda e, p=path, d=dlg, i=idx: self._on_thumbnail_double(p, d, i, on_select))
            except Exception:
                pass

            # Store thumbnail data for keyboard navigation
            thumb_data.append((lbl, path, idx, r, c))

            c += 1
            if c >= cols:
                c = 0
                r += 1

        # keep references on the dialog so images are not GC'd
        dlg._thumb_refs = thumb_refs
        dlg._thumb_data = thumb_data  # For keyboard navigation
        dlg._thumb_cols = cols  # For keyboard navigation

        # update scrollregion when contents change
        def _on_inner_config(e):
            try:
                canvas.configure(scrollregion=canvas.bbox('all'))
            except Exception:
                pass

        def _on_canvas_config(e):
            try:
                canvas.itemconfig(window_item, width=e.width)
            except Exception:
                pass

        inner.bind('<Configure>', _on_inner_config)
        canvas.bind('<Configure>', _on_canvas_config)

        # Mouse wheel scrolling support
        def _on_mousewheel(event):
            try:
                # Windows and MacOS
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass

        def _on_mousewheel_linux(event):
            try:
                # Linux uses button 4 (scroll up) and button 5 (scroll down)
                if event.num == 4:
                    canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    canvas.yview_scroll(1, "units")
            except Exception:
                pass

        # Bind mouse wheel events
        canvas.bind('<MouseWheel>', _on_mousewheel)  # Windows/MacOS
        canvas.bind('<Button-4>', _on_mousewheel_linux)  # Linux scroll up
        canvas.bind('<Button-5>', _on_mousewheel_linux)  # Linux scroll down

        # Also bind to the dialog and inner frame so scrolling works anywhere
        dlg.bind('<MouseWheel>', _on_mousewheel)
        dlg.bind('<Button-4>', _on_mousewheel_linux)
        dlg.bind('<Button-5>', _on_mousewheel_linux)
        inner.bind('<MouseWheel>', _on_mousewheel)
        inner.bind('<Button-4>', _on_mousewheel_linux)
        inner.bind('<Button-5>', _on_mousewheel_linux)

        # Keyboard navigation with arrow keys
        def _on_key_press(event):
            try:
                # Get current selection index from thumb_data list
                current_idx = None
                current_thumb = getattr(dlg, '_selected_thumb', None)

                if current_thumb is not None:
                    # Find current thumb in thumb_data
                    for i, (lbl, path, pos_idx, row, col) in enumerate(dlg._thumb_data):
                        if lbl is current_thumb:
                            current_idx = i
                            break

                # If nothing selected, start at first thumb
                if current_idx is None:
                    current_idx = 0

                num_thumbs = len(dlg._thumb_data)
                if num_thumbs == 0:
                    return

                new_idx = current_idx

                # Arrow key navigation
                if event.keysym == 'Right':
                    new_idx = min(current_idx + 1, num_thumbs - 1)
                elif event.keysym == 'Left':
                    new_idx = max(current_idx - 1, 0)
                elif event.keysym == 'Down':
                    # Move down one row (add cols)
                    new_idx = min(current_idx + dlg._thumb_cols, num_thumbs - 1)
                elif event.keysym == 'Up':
                    # Move up one row (subtract cols)
                    new_idx = max(current_idx - dlg._thumb_cols, 0)
                elif event.keysym == 'Return' or event.keysym == 'KP_Enter':
                    # Enter key: select current thumbnail
                    sel_path = getattr(dlg, '_selected_thumb_path', None)
                    sel_idx = getattr(dlg, '_selected_thumb_index', None)
                    if sel_path:
                        # Call the same logic as the Open button
                        if callable(on_select):
                            try:
                                on_select(sel_path, sel_idx)
                            except Exception:
                                pass
                        try:
                            self.selected_position_var.set(f'Position {sel_idx}')
                        except Exception:
                            pass
                        try:
                            if sel_idx is not None:
                                self._open_image_preview_by_position(sel_idx)
                            else:
                                self._open_image_preview(sel_path)
                        except Exception:
                            pass
                        try:
                            dlg.destroy()
                        except Exception:
                            pass
                    return
                else:
                    # Other keys - ignore
                    return

                # Select the new thumbnail
                if new_idx != current_idx and 0 <= new_idx < num_thumbs:
                    new_lbl, new_path, new_pos_idx, _, _ = dlg._thumb_data[new_idx]

                    # Clear previous highlight
                    prev = getattr(dlg, '_selected_thumb', None)
                    if prev is not None:
                        try:
                            prev.configure(highlightbackground='#f0f0f0')
                        except Exception:
                            pass

                    # Highlight new selection
                    try:
                        new_lbl.configure(highlightbackground='#1e90ff')
                    except Exception:
                        try:
                            new_lbl.configure(bg='#cfe6ff')
                        except Exception:
                            pass

                    # Update dialog state
                    dlg._selected_thumb = new_lbl
                    dlg._selected_thumb_path = new_path
                    dlg._selected_thumb_index = new_pos_idx

                    # Enable Open button
                    try:
                        if hasattr(dlg, '_open_btn'):
                            dlg._open_btn.configure(state='normal')
                    except Exception:
                        pass

                    # Scroll to make the new selection visible
                    try:
                        canvas.update_idletasks()
                        # Get the widget's position relative to the canvas window
                        widget_y = new_lbl.winfo_y()
                        widget_height = new_lbl.winfo_height() + 100  # Add padding for position label
                        canvas_height = canvas.winfo_height()

                        # Get scrollregion
                        scrollregion = canvas.cget('scrollregion')
                        if scrollregion:
                            parts = scrollregion.split()
                            if len(parts) == 4:
                                total_height = float(parts[3])
                                if total_height > 0:
                                    # Calculate visible area in canvas coordinates
                                    visible_top = canvas.yview()[0] * total_height
                                    visible_bottom = canvas.yview()[1] * total_height

                                    # Check if widget is outside visible area
                                    if widget_y < visible_top:
                                        # Scroll up to show widget at top
                                        canvas.yview_moveto(max(0, widget_y / total_height))
                                    elif (widget_y + widget_height) > visible_bottom:
                                        # Scroll down to show widget at bottom
                                        canvas.yview_moveto(max(0, (widget_y + widget_height - canvas_height) / total_height))
                    except Exception:
                        pass
            except Exception:
                pass

        # Bind keyboard events to the dialog
        dlg.bind('<Key>', _on_key_press)

        # Bottom action row (only Close button per request)
        action_row = ttk.Frame(dlg)
        action_row.pack(fill='x', pady=(6,8), padx=8)
        try:
            def _open_and_apply():
                try:
                    sel_path = getattr(dlg, '_selected_thumb_path', None)
                    sel_idx = getattr(dlg, '_selected_thumb_index', None)
                    if sel_path:
                        # call external callback if provided
                        if callable(on_select):
                            try:
                                on_select(sel_path, sel_idx)
                            except Exception:
                                pass
                        # update central selection var and preview in the main UI
                        try:
                            self.selected_position_var.set(f'Position {sel_idx}')
                        except Exception:
                            pass
                        try:
                            # prefer showing by position so preview reflects the live slot
                            if sel_idx is not None:
                                self._open_image_preview_by_position(sel_idx)
                            else:
                                self._open_image_preview(sel_path)
                        except Exception:
                            pass
                except Exception:
                    pass
                try:
                    dlg.destroy()
                except Exception:
                    pass

            btn = ttk.Button(action_row, text='Open', command=_open_and_apply, state='disabled')
            btn.pack(side='right')
            try:
                dlg._open_btn = btn
            except Exception:
                pass
        except Exception:
            try:
                ttk.Button(action_row, text='Open', command=dlg.destroy).pack(side='right')
            except Exception:
                pass

        try:
            dlg.wait_window()
        except Exception:
            pass

    def add_rule(self, rule_obj: dict | None = None, activate: bool = True):
        """Create and append a new rule box below existing ones."""
        try:
            idx = len(self.rule_boxes) + 1

            outer = tk.Frame(self.rules_container, bg='#e8f2ff', bd=0, highlightthickness=1,
                             highlightbackground='#cfe6ff', highlightcolor='#cfe6ff', takefocus=1)
            outer.pack(fill='both', expand=False, padx=8, pady=(4,8))

            # Bind click and focus so the user sees a subtle active background tint
            try:
                outer.bind('<Button-1>', lambda e, o=outer: self.set_active_rule(o))
                outer.bind('<FocusIn>', lambda e, o=outer: self.set_active_rule(o))
            except Exception:
                pass

            inner = ttk.Frame(outer, padding=10)
            inner.pack(fill='both', expand=True)

            # Bind inner area and header so clicks on child widgets make this rule active
            try:
                inner.bind('<Button-1>', lambda e, o=outer: self.set_active_rule(o))
            except Exception:
                pass

            # Header with label + delete
            header = ttk.Frame(inner)
            header.grid(row=0, column=0, columnspan=3, sticky='ew')
            header.grid_columnconfigure(0, weight=1)
            header.grid_columnconfigure(2, weight=1)
            rule_label = ttk.Label(header, text=f'Rule {idx}:', font=('Segoe UI', 12, 'bold'))
            rule_label.grid(row=0, column=0, sticky='w')

            # Delete button: prefer ttkbootstrap's styled button, fall back to tk.Button
            try:
                if TTB_AVAILABLE:
                    try:
                        del_btn = tb.Button(header, text='Delete', bootstyle='danger', command=lambda o=outer: self._remove_rule(o))
                    except Exception:
                        del_btn = tk.Button(header, text='Delete', bg='#ff4d6d', fg='white', bd=0, padx=8, pady=4, command=lambda o=outer: self._remove_rule(o))
                else:
                    del_btn = tk.Button(header, text='Delete', bg='#ff4d6d', fg='white', bd=0, padx=8, pady=4, command=lambda o=outer: self._remove_rule(o))
                del_btn.grid(row=0, column=1, padx=12, pady=2)
            except Exception:
                del_btn = None

            # Select Image button + per-rule position label
            sel_frame = ttk.Frame(inner)
            sel_frame.grid(row=1, column=0, columnspan=3, sticky='w', pady=(8,6))
            try:
                if TTB_AVAILABLE:
                    try:
                        select_btn = tb.Button(sel_frame, text='Select Image', bootstyle='secondary', command=lambda: self.open_source_images_dialog())
                    except Exception:
                        select_btn = tk.Button(sel_frame, text='Select Image', bg='#6c757d', fg='white', bd=0, padx=6, pady=4, command=lambda: self.open_source_images_dialog())
                else:
                    select_btn = tk.Button(sel_frame, text='Select Image', bg='#6c757d', fg='white', bd=0, padx=6, pady=4, command=lambda: self.open_source_images_dialog())
                select_btn.pack(side='left')
            except Exception:
                select_btn = None

            # per-rule position var and label
            try:
                pos_var = tk.StringVar(value='')
                pos_lbl = ttk.Label(sel_frame, textvariable=pos_var, font=('Segoe UI', 10))
                pos_lbl.pack(side='left', padx=(8,0))
            except Exception:
                pos_var = None

            # wire select_btn to update this rule's pos_var when an image is chosen
            try:
                if select_btn is not None and pos_var is not None:
                    select_btn.configure(command=lambda pv=pos_var, o=outer: self.open_source_images_dialog(on_select=lambda p, pos, pv=pv, o=o: self._on_rule_image_selected(o, p, pos, pv)))
            except Exception:
                pass

            # Aspect Ratio radios (per-rule variable)
            ar_label = ttk.Label(inner, text='Aspect Ratio:', font=('Segoe UI', 9, 'underline'))
            ar_label.grid(row=2, column=0, columnspan=3, sticky='w', pady=(6,2))
            ar_var = tk.StringVar(value='none')
            ar_values = [
                ('1:1','1:1'), ('3:4','3:4'), ('4:3','4:3'), ('4:5','4:5'),
                ('16:9','16:9'), ('9:16','9:16'), ('1.91:1','1.91:1'), ('3:2','3:2'),
                ('Custom','custom'), ('None','none')
            ]
            ar_frame = ttk.Frame(inner)
            ar_frame.grid(row=3, column=0, columnspan=3, sticky='w')
            for col in range(4):
                ar_frame.grid_columnconfigure(col, weight=1, minsize=65)
            for i, (lbl, val) in enumerate(ar_values):
                if i < 4:
                    r = 0
                    c = i
                elif i < 8:
                    r = 1
                    c = i - 4
                else:
                    r = 2
                    c = i - 8
                try:
                    rb = ttk.Radiobutton(ar_frame, text=lbl, value=val, variable=ar_var)
                    rb.grid(row=r, column=c, sticky='w', padx=2, pady=1)
                except Exception:
                    pass

            # Custom aspect ratio input field
            custom_ar_frame = ttk.Frame(inner)
            custom_ar_frame.grid(row=4, column=0, columnspan=3, sticky='w', pady=(4,0))
            ttk.Label(custom_ar_frame, text='').grid(row=0, column=0, sticky='w')
            custom_aspect_width_var = tk.StringVar(value='')
            custom_width_entry = ttk.Entry(custom_ar_frame, textvariable=custom_aspect_width_var, width=6)
            custom_width_entry.grid(row=0, column=1, sticky='w', padx=(4,0))
            ttk.Label(custom_ar_frame, text=':').grid(row=0, column=2, sticky='w', padx=(2,2))
            custom_aspect_height_var = tk.StringVar(value='')
            custom_height_entry = ttk.Entry(custom_ar_frame, textvariable=custom_aspect_height_var, width=6)
            custom_height_entry.grid(row=0, column=3, sticky='w', padx=(0,4))
            ttk.Label(custom_ar_frame, text='(e.g., 16:9)', font=('Segoe UI', 8), foreground='#666').grid(row=0, column=4, sticky='w', padx=(4,0))
            # Initially hide custom input
            custom_ar_frame.grid_remove()

            # Wire up aspect_var trace to snap crop to aspect when changed (for this rule)
            def _aspect_trace_rule(*args, av=ar_var, caf=custom_ar_frame, caw=custom_aspect_width_var, cah=custom_aspect_height_var):
                try:
                    val = av.get()
                    if val == 'custom':
                        # Show custom input field
                        caf.grid()
                        # Apply custom aspect if there's a value
                        custom_width = caw.get().strip()
                        custom_height = cah.get().strip()
                        if custom_width and custom_height:
                            try:
                                # Validate that both are numbers
                                float(custom_width)
                                float(custom_height)
                                custom_val = f"{custom_width}:{custom_height}"
                                self._apply_aspect_to_crop(custom_val)
                                self._draw_crop_rectangle()
                            except ValueError:
                                pass
                    else:
                        # Hide custom input field
                        caf.grid_remove()
                        if val and val != 'none':
                            self._apply_aspect_to_crop(val)
                            self._draw_crop_rectangle()
                except Exception:
                    pass

            # Also trace custom aspect input
            def _custom_aspect_trace_rule(*args, av=ar_var, caw=custom_aspect_width_var, cah=custom_aspect_height_var):
                try:
                    if av.get() == 'custom':
                        custom_width = caw.get().strip()
                        custom_height = cah.get().strip()
                        if custom_width and custom_height:
                            try:
                                # Validate that both are numbers
                                float(custom_width)
                                float(custom_height)
                                custom_val = f"{custom_width}:{custom_height}"
                                self._apply_aspect_to_crop(custom_val)
                                self._draw_crop_rectangle()
                            except ValueError:
                                pass
                except Exception:
                    pass

            try:
                ar_var.trace_add('write', _aspect_trace_rule)
            except Exception:
                try:
                    ar_var.trace('w', _aspect_trace_rule)
                except Exception:
                    pass

            try:
                custom_aspect_width_var.trace_add('write', _custom_aspect_trace_rule)
                custom_aspect_height_var.trace_add('write', _custom_aspect_trace_rule)
            except Exception:
                try:
                    custom_aspect_width_var.trace('w', _custom_aspect_trace_rule)
                    custom_aspect_height_var.trace('w', _custom_aspect_trace_rule)
                except Exception:
                    pass

            # Crop Area inputs (X1, Y1, X2, Y2)
            ca_label = ttk.Label(inner, text='Crop Area (in Pixels):', font=('Segoe UI', 9, 'underline'))
            ca_label.grid(row=5, column=0, columnspan=3, sticky='w', pady=(8,4))
            ca_frame = ttk.Frame(inner)
            ca_frame.grid(row=6, column=0, columnspan=3, sticky='w')
            ttk.Label(ca_frame, text='X1').grid(row=0, column=0)
            x1_var = tk.StringVar(value='0')
            ttk.Entry(ca_frame, textvariable=x1_var, width=6).grid(row=0, column=1, padx=(4,12))
            ttk.Label(ca_frame, text='Y1').grid(row=0, column=2)
            y1_var = tk.StringVar(value='0')
            ttk.Entry(ca_frame, textvariable=y1_var, width=6).grid(row=0, column=3, padx=(4,12))

            ttk.Label(ca_frame, text='X2').grid(row=1, column=0, pady=(6,0))
            x2_var = tk.StringVar(value='0')
            ttk.Entry(ca_frame, textvariable=x2_var, width=6).grid(row=1, column=1, padx=(4,12), pady=(6,0))
            ttk.Label(ca_frame, text='Y2').grid(row=1, column=2, pady=(6,0))
            y2_var = tk.StringVar(value='0')
            ttk.Entry(ca_frame, textvariable=y2_var, width=6).grid(row=1, column=3, padx=(4,12), pady=(6,0))

            # Add trace callbacks to crop area inputs to update the crop box in real-time
            def _on_crop_input_changed_rule(*args, av=ar_var, caw=custom_aspect_width_var, cah=custom_aspect_height_var,
                                           x1v=x1_var, y1v=y1_var, x2v=x2_var, y2v=y2_var, rule_outer=outer):
                # Check if we're already updating this rule's inputs
                if self._updating_rule_crop_inputs.get(rule_outer, False):
                    return
                try:
                    self._updating_rule_crop_inputs[rule_outer] = True

                    # Only update if this rule is active
                    if getattr(self, '_active_rule_outer', None) != rule_outer:
                        self._updating_rule_crop_inputs[rule_outer] = False
                        return

                    # Get current aspect ratio
                    aspect_val = av.get()

                    # If custom aspect, construct the aspect string from width/height inputs
                    if aspect_val == 'custom':
                        width_str = caw.get().strip()
                        height_str = cah.get().strip()
                        if width_str and height_str:
                            aspect_val = f"{width_str}:{height_str}"
                        else:
                            aspect_val = 'none'

                    # Get the original image dimensions for boundary checking
                    orig_w, orig_h = getattr(self, '_original_image_size', (1, 1))

                    # Get the values from input boxes (in original coordinates)
                    try:
                        x1_orig = int(float(x1v.get()))
                        y1_orig = int(float(y1v.get()))
                        x2_orig = int(float(x2v.get()))
                        y2_orig = int(float(y2v.get()))
                    except (ValueError, TypeError):
                        self._updating_rule_crop_inputs[rule_outer] = False
                        return

                    # ALWAYS clamp coordinates to image boundaries first
                    x1_orig = max(0, min(x1_orig, orig_w - 1))
                    y1_orig = max(0, min(y1_orig, orig_h - 1))
                    x2_orig = max(1, min(x2_orig, orig_w))
                    y2_orig = max(1, min(y2_orig, orig_h))

                    # Ensure x1 < x2 and y1 < y2
                    if x1_orig >= x2_orig:
                        x2_orig = min(x1_orig + 1, orig_w)
                    if y1_orig >= y2_orig:
                        y2_orig = min(y1_orig + 1, orig_h)

                    # If aspect ratio is set and not 'none', adjust coordinates to maintain aspect
                    if aspect_val and aspect_val != 'none':
                        asp = self._parse_aspect(aspect_val)
                        if asp is not None:
                            aw, ah = asp
                            r = aw / ah

                            # Calculate current dimensions
                            cur_w = x2_orig - x1_orig
                            cur_h = y2_orig - y1_orig

                            # Adjust to maintain aspect ratio - prefer width
                            target_h = int(round(cur_w / r))
                            if target_h != cur_h:
                                y2_orig = y1_orig + target_h
                                # Clamp y2 to image boundary
                                if y2_orig > orig_h:
                                    y2_orig = orig_h
                                    # Recalculate width to maintain aspect
                                    cur_h = y2_orig - y1_orig
                                    target_w = int(round(cur_h * r))
                                    x2_orig = x1_orig + target_w
                                    # Clamp x2 to image boundary
                                    if x2_orig > orig_w:
                                        x2_orig = orig_w

                    # Final boundary check after aspect ratio adjustment (ALWAYS applied)
                    x1_orig = max(0, min(x1_orig, orig_w - 1))
                    y1_orig = max(0, min(y1_orig, orig_h - 1))
                    x2_orig = max(x1_orig + 1, min(x2_orig, orig_w))
                    y2_orig = max(y1_orig + 1, min(y2_orig, orig_h))

                    # Update input boxes with clamped values
                    x1v.set(str(x1_orig))
                    y1v.set(str(y1_orig))
                    x2v.set(str(x2_orig))
                    y2v.set(str(y2_orig))

                    # Convert from original coordinates to preview coordinates
                    x1_prev, y1_prev = self._original_to_preview_coords(x1_orig, y1_orig)
                    x2_prev, y2_prev = self._original_to_preview_coords(x2_orig, y2_orig)

                    # Update the crop rectangle in preview space
                    if x1_prev is not None and y1_prev is not None and x2_prev is not None and y2_prev is not None:
                        self._crop_rect = [int(x1_prev), int(y1_prev), int(x2_prev), int(y2_prev)]
                        self._draw_crop_rectangle()
                except Exception:
                    pass
                finally:
                    self._updating_rule_crop_inputs[rule_outer] = False

            # Register trace callbacks
            try:
                x1_var.trace_add('write', _on_crop_input_changed_rule)
                y1_var.trace_add('write', _on_crop_input_changed_rule)
                x2_var.trace_add('write', _on_crop_input_changed_rule)
                y2_var.trace_add('write', _on_crop_input_changed_rule)
            except Exception:
                try:
                    x1_var.trace('w', _on_crop_input_changed_rule)
                    y1_var.trace('w', _on_crop_input_changed_rule)
                    x2_var.trace('w', _on_crop_input_changed_rule)
                    y2_var.trace('w', _on_crop_input_changed_rule)
                except Exception:
                    pass

            # Compression level
            comp_label = ttk.Label(inner, text='Compression Level:', font=('Segoe UI', 9, 'underline'))
            comp_label.grid(row=7, column=0, columnspan=3, sticky='w', pady=(12,2))
            comp_var = tk.StringVar(value='0')
            comp_frame = ttk.Frame(inner)
            comp_frame.grid(row=8, column=0, columnspan=3, sticky='w', pady=(0,2))
            comp_entry = ttk.Entry(comp_frame, textvariable=comp_var, width=8)
            comp_entry.grid(row=0, column=0, sticky='w')
            helper_lbl = ttk.Label(comp_frame, text='% (0 = no compression)', font=('Segoe UI', 9), foreground='#444')
            helper_lbl.grid(row=0, column=1, sticky='w', padx=(2,0))

            # Checkbox: Apply this rule to the remaining images
            apply_var = tk.BooleanVar(value=False)
            ttk.Checkbutton(inner, text='Apply this Rule to the Remaining Images.', variable=apply_var).grid(row=9, column=0, columnspan=3, sticky='w', pady=(6,0))

            # Append to tracking list so we can renumber/remove later and keep vars
            try:
                rb_entry = {
                    'outer': outer,
                    'label': rule_label,
                    'vars': {
                        'aspect': ar_var,
                        'x1': x1_var,
                        'y1': y1_var,
                        'x2': x2_var,
                        'y2': y2_var,
                        'comp': comp_var,
                        'apply': apply_var,
                        'pos': pos_var,
                        'custom_aspect_width': custom_aspect_width_var,
                        'custom_aspect_height': custom_aspect_height_var,
                    },
                    'img': None,
                }
                # If a rule object is provided (when loading a profile), populate the UI values
                if rule_obj and isinstance(rule_obj, dict):
                    try:
                        pos_raw = rule_obj.get('position') or rule_obj.get('position_number')
                        pos = None
                        try:
                            if isinstance(pos_raw, int):
                                pos = pos_raw
                            elif isinstance(pos_raw, str) and pos_raw.strip():
                                try:
                                    pos = int(float(pos_raw.strip()))
                                except Exception:
                                    pos = None
                            else:
                                pos = None
                        except Exception:
                            pos = None
                        if isinstance(pos, int) and pos > 0:
                            rb_entry['position'] = pos
                            try:
                                pos_var.set(f'Position {pos}')
                            except Exception:
                                pass
                        else:
                            # Legacy profiles may have stored a filename/path instead
                            # of a position. Try to resolve a filename to a current
                            # source image position so the preview can show it.
                            candidate_names = []
                            for k in ('file', 'filename', 'path', 'source', 'src', 'image', 'image_name', 'source_filename', 'example_path'):
                                v = rule_obj.get(k)
                                if v:
                                    try:
                                        candidate_names.append(os.path.basename(str(v)))
                                    except Exception:
                                        try:
                                            candidate_names.append(str(v))
                                        except Exception:
                                            pass
                            # also check for top-level keys that might contain a path
                            if not candidate_names:
                                for k, v in rule_obj.items():
                                    if isinstance(v, str) and (k.lower().find('file') != -1 or k.lower().find('path') != -1 or k.lower().find('image') != -1):
                                        try:
                                            candidate_names.append(os.path.basename(v))
                                        except Exception:
                                            pass
                            # collect original path-like values too for fallback if they exist
                            original_paths = []
                            for k in ('file', 'filename', 'path', 'source', 'src', 'image', 'image_name', 'source_filename', 'example_path'):
                                v = rule_obj.get(k)
                                if v and isinstance(v, str):
                                    original_paths.append(v)

                            if candidate_names:
                                try:
                                    src_list = self.get_source_images()
                                    # map basename -> index
                                    name_to_idx = {os.path.basename(p).lower(): i for i, p in enumerate(src_list)}
                                    found = None
                                    matched_cn = None
                                    for cn in candidate_names:
                                        if not cn:
                                            continue
                                        key = cn.lower()
                                        if key in name_to_idx:
                                            found = name_to_idx[key] + 1
                                            matched_cn = cn
                                            break
                                    if found:
                                        rb_entry['position'] = found
                                        try:
                                            pos_var.set(f'Position {found}')
                                        except Exception:
                                            pass
                                        # keep example path if available (use matched_cn)
                                        try:
                                            if matched_cn:
                                                rb_entry['img_temp'] = next((p for p in src_list if os.path.basename(p).lower() == os.path.basename(matched_cn).lower()), None)
                                        except Exception:
                                            pass
                                    else:
                                        # No position match; try any original absolute paths and pick the first that exists
                                        try:
                                            for op in original_paths:
                                                try:
                                                    if os.path.exists(op):
                                                        rb_entry['img_temp'] = op
                                                        break
                                                except Exception:
                                                    pass
                                        except Exception:
                                            pass
                                except Exception:
                                    pass
                        asp = rule_obj.get('aspect_ratio') or rule_obj.get('aspect')
                        if asp is not None:
                            try:
                                asp_str = str(asp)
                                # Check if it's a predefined aspect ratio
                                predefined = ['1:1', '3:4', '4:3', '4:5', '16:9', '9:16', '1.91:1', '3:2', 'none']
                                if asp_str in predefined:
                                    ar_var.set(asp_str)
                                else:
                                    # It's a custom aspect ratio - parse it into width:height
                                    ar_var.set('custom')
                                    # Try to parse the aspect ratio string
                                    if ':' in asp_str:
                                        parts = asp_str.split(':')
                                        if len(parts) == 2:
                                            custom_aspect_width_var.set(parts[0].strip())
                                            custom_aspect_height_var.set(parts[1].strip())
                                    else:
                                        # Fallback - just set it in width field
                                        custom_aspect_width_var.set(asp_str)
                            except Exception:
                                pass
                        crop = rule_obj.get('crop') or {}
                        try:
                            x1_var.set(str(int(crop.get('x1', 0))))
                            y1_var.set(str(int(crop.get('y1', 0))))
                            x2_var.set(str(int(crop.get('x2', 0))))
                            y2_var.set(str(int(crop.get('y2', 0))))
                        except Exception:
                            pass
                        try:
                            comp_var.set(str(int(rule_obj.get('compression', 0))))
                        except Exception:
                            pass
                        try:
                            apply_var.set(bool(rule_obj.get('apply_to_all_remaining', False)))
                        except Exception:
                            pass
                    except Exception:
                        pass

                self.rule_boxes.append(rb_entry)
                # mark the newly created rule active only when requested. During
                # bulk profile population we avoid activating each rule (to
                # prevent many preview attempts and user-facing error dialogs).
                if activate:
                    try:
                        self.set_active_rule(outer)
                    except Exception:
                        pass
            except Exception:
                pass

            # Scroll the canvas so the newly added rule is visible. Defer the
            # actual yview movement slightly (via `after`) so any geometry and
            # scrollregion updates from packing the new widgets complete first.
            try:
                # Only auto-scroll the new rule into view when the caller
                # explicitly requested activation. During bulk profile
                # population `activate` is False and we want the scroll to
                # remain at the top (showing Rule 1).
                if activate:
                    def _scroll_new_rule():
                        try:
                            self.rules_canvas.update_idletasks()
                            y = outer.winfo_y()
                            total = max(1, self.rules_container.winfo_height())
                            frac = min(1.0, max(0.0, y / total))
                            try:
                                self.rules_canvas.yview_moveto(frac)
                            except Exception:
                                self.rules_canvas.yview_moveto(1.0)
                        except Exception:
                            try:
                                self.rules_canvas.yview_moveto(1.0)
                            except Exception:
                                pass

                    try:
                        self.after(10, _scroll_new_rule)
                    except Exception:
                        _scroll_new_rule()
            except Exception:
                pass
        except Exception:
            pass

    def _populate_rules_from_profile(self, profile: dict):
        """Replace current rule boxes with those from the provided profile dict.

        profile is expected to be a dict with a 'rules' list of rule objects.
        """
        try:
            # destroy any existing rule widgets
            try:
                for rb in list(self.rule_boxes):
                    try:
                        o = rb.get('outer')
                        if o:
                            o.destroy()
                    except Exception:
                        pass
            except Exception:
                pass
            self.rule_boxes = []
            rules = profile.get('rules', []) if profile else []
            if not isinstance(rules, list):
                return
            for r in rules:
                try:
                    # create a new rule widget and populate its fields but do
                    # not activate its preview immediately
                    self.add_rule(rule_obj=r, activate=False)
                except Exception:
                    pass
            # activate the first rule (if any) so the preview is shown once
            try:
                if self.rule_boxes:
                    first_outer = self.rule_boxes[0].get('outer')
                    if first_outer:
                        # Defer activation so widgets and canvas have a chance to layout
                        try:
                            self.after(50, lambda o=first_outer: self.set_active_rule(o))
                        except Exception:
                            try:
                                self.set_active_rule(first_outer)
                            except Exception:
                                pass
            except Exception:
                pass
        except Exception:
            pass

    def set_active_rule(self, outer):
        """Mark the given rule outer Frame as active and apply a subtle tint.

        This updates the previously active rule to its default visuals and
        applies a slightly stronger background/highlight to the newly active
        outer frame. Robust to exceptions so it never breaks the UI.
        """
        try:
            prev = getattr(self, '_active_rule_outer', None)
            if prev is outer:
                # already active
                try:
                    # still ensure keyboard focus is on it
                    outer.focus_set()
                except Exception:
                    pass
                return

            # restore previous visuals
            if prev is not None:
                try:
                    prev.configure(bg='#e8f2ff', highlightbackground='#cfe6ff', highlightcolor='#cfe6ff')
                except Exception:
                    pass

            # apply subtle active tint to the new outer frame
            try:
                # Darker active tint for better visibility while remaining subtle
                outer.configure(bg='#b1d7ff', highlightbackground='#7bcfff', highlightcolor='#7bcfff')
                try:
                    outer.focus_set()
                except Exception:
                    pass
                self._active_rule_outer = outer

                # update the preview to show the image assigned to this rule (if any)
                try:
                    img_path = None
                    pos = None
                    aspect_var = None
                    has_saved_coords = False
                    for rb in self.rule_boxes:
                        try:
                            if rb.get('outer') is outer:
                                pos = rb.get('position')
                                # prefer explicit img but allow temporary example path
                                img_path = rb.get('img') or rb.get('img_temp')
                                aspect_var = rb.get('vars', {}).get('aspect')
                                # Check if this rule has saved crop coordinates
                                try:
                                    rvars = rb.get('vars', {}) or {}
                                    x1_val = rvars.get('x1')
                                    y1_val = rvars.get('y1')
                                    x2_val = rvars.get('x2')
                                    y2_val = rvars.get('y2')
                                    if x1_val and y1_val and x2_val and y2_val:
                                        x1_saved = int(x1_val.get() or 0)
                                        y1_saved = int(y1_val.get() or 0)
                                        x2_saved = int(x2_val.get() or 0)
                                        y2_saved = int(y2_val.get() or 0)
                                        # Consider coords "saved" if they're non-default (not full image)
                                        if x2_saved > x1_saved and y2_saved > y1_saved and not (x1_saved == 0 and y1_saved == 0):
                                            has_saved_coords = True
                                except Exception:
                                    pass
                                break
                        except Exception:
                            pass
                    if pos is not None:
                        # show by position (preferred) so preview resolves the live file for that slot
                        try:
                            try:
                                self.preview_status_var.set(f'Resolving Position {int(pos)}...')
                            except Exception:
                                pass
                            self._open_image_preview_by_position(int(pos))
                        except Exception:
                            # fallback to explicit path if position resolution fails
                            try:
                                if img_path:
                                    try:
                                        self.preview_status_var.set('Showing fallback file: ' + os.path.basename(img_path))
                                    except Exception:
                                        pass
                                    self._open_image_preview(img_path)
                            except Exception:
                                pass
                        # Only apply aspect if there are no saved coordinates
                        # (saved coords already respect the aspect ratio and shouldn't be recalculated)
                        try:
                            if not has_saved_coords and aspect_var is not None and aspect_var.get() != 'none':
                                self._apply_aspect_to_crop(aspect_var.get())
                                self._draw_crop_rectangle()
                        except Exception:
                            pass
                    elif img_path:
                        # legacy/path-only rules: show by path
                        try:
                            try:
                                self.preview_status_var.set('Showing file: ' + (os.path.basename(img_path) or '(file)'))
                            except Exception:
                                pass
                            self._open_image_preview(img_path)
                        except Exception:
                            pass
                        try:
                            if not has_saved_coords and aspect_var is not None and aspect_var.get() != 'none':
                                self._apply_aspect_to_crop(aspect_var.get())
                                self._draw_crop_rectangle()
                        except Exception:
                            pass
                    else:
                        # clear preview and show a placeholder text
                        try:
                            self.canvas.delete('all')
                            # center placeholder text
                            w = max(1, self.canvas.winfo_width())
                            h = max(1, self.canvas.winfo_height())
                            self.canvas.create_text(w//2, h//2, text=f'No image selected for this rule', fill='#666', font=('Segoe UI', 12), anchor='center')
                            try:
                                self.preview_status_var.set('No image selected for this rule')
                            except Exception:
                                pass
                        except Exception:
                            pass
                except Exception:
                    pass
                # Note: Crop value syncing is now handled entirely by _open_image_preview
                # via _sync_crop_values, so we don't duplicate it here. This prevents
                # stale _crop_rect values from overwriting the correct per-rule values.
            except Exception:
                pass
        except Exception:
            pass

    def _remove_rule(self, outer):
        """Safely remove the given rule outer Frame from the UI and tracking list.

        - Destroys the widget, removes entry from self.rule_boxes
        - Renumbers remaining rules' labels.
        - If removed rule was active, move active to the next rule (or previous)
        """
        try:
            idx = None
            for i, rb in enumerate(self.rule_boxes):
                try:
                    if rb.get('outer') is outer:
                        idx = i
                        break
                except Exception:
                    pass
            if idx is None:
                # not found; still try to destroy
                try:
                    outer.destroy()
                except Exception:
                    pass
                return

            # destroy the widget
            try:
                outer.destroy()
            except Exception:
                pass

            # remove from tracking list
            try:
                self.rule_boxes.pop(idx)
            except Exception:
                pass

            # renumber labels
            try:
                for i, rb in enumerate(self.rule_boxes, start=1):
                    try:
                        rb.get('label').configure(text=f'Rule {i}:')
                    except Exception:
                        pass
            except Exception:
                pass

            # update active selection if needed
            try:
                if getattr(self, '_active_rule_outer', None) is outer:
                    new_outer = None
                    if idx < len(self.rule_boxes):
                        new_outer = self.rule_boxes[idx]['outer']
                    elif idx - 1 >= 0 and len(self.rule_boxes) > 0:
                        new_outer = self.rule_boxes[idx - 1]['outer']
                    if new_outer:
                        self.set_active_rule(new_outer)
                    else:
                        self._active_rule_outer = None
            except Exception:
                pass

            # refresh scrollregion
            try:
                self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox('all'))
            except Exception:
                pass
        except Exception:
            pass


if __name__ == "__main__":
    # Support optional profile path as first arg
    arg = sys.argv[1] if len(sys.argv) > 1 else None
    open_cropping_window(parent=None, profile_file_path=arg)
