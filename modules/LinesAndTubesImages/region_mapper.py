r"""
Region Mapper - click-region calibration tool for image-map device pickers.

Load a line drawing, trace polygon regions with the mouse, name each one,
and export the pixel coordinates as JSON and as an AHK v2 Map() literal
ready to paste into LinesAndTubes.ahk.

Usage:
    python region_mapper.py
    (a file-open dialog will prompt you for the image)

Or from the command line:
    python region_mapper.py path\to\image.png

Controls:
    Left click         add a vertex to the region you're currently tracing
    Enter / dbl-click  finish the current region (prompts for a name)
    Backspace / u      undo the last vertex
    Mouse wheel        zoom in/out, centered on the cursor
    Right-click drag   pan around
    Middle click       reset zoom/pan back to 100%, top-left
    Click a region in the list to select it, then:
        Delete           delete the selected region
        F2 / r           rename the selected region
    Ctrl+S             export now (also happens automatically on quit)
    Esc / close window quit (auto-exports first)

Zoom/pan are purely a viewing aid -- every stored and exported coordinate
is always in the original image's native pixel space, never the zoomed
screen space. Clicking at high zoom just lets you land more precisely.

Requires: Python 3 with tkinter (standard). Pillow is used if present, which
adds support for JPG/BMP/etc. and smooth continuous zoom. Without Pillow,
only PNG/GIF/PGM/PPM work, and zoom snaps to whole integer ratios (tkinter's
built-in PhotoImage can only zoom/subsample by whole numbers).
"""

import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

try:
    from PIL import Image, ImageTk
    HAVE_PIL = True
except ImportError:
    HAVE_PIL = False


REGION_COLORS = [
    "#e6194b", "#3cb44b", "#4363d8", "#f58231", "#911eb4",
    "#42d4f4", "#f032e6", "#bfef45", "#fabed4", "#469990",
    "#dcbeff", "#9a6324", "#800000", "#aaffc3", "#000075",
]
SELECTED_OUTLINE = "#000000"


class RegionMapper:
    MIN_ZOOM = 0.1
    MAX_ZOOM = 8.0
    ZOOM_STEP = 1.15  # multiplicative change per wheel notch

    def __init__(self, root, image_path):
        self.root = root
        self.image_path = image_path
        self.root.title(f"Region Mapper - {os.path.basename(image_path)}")

        self.regions = []          # list of {"name": str, "points": [[x,y],...]} -- always native pixels
        self.current_points = []   # in-progress polygon -- always native pixels
        self.exported = False
        self.zoom = 1.0

        if HAVE_PIL:
            self.pil_image = Image.open(image_path)
            self.img_w, self.img_h = self.pil_image.size
        else:
            self.pil_image = None
            self.native_photo = tk.PhotoImage(file=image_path)
            self.img_w, self.img_h = self.native_photo.width(), self.native_photo.height()

        self.photo = None
        self.canvas_image_id = None

        # Main layout: canvas on the left, region list + status on the right
        outer = tk.Frame(root)
        outer.pack(fill="both", expand=True)

        canvas_frame = tk.Frame(outer)
        canvas_frame.pack(side="left", fill="both", expand=True)

        self.canvas = tk.Canvas(
            canvas_frame, width=min(self.img_w, 1200), height=min(self.img_h, 900),
            bg="#dddddd"
        )
        hbar = tk.Scrollbar(canvas_frame, orient="horizontal", command=self.canvas.xview)
        vbar = tk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        hbar.pack(side="bottom", fill="x")
        vbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        side = tk.Frame(outer, width=260)
        side.pack(side="right", fill="y")
        side.pack_propagate(False)

        tk.Label(side, text="Regions", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(8, 0))
        self.listbox = tk.Listbox(side, width=32, exportselection=False)
        self.listbox.pack(fill="both", expand=True, padx=8, pady=4)
        self.listbox.bind("<<ListboxSelect>>", lambda e: self.redraw())
        self.listbox.bind("<Delete>", lambda e: self.delete_selected_region())
        self.listbox.bind("<F2>", lambda e: self.rename_selected_region())
        self.listbox.bind("<Double-Button-1>", lambda e: self.rename_selected_region())

        btn_row = tk.Frame(side)
        btn_row.pack(fill="x", padx=8)
        tk.Button(btn_row, text="Rename", command=self.rename_selected_region).pack(side="left", expand=True, fill="x")
        tk.Button(btn_row, text="Delete", command=self.delete_selected_region).pack(side="left", expand=True, fill="x", padx=(6, 0))

        tk.Label(side, text="Instructions", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=8, pady=(10, 0))
        instructions = (
            "Left click: add point\n"
            "Enter / dbl-click on canvas: finish region\n"
            "Backspace: undo last point\n"
            "Wheel: zoom (toward cursor)\n"
            "Right-drag: pan\n"
            "Middle click: reset view\n"
            "Click a region in the list, then:\n"
            "  Delete: remove it\n"
            "  F2 / double-click list item: rename it\n"
            "Ctrl+S: export now\n"
            "Esc: quit (auto-exports)"
        )
        tk.Label(side, text=instructions, justify="left", font=("Segoe UI", 9)).pack(anchor="w", padx=8)

        self.status = tk.StringVar(value="Click to start a region.")
        tk.Label(side, textvariable=self.status, wraplength=240, justify="left",
                 fg="#333333", font=("Segoe UI", 9)).pack(anchor="w", padx=8, pady=8)

        self.coord_label = tk.StringVar(value="x=0, y=0  |  zoom=100%")
        tk.Label(side, textvariable=self.coord_label, font=("Consolas", 10)).pack(anchor="w", padx=8)

        export_btn = tk.Button(side, text="Export now", command=self.export)
        export_btn.pack(fill="x", padx=8, pady=(8, 8))

        # Canvas bindings: drawing
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<Double-Button-1>", lambda e: self.finish_region())
        self.canvas.bind("<Motion>", self.on_motion)

        # Canvas bindings: zoom / pan / reset
        self.canvas.bind("<MouseWheel>", self.on_wheel)     # Windows
        self.canvas.bind("<Button-4>", self.on_wheel)        # Linux scroll up
        self.canvas.bind("<Button-5>", self.on_wheel)        # Linux scroll down
        self.canvas.bind("<ButtonPress-3>", self.on_pan_start)
        self.canvas.bind("<B3-Motion>", self.on_pan_move)
        self.canvas.bind("<Button-2>", self.on_reset_view)

        # Window-level bindings (only act when the canvas -- not the listbox
        # -- has focus, so e.g. renaming doesn't also undo a canvas point)
        root.bind("<Return>", self._on_return)
        root.bind("<BackSpace>", self._on_backspace)
        root.bind("u", self._on_undo_key)
        root.bind("<Control-s>", lambda e: self.export())
        root.bind("<Escape>", lambda e: self.quit())
        root.protocol("WM_DELETE_WINDOW", self.quit)

        self._render_image()

    # -- focus-aware key handlers --
    def _widget_has_focus(self, widget):
        return self.root.focus_get() == widget

    def _on_return(self, event):
        if not self._widget_has_focus(self.listbox):
            self.finish_region()

    def _on_backspace(self, event):
        if not self._widget_has_focus(self.listbox):
            self.undo_point()

    def _on_undo_key(self, event):
        if not self._widget_has_focus(self.listbox):
            self.undo_point()

    # -- image rendering at current zoom --
    def _render_image(self):
        if self.pil_image is not None:
            disp_w = max(1, round(self.img_w * self.zoom))
            disp_h = max(1, round(self.img_h * self.zoom))
            resized = self.pil_image.resize((disp_w, disp_h), Image.LANCZOS)
            self.photo = ImageTk.PhotoImage(resized)
        else:
            self.photo = self._tk_zoom_native()
            disp_w, disp_h = self.photo.width(), self.photo.height()

        if self.canvas_image_id is None:
            self.canvas_image_id = self.canvas.create_image(0, 0, anchor="nw", image=self.photo)
        else:
            self.canvas.itemconfig(self.canvas_image_id, image=self.photo)
        self.canvas.config(scrollregion=(0, 0, disp_w, disp_h))
        self.redraw()

    def _tk_zoom_native(self):
        # No Pillow: tkinter's PhotoImage can only zoom/subsample by whole
        # integers, so snap self.zoom to the nearest achievable ratio.
        if self.zoom >= 1:
            factor = max(1, round(self.zoom))
            self.zoom = float(factor)
            return self.native_photo.zoom(factor, factor)
        else:
            factor = max(1, round(1 / self.zoom))
            self.zoom = 1.0 / factor
            return self.native_photo.subsample(factor, factor)

    # -- coordinate conversion: canvas/screen space <-> native image space --
    def canvas_xy(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        return cx / self.zoom, cy / self.zoom

    def on_motion(self, event):
        x, y = self.canvas_xy(event)
        self.coord_label.set(f"x={x:.1f}, y={y:.1f}  |  zoom={self.zoom * 100:.0f}%")

    def on_click(self, event):
        x, y = self.canvas_xy(event)
        self.current_points.append([round(x), round(y)])
        self.redraw()
        self.status.set(f"Region in progress: {len(self.current_points)} point(s). "
                         f"Enter/dbl-click to finish, Backspace to undo.")

    def undo_point(self):
        if self.current_points:
            self.current_points.pop()
            self.redraw()
            self.status.set(f"Region in progress: {len(self.current_points)} point(s).")

    # -- zoom / pan / reset --
    def on_wheel(self, event):
        if getattr(event, "num", None) == 4:
            factor = self.ZOOM_STEP
        elif getattr(event, "num", None) == 5:
            factor = 1 / self.ZOOM_STEP
        elif getattr(event, "delta", 0) > 0:
            factor = self.ZOOM_STEP
        elif getattr(event, "delta", 0) < 0:
            factor = 1 / self.ZOOM_STEP
        else:
            return

        new_zoom = max(self.MIN_ZOOM, min(self.MAX_ZOOM, self.zoom * factor))
        if new_zoom == self.zoom:
            return

        # image-space point currently under the cursor
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        img_x, img_y = cx / self.zoom, cy / self.zoom

        self.zoom = new_zoom
        self._render_image()

        # scroll so that same image-space point lands back under the cursor
        disp_w = max(1, round(self.img_w * self.zoom))
        disp_h = max(1, round(self.img_h * self.zoom))
        target_cx = img_x * self.zoom
        target_cy = img_y * self.zoom
        frac_x = (target_cx - event.x) / disp_w if disp_w else 0
        frac_y = (target_cy - event.y) / disp_h if disp_h else 0
        self.canvas.xview_moveto(max(0, min(1, frac_x)))
        self.canvas.yview_moveto(max(0, min(1, frac_y)))

        self.status.set(f"Zoom: {self.zoom * 100:.0f}%")

    def on_pan_start(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def on_pan_move(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def on_reset_view(self, event):
        self.zoom = 1.0
        self._render_image()
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)
        self.status.set("View reset to 100%.")

    # -- region finish / select / delete / rename --
    def finish_region(self):
        if len(self.current_points) < 3:
            self.status.set("Need at least 3 points to close a region.")
            return
        name = simpledialog.askstring("Region name", "Name this region (e.g. 'Right apex'):",
                                       parent=self.root)
        if not name:
            self.status.set("Cancelled -- region not saved. Points kept, keep clicking or undo.")
            return
        self.regions.append({"name": name, "points": self.current_points})
        self.listbox.insert("end", f"{name}  ({len(self.current_points)} pts)")
        self.current_points = []
        self.redraw()
        self.status.set(f"Saved '{name}'. {len(self.regions)} region(s) total.")

    def _selected_index(self):
        if not hasattr(self, "listbox"):
            return None
        sel = self.listbox.curselection()
        return sel[0] if sel else None

    def delete_selected_region(self):
        idx = self._selected_index()
        if idx is None:
            self.status.set("Select a region in the list first, then Delete.")
            return
        removed = self.regions.pop(idx)
        self.listbox.delete(idx)
        self.redraw()
        self.status.set(f"Deleted '{removed['name']}'. {len(self.regions)} region(s) left.")

    def rename_selected_region(self):
        idx = self._selected_index()
        if idx is None:
            self.status.set("Select a region in the list first, then Rename.")
            return
        old_name = self.regions[idx]["name"]
        new_name = simpledialog.askstring("Rename region", "New name:",
                                           initialvalue=old_name, parent=self.root)
        if not new_name or new_name == old_name:
            self.status.set("Rename cancelled.")
            return
        self.regions[idx]["name"] = new_name
        pt_count = len(self.regions[idx]["points"])
        self.listbox.delete(idx)
        self.listbox.insert(idx, f"{new_name}  ({pt_count} pts)")
        self.listbox.selection_set(idx)
        self.redraw()
        self.status.set(f"Renamed '{old_name}' -> '{new_name}'.")

    # -- drawing --
    def redraw(self):
        self.canvas.delete("overlay")
        selected_idx = self._selected_index()
        for i, region in enumerate(self.regions):
            color = REGION_COLORS[i % len(REGION_COLORS)]
            is_selected = (i == selected_idx)
            pts = region["points"]
            flat = [c * self.zoom for p in pts for c in p]
            self.canvas.create_polygon(
                flat, outline=(SELECTED_OUTLINE if is_selected else color), fill=color,
                stipple="gray25", width=(3 if is_selected else 2), tags="overlay"
            )
            cx = sum(p[0] for p in pts) / len(pts) * self.zoom
            cy = sum(p[1] for p in pts) / len(pts) * self.zoom
            self.canvas.create_text(
                cx, cy, text=region["name"], fill="black",
                font=("Segoe UI", 9, "bold" if is_selected else "normal"), tags="overlay"
            )
        if self.current_points:
            for (x, y) in self.current_points:
                dx, dy = x * self.zoom, y * self.zoom
                self.canvas.create_oval(dx - 3, dy - 3, dx + 3, dy + 3, fill="red", tags="overlay")
            if len(self.current_points) > 1:
                flat = [c * self.zoom for p in self.current_points for c in p]
                self.canvas.create_line(flat, fill="red", width=2, tags="overlay")

    # -- export (always native pixel coordinates, unaffected by zoom) --
    def export(self):
        if not self.regions:
            self.status.set("No completed regions to export yet.")
            return

        base = os.path.splitext(self.image_path)[0]
        json_path = base + "_regions.json"
        ahk_path = base + "_regions.ahk.txt"

        with open(json_path, "w", encoding="utf-8") as f:
            json.dump({
                "image": os.path.basename(self.image_path),
                "regions": self.regions
            }, f, indent=2)

        with open(ahk_path, "w", encoding="utf-8") as f:
            f.write(self._build_ahk_snippet())

        self.exported = True
        self.status.set(f"Exported:\n{os.path.basename(json_path)}\n{os.path.basename(ahk_path)}")

    def _build_ahk_snippet(self):
        lines = []
        lines.append("; Auto-generated by region_mapper.py -- paste this Map() into the AHK module.")
        lines.append(f"; Source image: {os.path.basename(self.image_path)}")
        lines.append("LT_ImageRegions := [")
        for region in self.regions:
            pts_str = ", ".join(f"[{x}, {y}]" for x, y in region["points"])
            safe_name = region["name"].replace('"', '""')
            lines.append(f'    Map("name", "{safe_name}", "points", [{pts_str}]),')
        lines.append("]")
        return "\n".join(lines)

    def quit(self):
        if self.regions and not self.exported:
            if messagebox.askyesno("Export before quitting?",
                                    "You have unsaved regions. Export them now?"):
                self.export()
        self.root.destroy()


def main():
    root = tk.Tk()
    root.withdraw()

    if len(sys.argv) > 1:
        image_path = sys.argv[1]
    else:
        image_path = filedialog.askopenfilename(
            title="Choose the line drawing image",
            filetypes=[("Images", "*.png *.gif *.pgm *.ppm *.jpg *.jpeg *.bmp"), ("All files", "*.*")]
        )
        if not image_path:
            return

    root.deiconify()
    app = RegionMapper(root, image_path)
    root.mainloop()


if __name__ == "__main__":
    main()
