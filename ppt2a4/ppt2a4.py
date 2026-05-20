"""
PowerPoint to A4 sheet exporter - Simple tkinter version
All features: auto-detection, merge, preview, PDF, footer, margins, numbering.
Plain white window, standard widgets.
"""

import os
import tempfile
import shutil
import threading
import time
import re
from PIL import Image, ImageDraw, ImageFont, ImageTk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import fitz
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

# ========== Constants ==========
A4_W_MM = 210
A4_H_MM = 297
DPI = 300
A4_W_PX = int(A4_W_MM / 25.4 * DPI)
A4_H_PX = int(A4_H_MM / 25.4 * DPI)
DEFAULT_MARGIN_PX = 10

# ========== Core Functions ==========
def detect_slide_orientation(img):
    w, h = img.size
    return 'landscape' if w >= h else 'portrait'

def draw_slide_number_on_cell(sheet, cell_x, cell_y, cell_w, cell_h, number):
    draw = ImageDraw.Draw(sheet)
    text = str(number)
    font_size = 28
    font = None
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except:
        try:
            font = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", font_size)
        except:
            font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    margin = 15
    x = cell_x + cell_w - text_w - margin
    y = cell_y + cell_h - text_h - margin
    for dx, dy in [(-1,-1),(-1,1),(1,-1),(1,1)]:
        draw.text((x+dx, y+dy), text, fill='black', font=font)
    draw.text((x, y), text, fill='white', font=font)

def draw_footer_text(sheet, text, margin_px):
    if not text.strip(): return
    draw = ImageDraw.Draw(sheet)
    font_size = 24
    font = None
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except:
        try:
            font = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", font_size)
        except:
            font = ImageFont.load_default()
    bbox = draw.textbbox((0,0), text, font=font)
    text_w = bbox[2]-bbox[0]
    x = (A4_W_PX - text_w)//2
    y = A4_H_PX - margin_px - 30
    for dx,dy in [(-1,-1),(-1,1),(1,-1),(1,1)]:
        draw.text((x+dx, y+dy), text, fill='black', font=font)
    draw.text((x, y), text, fill='white', font=font)

def generate_output_filename(template, pptx_path, sheet_number):
    base = os.path.splitext(os.path.basename(pptx_path))[0]
    result = template.replace("{name}", base).replace("{n}", str(sheet_number))
    result = re.sub(r'[\\/*?:"<>|]', "_", result)
    if not result.lower().endswith('.png'): result += '.png'
    return result

def export_pptx_to_png(pptx_path, output_folder, max_slides=None):
    """Export PPTX slides to PNG via PDF bridge with robust error handling."""
    os.makedirs(output_folder, exist_ok=True)
    pdf_path = os.path.join(output_folder, "_temp_slides.pdf")
    powerpoint = None
    presentation = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True
        powerpoint.WindowState = 2
        presentation = powerpoint.Presentations.Open(os.path.abspath(pptx_path), ReadOnly=True, WithWindow=False)
        time.sleep(0.5)  # allow full load

        slide_count = presentation.Slides.Count
        # Save as PDF
        presentation.SaveAs(pdf_path, 32)  # 32 = ppSaveAsPDF
        time.sleep(2)  # give time to write the file

        presentation.Close()
        presentation = None
        powerpoint.Quit()
        powerpoint = None

        # Verify PDF exists
        if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) == 0:
            raise RuntimeError("PDF file was not created or is empty")

        # Open PDF with fitz
        doc = fitz.open(pdf_path)
        num = len(doc) if max_slides is None else min(max_slides, len(doc))
        for i in range(num):
            page = doc.load_page(i)
            pix = page.get_pixmap(dpi=150)
            png_path = os.path.join(output_folder, f"slide_{i+1:03d}.png")
            pix.save(png_path)
        doc.close()
        os.remove(pdf_path)
        return num if max_slides is not None else slide_count

    except Exception as e:
        raise RuntimeError(f"Export failed: {e}")
    finally:
        if presentation:
            try: presentation.Close()
            except: pass
        if powerpoint:
            try: powerpoint.Quit()
            except: pass
        if os.path.exists(pdf_path):
            try: os.remove(pdf_path)
            except: pass

def auto_detect_slides_per_sheet(pptx_path):
    temp_dir = tempfile.mkdtemp(prefix="autodetect_")
    try:
        export_pptx_to_png(pptx_path, temp_dir, max_slides=1)
        slide_files = [f for f in os.listdir(temp_dir) if f.lower().endswith('.png')]
        if not slide_files: return 6
        img = Image.open(os.path.join(temp_dir, slide_files[0]))
        w, h = img.size
        ratio = w/h
        return 9 if ratio <= 1.2 else 6
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

def create_sheet(slide_images, cells_per_row=3, rows_per_sheet=2, output_path=None,
                 start_index=0, show_numbers=True, margin_px=DEFAULT_MARGIN_PX,
                 footer_text=""):
    total_cells = cells_per_row * rows_per_sheet
    available_w = A4_W_PX - 2*margin_px
    available_h = A4_H_PX - 2*margin_px
    cell_w = (available_w - (cells_per_row-1)*margin_px)//cells_per_row
    cell_h = (available_h - (rows_per_sheet-1)*margin_px)//rows_per_sheet
    sheet = Image.new('RGB', (A4_W_PX, A4_H_PX), 'white')
    for idx, img in enumerate(slide_images):
        if idx >= total_cells: break
        orientation = detect_slide_orientation(img)
        rotated = img.rotate(-90, expand=True) if orientation == 'landscape' else img.copy()
        rotated.thumbnail((cell_w, cell_h), Image.Resampling.LANCZOS)
        x_offset = (cell_w - rotated.width)//2
        y_offset = (cell_h - rotated.height)//2
        row = idx // cells_per_row
        col = idx % cells_per_row
        rev_col = cells_per_row - 1 - col
        x = margin_px + rev_col*(cell_w+margin_px) + x_offset
        y = margin_px + row*(cell_h+margin_px) + y_offset
        sheet.paste(rotated, (x,y))
        if show_numbers:
            global_slide_number = start_index + idx + 1
            cell_abs_x = margin_px + rev_col*(cell_w+margin_px)
            cell_abs_y = margin_px + row*(cell_h+margin_px)
            draw_slide_number_on_cell(sheet, cell_abs_x, cell_abs_y, cell_w, cell_h, global_slide_number)
    if footer_text.strip():
        draw_footer_text(sheet, footer_text, margin_px)
    if output_path:
        sheet.save(output_path, dpi=(DPI,DPI))
    return sheet

def convert_pngs_to_pdf(png_paths, pdf_output_path):
    c = canvas.Canvas(pdf_output_path, pagesize=A4)
    for png_path in png_paths:
        img = ImageReader(png_path)
        c.setPageSize((595.28, 841.89))
        c.drawImage(img, 0, 0, width=595.28, height=841.89, preserveAspectRatio=True, anchor='c')
        c.showPage()
    c.save()

def process_single_presentation(pptx_path, slides_per_sheet, output_dir, show_numbers,
                                margin_px, filename_template, footer_text, export_pdf,
                                progress_callback=None):
    temp_dir = tempfile.mkdtemp(prefix="ppt_export_")
    try:
        if progress_callback: progress_callback("Exporting slides...", 0)
        slide_count = export_pptx_to_png(pptx_path, temp_dir)
        if slide_count == 0: raise ValueError("No slides found")
        slide_files = sorted([f for f in os.listdir(temp_dir) if f.lower().endswith('.png')])
        if not slide_files:
            raise RuntimeError("No PNG files created from presentation")
        slide_images = [Image.open(os.path.join(temp_dir,f)) for f in slide_files]
        cells_per_row, rows_per_sheet = 3, (2 if slides_per_sheet==6 else 3)
        per_sheet = cells_per_row*rows_per_sheet
        total_sheets = (len(slide_images)+per_sheet-1)//per_sheet
        output_sheets = []
        for sheet_idx in range(total_sheets):
            start = sheet_idx*per_sheet
            end = min(start+per_sheet, len(slide_images))
            sheet_slides = slide_images[start:end]
            out_path = os.path.join(output_dir, generate_output_filename(filename_template, pptx_path, sheet_idx+1))
            create_sheet(sheet_slides, cells_per_row, rows_per_sheet, out_path,
                         start_index=start, show_numbers=show_numbers,
                         margin_px=margin_px, footer_text=footer_text)
            output_sheets.append(out_path)
            if progress_callback:
                progress_callback(f"Sheet {sheet_idx+1}/{total_sheets}", int((sheet_idx+1)/total_sheets*80))
        if export_pdf:
            pdf_path = os.path.join(output_dir, os.path.splitext(os.path.basename(pptx_path))[0]+"_combined.pdf")
            convert_pngs_to_pdf(output_sheets, pdf_path)
            if progress_callback: progress_callback("PDF done", 100)
            return output_sheets, pdf_path
        return output_sheets, None
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

def process_presentations(pptx_paths, slides_per_sheet, output_dir, show_numbers,
                          margin_px, filename_template, footer_text, export_pdf,
                          merge=False, progress_callback=None):
    if merge:
        temp_dirs = []
        all_slide_images = []
        try:
            for idx, p in enumerate(pptx_paths):
                td = tempfile.mkdtemp(prefix=f"merge_{idx}_")
                temp_dirs.append(td)
                if progress_callback: progress_callback(f"Exporting {os.path.basename(p)}...", int(idx/len(pptx_paths)*50))
                export_pptx_to_png(p, td)
                slide_files = sorted([f for f in os.listdir(td) if f.lower().endswith('.png')])
                for f in slide_files:
                    all_slide_images.append(Image.open(os.path.join(td,f)))
            cells_per_row, rows_per_sheet = 3, (2 if slides_per_sheet==6 else 3)
            per_sheet = cells_per_row*rows_per_sheet
            total_sheets = (len(all_slide_images)+per_sheet-1)//per_sheet
            output_sheets = []
            for sheet_idx in range(total_sheets):
                start = sheet_idx*per_sheet
                end = min(start+per_sheet, len(all_slide_images))
                sheet_slides = all_slide_images[start:end]
                base_name = os.path.splitext(os.path.basename(pptx_paths[0]))[0]+"_merged"
                out_name = filename_template.replace("{name}", base_name).replace("{n}", str(sheet_idx+1))
                out_name = re.sub(r'[\\/*?:"<>|]', "_", out_name)
                if not out_name.endswith('.png'): out_name += '.png'
                out_path = os.path.join(output_dir, out_name)
                create_sheet(sheet_slides, cells_per_row, rows_per_sheet, out_path,
                             start_index=start, show_numbers=show_numbers,
                             margin_px=margin_px, footer_text=footer_text)
                output_sheets.append(out_path)
                if progress_callback:
                    progress_callback(f"Creating sheet {sheet_idx+1}/{total_sheets}", int(50+(sheet_idx+1)/total_sheets*40))
            if export_pdf:
                pdf_path = os.path.join(output_dir, "merged_combined.pdf")
                convert_pngs_to_pdf(output_sheets, pdf_path)
                if progress_callback: progress_callback("PDF created", 100)
                return output_sheets, pdf_path
            return output_sheets, None
        finally:
            for td in temp_dirs:
                shutil.rmtree(td, ignore_errors=True)
    else:
        all_sheets, last_pdf = [], None
        for i, p in enumerate(pptx_paths):
            if progress_callback:
                progress_callback(f"Processing {os.path.basename(p)} ({i+1}/{len(pptx_paths)})", int(i/len(pptx_paths)*100))
            sheets, pdf = process_single_presentation(p, slides_per_sheet, output_dir, show_numbers,
                                                      margin_px, filename_template, footer_text, export_pdf,
                                                      progress_callback)
            all_sheets.extend(sheets)
            if pdf: last_pdf = pdf
        return all_sheets, last_pdf

# ========== Simple Tkinter GUI ==========
class App:
    def __init__(self, root):
        self.root = root
        
        root.title("PPT to A4 Sheet Generator")
        
        root.geometry("620x700")
        root.resizable(False, False)

        # File list
        tk.Label(root, text="Presentations (.pptx):").pack(pady=(10,0))
        self.file_listbox = tk.Listbox(root, height=5, selectmode=tk.EXTENDED)
        self.file_listbox.pack(fill=tk.X, padx=20, pady=(5,5))
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=(0,5))
        tk.Button(btn_frame, text="Add Files", command=self.add_files).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Remove Selected", command=self.remove_selected).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Clear All", command=self.clear_files).pack(side=tk.LEFT, padx=5)

        # Merge option
        self.merge_var = tk.BooleanVar(value=False)
        tk.Checkbutton(root, text="Merge all presentations into one sequence (global numbering)", variable=self.merge_var).pack(pady=(5,5))

        # Slides per sheet
        tk.Label(root, text="Slides per A4 sheet:").pack(pady=(5,0))
        self.slides_mode = tk.StringVar(value="auto")
        mode_frame = tk.Frame(root)
        mode_frame.pack()
        tk.Radiobutton(mode_frame, text="Auto", variable=self.slides_mode, value="auto").pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(mode_frame, text="6 slides", variable=self.slides_mode, value="6").pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(mode_frame, text="9 slides", variable=self.slides_mode, value="9").pack(side=tk.LEFT, padx=10)

        # Numbering
        self.numbering_var = tk.BooleanVar(value=True)
        tk.Checkbutton(root, text="Show slide numbers", variable=self.numbering_var).pack(pady=(5,0))

        # Margin
        tk.Label(root, text="Margin between slides (pixels):").pack(pady=(5,0))
        self.margin_var = tk.IntVar(value=DEFAULT_MARGIN_PX)
        tk.Spinbox(root, from_=0, to=50, textvariable=self.margin_var, width=10).pack()

        # Footer text
        tk.Label(root, text="Footer text (optional):").pack(pady=(5,0))
        self.footer_var = tk.StringVar()
        tk.Entry(root, textvariable=self.footer_var, width=60).pack(pady=(2,5))

        # Filename template
        tk.Label(root, text="Output filename template (use {name}, {n}):").pack(pady=(5,0))
        self.template_var = tk.StringVar(value="{name}_sheet_{n}")
        tk.Entry(root, textvariable=self.template_var, width=50).pack(pady=(2,5))

        # PDF export
        self.pdf_var = tk.BooleanVar(value=True)
        tk.Checkbutton(root, text="Also create a single PDF (all sheets combined)", variable=self.pdf_var).pack(pady=(5,5))

        # Output folder
        tk.Label(root, text="Output Folder:").pack(pady=(5,0))
        out_frame = tk.Frame(root)
        out_frame.pack(fill=tk.X, padx=20)
        self.out_path_var = tk.StringVar()
        tk.Entry(out_frame, textvariable=self.out_path_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(out_frame, text="Browse", command=self.browse_output).pack(side=tk.RIGHT, padx=(5,0))

        # Preview button
        tk.Button(root, text="Preview First Sheet (for selected presentation)", command=self.preview_selected).pack(pady=(10,5))

        # Progress bar
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.progress.pack(pady=(10,5))
        self.status_label = tk.Label(root, text="Ready")
        self.status_label.pack()

        # Start button
        self.start_button = tk.Button(root, text="Generate Sheets", command=self.start_processing, bg="lightgreen", width=20)
        self.start_button.pack(pady=(15,10))

        self.file_paths = []

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PowerPoint", "*.pptx")])
        for f in files:
            if f not in self.file_paths:
                self.file_paths.append(f)
                self.file_listbox.insert(tk.END, os.path.basename(f))

    def remove_selected(self):
        selected = self.file_listbox.curselection()
        for i in reversed(selected):
            self.file_listbox.delete(i)
            del self.file_paths[i]

    def clear_files(self):
        self.file_listbox.delete(0, tk.END)
        self.file_paths.clear()

    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.out_path_var.set(folder)

    def preview_selected(self):
        if not self.file_paths:
            messagebox.showerror("Error", "No presentations added")
            return
        selected = self.file_listbox.curselection()
        idx = selected[0] if selected else 0
        pptx_path = self.file_paths[idx]
        mode = self.slides_mode.get()
        if mode == "auto":
            try:
                slides_per_sheet = auto_detect_slides_per_sheet(pptx_path)
            except:
                slides_per_sheet = 6
        else:
            slides_per_sheet = int(mode)
        margin_px = self.margin_var.get()
        show_numbers = self.numbering_var.get()
        footer_text = self.footer_var.get().strip()
        preview_first_sheet(pptx_path, slides_per_sheet, margin_px, show_numbers, footer_text, self.root)

    def update_progress(self, message, percent):
        self.status_label.config(text=message)
        self.progress['value'] = percent
        self.root.update_idletasks()

    def start_processing(self):
        if not self.file_paths:
            messagebox.showerror("Error", "Please add at least one .pptx file")
            return
        output_dir = self.out_path_var.get()
        if not output_dir:
            messagebox.showerror("Error", "Please select an output folder")
            return
        mode = self.slides_mode.get()
        if mode == "auto":
            try:
                slides_per_sheet = auto_detect_slides_per_sheet(self.file_paths[0])
                self.status_label.config(text=f"Auto-detected: {slides_per_sheet} slides per sheet")
                self.root.update()
            except Exception as e:
                messagebox.showerror("Error", f"Auto-detection failed: {e}\nUsing 6")
                slides_per_sheet = 6
        else:
            slides_per_sheet = int(mode)
        show_numbers = self.numbering_var.get()
        margin_px = self.margin_var.get()
        footer_text = self.footer_var.get().strip()
        filename_template = self.template_var.get().strip()
        if not filename_template:
            filename_template = "{name}_sheet_{n}"
        export_pdf = self.pdf_var.get()
        merge = self.merge_var.get()

        self.start_button.config(state=tk.DISABLED, text="Processing...")
        self.progress['value'] = 0
        self.status_label.config(text="Starting...")

        def process():
            try:
                sheets, pdf_path = process_presentations(
                    pptx_paths=self.file_paths,
                    slides_per_sheet=slides_per_sheet,
                    output_dir=output_dir,
                    show_numbers=show_numbers,
                    margin_px=margin_px,
                    filename_template=filename_template,
                    footer_text=footer_text,
                    export_pdf=export_pdf,
                    merge=merge,
                    progress_callback=self.update_progress
                )
                msg = f"Generated {len(sheets)} sheet(s) in:\n{output_dir}"
                if pdf_path:
                    msg += f"\n\nPDF saved as:\n{pdf_path}"
                messagebox.showinfo("Success", msg)
                self.update_progress("Done", 100)
            except Exception as e:
                messagebox.showerror("Error", str(e))
                self.update_progress("Error", 0)
            finally:
                self.start_button.config(state=tk.NORMAL, text="Generate Sheets")

        threading.Thread(target=process, daemon=True).start()

def preview_first_sheet(pptx_path, slides_per_sheet, margin_px, show_numbers, footer_text, parent):
    temp_dir = tempfile.mkdtemp(prefix="preview_")
    try:
        max_slides = slides_per_sheet
        slide_count = export_pptx_to_png(pptx_path, temp_dir, max_slides=max_slides)
        if slide_count == 0:
            messagebox.showerror("Error", "No slides found")
            return
        slide_files = sorted([f for f in os.listdir(temp_dir) if f.lower().endswith('.png')])
        slide_images = [Image.open(os.path.join(temp_dir, f)) for f in slide_files]
        cells_per_row, rows_per_sheet = 3, (2 if slides_per_sheet == 6 else 3)
        sheet_img = create_sheet(slide_images, cells_per_row, rows_per_sheet,
                                 start_index=0, show_numbers=show_numbers,
                                 margin_px=margin_px, footer_text=footer_text)
        preview_w = min(800, A4_W_PX)
        preview_h = int(A4_H_PX * (preview_w / A4_W_PX))
        preview_img = sheet_img.resize((preview_w, preview_h), Image.Resampling.LANCZOS)
        preview_window = tk.Toplevel(parent)
        preview_window.title("Preview - First Sheet")
        preview_window.geometry(f"{preview_w+20}x{preview_h+60}")
        photo = ImageTk.PhotoImage(preview_img)
        label = tk.Label(preview_window, image=photo)
        label.image = photo
        label.pack(padx=10, pady=10)
        tk.Button(preview_window, text="Close", command=preview_window.destroy).pack(pady=10)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()