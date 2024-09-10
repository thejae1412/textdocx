from win32com.client import DispatchEx
from tkinter import filedialog, messagebox, ttk, Canvas, Scrollbar
from datetime import datetime
from pywintypes import com_error
import tkinter as tk
import os
import sys
import pythoncom
import locale
import win32print
import tempfile
import tkinter.font as tkfont
import threading
import subprocess

def load_fitz():
    import fitz
    return fitz

def load_image():
    from PIL import Image
    return Image

def load_pdf_viewer():
    from tkPDFViewer import tkPDFViewer as pdf
    return pdf
def load_imagetk():
    from PIL import ImageTk
    return ImageTk


fonts = 'TH SarabunIT๙',16
fontBT = 'TH SarabunIT๙',24,'bold'
locale.setlocale(locale.LC_ALL, 'th_TH.utf8')
current_date = datetime.now()
buddhist_year = current_date.year + 543
formatted_date = current_date.strftime(f'%B {buddhist_year}')
place_date = current_date.strftime(f'{buddhist_year}')
thai_date = current_date.strftime(f'%d %B {buddhist_year}')

def print_pdf(pdf_path):
    try:
        # แสดงหน้าต่างเลือกเครื่องพิมพ์
        printer_name = win32print.GetDefaultPrinter()
        printer_info = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
        printer_names = [printer[2] for printer in printer_info]
        
        printer_window = tk.Toplevel()
        printer_window.title("เลือกเครื่องพิมพ์")
        printer_window.geometry("300x200")
        
        tk.Label(printer_window, text="เลือกเครื่องพิมพ์:", font=fonts).pack(pady=10)
        
        printer_var = tk.StringVar(printer_window)
        printer_var.set(printer_name)
        
        printer_menu = ttk.OptionMenu(printer_window, printer_var, printer_name, *printer_names)
        printer_menu.pack(pady=10)
        
        def on_print():
            selected_printer = printer_var.get()
            win32print.SetDefaultPrinter(selected_printer)
            os.startfile(pdf_path, "print")
            printer_window.destroy()
        
        ttk.Button(printer_window, text="พิมพ์", command=on_print).pack(pady=10)
        
    except com_error as e:
        messagebox.showerror("Error", f"เกิดข้อผิดพลาดในการพิมพ์: {str(e)}")

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

TEMPLATES_FOLDER = r"\\10.22.36.31\templates\government"

# ปรับขนาด canvas window เมื่อ form_frame เปลี่ยนขนาด
def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))
    # ปรับความกว้างของ canvas window ให้เท่ากับความกว้างของ canvas
    if event.width > canvas.winfo_width():
        canvas.itemconfig(canvas_window, width=event.width)
    canvas.itemconfig(canvas_window, width=canvas.winfo_width())

def on_canvas_configure(event):
    # ปรับความกว้างของ frame ให้เท่ากับความกว้างของ canvas ถ้า canvas กว้างกว่า frame
    if event.width > form_frame.winfo_width():
        canvas.itemconfig(canvas_window, width=event.width)

def update_progress(progress_var, progress_bar, value):
    progress_var.set(value)
    progress_bar.update()

def on_mousewheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
   
def create_scrolled_text(parent, height=4, **kwargs):
    frame = ttk.Frame(parent)
    text = tk.Text(frame, height=height, font=fonts, bd=1, relief="solid", padx=5, pady=5, **kwargs)
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
    text.configure(yscrollcommand=scrollbar.set)
    text.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(0, weight=1)
    
    # เพิ่ม event binding สำหรับ mouse wheel
    text.bind("<MouseWheel>", lambda e: _on_mousewheel(e, text))
    
    return frame, text

def _on_mousewheel(event, widget):
    widget.yview_scroll(int(-1*(event.delta/120)), "units")
    # ป้องกันการเลื่อน event ไปยัง widget อื่น
    return "break"

def show_pdf_preview(pdf_path):
    fitz = load_fitz()  # เรียกใช้ฟังก์ชัน lazy load
    Image = load_image()
    # โหลดและปรับขนาดไอคอน
    zoom_in_icon = Image.open(resource_path("img/zoomin.png"))
    zoom_in_icon = zoom_in_icon.resize((24, 24), Image.LANCZOS)
    zoom_in_photo = ImageTk.PhotoImage(zoom_in_icon)

    zoom_out_icon = Image.open(resource_path("img/zoomout.png"))
    zoom_out_icon = zoom_out_icon.resize((24, 24), Image.LANCZOS)
    zoom_out_photo = ImageTk.PhotoImage(zoom_out_icon)

    # โหลดและปรับขนาดไอคอนสำหรับปุ่ม Print
    print_icon = Image.open(resource_path("img/printer.png"))
    print_icon = print_icon.resize((24, 24), Image.LANCZOS)
    print_photo = ImageTk.PhotoImage(print_icon)
    
    preview_window = tk.Toplevel()
    preview_window.title("Preview")
    preview_window.geometry("1024x768")
    
    
    button_frame = tk.Frame(preview_window)
    button_frame.pack(side=tk.TOP, fill=tk.X, pady=(10,0))
    # จัดการ Layout ของ button_frame
    button_frame.grid_columnconfigure(0, weight=1)
    button_frame.grid_columnconfigure(1, weight=1)
    style = ttk.Style()
    style.configure('Zoom.TButton', padding=5)
  
    main_frame = ttk.Frame(preview_window)
    main_frame.pack(fill=tk.BOTH, expand=True)

    canvas = tk.Canvas(main_frame)
    canvas.grid(row=0, column=0, sticky="nsew")

    v_scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
    v_scrollbar.grid(row=0, column=1, sticky="ns")

    main_frame.grid_rowconfigure(0, weight=1)
    main_frame.grid_columnconfigure(0, weight=1)

    print_frame = ttk.Frame(preview_window)
    print_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 10))
    
    frame = tk.Frame(canvas)
    canvas_window = canvas.create_window((0, 0), window=frame, anchor="nw")


    doc = fitz.open(pdf_path)
    zoom = 1.0
    rotation = 0
    images = []  # เก็บภาพทั้งหมดไว้ในลิสต์

    def update_page(zoom_factor):
        nonlocal zoom, doc, images
        zoom *= zoom_factor
        for widget in frame.winfo_children():
            widget.destroy()
        images.clear()
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom).prerotate(rotation))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            photo = ImageTk.PhotoImage(img)
            images.append(photo)
        draw_images()

    def draw_images():
        for photo in images:
            label = tk.Label(frame, image=photo)
            label.image = photo
            label.pack()
        frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    def zoom_in():
        update_page(1.2)

    def zoom_out():
        update_page(0.8)

    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
        # ปรับความกว้างของ canvas window ให้เท่ากับความกว้างของ frame ถ้า frame กว้างกว่า canvas
        if event.width > canvas.winfo_width():
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.itemconfig(canvas_window, width=canvas.winfo_width())
        
    def on_canvas_configure(event):
         # ปรับความกว้างของ frame ให้เท่ากับความกว้างของ canvas ถ้า canvas กว้างกว่า frame
        if event.width > frame.winfo_width():
            canvas.itemconfig(canvas_window, width=event.width)

    def print_current_pdf():
        print_pdf(pdf_path)



    frame.bind("<Configure>", on_frame_configure)
    canvas.bind("<Configure>", on_canvas_configure)

    def on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        return "break"
    def on_shift_mousewheel(event):
        canvas.xview_scroll(int(-1*(event.delta/120)), "units")

    frame.bind("<Configure>", on_frame_configure)
    canvas.bind("<Configure>", on_canvas_configure)
    canvas.bind_all("<MouseWheel>", on_mousewheel)
    canvas.bind_all("<Shift-MouseWheel>", on_shift_mousewheel)


     # สร้างปุ่ม Zoom In
    zoom_in_button = ttk.Button(button_frame, image=zoom_in_photo, command=zoom_in, style='Zoom.TButton')
    zoom_in_button.image = zoom_in_photo
    zoom_in_button.grid(row=0, column=0, sticky="e", padx=(0, 0))

    # สร้างปุ่ม Zoom Out
    zoom_out_button = ttk.Button(button_frame, image=zoom_out_photo, command=zoom_out, style='Zoom.TButton')
    zoom_out_button.image = zoom_out_photo
    zoom_out_button.grid(row=0, column=1, sticky="w", padx=(0, 0))

     # สร้างปุ่ม Print
    print_button = ttk.Button(print_frame, image=print_photo, text="Print", compound="left", command=print_current_pdf, style='Print.TButton')
    print_button.image = print_photo
    print_button.pack(pady=10)
    
    update_page(1)  # แสดงหน้าแรกของ PDF

    preview_window.protocol("WM_DELETE_WINDOW", lambda: (doc.close, preview_window.destroy()))

def show_preview():
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
 
    def preview_task():
        try:
            # สร้างไฟล์ Word ชั่วคราว
            update_progress(progress_var, progress_bar, 10)
            temp_docx = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
            temp_docx_path = temp_docx.name
            temp_docx.close()

            # บันทึกข้อมูลลงในไฟล์ Word ชั่วคราว
            update_progress(progress_var, progress_bar, 30)
            template_name = selected_template.get()
            template_path = os.path.join(TEMPLATES_FOLDER, template_name)
            data = get_form_data()
            fill_word_template(template_path, temp_docx_path, data, progress_var, progress_bar, start=30, end=60)

            # แปลงไฟล์ Word เป็น PDF โดยไม่ต้องใช้ Word Application
            update_progress(progress_var, progress_bar, 70)
            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_pdf_path = temp_pdf.name
            temp_pdf.close()
            save_as_pdf(temp_docx_path, temp_pdf_path)

            # แสดง Print Preview
            update_progress(progress_var, progress_bar, 90)
            show_pdf_preview(temp_pdf_path)

            # ลบไฟล์ชั่วคราว
            os.unlink(temp_docx_path)
            #os.unlink(temp_pdf_path)

            update_progress(progress_var, progress_bar, 100)
        finally:
            progress_bar.grid_remove()

    threading.Thread(target=preview_task, daemon=True).start()

def create_entry(label_text, row, initial_text=None, readonly=False):
    tk.Label(form_frame, text=label_text, font=fonts).grid(row=row, column=0, padx=10, pady=10, sticky="w")
    entry = ttk.Entry(form_frame, font=fonts, style='Custom.TEntry')
    if initial_text:
        entry.insert(0, initial_text)
    if readonly:
        entry.config(state='readonly')
    entry.grid(row=row, column=1, padx=10, pady=10, sticky="ew")
    return entry

def create_scrolled_text(label_text, row, height=4):
    tk.Label(form_frame, text=label_text, font=fonts).grid(row=row, column=0, padx=10, pady=10, sticky="w")
    frame = ttk.Frame(form_frame)
    text = tk.Text(frame, height=height, font=fonts, bd=1, relief="solid", padx=5, pady=5)
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
    text.configure(yscrollcommand=scrollbar.set)
    text.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")
    frame.grid(row=row, column=1, padx=10, pady=10, sticky="ew")
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(0, weight=1)
    return text
 
def get_form_data(template_name):
    template_name = selected_template.get()
    data = {}
    
    if template_name == "1.หนังสือภายใน รพ.ราชวิถี2(รังสิต).docx":
        data = {
            "{name}": entry_name.get(),
            "{place}": place_date,
            "{date}": formatted_date,
            "{topic}": entry_topic.get(),
            "{datai}": entry_datai.get("1.0", tk.END).strip(),
            "{dataii}": entry_dataii.get("1.0", tk.END).strip(),
            "{dataiii}": entry_dataiii.get("1.0", tk.END).strip(),
            "{name1}": entry_name1.get(),
            "{rank1}": entry_rank1.get(),
            "{name2}": entry_name2.get(),
            "{rank2}": entry_rank2.get(),
            "{name3}": entry_name3.get(),
            "{rank3}": entry_rank3.get()
        }
    elif template_name == "2.หนังสือภายใน รพ.ราชวิถี2(รังสิต) สธ1.docx":
        data = {
            "{name}": entry_name.get(),
            "{place}": entry_place.get(),
            "{date}": formatted_date,
            "{topic}": entry_topic.get(),
            "{datai}": entry_datai.get("1.0", tk.END).strip(),
            "{dataii}": entry_dataii.get("1.0", tk.END).strip(),
            "{name1}": entry_name1.get(),
            "{rank1}": entry_rank1.get(),
            "{name2}": entry_name2.get(),
            "{rank2}": entry_rank2.get(),
            "{name3}": entry_name3.get(),
            "{rank3}": entry_rank3.get()
        }
    elif template_name == "3.หนังสือภายใน รพ.ราชวิถี2(รังสิต) สธ2 + สำเนา.docx":
        data = {
            "{name}": entry_name.get(),
            "{to}": entry_to.get(),
            "{date}": formatted_date,
            "{topic}": entry_topic.get(),
            "{datai}": entry_datai.get("1.0", tk.END).strip(),
            "{dataii}": entry_dataii.get("1.0", tk.END).strip(),
            "{dataiii}": entry_dataiii.get("1.0", tk.END).strip(),
            "{name1}": entry_name1.get(),
            "{rank1}": entry_rank1.get(),
        }
    elif template_name == "4.หนังสือภายนอก รพ.ราชวิถี2(รังสิต) สธ1.docx":
        data = {
            "{name}": entry_name.get(),
            "{place}": entry_place.get(),
            "{date}": formatted_date,
            "{topic}": entry_topic.get(),
            "{datai}": entry_datai.get("1.0", tk.END).strip(),
            "{dataii}": entry_dataii.get("1.0", tk.END).strip(),
            "{name1}": entry_name1.get(),
            "{rank1}": entry_rank1.get(),
            "{name2}": entry_name2.get(),
            "{rank2}": entry_rank2.get(),
            "{name3}": entry_name3.get(),
            "{rank3}": entry_rank3.get()
        }
    elif template_name == "5.หนังสือภายนอก รพ.ราชวิถี2(รังสิต) สธ2+สำเนา.docx":
        data = {
            "{date}": formatted_date,
            "{topic}": entry_topic.get(),
            "{to}": entry_to.get(),
            "{ref}": entry_ref.get(),
            "{attach}": entry_attach.get(),
            "{datai}": entry_datai.get("1.0", tk.END).strip(),
            "{dataii}": entry_dataii.get("1.0", tk.END).strip(),
            "{dataiii}": entry_dataiii.get("1.0", tk.END).strip(),
            "{name1}": entry_name1.get(),
            "{rank1}": entry_rank1.get()
        }

    return data
    
def on_submit(file_format):
    global last_saved_file
    template_name = selected_template.get()
    if not template_name:
        messagebox.showerror("Error", "กรุณาเลือกเทมเพลต.")
        return

    template_path = os.path.join(TEMPLATES_FOLDER, template_name)
    temp_docx_path = os.path.join(os.environ['TEMP'], "temp.docx")
    data = get_form_data(template_name)

    # ตรวจสอบว่าได้ข้อมูลครบถ้วนหรือไม่
    if not data:
        messagebox.showerror("Error", "กรุณากรอกข้อมูลให้ครบถ้วน.")
        return
    
    template_path = os.path.join(TEMPLATES_FOLDER, template_name)
    temp_docx_path = os.path.join(os.environ['TEMP'], "temp.docx")
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
    def save_file():
        try:
            update_progress(progress_var, progress_bar, 30)
            if file_format == 'docx':
                save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
                if save_path:
                    fill_word_template(template_path, save_path, data, progress_var, progress_bar)
                    messagebox.showinfo("Success", "บันทึกไฟล์เป็น DOC สำเร็จ.")
            elif file_format == 'pdf':
                update_progress(progress_var, progress_bar, 60)
                fill_word_template(template_path, temp_docx_path, data, progress_var, progress_bar)
                save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
                if save_path:
                    save_as_pdf(temp_docx_path, save_path)
                    os.remove(temp_docx_path)
                    messagebox.showinfo("Success", "บันทึกไฟล์เป็น PDF สำเร็จ.")
            update_progress(progress_var, progress_bar, 100)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            progress_bar.grid_remove()

    threading.Thread(target=save_file, daemon=True).start()

def fill_word_template(source_filepath, target_filepath, data, progress_var, progress_bar, start=0, end=100):
    pythoncom.CoInitialize()
    word = DispatchEx('Word.Application')
    word.visible = False
    word.Documents.Add()
    doc = None
    doc = word.Documents.Open(source_filepath)
    try:
        shapes = doc.Shapes
        name_shapes = {}
        rank_shapes = {}
        total_shapes = len(shapes)
        
        # Avoid division by zero
        if total_shapes > 0:
            for i, shape in enumerate(shapes):
                if shape.TextFrame.HasText:
                    for j in range(1, 4):  # สำหรับ name1-3 และ rank1-3
                        if f"{{name{j}}}" in shape.TextFrame.TextRange.Text:
                            if j not in name_shapes:
                                name_shapes[j] = []
                            name_shapes[j].append(shape)
                        elif f"{{rank{j}}}" in shape.TextFrame.TextRange.Text:
                            if j not in rank_shapes:
                                rank_shapes[j] = []
                            rank_shapes[j].append(shape)
                update_progress(progress_var, progress_bar, start + (i / total_shapes) * ((end - start) / 2))
            
        # แทนที่ข้อความและจัดตำแหน่ง
        total_replacements = sum(len(shapes) for shapes in name_shapes.values()) + sum(len(shapes) for shapes in rank_shapes.values())
        replacement_count = 0
        
        for i in range(1, 4):
            name_key = f"{{name{i}}}"
            rank_key = f"{{rank{i}}}"
                
            if i in name_shapes and i in rank_shapes:
                for name_shape, rank_shape in zip(name_shapes[i], rank_shapes[i]):
                    if name_key in data and rank_key in data:
                        name_shape.TextFrame.TextRange.Text = name_shape.TextFrame.TextRange.Text.replace(name_key, data[name_key])
                        rank_shape.TextFrame.TextRange.Text = rank_shape.TextFrame.TextRange.Text.replace(rank_key, data[rank_key])
                            
                        # จัดตำแหน่ง rank_shape ให้อยู่กึ่งกลางใต้ name_shape
                        rank_shape.Left = name_shape.Left + (name_shape.Width - rank_shape.Width) / 2
            replacement_count += 2
            if total_replacements > 0:
                update_progress(progress_var, progress_bar, 50 + (replacement_count / total_replacements) * 25)           

        # ส่วนที่เหลือของการแทนที่ข้อความในส่วนอื่นๆ ของเอกสาร
        if total_shapes > 0:
            for i, shape in enumerate(shapes):
                if shape.Type == 6:  # Group shape
                    for sub_shape in shape.GroupItems:
                        if sub_shape.TextFrame.HasText:
                            text = sub_shape.TextFrame.TextRange.Text
                            for key, value in data.items():
                                if key in text:
                                    text = text.replace(key, value)
                            sub_shape.TextFrame.TextRange.Text = text
                elif shape.TextFrame.HasText:
                    text = shape.TextFrame.TextRange.Text
                    for key, value in data.items():
                        if key in text:
                            text = text.replace(key, value)
                    shape.TextFrame.TextRange.Text = text
                update_progress(progress_var, progress_bar, start + ((end - start) / 2) + (i / total_shapes) * ((end - start) / 2))
            
        for para in doc.Paragraphs:
            for key, value in data.items():
                if key in para.Range.Text:
                    para.Range.Text = para.Range.Text.replace(key, value)
                    para.Range.ParagraphFormat.Alignment = 3  # 3 คือ wdAlignParagraphCenter
                    

        file_extension = os.path.splitext(target_filepath)[1].lower()
        if file_extension == '.docx':
            doc.SaveAs(target_filepath, FileFormat=16)
        else:
            doc.SaveAs(target_filepath)
    except Exception as e:
        raise RuntimeError(f"Failed to fill Word template: {str(e)}")
    
    finally:
        if doc:
            doc.Close()
    
def save_as_pdf(docx_path, pdf_path):
    try:
        # สร้างคำสั่งแปลงไฟล์โดยใช้ docx2pdf แบบไม่แสดงหน้าต่าง console
        convert_command = f'docx2pdf "{docx_path}" "{pdf_path}"'
        
        # ใช้ subprocess เพื่อเรียกคำสั่งและซ่อนหน้าต่าง console
        subprocess.run(convert_command, shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
    
    except Exception as e:
        raise RuntimeError(f"Failed to save as PDF: {str(e)}")

def load_templates():
    if not os.path.exists(TEMPLATES_FOLDER):
        os.makedirs(TEMPLATES_FOLDER)
    templates = [f for f in os.listdir(TEMPLATES_FOLDER) if f.endswith(".docx")]
    return templates

def update_form(form_frame, event=None):

    global entry_name, entry_place, entry_date, entry_topic, entry_ref, entry_attach
    global entry_to, entry_datai, entry_dataii, entry_dataiii
    global entry_name1, entry_name2, entry_name3
    global entry_rank1, entry_rank2, entry_rank3

    
    for widget in form_frame.winfo_children():
        widget.destroy()
    
    template_name = selected_template.get()

    if template_name == "1.หนังสือภายใน รพ.ราชวิถี2(รังสิต).docx":
        entry_name = create_entry("ส่วนราชการ:", 0)
        entry_place = create_entry("ที่:", 1, initial_text=place_date, readonly=True)
        entry_date = create_entry("วันที่:", 2, initial_text=formatted_date, readonly=True)
        entry_topic = create_entry("เรื่อง:", 3)
        entry_datai = create_scrolled_text("ภาคเหตุ:", 4)
        entry_dataii = create_scrolled_text("ภาคความประสงค์:", 5)
        entry_dataiii = create_scrolled_text("ภาคสรุป:", 6)
        entry_name1 = create_entry("ชื่อ - สกุล:", 7)
        entry_rank1 = create_entry("หัวหน้างาน/หัวหน้าหน่วยงาน:", 8)
        entry_name2 = create_entry("ชื่อ - สกุล:", 9)
        entry_rank2 = create_entry("หัวหน้ากลุ่มงาน:", 10)
        entry_name3 = create_entry("ชื่อ - สกุล:", 11)
        entry_rank3 = create_entry("รองผู้อำนวยการฯ:", 12)
        
    elif template_name == "2.หนังสือภายใน รพ.ราชวิถี2(รังสิต) สธ1.docx":
        entry_name = create_entry("ส่วนราชการ:", 0)
        entry_place = create_entry("ที่:", 1, initial_text=place_date, readonly=True)
        entry_date = create_entry("วันที่:", 2, initial_text=formatted_date, readonly=True)
        entry_topic = create_entry("เรื่อง:", 3)
        entry_datai = create_scrolled_text("ภาคเหตุ:", 4)
        entry_dataii = create_scrolled_text("ภาคความประสงค์:", 5)
        entry_name1 = create_entry("ชื่อ - สกุล:", 7)
        entry_rank1 = create_entry("หัวหน้างาน/หัวหน้าหน่วยงาน:", 8)
        entry_name2 = create_entry("ชื่อ - สกุล:", 9)
        entry_rank2 = create_entry("หัวหน้ากลุ่มงาน:", 10)
        entry_name3 = create_entry("ชื่อ - สกุล:", 11)
        entry_rank3 = create_entry("รองผู้อำนวยการฯ:", 12)
        
    elif template_name == "3.หนังสือภายใน รพ.ราชวิถี2(รังสิต) สธ2 + สำเนา.docx":
        entry_name = create_entry("ส่วนราชการ:", 0)
        entry_date = create_entry("วันที่:", 1, initial_text=formatted_date, readonly=True)
        entry_topic = create_entry("เรื่อง:", 2)
        entry_to = create_entry("เรียน:", 3)
        entry_datai = create_scrolled_text("ภาคเหตุ:", 4)
        entry_dataii = create_scrolled_text("ภาคความประสงค์:", 5)
        entry_name1 = create_entry("ชื่อ - สกุล:", 6)
        entry_rank1 = create_entry("หัวหน้างาน/หัวหน้าหน่วยงาน:", 7)
        
    elif template_name == "4.หนังสือภายนอก รพ.ราชวิถี2(รังสิต) สธ1.docx":
        entry_name = create_entry("ส่วนราชการ:", 0)
        entry_place = create_entry("ที่:", 1, initial_text=place_date, readonly=True)
        entry_date = create_entry("วันที่:", 2, initial_text=formatted_date, readonly=True)
        entry_topic = create_entry("เรื่อง:", 3)
        entry_datai = create_scrolled_text("ภาคเหตุ:", 4)
        entry_dataii = create_scrolled_text("ภาคความประสงค์:", 5)
        entry_name1 = create_entry("ชื่อ - สกุล:", 7)
        entry_rank1 = create_entry("หัวหน้างาน/หัวหน้าหน่วยงาน:", 8)
        entry_name2 = create_entry("ชื่อ - สกุล:", 9)
        entry_rank2 = create_entry("หัวหน้ากลุ่มงาน:", 10)
        entry_name3 = create_entry("ชื่อ - สกุล:", 11)
        entry_rank3 = create_entry("รองผู้อำนวยการฯ:", 12)

    elif template_name == "5.หนังสือภายนอก รพ.ราชวิถี2(รังสิต) สธ2+สำเนา.docx":
        entry_date = create_entry("วันที่:", 0, initial_text=formatted_date, readonly=True)
        entry_topic = create_entry("เรื่อง:", 1)
        entry_to = create_entry("เรียน:", 2)
        entry_ref = create_entry("อ้างอึง:", 3)   
        entry_attach = create_entry("สิ่งที่แนบมาด้วย:", 4)   
        entry_datai = create_scrolled_text("ภาคเหตุ:", 5)
        entry_dataii = create_scrolled_text("ภาคความประสงค์:", 6)
        entry_dataiii = create_scrolled_text("ภาคสรุป:", 7)
        entry_name1 = create_entry("ชื่อ - สกุล:", 8)
        entry_rank1 = create_entry("หัวหน้างาน/หัวหน้าหน่วยงาน:", 9)
        
    else:
        tk.Label(form_frame, text="ไม่พบฟอร์มสำหรับเทมเพลตนี้", font=fonts).grid(row=0, column=0, padx=10, pady=10)

    root.after(100, lambda: canvas.configure(scrollregion=canvas.bbox("all")))

ImageTk = load_imagetk()
Image = load_image()    
# สร้างหน้าต่างหลักของ GUI
root = tk.Tk()
root.title("Auto Fill Word Files")
icon = Image.open(resource_path("img/program_ico.ico"))
photo = ImageTk.PhotoImage(icon)
root.wm_iconphoto(True, photo)
root.geometry("1024x768")
root.minsize(1024, 768)
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)
# สร้าง main frame เพื่อเป็นคอนเทนเนอร์หลัก
main_frame = ttk.Frame(root)
main_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
main_frame.grid_rowconfigure(1, weight=1)
main_frame.grid_columnconfigure(0, weight=1)

# ตัวแปรเก็บไฟล์ล่าสุดที่บันทึก
last_saved_file = None

# โหลดรายชื่อไฟล์เทมเพลต
templates = load_templates()
selected_template = tk.StringVar(root)
selected_template.set(templates[0] if templates else "")
option_font = tkfont.Font(family="TH SarabunIT๙", size=16)
option_frame = ttk.Frame(main_frame)
option_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
option_frame.grid_columnconfigure(1, weight=1)
ttk.Label(option_frame, text="เลือกเทมเพลต:", font=fonts).grid(row=0, column=0, padx=10, pady=10, sticky="w")
option_menu = ttk.OptionMenu(option_frame, selected_template, templates[0] if templates else "", *templates, style='Custom.TMenubutton')
option_menu.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
menu = option_menu['menu']
menu.config(font=("TH SarabunIT๙", 14), bg="#f0f0f0", activebackground="#4a7abc", activeforeground="white")

# สร้าง Canvas และ Scrollbar
canvas_frame = ttk.Frame(main_frame)
canvas_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 10))
canvas_frame.grid_rowconfigure(0, weight=1)
canvas_frame.grid_columnconfigure(0, weight=1)

btn_frame = ttk.Frame(main_frame)
btn_frame.grid(row=2, column=0, sticky="ew", pady=(20, 0))
btn_frame.grid_columnconfigure(0, weight=1)
btn_frame.grid_columnconfigure(1, weight=1)

canvas = Canvas(canvas_frame)
canvas.grid(row=0,column=0,sticky="nsew")
scrollbar = Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
scrollbar.grid(row=0, column=1, sticky="ns")
canvas.configure(yscrollcommand=scrollbar.set)

# สร้าง frame ภายใน canvas สำหรับใส่ widget ของฟอร์ม
form_frame = tk.Frame(canvas)
form_frame.grid_columnconfigure(1,weight=1)


docx_path = resource_path("img/docx_ico.png")
docx_ico = Image.open(docx_path)
docx_ico = docx_ico.resize((40,40),Image.ADAPTIVE)
docx_bt = ImageTk.PhotoImage(docx_ico)
btn_submit_docx = ttk.Button(btn_frame,image=docx_bt , compound="left",text="บันทึกเป็น DOCX", command=lambda: on_submit('docx'), width=20, style='DOCX.TButton')
btn_submit_docx.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

pdf_path = resource_path("img/pdf_ico.png")
pdf_ico = Image.open(pdf_path)
pdf_ico = pdf_ico.resize((40,40),Image.ADAPTIVE)
pdf_bt = ImageTk.PhotoImage(pdf_ico)
btn_submit_pdf = ttk.Button(btn_frame, image=pdf_bt , compound="left",text="บันทึกเป็น PDF", command=lambda: on_submit('pdf'), width=20, style='PDF.TButton')
btn_submit_pdf.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

pv_path = resource_path("img/pv_ico.png")
pv_ico = Image.open(pv_path)
pv_ico = pv_ico.resize((40,40),Image.ADAPTIVE)
pv_bt = ImageTk.PhotoImage(pv_ico)
btn_preview = ttk.Button(btn_frame, image=pv_bt , compound="left", text="Preview", command=show_preview, width=20, style='Preview.TButton')
btn_preview.grid(row=0, column=2, padx=10, pady=10, sticky="ew")

# สร้างฟอร์มสำหรับกรอกข้อมูล
style = ttk.Style()
style.theme_use('clam')
style.configure("TScrollbar", background="#f0f0f0", troughcolor="#d0d0d0", width=10, arrowsize=13)
style.configure('Custom.TMenubutton', 
                background="#f0f0f0",
                foreground="black",
                padding=10,
                font=fonts,
                relief="flat",
                width=45)

style.map('Custom.TMenubutton',
          background=[('active', '#e0e0e0'), ('pressed', '#d0d0d0')],
          relief=[('pressed', 'groove'), ('!pressed', 'ridge')])


style.configure('Custom.TEntry',
                foreground = 'black',
                background = 'white',
                fieldbackground = 'white',
                borderwidth = 5,
                relief = 'flat',
                font = fonts,
                padding = 5)

style.map('Custom.TEntry',
          foreground = [('disabled', 'gray')],
          fieldbackground = [('disabled', '#f0f0f0')])

style.configure('DOCX.TButton',
                background='#1976D2',
                foreground='black',
                font=fontBT,
                padding=10)

style.map('DOCX.TButton',
          background=[('active', '#1976D2'), ('disabled', '#a0a0a0')],
          foreground=[('disabled', '#d0d0d0')])


style.configure('PDF.TButton',
                background='#AA3939',
                foreground='black',
                font=fontBT,
                padding=10)

style.map('PDF.TButton',
          background=[('active', '#AA3939'), ('disabled', '#a0a0a0')],
          foreground=[('disabled', '#d0d0d0')])


style.configure('Custom.DateEntry',
                foreground = 'black',
                background = 'white',
                fieldbackground = 'white',
                borderwidth = 1,
                relief = 'solid',
                arrowcolor = 'black',
                font = fonts)

style.configure('Preview.TButton',
                background='#F57C00',  # สีส้ม
                foreground='black',
                font=fontBT,
                padding=10)

style.map('Preview.TButton',
          background=[('active', '#F57C00'), ('disabled', '#a0a0a0')],
          foreground=[('disabled', '#d0d0d0')])

style.configure('Print.TButton',
                background='#4CAF50',  # สีเขียว
                foreground='black',
                font=fontBT,
                padding=10)

style.map('Print.TButton',
          background=[('active', '#45a049'), ('disabled', '#a0a0a0')],
          foreground=[('disabled', '#d0d0d0')])

style.configure("TCombobox", font=fonts)
style.configure("TButton", font=fontBT)


form_frame.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
selected_template.trace("w", lambda *args: update_form(form_frame))
update_form(form_frame)
root.after(100, lambda: canvas.configure(scrollregion=canvas.bbox("all")))
root.update_idletasks()
form_frame.update_idletasks()
canvas.config(scrollregion=canvas.bbox("all"))
canvas.bind("<MouseWheel>", on_mousewheel)
canvas.bind("<Configure>", on_canvas_configure)
canvas_window = canvas.create_window((0, 0), window=form_frame, anchor="nw")

root.mainloop()