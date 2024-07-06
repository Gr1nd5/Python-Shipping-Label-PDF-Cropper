import fitz
import tkinter as tk
from tkinter import filedialog, messagebox, Scale
from PIL import Image, ImageTk
import os

class PDFCropper:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Cropper")

        self.input_pdf_path = None
        self.offset_x = 0
        self.offset_y = 0
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.crop_rect_id = None
        self.drag_offset_x = 0
        self.drag_offset_y = 0

        self.create_widgets()
        self.show_status("Ready")

    def create_widgets(self):
        self.canvas_frame = tk.Frame(self.root)
        self.canvas_frame.pack(expand=True, fill="both")

        self.canvas = tk.Canvas(self.canvas_frame, bg="white", highlightthickness=0)
        self.canvas.pack(expand=True, fill="both")

        self.status_bar = tk.Label(self.root, text="Ready", bd=1, relief="sunken", anchor="w", bg="lightgrey")
        self.status_bar.pack(side="bottom", fill="x")

        self.crop_width = Scale(self.root, from_=1, to=8, orient=tk.HORIZONTAL, label="Width (inches)", resolution=0.1, command=self.update_preview)
        self.crop_width.set(4)
        self.crop_width.pack()

        self.crop_height = Scale(self.root, from_=1, to=10, orient=tk.HORIZONTAL, label="Height (inches)", resolution=0.1, command=self.update_preview)
        self.crop_height.set(6)
        self.crop_height.pack()

        self.open_button = tk.Button(self.root, text="Open PDF", command=self.open_file)
        self.open_button.pack()

        self.save_button = tk.Button(self.root, text="Save Cropped PDF", command=self.save_file)
        self.save_button.pack()

        self.canvas.bind("<Button-1>", self.start_drag)
        self.canvas.bind("<B1-Motion>", self.drag_pdf)
        self.canvas.bind("<ButtonRelease-1>", self.update_position)

    def start_drag(self, event):
        try:
            self.drag_start_x = event.x
            self.drag_start_y = event.y

            crop_x1 = self.offset_x * 72
            crop_y1 = self.offset_y * 72
            crop_x2 = (self.offset_x + self.crop_width.get()) * 72
            crop_y2 = (self.offset_y + self.crop_height.get()) * 72

            center_x = (crop_x1 + crop_x2) / 2
            center_y = (crop_y1 + crop_y2) / 2

            self.drag_offset_x = self.drag_start_x - center_x
            self.drag_offset_y = self.drag_start_y - center_y
        except Exception as e:
            self.show_status(f"Failed to start drag: {str(e)}", error=True)

    def open_file(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
            if file_path:
                self.input_pdf_path = file_path
                self.update_preview()
                self.show_status("PDF opened successfully")
        except Exception as e:
            self.show_status(f"Failed to open PDF: {str(e)}", error=True)

    def save_file(self):
        try:
            if self.input_pdf_path:
                input_file_name = os.path.basename(self.input_pdf_path)
                input_file_dir = os.path.dirname(self.input_pdf_path)
                output_file_path = os.path.join(input_file_dir, f"cropped_{input_file_name}")

                if os.path.exists(output_file_path):
                    overwrite = messagebox.askyesno("File Exists", f"File '{output_file_path}' already exists. Overwrite?")
                    if not overwrite:
                        return

                self.crop_pdf(self.input_pdf_path, output_file_path, self.crop_width.get(), self.crop_height.get(), self.offset_x, self.offset_y)
                self.show_status(f"PDF cropped and saved successfully to {output_file_path}")

        except Exception as e:
            self.show_status(f"Failed to save PDF: {str(e)}", error=True)

    def update_preview(self, event=None):
        try:
            if self.input_pdf_path:
                pix = self.show_cropped_page(self.input_pdf_path, self.crop_width.get(), self.crop_height.get(), self.offset_x, self.offset_y)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                self.img_tk = ImageTk.PhotoImage(img)
                self.canvas.create_image(0, 0, anchor="nw", image=self.img_tk)
                self.canvas.config(scrollregion=self.canvas.bbox("all"))

                if self.crop_rect_id:
                    self.canvas.delete(self.crop_rect_id)
                crop_x1 = self.offset_x * 72
                crop_y1 = self.offset_y * 72
                crop_x2 = (self.offset_x + self.crop_width.get()) * 72
                crop_y2 = (self.offset_y + self.crop_height.get()) * 72
                self.crop_rect_id = self.canvas.create_rectangle(crop_x1, crop_y1, crop_x2, crop_y2, outline="black", width=2)
        except Exception as e:
            self.show_status(f"Failed to update preview: {str(e)}", error=True)

    def drag_pdf(self, event):
        try:
            new_center_x = event.x - self.drag_offset_x
            new_center_y = event.y - self.drag_offset_y

            self.offset_x = new_center_x / 72 - self.crop_width.get() / 2
            self.offset_y = new_center_y / 72 - self.crop_height.get() / 2

            self.constrain_offsets()
            self.update_preview()
        except Exception as e:
            self.show_status(f"Failed to drag PDF: {str(e)}", error=True)

    def update_position(self, event):
        try:
            self.drag_start_x = 0
            self.drag_start_y = 0
        except Exception as e:
            self.show_status(f"Failed to update position: {str(e)}", error=True)

    def constrain_offsets(self):
        try:
            if self.input_pdf_path:
                pdf_document = fitz.open(self.input_pdf_path)
                page = pdf_document.load_page(0)
                media_box = page.rect
                max_x = media_box.width / 72 - self.crop_width.get()
                max_y = media_box.height / 72 - self.crop_height.get()
                self.offset_x = max(0, min(self.offset_x, max_x))
                self.offset_y = max(0, min(self.offset_y, max_y))
                pdf_document.close()
        except Exception as e:
            self.show_status(f"Failed to constrain offsets: {str(e)}", error=True)

    def show_cropped_page(self, input_pdf_path, crop_width, crop_height, offset_x, offset_y, page_number=0):
        try:
            pdf_document = fitz.open(input_pdf_path)
            page = pdf_document.load_page(page_number)
            crop_rect = fitz.Rect(offset_x * 72, offset_y * 72, (offset_x + crop_width) * 72, (offset_y + crop_height) * 72)
            page.set_cropbox(crop_rect)
            pix = page.get_pixmap()
            pdf_document.close()
            return pix
        except Exception as e:
            self.show_status(f"Failed to show cropped page: {str(e)}", error=True)

    def crop_pdf(self, input_pdf_path, output_pdf_path, crop_width, crop_height, offset_x, offset_y):
        try:
            pdf_document = fitz.open(input_pdf_path)
            crop_rect = fitz.Rect(offset_x * 72, offset_y * 72, (offset_x + crop_width) * 72, (offset_y + crop_height) * 72)
            for page_number in range(len(pdf_document)):
                page = pdf_document.load_page(page_number)
                page.set_cropbox(crop_rect)
            pdf_document.save(output_pdf_path)
            pdf_document.close()
        except Exception as e:
            self.show_status(f"Failed to crop PDF: {str(e)}", error=True)

    def show_status(self, message, error=False):
        if error:
            messagebox.showerror("Error", message)
        else:
            self.status_bar.config(text=message)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCropper(root)
    root.mainloop()
