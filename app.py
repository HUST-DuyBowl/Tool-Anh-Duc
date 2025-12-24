# app.py
# v21.0 - Thêm cấu hình Vector Store ID động trên giao diện.

import tkinter as tk
from tkinter import filedialog, Text, END, messagebox
import os
import threading
import shutil
import time

# Import các hàm từ filechinh.py (File này giữ nguyên)
from filechinh import (
    analyze_and_sort_folder,
    taopdftudocx,
    taofile_from_images,
    append_row_to_xlsx
)

# Import hàm phantichfile từ chatgpt.py
from chatgpt import phantichfile

class AutoFileClassifierApp:

    def __init__(self, master):
        self.master = master
        self.master.title('Tool Phân Loại')
        self.master.geometry('800x720') # Tăng chiều cao một chút để chứa thêm ô input
        
        self.folder_path = tk.StringVar()
        self.destination_path = tk.StringVar()
        self.xlsx_path = tk.StringVar()
        
        # *** MỚI: Biến lưu Vector Store ID ***
        self.vector_id = tk.StringVar()
        self.vector_id.set("vs_68277a920bc88191b52d03839b4c71f7") # Giá trị mặc định cũ của bạn
        
        self.stop_event = threading.Event()
        
        self._create_widgets()

    def log_message(self, message):
        if self.master.winfo_exists():
            self.log_area.config(state=tk.NORMAL)
            self.log_area.insert(END, message + '\n')
            self.log_area.see(END)
            self.log_area.config(state=tk.DISABLED)
            self.master.update_idletasks()
            
    def _organize_and_move_folder(self, destination_base_path, item_folder_path, lop, kythi, chuong):
        if not lop:
            self.log_message("  -> Không có thông tin Lớp, không thể phân loại.")
            return item_folder_path 

        class_folder_name = f"Toán {lop}"
        class_folder_path = os.path.join(destination_base_path, class_folder_name)

        subfolder_name = ""
        kythi_normalized = kythi.upper() if kythi else ""
        exam_map = { "GK I": "GK1", "GK II": "GK2", "CK I": "CK1", "CK II": "CK2" }

        if kythi_normalized in exam_map:
            subfolder_name = exam_map[kythi_normalized]
        elif "KSCL" in kythi_normalized:
            subfolder_name = "KSCL"
        elif chuong:
            subfolder_name = f"Chuong {chuong}"
        else:
            subfolder_name = "Chua Phan Loai"

        destination_path = os.path.join(class_folder_path, subfolder_name)
        
        if not os.path.isdir(destination_path):
            destination_path = os.path.join(class_folder_path, "Chua Phan Loai")
            
        try:
            item_folder_name = os.path.basename(item_folder_path)
            final_destination_path = os.path.join(destination_path, item_folder_name)

            shutil.move(item_folder_path, destination_path)
            time.sleep(0.5)
            self.log_message(f"  -> Đã di chuyển vào: '{os.path.relpath(destination_path, destination_base_path)}'")
            
            return final_destination_path 

        except Exception as move_error:
            self.log_message(f"!! Lỗi khi di chuyển vào thư mục đích: {move_error}")
            return item_folder_path

    def _setup_destination_folders(self, base_path):
        self.log_message("Đang kiểm tra/tạo cấu trúc thư mục đích...")
        try:
            classes = [f"Toán {i}" for i in range(6, 13)] 
            subfolders = ["GK1", "GK2", "CK1", "CK2", "KSCL", "Chua Phan Loai"] + \
                         [f"Chuong {i}" for i in range(1, 11)]
            
            for class_name in classes:
                class_path = os.path.join(base_path, class_name)
                os.makedirs(class_path, exist_ok=True)
                for sub_name in subfolders:
                    os.makedirs(os.path.join(class_path, sub_name), exist_ok=True)
            
            self.log_message("Cấu trúc thư mục đã sẵn sàng.")
            return True
        except Exception as e:
            self.log_message(f"!! Lỗi nghiêm trọng khi tạo thư mục: {e}")
            messagebox.showerror("Lỗi Tạo Thư Mục", f"Không thể tạo cấu trúc thư mục tại:\n{base_path}\nLỗi: {e}")
            return False

    def process_files_logic(self):
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete('1.0', END)
        self.log_area.config(state=tk.DISABLED)
        
        try:
            apikey = self.apikeychatgpt_entry.get()
            source_folder = self.folder_path.get()
            destination_folder = self.destination_path.get()
            local_xlsx_file = self.xlsx_path.get()
            
            # *** MỚI: Lấy Vector ID từ giao diện ***
            current_vector_id = self.vector_id.get().strip()

            if not all([source_folder, destination_folder, local_xlsx_file, apikey, current_vector_id]):
                messagebox.showerror('Lỗi', 'Vui lòng điền đủ 5 thông tin (bao gồm Vector Store ID).')
                return

            if not self._setup_destination_folders(destination_folder):
                self.log_message("Tác vụ bị hủy do không thể tạo thư mục đích.")
                return 

            if self.stop_event.is_set(): return 

            self.log_message("Bắt đầu quá trình phân loại file...")
            pdf_doc_pairs, pdf_only, doc_only, archive_files, jpg_groups, single_jpgs, skipped_files = analyze_and_sort_folder(source_folder)
            
            self.log_message(f"Phân loại hoàn tất. Bắt đầu tổ chức và xử lý...")
            self.log_message("-" * 40)

            # --- Vòng lặp xử lý ---
            for filename in pdf_only:
                if self.stop_event.is_set(): break
                self.log_message(f"\n[PDF] Đang xử lý: {filename}.pdf")
                try:
                    new_folder_path = os.path.join(source_folder, filename)
                    os.makedirs(new_folder_path, exist_ok=True)
                    source_path = os.path.join(source_folder, f"{filename}.pdf")
                    dest_path = os.path.join(new_folder_path, f"{filename}.pdf")
                    shutil.move(source_path, dest_path)
                    
                    # [TRUYỀN VECTOR ID VÀO HÀM]
                    lop, chuong, bai, namhoc, kythi = phantichfile(apikey, dest_path, current_vector_id)
                    
                    final_folder_path = self._organize_and_move_folder(destination_folder, new_folder_path, lop, kythi, chuong)
                    final_pdf_path = os.path.join(final_folder_path, os.path.basename(dest_path))
                    row = ['', filename, final_pdf_path, '', '', '', namhoc, kythi, lop, chuong, bai]
                    append_row_to_xlsx(local_xlsx_file, row)
                    self.log_message(f"=> Đã phân tích và lưu vào Excel.")
                except Exception as e:
                    self.log_message(f"!! Lỗi khi xử lý {filename}.pdf: {e}")
            
            if self.stop_event.is_set(): return

            for filename, ext in doc_only:
                 if self.stop_event.is_set(): break
                 self.log_message(f"\n[DOC] Đang xử lý: {filename}{ext}")
                 try:
                    new_folder_path = os.path.join(source_folder, filename)
                    os.makedirs(new_folder_path, exist_ok=True)
                    source_path = os.path.join(source_folder, f"{filename}{ext}")
                    dest_path_doc = os.path.join(new_folder_path, f"{filename}{ext}")
                    shutil.move(source_path, dest_path_doc)
                    temp_pdf_path = taopdftudocx(filename, dest_path_doc)
                    if temp_pdf_path and os.path.exists(temp_pdf_path):
                        # [TRUYỀN VECTOR ID VÀO HÀM]
                        lop, chuong, bai, namhoc, kythi = phantichfile(apikey, temp_pdf_path, current_vector_id)
                        
                        final_folder_path = self._organize_and_move_folder(destination_folder, new_folder_path, lop, kythi, chuong)
                        final_pdf_path = os.path.join(final_folder_path, os.path.basename(temp_pdf_path))
                        final_doc_path = os.path.join(final_folder_path, os.path.basename(dest_path_doc))
                        row = ['', filename, final_pdf_path, final_doc_path, '', '', namhoc, kythi, lop, chuong, bai]
                        append_row_to_xlsx(local_xlsx_file, row)
                        self.log_message(f"=> Đã phân tích và lưu vào Excel.")
                 except Exception as e:
                    self.log_message(f"!! Lỗi khi xử lý {filename}{ext}: {e}")

            if self.stop_event.is_set(): return

            for filename, ext in pdf_doc_pairs:
                if self.stop_event.is_set(): break
                self.log_message(f"\n[Cặp PDF+DOC] Đang xử lý: {filename}")
                try:
                    new_folder_path = os.path.join(source_folder, filename)
                    os.makedirs(new_folder_path, exist_ok=True)
                    source_pdf = os.path.join(source_folder, f"{filename}.pdf")
                    dest_pdf = os.path.join(new_folder_path, f"{filename}.pdf")
                    shutil.move(source_pdf, dest_pdf)
                    source_doc = os.path.join(source_folder, f"{filename}{ext}")
                    dest_doc = os.path.join(new_folder_path, f"{filename}{ext}")
                    shutil.move(source_doc, dest_doc)
                    
                    # [TRUYỀN VECTOR ID VÀO HÀM]
                    lop, chuong, bai, namhoc, kythi = phantichfile(apikey, dest_pdf, current_vector_id)
                    
                    final_folder_path = self._organize_and_move_folder(destination_folder, new_folder_path, lop, kythi, chuong)
                    final_pdf_path = os.path.join(final_folder_path, os.path.basename(dest_pdf))
                    final_doc_path = os.path.join(final_folder_path, os.path.basename(dest_doc))
                    row = ['', filename, final_pdf_path, final_doc_path, '', '', namhoc, kythi, lop, chuong, bai]
                    append_row_to_xlsx(local_xlsx_file, row)
                    self.log_message(f"=> Đã phân tích và lưu vào Excel.")
                except Exception as e:
                    self.log_message(f"!! Lỗi khi xử lý cặp {filename}: {e}")

            if self.stop_event.is_set(): return

            for filename in archive_files:
                if self.stop_event.is_set(): break
                self.log_message(f"\n[File Nén] Đang xử lý: {filename}")
                try:
                    base_name = os.path.splitext(filename)[0]
                    new_folder_path = os.path.join(source_folder, base_name)
                    os.makedirs(new_folder_path, exist_ok=True)
                    source_path = os.path.join(source_folder, filename)
                    dest_path = os.path.join(new_folder_path, filename)
                    shutil.move(source_path, dest_path)
                    lop_guess = next((str(i) for i in range(6, 13) if str(i) in filename), "")
                    final_folder_path = self._organize_and_move_folder(destination_folder, new_folder_path, lop_guess, "", "")
                    final_archive_path = os.path.join(final_folder_path, os.path.basename(dest_path))
                    row = ['', base_name, '', '', final_archive_path, '', '', '', lop_guess, '', '']
                    append_row_to_xlsx(local_xlsx_file, row)
                    self.log_message(f"=> Đã lưu vào Excel.")
                except Exception as e:
                    self.log_message(f"!! Lỗi khi xử lý {filename}: {e}")

            if self.stop_event.is_set(): return

            all_jpg_items = {**jpg_groups, **{os.path.splitext(f)[0]: [f] for f in single_jpgs}}
            for base_name, files in all_jpg_items.items():
                if self.stop_event.is_set(): break
                log_prefix = "[Nhóm JPG]" if len(files) > 1 else "[Ảnh JPG đơn]"
                self.log_message(f"\n{log_prefix} Đang xử lý: {base_name}")
                try:
                    new_folder_path = os.path.join(source_folder, base_name)
                    os.makedirs(new_folder_path, exist_ok=True)
                    image_paths = [os.path.join(source_folder, f) for f in files]
                    new_pdf_path, new_docx_path = taofile_from_images(base_name, image_paths, new_folder_path)
                    if new_pdf_path and new_docx_path:
                        for img_path in image_paths:
                            shutil.move(img_path, new_folder_path)
                        
                        # [TRUYỀN VECTOR ID VÀO HÀM]
                        lop, chuong, bai, namhoc, kythi = phantichfile(apikey, new_pdf_path, current_vector_id)
                        
                        final_folder_path = self._organize_and_move_folder(destination_folder, new_folder_path, lop, kythi, chuong)
                        final_pdf_path = os.path.join(final_folder_path, os.path.basename(new_pdf_path))
                        final_doc_path = os.path.join(final_folder_path, os.path.basename(new_docx_path))
                        row = ['', base_name, final_pdf_path, final_doc_path, '', '', namhoc, kythi, lop, chuong, bai]
                        append_row_to_xlsx(local_xlsx_file, row)
                        self.log_message(f"=> Đã phân tích và lưu vào Excel.")
                except Exception as e:
                    self.log_message(f"!! Lỗi khi xử lý ảnh {base_name}: {e}")

            if self.stop_event.is_set():
                self.log_message("\n--- TÁC VỤ ĐÃ DỪNG ---")
            else:
                if skipped_files:
                    self.log_message("-" * 40)
                    self.log_message("\nCác file sau đã bị bỏ qua (không di chuyển):")
                    for f in skipped_files:
                        self.log_message(f"- {f}")
                self.log_message("\n--- HOÀN TẤT TOÀN BỘ QUÁ TRÌNH ---")

        except Exception as e:
            self.log_message(f'\n!!! LỖI NGHIÊM TRỌNG: {e} !!!')
        finally:
            if self.master.winfo_exists():
                self.process_button.config(state=tk.NORMAL)
                self.stop_button.config(state=tk.DISABLED)

    def _create_widgets(self):
        input_frame = tk.LabelFrame(self.master, text='1. Chọn Đường Dẫn', padx=10, pady=10)
        input_frame.pack(fill='x', padx=10, pady=5)
        input_frame.grid_columnconfigure(1, weight=1)
        
        tk.Button(input_frame, text='Chọn Thư Mục Chứa File', command=self.select_source_folder).grid(row=0, column=0, sticky='ew', pady=2)
        tk.Entry(input_frame, textvariable=self.folder_path, state='readonly').grid(row=0, column=1, sticky='ew', padx=5)
        
        tk.Button(input_frame, text='Chọn Thư Mục Lưu Trữ', command=self.select_destination_folder).grid(row=1, column=0, sticky='ew', pady=2)
        tk.Entry(input_frame, textvariable=self.destination_path, state='readonly').grid(row=1, column=1, sticky='ew', padx=5)
        
        tk.Button(input_frame, text='Chọn File Excel Lưu Log', command=self.select_xlsx_file).grid(row=2, column=0, sticky='ew', pady=2)
        tk.Entry(input_frame, textvariable=self.xlsx_path, state='readonly').grid(row=2, column=1, sticky='ew', padx=5)
        
        config_frame = tk.LabelFrame(self.master, text='2. Điền Thông Tin Cấu Hình', padx=10, pady=10)
        config_frame.pack(fill='x', padx=10, pady=5)
        config_frame.grid_columnconfigure(1, weight=1)
        
        # Dòng 1: API Key
        tk.Label(config_frame, text='ChatGPT API Key:').grid(row=0, column=0, sticky='w', pady=2)
        self.apikeychatgpt_entry = tk.Entry(config_frame, width=70, show='*')
        self.apikeychatgpt_entry.grid(row=0, column=1, sticky='ew', padx=5)
        
        # *** MỚI: Dòng 2: Vector Store ID ***
        tk.Label(config_frame, text='Vector Store ID:').grid(row=1, column=0, sticky='w', pady=2)
        tk.Entry(config_frame, textvariable=self.vector_id, width=70).grid(row=1, column=1, sticky='ew', padx=5)

        button_frame = tk.Frame(self.master)
        button_frame.pack(pady=10)

        self.process_button = tk.Button(button_frame, text='BẮT ĐẦU PHÂN LOẠI', command=self.start_processing_thread, font=('Helvetica', 12, 'bold'), bg='#4CAF50', fg='white', padx=20, pady=10, width=25)
        self.process_button.pack(side=tk.LEFT, padx=10)
        
        self.stop_button = tk.Button(button_frame, text='DỪNG LẠI', command=self.stop_processing, font=('Helvetica', 12, 'bold'), bg='#F44336', fg='white', padx=20, pady=10, state=tk.DISABLED, width=25)
        self.stop_button.pack(side=tk.LEFT, padx=10)
        
        log_frame = tk.LabelFrame(self.master, text='Nhật Ký Xử Lý', padx=10, pady=10)
        log_frame.pack(fill='both', expand=True, padx=10, pady=10)
        self.log_area = Text(log_frame, wrap=tk.WORD, state=tk.DISABLED, height=15, bg='#2E2E2E', fg='#E0E0E0')
        self.log_area.pack(side=tk.LEFT, fill='both', expand=True)
        scrollbar = tk.Scrollbar(log_frame, command=self.log_area.yview)
        scrollbar.pack(side=tk.RIGHT, fill='y')
        self.log_area.config(yscrollcommand=scrollbar.set)

    def select_source_folder(self):
        path = filedialog.askdirectory(title='Chọn thư mục chứa file cần xử lý')
        if path:
            self.folder_path.set(os.path.normpath(path))

    def select_destination_folder(self):
        path = filedialog.askdirectory(title='Chọn thư mục để lưu trữ (chứa các folder Toán 6, 7,...)')
        if path:
            self.destination_path.set(os.path.normpath(path))

    def select_xlsx_file(self):
        path = filedialog.asksaveasfilename(
            title='Chọn hoặc tạo file Excel để lưu kết quả',
            defaultextension=".xlsx",
            filetypes=[('Excel Files', '*.xlsx')]
        )
        if path:
            self.xlsx_path.set(os.path.normpath(path))

    def stop_processing(self):
        self.log_message("\n!!! ĐÃ NHẬN LỆNH DỪNG... (Sẽ dừng sau khi hoàn tất file hiện tại) !!!")
        self.stop_event.set()
        self.stop_button.config(state=tk.DISABLED)

    def start_processing_thread(self):
        self.stop_event.clear()
        self.process_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        processing_thread = threading.Thread(target=self.process_files_logic, daemon=True)
        processing_thread.start()

if __name__ == '__main__':
    main_root = tk.Tk()
    app = AutoFileClassifierApp(main_root)
    main_root.mainloop()