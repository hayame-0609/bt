import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import asksaveasfilename
import csv
from datetime import datetime

# Hàm lưu thông tin vào file CSV
def save_to_csv():
    data = {
        "Mã": entry_id.get(),
        "Tên": entry_name.get(),
        "Đơn vị": entry_department.get(),
        "Chức danh": entry_position.get(),
        "Ngày sinh": entry_birth.get(),
        "Giới tính": gender_var.get(),
        "Số CMND": entry_id_number.get(),
        "Ngày cấp": entry_issue_date.get(),
        "Nơi cấp": entry_issue_place.get()
    }

    if all(data.values()):
        with open("employee_data.csv", mode="a", newline='', encoding="utf-8") as file:
            writer = csv.DictWriter(file, fieldnames=data.keys())
            if file.tell() == 0:
                writer.writeheader()
            writer.writerow(data)
        messagebox.showinfo("Thành công", "Dữ liệu đã được lưu thành công.")
    else:
        messagebox.showerror("Lỗi", "Vui lòng điền đầy đủ thông tin.")

# Hàm hiển thị nhân viên có sinh nhật hôm nay
def show_today_birthdays():
    try:
        today = datetime.today().strftime('%d/%m/%Y')
        with open("employee_data.csv", mode="r", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            birthdays = [row for row in reader if row['Ngày sinh'] == today]

        if birthdays:
            result = "\n".join([f"Mã: {row['Mã']}, Tên: {row['Tên']}, Ngày sinh: {row['Ngày sinh']}" for row in birthdays])
            messagebox.showinfo("Danh sách sinh nhật hôm nay", result)
        else:
            messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay.")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu. Vui lòng nhập thông tin trước.")

# Hàm xuất toàn bộ danh sách ra Excel
def export_to_excel():
    try:
        df = pd.read_csv("employee_data.csv")
        df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'], format='%d/%m/%Y')
        df = df.sort_values(by=['Ngày sinh'], ascending=False)

        filepath = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if filepath:
            df.to_excel(filepath, index=False)
            messagebox.showinfo("Thành công", f"Dữ liệu đã được xuất ra file {filepath}")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu để xuất.")

# Giao diện chính
root = tk.Tk()
root.title("Thông tin nhân viên")
root.geometry('1200x400')

# Các biến lưu trữ thông tin
entry_id = tk.StringVar()
entry_name = tk.StringVar()
entry_department = tk.StringVar()
entry_position = tk.StringVar()
entry_birth = tk.StringVar()
gender_var = tk.StringVar(value="Nam")
entry_id_number = tk.StringVar()
entry_issue_date = tk.StringVar()
entry_issue_place = tk.StringVar()

# Tạo các widget
frame = ttk.Frame(root, padding=10)
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Các trường nhập liệu
labels = ["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Ngày cấp", "Nơi cấp"]
entries = [entry_id, entry_name, entry_department, entry_position, entry_birth, None, entry_id_number, entry_issue_date, entry_issue_place]

gender_frame = ttk.Frame(frame)
for i, label in enumerate(labels):
    ttk.Label(frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=5)
    if i == 5:  # Giới tính
        ttk.Radiobutton(gender_frame, text="Nam", variable=gender_var, value="Nam").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(gender_frame, text="Nữ", variable=gender_var, value="Nữ").pack(side=tk.LEFT, padx=5)
        gender_frame.grid(row=i, column=1, sticky=tk.W)
    else:
        ttk.Entry(frame, textvariable=entries[i]).grid(row=i, column=1, sticky=(tk.W, tk.E), pady=5)

# Nút chức năng
button_frame = ttk.Frame(root, padding=10)
button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))

ttk.Button(button_frame, text="Lưu thông tin", command=save_to_csv).grid(row=0, column=0, padx=5)
ttk.Button(button_frame, text="Sinh nhật hôm nay", command=show_today_birthdays).grid(row=0, column=1, padx=5)
ttk.Button(button_frame, text="Xuất danh sách", command=export_to_excel).grid(row=0, column=2, padx=5)

root.mainloop()
