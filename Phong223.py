import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime
import csv
import pandas as pd
from pathlib import Path

def save_to_csv():
    data = {
        "Mã": entry_ma.get(),
        "Tên": entry_ten.get(),
        "Ngày sinh": entry_ngaysinh.get(),
        "Giới tính": "Nam" if gender_var.get() == 1 else "Nữ",
        "Số CMND": entry_cmnd.get(),
        "Ngày cấp": entry_ngaycap.get(),
        "Nơi cấp": entry_noicap.get(),
        "Đơn vị": entry_donvi.get(),
        "Chức danh": entry_chucdanh.get()
    }
    file_path = Path("data.csv")
    with file_path.open("a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=data.keys())
        if file_path.stat().st_size == 0:
            writer.writeheader()
        writer.writerow(data)
    messagebox.showinfo("Thông báo", "Dữ liệu đã được lưu")
    clear_form()
    load_data()

def clear_form():
    entry_ma.delete(0, tk.END)
    entry_ten.delete(0, tk.END)
    entry_ngaysinh.set_date(datetime.now())
    gender_var.set(1)
    entry_cmnd.delete(0, tk.END)
    entry_ngaycap.set_date(datetime.now())
    entry_noicap.delete(0, tk.END)
    entry_donvi.delete(0, tk.END)
    entry_chucdanh.delete(0, tk.END)

def load_data():
    for row in tree.get_children():
        tree.delete(row)
    file_path = Path("data.csv")
    if file_path.exists():
        with file_path.open("r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                tree.insert("", "end", values=tuple(row.values()))

def check_birthdays():
    today = datetime.now().strftime("%d/%m/%Y")
    file_path = Path("data.csv")
    if not file_path.exists():
        messagebox.showinfo("Sinh nhật hôm nay", "Không có dữ liệu để kiểm tra.")
        return
    with file_path.open("r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        birthdays = [row for row in reader if row["Ngày sinh"] == today]
    if birthdays:
        result = "\n".join([f"{row['Tên']} ({row['Mã']})" for row in birthdays])
    else:
        result = "Không có nhân viên nào sinh nhật hôm nay."
    messagebox.showinfo("Sinh nhật hôm nay", result)

def export_sorted_list():
    file_path = Path("data.csv")
    if not file_path.exists():
        messagebox.showinfo("Thông báo", "Không có dữ liệu để xuất.")
        return
    df = pd.read_csv(file_path)
    df["Ngày sinh"] = pd.to_datetime(df["Ngày sinh"], format="%d/%m/%Y")
    sorted_df = df.sort_values(by="Ngày sinh", ascending=False)
    sorted_df.to_excel("sorted_list.xlsx", index=False)
    messagebox.showinfo("Thông báo", "Danh sách đã được xuất")

root = tk.Tk()
root.title("Thông tin nhân viên")

gender_var = tk.IntVar(value=1)

# Form input
form_frame = tk.LabelFrame(root, text="Thông tin nhân viên", padx=10, pady=10)
form_frame.pack(pady=10, fill="x")

tk.Label(form_frame, text="Mã").grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_ma = tk.Entry(form_frame)
entry_ma.grid(row=0, column=1, padx=5, pady=5)

tk.Label(form_frame, text="Tên").grid(row=0, column=2, padx=5, pady=5, sticky="w")
entry_ten = tk.Entry(form_frame)
entry_ten.grid(row=0, column=3, padx=5, pady=5)

tk.Label(form_frame, text="Ngày sinh").grid(row=1, column=0, padx=5, pady=5, sticky="w")
entry_ngaysinh = DateEntry(form_frame, date_pattern="dd/mm/yyyy")
entry_ngaysinh.grid(row=1, column=1, padx=5, pady=5)

tk.Label(form_frame, text="Giới tính").grid(row=1, column=2, padx=5, pady=5, sticky="w")
tk.Radiobutton(form_frame, text="Nam", variable=gender_var, value=1).grid(row=1, column=3, sticky="w")
tk.Radiobutton(form_frame, text="Nữ", variable=gender_var, value=2).grid(row=1, column=3, sticky="e")

tk.Label(form_frame, text="Số CMND").grid(row=2, column=0, padx=5, pady=5, sticky="w")
entry_cmnd = tk.Entry(form_frame)
entry_cmnd.grid(row=2, column=1, padx=5, pady=5)

tk.Label(form_frame, text="Ngày cấp").grid(row=2, column=2, padx=5, pady=5, sticky="w")
entry_ngaycap = DateEntry(form_frame, date_pattern="dd/mm/yyyy")
entry_ngaycap.grid(row=2, column=3, padx=5, pady=5)

tk.Label(form_frame, text="Nơi cấp").grid(row=3, column=0, padx=5, pady=5, sticky="w")
entry_noicap = tk.Entry(form_frame)
entry_noicap.grid(row=3, column=1, padx=5, pady=5)

tk.Label(form_frame, text="Đơn vị").grid(row=3, column=2, padx=5, pady=5, sticky="w")
entry_donvi = tk.Entry(form_frame)
entry_donvi.grid(row=3, column=3, padx=5, pady=5)

tk.Label(form_frame, text="Chức danh").grid(row=4, column=0, padx=5, pady=5, sticky="w")
entry_chucdanh = tk.Entry(form_frame)
entry_chucdanh.grid(row=4, column=1, padx=5, pady=5)

tk.Button(form_frame, text="Lưu thông tin", command=save_to_csv).grid(row=5, column=0, pady=10)
tk.Button(form_frame, text="Sinh nhật hôm nay", command=check_birthdays).grid(row=5, column=1, pady=10)
tk.Button(form_frame, text="Xuất danh sách", command=export_sorted_list).grid(row=5, column=2, pady=10)

# Table
table_frame = tk.Frame(root)
table_frame.pack(pady=10, fill="both", expand=True)

columns = ("Mã", "Tên", "Ngày sinh", "Giới tính", "Số CMND", "Ngày cấp", "Nơi cấp", "Đơn vị", "Chức danh")
tree = ttk.Treeview(table_frame, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=120)
tree.pack(fill="both", expand=True)

load_data()
root.mainloop()
