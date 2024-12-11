import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import csv
import pandas as pd

# Khởi tạo giao diện
root = tk.Tk()
root.title("Thông tin nhân viên")
root.geometry("700x400")

# Hàm lưu dữ liệu vào file CSV
def save_data():
    data = {
        "Mã": entry_id.get(),
        "Tên": entry_name.get(),
        "Đơn vị": entry_unit.get(),
        "Chức danh": entry_position.get(),
        "Ngày sinh": entry_dob.get(),
        "Giới tính": gender.get(),
        "Số CMND": entry_id_number.get(),
        "Nơi cấp": entry_place.get(),
        "Ngày cấp": entry_issue_date.get(),
        "Là khách hàng": customer_var.get(),
        "Là nhà cung cấp": supplier_var.get()
    }

    if not all(data.values()):
        messagebox.showwarning("Cảnh báo", "Vui lòng điền đầy đủ thông tin.")
        return

    with open("employees.csv", mode="a", newline='', encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=data.keys())
        if file.tell() == 0:
            writer.writeheader()
        writer.writerow(data)

    messagebox.showinfo("Thành công", "Lưu thông tin nhân viên thành công!")

# Hàm hiển thị danh sách nhân viên có sinh nhật hôm nay
def show_today_birthdays():
    today = datetime.now().strftime("%d/%m")
    employees = []

    try:
        with open("employees.csv", mode="r", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            for row in reader:
                dob = row["Ngày sinh"]
                if dob and today in dob:
                    employees.append(row)
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu nhân viên.")
        return

    if employees:
        msg = "\n".join([f"{emp['Tên']} ({emp['Mã']})" for emp in employees])
        messagebox.showinfo("Sinh nhật hôm nay", msg)
    else:
        messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay.")

# Hàm xuất danh sách nhân viên ra file Excel
def export_to_excel():
    try:
        df = pd.read_csv("employees.csv")
        df["Tuổi"] = df["Ngày sinh"].apply(lambda x: (datetime.now() - datetime.strptime(x, "%d/%m/%Y")).days // 365)
        df = df.sort_values(by="Tuổi", ascending=False)
        df.to_excel("employees.xlsx", index=False, encoding="utf-8", engine="openpyxl")
        messagebox.showinfo("Thành công", "Xuất file Excel thành công!")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất file Excel: {e}")

# Tạo các widget giao diện
tk.Label(root, text="Thông tin nhân viên", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=4, pady=10)

entry_id = tk.Entry(root, width=30)
entry_name = tk.Entry(root, width=30)
entry_unit = tk.Entry(root, width=30)
entry_position = tk.Entry(root, width=30)
entry_dob = tk.Entry(root, width=30)
entry_id_number = tk.Entry(root, width=30)
entry_place = tk.Entry(root, width=30)
entry_issue_date = tk.Entry(root, width=30)

gender = tk.StringVar()
gender.set("Nam")

customer_var = tk.StringVar()
customer_var.set("Không")

supplier_var = tk.StringVar()
supplier_var.set("Không")

# Các trường nhập liệu
tk.Label(root, text="Mã *").grid(row=1, column=0, sticky="w", padx=10, pady=5)
entry_id.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Tên *").grid(row=2, column=0, sticky="w", padx=10, pady=5)
entry_name.grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="Đơn vị *").grid(row=3, column=0, sticky="w", padx=10, pady=5)
entry_unit.grid(row=3, column=1, padx=10, pady=5)

tk.Label(root, text="Chức danh").grid(row=4, column=0, sticky="w", padx=10, pady=5)
entry_position.grid(row=4, column=1, padx=10, pady=5)

tk.Label(root, text="Ngày sinh (DD/MM/YYYY) *").grid(row=5, column=0, sticky="w", padx=10, pady=5)
entry_dob.grid(row=5, column=1, padx=10, pady=5)

tk.Label(root, text="Giới tính").grid(row=6, column=0, sticky="w", padx=10, pady=5)
tk.Radiobutton(root, text="Nam", variable=gender, value="Nam").grid(row=6, column=1, sticky="w")
tk.Radiobutton(root, text="Nữ", variable=gender, value="Nữ").grid(row=6, column=2, sticky="w")

tk.Label(root, text="Số CMND").grid(row=7, column=0, sticky="w", padx=10, pady=5)
entry_id_number.grid(row=7, column=1, padx=10, pady=5)

tk.Label(root, text="Nơi cấp").grid(row=8, column=0, sticky="w", padx=10, pady=5)
entry_place.grid(row=8, column=1, padx=10, pady=5)

tk.Label(root, text="Ngày cấp (DD/MM/YYYY)").grid(row=9, column=0, sticky="w", padx=10, pady=5)
entry_issue_date.grid(row=9, column=1, padx=10, pady=5)

tk.Checkbutton(root, text="Là khách hàng", variable=customer_var, onvalue="Có", offvalue="Không").grid(row=1, column=2, sticky="w", padx=10, pady=5)
tk.Checkbutton(root, text="Là nhà cung cấp", variable=supplier_var, onvalue="Có", offvalue="Không").grid(row=2, column=2, sticky="w", padx=10, pady=5)

# Các nút chức năng
tk.Button(root, text="Lưu", command=save_data).grid(row=11, column=0, pady=10)
tk.Button(root, text="Sinh nhật hôm nay", command=show_today_birthdays).grid(row=11, column=1, pady=10)
tk.Button(root, text="Xuất danh sách", command=export_to_excel).grid(row=11, column=2, pady=10)

root.mainloop()
