import pandas as pd
import os
from tkinter import Tk, Label, Button, filedialog, Listbox, END, MULTIPLE
from tkinter.messagebox import showinfo

pd.set_option('display.max_columns', 100)

def calc_revenue(df, month=None, store_name=None):
    df.columns = [col.replace('\n', '') for col in df.columns]
    condition = df['狀態'].apply(lambda x: x.split('\n')[0] not in ['取消訂單', '逾期未取件']) 
    df = df.loc[condition]
    
    if month:
        condition = df['訂購日期'].apply(lambda x: x.split('/')[1] in month)
        df = df.loc[condition]

    if store_name:
        condition = df['賣場名稱'].apply(lambda x: x in store_name)
        df = df.loc[condition]

    revenue = df['小計(A)'].apply(lambda x: x.replace(',', '')).astype(int).sum()
    return revenue

def read_data(df_paths):
    df_list = [pd.read_excel(df_path, skiprows=2) for df_path in df_paths]
    return pd.concat(df_list, axis=0)

def select_files():
    file_paths = filedialog.askopenfilenames(title="選擇檔案", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_paths:
        file_listbox.delete(0, END)  # 清除舊的檔案列表
        selected_files.clear()       # 清除先前選擇的檔案
        selected_files.extend(file_paths)  # 添加新的檔案路徑

        for file in file_paths:
            file_listbox.insert(END, os.path.basename(file))  # 列出檔案名稱

def calculate_revenue():
    if not selected_files:
        showinfo("提示", "未選擇任何檔案")
        return
    
    # 獲取月份過濾條件
    selected_months = [month_listbox.get(i) for i in month_listbox.curselection()]
    month = [str(month).zfill(2) for month in selected_months]  # 格式化成兩位數月份

    store_name = None
    df = read_data(selected_files)
    revenue = calc_revenue(df=df, month=month, store_name=store_name)
    showinfo("營收結果", f"總營收: {revenue}")

# 建立介面
root = Tk()
root.title("營收計算工具")

Label(root, text="請選擇月份:").pack(pady=10)

# 月份選擇 Listbox，啟用多選模式
month_listbox = Listbox(root, selectmode='extended' , width=20, height=12, exportselection=False)
months = [str(i) for i in range(1, 13)]
for month in months:
    month_listbox.insert(END, month)
month_listbox.pack()

Button(root, text="選擇檔案", command=select_files).pack(pady=10)

Label(root, text="已選擇的檔案:").pack()
file_listbox = Listbox(root, width=50, height=8)
file_listbox.pack()

Button(root, text="開始計算", command=calculate_revenue).pack(pady=20)

selected_files = []

# 自動調整視窗大小
root.update_idletasks()
root.minsize(root.winfo_reqwidth(), root.winfo_reqheight())
root.mainloop()
