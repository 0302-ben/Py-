import google.generativeai as genai
import datetime
import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
from tkcalendar import DateEntry

# ===== 基本設定與資料檔 =====
EXCEL_FILENAME = "records.xlsx"          # Excel 檔案名稱
CATEGORY_FILE = "categories.txt"
GEMINI_API_KEY = "AIzaSyD1028NeaHQl_-VaZFjN8LLCo4nnFiPMQk"  # 請確認此為你的 API 金鑰
records = []  # 全域清單，存放所有收支記錄字典

genai.configure(api_key=GEMINI_API_KEY)

# ===== 資料與類別操作 =====
def load_categories():
    if os.path.exists(CATEGORY_FILE):
        with open(CATEGORY_FILE, "r", encoding="utf-8") as f:
            return [line.strip() for line in f if line.strip()]
    else:
        # 預設類別清單
        return ["飲食", "交通", "娛樂", "生活用品", "醫療", "房租", "水電", "收入", "投資", "服飾"]

def save_categories(categories):
    with open(CATEGORY_FILE, "w", encoding="utf-8") as f:
        for cat in categories:
            f.write(cat + "\n")

# ===== 從 Excel 讀取記錄，放入 records 清單 =====
def load_records():
    """
    從 Excel 檔讀取記錄，並存在 records 清單中
    """
    if not os.path.exists(EXCEL_FILENAME):
        return

    records.clear()  # 清空舊資料，避免重複

    try:
        # 用 pandas 讀 Excel
        df = pd.read_excel(EXCEL_FILENAME, engine='openpyxl')
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取 Excel 檔案失敗：{e}")
        return

    # 逐行轉成字典放入 records
    for _, row in df.iterrows():
        try:
            amount = float(row["amount"])
        except Exception:
            amount = 0.0
        record = {
            "date": str(row["date"])[:10],  # 取前10字元 YYYY-MM-DD
            "category": str(row["category"]),
            "amount": amount,
            "note": str(row.get("note", ""))
        }
        records.append(record)

# ===== 將 records 寫入 Excel 檔案 =====
def save_records():
    """
    將 records 清單寫入 Excel 檔案，確保資料儲存
    """
    try:
        df = pd.DataFrame(records)
        df.to_excel(EXCEL_FILENAME, index=False, engine='openpyxl')
    except Exception as e:
        messagebox.showerror("錯誤", f"儲存 Excel 檔案失敗：{e}")

# ===== 記錄操作功能 =====
def add_record():
    date = entry_date.get()
    category = entry_category.get()
    amount_str = entry_amount.get()
    note = entry_note.get()

    if not date or not category or not amount_str:
        messagebox.showwarning("提示", "請輸入所有必要欄位")
        return
    try:
        amount = float(amount_str)
    except ValueError:
        messagebox.showerror("錯誤", "請輸入有效的數字金額")
        return

    record = {"date": date, "category": category, "amount": amount, "note": note}
    records.append(record)
    save_records()
    refresh_treeview()
    clear_inputs()
    messagebox.showinfo("成功", "紀錄已新增！")

def update_record():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("提醒", "請先選擇要修改的記錄")
        return
    idx = int(tree.item(selected[0])['values'][0]) - 1
    date = entry_date.get()
    category = entry_category.get()
    amount_str = entry_amount.get()
    note = entry_note.get()

    try:
        amount = float(amount_str)
    except ValueError:
        messagebox.showerror("錯誤", "請輸入有效的數字金額")
        return

    records[idx] = {"date": date, "category": category, "amount": amount, "note": note}
    save_records()
    refresh_treeview()
    clear_inputs()
    messagebox.showinfo("成功", "紀錄已更新！")

def delete_record():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("提醒", "請先選取要刪除的記錄")
        return
    confirm = messagebox.askyesno("確認", "確定要刪除選取的記錄嗎？")
    if not confirm:
        return
    indices = [int(tree.item(sel)['values'][0]) - 1 for sel in selected]
    for idx in sorted(indices, reverse=True):
        records.pop(idx)
    save_records()
    refresh_treeview()
    messagebox.showinfo("完成", "選取記錄已刪除")

def clear_inputs():
    entry_date.delete(0, tk.END)
    entry_category.set("")
    entry_amount.delete(0, tk.END)
    entry_note.delete(0, tk.END)

def fill_inputs(event):
    selected = tree.selection()
    if not selected:
        return
    idx = int(tree.item(selected[0])['values'][0]) - 1
    r = records[idx]
    entry_date.delete(0, tk.END)
    entry_date.insert(0, r["date"])
    entry_category.set(r["category"])
    entry_amount.delete(0, tk.END)
    entry_amount.insert(0, str(r["amount"]))
    entry_note.delete(0, tk.END)
    entry_note.insert(0, r["note"])

def refresh_treeview():
    for item in tree.get_children():
        tree.delete(item)
    for i, r in enumerate(records, start=1):
        tree.insert("", "end", values=(i, r["date"], r["category"], r["amount"], r["note"]))
    update_total_amount()

def update_total_amount():
    total = sum(r['amount'] for r in records)
    income = sum(r['amount'] for r in records if r['amount'] > 0)
    expense = sum(-r['amount'] for r in records if r['amount'] < 0)
    net_asset = income - expense
    total_amount_var.set(f"總金額：{total:.0f} 元")
    income_total_var.set(f"收入總額：{income:.0f} 元")
    expense_total_var.set(f"支出總額：{expense:.0f} 元")
    net_asset_var.set(f"💰 淨資產（收入-支出）：{net_asset:.0f} 元")

def search_records():
    keyword = entry_search.get().strip()
    if not keyword:
        refresh_treeview()
        return
    filtered = [r for r in records if keyword in r['category'] or keyword in r['note']]
    tree.delete(*tree.get_children())
    for i, r in enumerate(filtered, start=1):
        tree.insert("", "end", values=(i, r["date"], r["category"], r["amount"], r["note"]))

def summary_by_category():
    summary = {}
    for r in records:
        summary[r['category']] = summary.get(r['category'], 0) + r['amount']
    msg = "\n".join([f"{cat}: {amt:.0f} 元" for cat, amt in summary.items()])
    messagebox.showinfo("類別統計", msg)

# ===== Gemini 理財建議 =====
def get_financial_advice():
    if not records:
        messagebox.showinfo("提示", "目前沒有任何記錄，無法產生建議")
        return

    model = genai.GenerativeModel("models/gemini-1.5-pro-latest")

    summary_lines = [f"{r['date']} 類別: {r['category']} 金額: {r['amount']} 備註: {r['note']}" for r in records]
    summary_text = "\n".join(summary_lines)

    prompt = (
        "你是一位專業理財顧問，以下是使用者的收支紀錄：\n"
        f"{summary_text}\n"
        "請提供三點具體的理財建議，使用中文，簡明扼要，麻煩用簡單點的語氣。"
    )

    try:
        response = model.generate_content(prompt)
        messagebox.showinfo("理財建議", response.text.strip())
    except Exception as e:
        messagebox.showerror("錯誤", f"取得理財建議失敗：{e}")

# ===== 繪圖 =====
matplotlib.rcParams['font.family'] = 'Microsoft JhengHei'  # 中文字體設定
matplotlib.rcParams['axes.unicode_minus'] = False          # 負號正常顯示

def plot_all_charts():
    fig.clear()
    axs = fig.subplots(1, 2)

    # 支出類別比例圓餅圖 (負數金額視為支出)
    summary_pie = {}
    for r in records:
        if r['amount'] < 0:
            summary_pie[r['category']] = summary_pie.get(r['category'], 0) + abs(r['amount'])
    if summary_pie:
        axs[0].pie(summary_pie.values(), labels=summary_pie.keys(), autopct='%1.1f%%', startangle=140)
        axs[0].set_title("支出比例圓餅圖")
    else:
        axs[0].text(0.5, 0.5, "無支出資料", ha='center', va='center')

    # 每月收支折線圖
    monthly_summary = {}
    for r in records:
        month = r['date'][:7]  # 取 YYYY-MM
        monthly_summary[month] = monthly_summary.get(month, 0) + r['amount']
    if monthly_summary:
        months = sorted(monthly_summary.keys())
        amounts = [monthly_summary[m] for m in months]
        axs[1].plot(months, amounts, marker='o')
        axs[1].set_title("每月收支折線圖")
        axs[1].set_xlabel("月份")
        axs[1].set_ylabel("淨收支 (元)")
        axs[1].tick_params(axis='x', rotation=45)
    else:
        axs[1].text(0.5, 0.5, "無資料", ha='center', va='center')

    fig.tight_layout()
    canvas.draw()

# ===== GUI =====
root = tk.Tk()
root.title("Gemini 理財管理系統")
# ===== 新增聊天視窗功能 =====
def open_chat_window():
    chat_win = tk.Toplevel(root)
    chat_win.title("Gemini 理財聊天")
    chat_win.geometry("600x400")

    chat_text = tk.Text(chat_win, width=70, height=20, state=tk.DISABLED)
    chat_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    entry_chat = ttk.Entry(chat_win, width=50)
    entry_chat.pack(side=tk.LEFT, padx=10, pady=5, fill=tk.X, expand=True)

    def send_message():
        user_msg = entry_chat.get().strip()
        if not user_msg:
            return
        entry_chat.delete(0, tk.END)
        chat_text.config(state=tk.NORMAL)
        chat_text.insert(tk.END, f"你: {user_msg}\n")
        chat_text.config(state=tk.DISABLED)
        chat_text.see(tk.END)

        summary_lines = [f"{r['date']} {r['category']} {r['amount']} 元 - {r['note']}" for r in records]
        summary_text = "\n".join(summary_lines) if summary_lines else "目前無記錄"

        prompt = (
            "你是一位友善的中文理財顧問。以下是使用者的歷史收支紀錄：\n"
            f"{summary_text}\n\n"
            f"使用者提出的問題是：{user_msg}\n"
            "請根據收支情況給出回應，若資料不足，也請誠實說明。"
        )

        try:
            model = genai.GenerativeModel("models/gemini-1.5-pro-latest")
            response = model.generate_content(prompt)
            answer = response.text.strip()
        except Exception as e:
            answer = f"錯誤：無法取得回覆 ({e})"

        chat_text.config(state=tk.NORMAL)
        chat_text.insert(tk.END, f"Gemini: {answer}\n\n")
        chat_text.config(state=tk.DISABLED)
        chat_text.see(tk.END)

    btn_send = ttk.Button(chat_win, text="送出", command=send_message)
    btn_send.pack(side=tk.RIGHT, padx=10, pady=5)

# ===== 在主視窗加入開啟聊天的按鈕 =====
btn_chat = ttk.Button(root, text="開啟 Gemini 聊天", command=open_chat_window)
btn_chat.pack(pady=5)


# Frame: 輸入區
frame_input = ttk.Frame(root)
frame_input.pack(pady=10, padx=10, fill=tk.X)

ttk.Label(frame_input, text="日期 (YYYY-MM-DD):").grid(row=0, column=0)
entry_date = DateEntry(frame_input, date_pattern='yyyy-MM-dd', locale='zh_TW')
entry_date.grid(row=0, column=1)

ttk.Label(frame_input, text="類別:").grid(row=0, column=2)
category_list = load_categories()
entry_category = ttk.Combobox(frame_input, values=category_list)
entry_category.grid(row=0, column=3)

ttk.Label(frame_input, text="金額 (正收入/負支出):").grid(row=1, column=0)
entry_amount = ttk.Entry(frame_input)
entry_amount.grid(row=1, column=1)

ttk.Label(frame_input, text="備註:").grid(row=1, column=2)
entry_note = ttk.Entry(frame_input)
entry_note.grid(row=1, column=3)

# 按鈕
btn_add = ttk.Button(frame_input, text="新增", command=add_record)
btn_add.grid(row=2, column=0, pady=5)
btn_update = ttk.Button(frame_input, text="修改", command=update_record)
btn_update.grid(row=2, column=1)
btn_delete = ttk.Button(frame_input, text="刪除", command=delete_record)
btn_delete.grid(row=2, column=2)
btn_clear = ttk.Button(frame_input, text="清空欄位", command=clear_inputs)
btn_clear.grid(row=2, column=3)

# Frame: 搜尋區
frame_search = ttk.Frame(root)
frame_search.pack(pady=5, padx=10, fill=tk.X)

ttk.Label(frame_search, text="搜尋（類別或備註）:").pack(side=tk.LEFT)
entry_search = ttk.Entry(frame_search)
entry_search.pack(side=tk.LEFT, fill=tk.X, expand=True)
btn_search = ttk.Button(frame_search, text="搜尋", command=search_records)
btn_search.pack(side=tk.LEFT, padx=5)
btn_summary = ttk.Button(frame_search, text="類別統計", command=summary_by_category)
btn_summary.pack(side=tk.LEFT)

# Treeview 顯示清單
columns = ("序號", "日期", "類別", "金額", "備註")
# 新增類別輸入框和按鈕
ttk.Label(frame_input, text="新增類別:").grid(row=3, column=0, sticky="w", pady=5)
entry_new_category = ttk.Entry(frame_input)
entry_new_category.grid(row=3, column=1, pady=5)

def add_category():
    new_cat = entry_new_category.get().strip()
    if not new_cat:
        messagebox.showwarning("提醒", "請輸入類別名稱")
        return
    categories = list(entry_category['values'])
    if new_cat in categories:
        messagebox.showwarning("提醒", "類別已存在")
        return
    categories.append(new_cat)
    save_categories(categories)
    entry_category['values'] = categories
    entry_new_category.delete(0, tk.END)
    messagebox.showinfo("成功", f"類別「{new_cat}」已新增")

def delete_category():
    cat_to_delete = entry_category.get()
    if not cat_to_delete:
        messagebox.showwarning("提醒", "請先選擇要刪除的類別")
        return
    confirm = messagebox.askyesno("確認", f"確定要刪除類別「{cat_to_delete}」嗎？")
    if not confirm:
        return
    categories = list(entry_category['values'])
    if cat_to_delete in categories:
        categories.remove(cat_to_delete)
        save_categories(categories)
        entry_category['values'] = categories
        entry_category.set('')
        messagebox.showinfo("成功", f"類別「{cat_to_delete}」已刪除")
    else:
        messagebox.showwarning("提醒", "該類別不存在")

btn_add_category = ttk.Button(frame_input, text="新增類別", command=add_category)
btn_add_category.grid(row=3, column=2, padx=5)

btn_delete_category = ttk.Button(frame_input, text="刪除類別", command=delete_category)
btn_delete_category.grid(row=3, column=3, padx=5)

tree = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=100)
tree.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

tree.bind("<<TreeviewSelect>>", fill_inputs)

# 顯示總計
frame_total = ttk.Frame(root)
frame_total.pack(padx=10, pady=5, fill=tk.X)
total_amount_var = tk.StringVar(value="總金額：0 元")
income_total_var = tk.StringVar(value="收入總額：0 元")
expense_total_var = tk.StringVar(value="支出總額：0 元")
net_asset_var = tk.StringVar(value="💰 淨資產（收入-支出）：0 元")

ttk.Label(frame_total, textvariable=total_amount_var).pack(side=tk.LEFT, padx=5)
ttk.Label(frame_total, textvariable=income_total_var).pack(side=tk.LEFT, padx=5)
ttk.Label(frame_total, textvariable=expense_total_var).pack(side=tk.LEFT, padx=5)
ttk.Label(frame_total, textvariable=net_asset_var).pack(side=tk.LEFT, padx=5)

# Gemini 建議按鈕
btn_gemini = ttk.Button(root, text="取得 Gemini 理財建議", command=get_financial_advice)
btn_gemini.pack(pady=5)

# 繪圖區
fig = plt.Figure(figsize=(8, 3))
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack()

btn_plot = ttk.Button(root, text="繪製圖表", command=plot_all_charts)
btn_plot.pack(pady=5)

# 啟動時讀取資料與顯示
load_records()
refresh_treeview()

root.mainloop()