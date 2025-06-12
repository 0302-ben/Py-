import datetime
import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from tkcalendar import DateEntry
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import google.generativeai as genai


# ===== 基本設定與資料檔 =====
EXCEL_FILENAME = "records.xlsx"
CATEGORY_FILE = "categories.txt"
GEMINI_API_KEY = "AIzaSyD1028NeaHQl_-VaZFjN8LLCo4nnFiPMQk"
records = []

genai.configure(api_key=GEMINI_API_KEY)

# ===== 資料與類別操作 =====
def load_categories():
    if os.path.exists(CATEGORY_FILE):
        with open(CATEGORY_FILE, "r", encoding="utf-8") as f:
            return [line.strip() for line in f if line.strip()]
    else:
        return ["飲食", "交通", "娛樂", "生活用品", "醫療", "房租", "水電", "收入", "投資", "服飾"]

def save_categories(categories):
    with open(CATEGORY_FILE, "w", encoding="utf-8") as f:
        for cat in categories:
            f.write(cat + "\n")

# ===== 從 Excel 讀取記錄，放入 records 清單 =====
def load_records():
    if not os.path.exists(EXCEL_FILENAME):
        return
    records.clear()
    try:
        df = pd.read_excel(EXCEL_FILENAME, engine='openpyxl')
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取 Excel 檔案失敗：{e}")
        return
    for _, row in df.iterrows():
        try:
            amount = float(row["amount"])
        except Exception:
            amount = 0.0
        record = {
            "date": str(row["date"])[:10],
            "category": str(row["category"]),
            "amount": amount,
            "note": str(row.get("note", ""))
        }
        records.append(record)

# ===== 將 records 寫入 Excel 檔案 =====
def save_records():
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

# ===== 繪圖設定 =====
matplotlib.rcParams['font.family'] = 'Microsoft JhengHei'
matplotlib.rcParams['axes.unicode_minus'] = False

def plot_all_charts():
    fig.clear()
    axs = fig.subplots(1, 2)

    summary_pie = {}
    for r in records:
        if r['amount'] < 0:
            summary_pie[r['category']] = summary_pie.get(r['category'], 0) + abs(r['amount'])
    if summary_pie:
        axs[0].pie(summary_pie.values(), labels=summary_pie.keys(), autopct='%1.1f%%', startangle=140)
        axs[0].set_title("支出比例圓餅圖")
    else:
        axs[0].text(0.5, 0.5, "無支出資料", ha='center', va='center')

    df = pd.DataFrame(records)
    if not df.empty:
        df['date'] = pd.to_datetime(df['date'])
        monthly = df.groupby(pd.Grouper(key='date', freq='M'))['amount'].sum()
        axs[1].plot(monthly.index, monthly.values, marker='o')
        axs[1].set_title("每月收支趨勢")
        axs[1].set_xlabel("月份")
        axs[1].set_ylabel("金額 (元)")
    else:
        axs[1].text(0.5, 0.5, "無紀錄資料", ha='center', va='center')

    canvas.draw()

# ===== 預算分析 & 寄送 Email =====
def analyze_and_email_budget_report(budget=20000, send_email=False, recipient="example@gmail.com"):
    today = datetime.date.today()
    current_month = today.strftime("%Y-%m")
    monthly_expenses = [r for r in records if r["date"].startswith(current_month) and r["amount"] < 0]

    total_spent = -sum(r["amount"] for r in monthly_expenses)
    over_budget = total_spent - budget
    status = "✅ 在預算內" if over_budget <= 0 else "⚠️ 超出預算"

    category_summary = {}
    for r in monthly_expenses:
        category_summary[r["category"]] = category_summary.get(r["category"], 0) + abs(r["amount"])

    report = f"📆 {current_month} 開銷預算分析報告\n"
    report += f"📌 預算金額：{budget} 元\n"
    report += f"📉 總支出：{total_spent:.0f} 元\n"
    report += f"{status}（差額：{abs(over_budget):.0f} 元）\n\n"
    report += "📊 類別支出明細：\n"
    for cat, amt in category_summary.items():
        report += f"- {cat}: {amt:.0f} 元\n"

    messagebox.showinfo("本月預算分析報告", report)

    if send_email:
        try:
            msg = MIMEMultipart()
            msg["Subject"] = f"{current_month} 預算分析報告"
            msg["From"] = "your_email@gmail.com"
            msg["To"] = recipient
            msg.attach(MIMEText(report, "plain", "utf-8"))

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login("your_email@gmail.com", "your_app_password")
                server.send_message(msg)

            messagebox.showinfo("Email 寄送成功", f"已寄送報告至 {recipient}")
        except Exception as e:
            messagebox.showerror("Email 錯誤", f"寄送失敗：{e}")

def ask_budget_and_send():
    budget_win = tk.Toplevel(root)
    budget_win.title("本月預算分析")

    tk.Label(budget_win, text="預算金額：").grid(row=0, column=0, padx=5, pady=5)
    budget_entry = tk.Entry(budget_win)
    budget_entry.insert(0, "20000")
    budget_entry.grid(row=0, column=1, padx=5, pady=5)

    send_email_var = tk.BooleanVar()
    send_email_check = tk.Checkbutton(budget_win, text="寄送報告 Email", variable=send_email_var)
    send_email_check.grid(row=1, column=0, columnspan=2, pady=5)

    tk.Label(budget_win, text="收件人 Email：").grid(row=2, column=0, padx=5, pady=5)
    email_entry = tk.Entry(budget_win)
    email_entry.grid(row=2, column=1, padx=5, pady=5)

    def on_confirm():
        try:
            budget = float(budget_entry.get())
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的數字預算")
            return
        recipient = email_entry.get().strip()
        if send_email_var.get() and not recipient:
            messagebox.showerror("錯誤", "請輸入收件人 Email")
            return
        budget_win.destroy()
        analyze_and_email_budget_report(budget=budget, send_email=send_email_var.get(), recipient=recipient)

    btn_confirm = ttk.Button(budget_win, text="開始分析", command=on_confirm)
    btn_confirm.grid(row=3, column=0, columnspan=2, pady=10)

def prompt_email_and_send():
    email = simpledialog.askstring("寄送報表", "請輸入收件人 Email：")
    if email:
        analyze_and_email_budget_report(send_email=True, recipient=email)

# ===== GUI 介面 =====
root = tk.Tk()
root.title("個人理財記帳軟體")

frame = ttk.Frame(root)
frame.pack(padx=10, pady=10)

ttk.Label(frame, text="日期：").grid(row=0, column=0, sticky=tk.W)
entry_date = DateEntry(frame, date_pattern='yyyy-mm-dd')
entry_date.grid(row=0, column=1, sticky=tk.W)

ttk.Label(frame, text="類別：").grid(row=1, column=0, sticky=tk.W)
categories = load_categories()
entry_category = ttk.Combobox(frame, values=categories)
entry_category.grid(row=1, column=1, sticky=tk.W)

ttk.Label(frame, text="金額：").grid(row=2, column=0, sticky=tk.W)
entry_amount = ttk.Entry(frame)
entry_amount.grid(row=2, column=1, sticky=tk.W)

ttk.Label(frame, text="備註：").grid(row=3, column=0, sticky=tk.W)
entry_note = ttk.Entry(frame)
entry_note.grid(row=3, column=1, sticky=tk.W)

btn_add = ttk.Button(frame, text="新增", command=add_record)
btn_add.grid(row=4, column=0, pady=5)
btn_update = ttk.Button(frame, text="修改", command=update_record)
btn_update.grid(row=4, column=1, pady=5)
btn_delete = ttk.Button(frame, text="刪除", command=delete_record)
btn_delete.grid(row=4, column=2, pady=5)

entry_search = ttk.Entry(frame)
entry_search.grid(row=5, column=0, pady=5)
btn_search = ttk.Button(frame, text="搜尋", command=search_records)
btn_search.grid(row=5, column=1, pady=5)
btn_summary = ttk.Button(frame, text="類別統計", command=summary_by_category)
btn_summary.grid(row=5, column=2, pady=5)

btn_advice = ttk.Button(frame, text="理財建議", command=get_financial_advice)
btn_advice.grid(row=6, column=0, pady=5)

btn_budget = ttk.Button(frame, text="本月預算分析", command=ask_budget_and_send)
btn_budget.grid(row=6, column=1, pady=5)

btn_email = ttk.Button(frame, text="寄送本月報表到 Email", command=prompt_email_and_send)
btn_email.grid(row=6, column=2, pady=5)

columns = ("編號", "日期", "類別", "金額", "備註")
tree = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col)
tree.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
tree.bind("<<TreeviewSelect>>", fill_inputs)

# 總金額、收入、支出、淨資產顯示
total_amount_var = tk.StringVar()
income_total_var = tk.StringVar()
expense_total_var = tk.StringVar()
net_asset_var = tk.StringVar()
lbl_total = ttk.Label(root, textvariable=total_amount_var)
lbl_total.pack()
lbl_income = ttk.Label(root, textvariable=income_total_var)
lbl_income.pack()
lbl_expense = ttk.Label(root, textvariable=expense_total_var)
lbl_expense.pack()
lbl_net = ttk.Label(root, textvariable=net_asset_var)
lbl_net.pack()

# 繪圖區
fig = plt.Figure(figsize=(8, 4))
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack()

# 初始化
load_records()
refresh_treeview()
plot_all_charts()

root.mainloop()
