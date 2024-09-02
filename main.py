import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter import font as tkfont
from tkinter import messagebox
import datetime
from tkinter import filedialog


def open_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls;*.xlsm')])  # 打开文件对话框
    if file_path:
        entry1.delete(0, tk.END)  # 清空输入框
        entry1.insert(0, file_path)  # 显示文件路径



def show_info():
    try:

        file_path = entry1.get()  # 获取文件路径
        if not file_path:
            messagebox.showinfo("提示", "请先选择要导入的Excel文件")
            return

        # 读取 Excel 文件
        df = pd.read_excel(file_path)

        # 需要保留的列
        keep_columns = ['日期', '序号', '老板名称', '员工名称', '单子单情', '单类价格', '消存单/现消', '结单方式',
                        '结单接待']

        # 删除除需保留的列之外的所有其他列
        df = df[keep_columns]
        # 获取选择的条件
        selected_option = var.get()

        if selected_option == 1:  # 老板名称
            name = entry.get()
            result = df[df['老板名称'] == name]
        elif selected_option == 2:  # 员工名称
            name = entry.get()
            result = df[df['员工名称'] == name]
        elif selected_option == 3:  # 结单接待名称
            name = entry.get()
            result = df[df['结单接待'] == name]
        else:
            result = None

        if result is not None:
            # 进行日期过滤
            start_date = cal_start.get_date()
            end_date = cal_end.get_date()
            if (start_date is not None) and (end_date is not None):
                start_date = datetime.datetime.combine(start_date, datetime.time.min)
                end_date = datetime.datetime.combine(end_date, datetime.time.max)
                result = result[(result['日期'] >= start_date) & (result['日期'] <= end_date)]

            if len(result) > 0:
                # 清空表格数据
                table_data.delete(*table_data.get_children())

                # 将查询结果转换为列表
                table_data_list = result.values.tolist()

                # 填充表格数据
                for row in table_data_list:
                    table_data.insert('', tk.END, values=row)

                # 计算单类价格总和
                total_price = result['单类价格'].sum()

                # 显示单类价格总和
                total_label.config(text=f"单类价格总和：{total_price}")
            else:
                messagebox.showinfo("提示", "未找到对应信息")
        else:
            messagebox.showinfo("提示", "未找到对应信息")
    except pd.errors.EmptyDataError:
        messagebox.showerror("错误", "Excel文件为空或读取失败")
    except FileNotFoundError:
        messagebox.showerror("错误", "找不到指定的Excel文件")



# 创建主窗口
window = tk.Tk()
window.title("信息查询")
window.geometry("800x500")




# 创建导入文件按钮
import_button = ttk.Button(window, text="导入文件", command=open_file)
import_button.place(x=650, y=30)



# 设置字体样式
font_style = tkfont.Font(family="Arial", size=12)

# 创建单选框
var = tk.IntVar()
radio1 = ttk.Radiobutton(window, text="老板名称", variable=var, value=1)
radio1.place(x=50, y=30)
radio1['style'] = 'TRadiobutton'
radio2 = ttk.Radiobutton(window, text="员工名称", variable=var, value=2)
radio2.place(x=150, y=30)
radio2['style'] = 'TRadiobutton'
radio3 = ttk.Radiobutton(window, text="结单接待名称", variable=var, value=3)
radio3.place(x=230, y=30)
radio3['style'] = 'TRadiobutton'

# 创建输入框
entry = ttk.Entry(window)
entry.place(x=350, y=30)
entry['font'] = font_style

entry1 = ttk.Entry(window)
entry1.place(x=550, y=70)
entry1['font'] = font_style


# 创建起始日期选择框
cal_start = DateEntry(window, width=12, background='darkblue', foreground='white',
                      borderwidth=2, date_pattern='yyyy-mm-dd')
cal_start.place(x=50, y=70)

# 创建结束日期选择框
cal_end = DateEntry(window, width=12, background='darkblue', foreground='white',
                    borderwidth=2, date_pattern='yyyy-mm-dd')
cal_end.place(x=200, y=70)

# 创建按钮
button = ttk.Button(window, text="查询", command=show_info)
button.place(x=350, y=70)

# 创建一个框架，以便在窗口中放置表格
frame = tk.Frame(window)
frame.place(x=50, y=120)

# 创建一个画布并将其放置在框架左侧，以便我们可以滚动表格
canvas = tk.Canvas(frame)
canvas.pack(side=tk.LEFT, fill=tk.BOTH)

# 创建一个垂直滚动条并将其放置在框架右侧，以便我们可以滚动表格
vsb = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
vsb.pack(side=tk.RIGHT, fill=tk.Y)

# 将画布的yview设置为Scrollbar的set方法
canvas.configure(yscrollcommand=vsb.set)

# 在画布上创建一个框架，以便我们可以将表格放置在其中
table_frame = tk.Frame(canvas)
table_frame.pack()



# 创建TreeView表格，并包括其他列
table_data = ttk.Treeview(table_frame, columns=( '日期','序号', '老板名称', '员工名称', '单子单情', '单类价格', '消存单/现消', '结单方式' , '结单接待'), show='headings')


# 设置原有列的列宽
table_data.column('日期', width=75)
table_data.column('序号', width=40)
table_data.column('老板名称', width=80)
table_data.column('员工名称', width=80)
table_data.column('单子单情', width=80)
table_data.column('单类价格', width=80)
table_data.column('消存单/现消', width=100)
table_data.column('结单方式', width=80)
table_data.column('结单接待', width=80)

# 设置原有列的列标题
table_data.heading('日期', text='日期')
table_data.heading('序号', text='序号')
table_data.heading('老板名称', text='老板名称')
table_data.heading('员工名称', text='员工名称')
table_data.heading('单子单情', text='单子单情')
table_data.heading('单类价格', text='单类价格')
table_data.heading('消存单/现消', text='消存单/现消')
table_data.heading('结单方式', text='结单方式')
table_data.heading('结单接待', text='结单接待')

# 将TreeView表格放置在左侧并填充整个框架
table_data.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

# 将Scrollbar的command设置为TreeView表格的yview方法
vsb.configure(command=table_data.yview)

# 将TreeView表格放置在左侧并填充整个框架
table_data.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

# 将Scrollbar的command设置为TreeView表格的yview方法
vsb.configure(command=table_data.yview)

# 配置画布的滚动区域
canvas.configure(scrollregion=canvas.bbox(tk.ALL))
# 设置表格样式
style = ttk.Style(window)
style.configure("Treeview", font=("Arial", 10), rowheight=25)
style.configure("Treeview.Heading", font=("Arial", 10, "bold"))

# 创建标签用于显示单类价格总和
total_label = ttk.Label(window, text="单类价格总和：")
total_label.place(x=50, y=400)

# 开启主循环
window.mainloop()
