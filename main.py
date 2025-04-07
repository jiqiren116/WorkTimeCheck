import pandas as pd
import tkinter as tk
import os
from tkinter import filedialog
from tkinter import messagebox

DIFF_UMS_FILE_PATH = ""
DIFF_EW_FILE_PATH = ""

def calculate_hours(file_path, column_index, keyword, pattern):
    try:
        data = pd.read_excel(file_path, header=None, sheet_name=0)  # 读取第一个sheet
        # 获取data中有几行
        num_rows = data.shape[0]
        temp_data = data.loc[4: num_rows - 1, column_index]

        # 筛选包含特定关键词的数据
        filtered_data = temp_data[temp_data.str.contains(keyword, na=False)]

        # 提取小时数并转换为数值类型
        hours_data = filtered_data.str.extract(pattern).astype(float)

        # 保留整数部分
        hours_data = hours_data[0].astype(int)

        # 求和
        result_hours = hours_data.sum()

        return result_hours
    except Exception as e:
        messagebox.showerror("错误", f"处理文件时发生错误: {e}")
        return None

def calculate_total_hours(file_path):
    # 第12列 列名为 实际工作时长
    return calculate_hours(file_path, 12, "小时", r'(\d+)\.?\d*')

def calculate_overtime_hours(file_path):
    # 第13列 列名为 "假勤申请"
    return calculate_hours(file_path, 13, "加班", r'加班(\d+\.\d+)小时')

def select_file(filetypes, startswith, endswith):
    # 打开文件选择对话框
    file_path = filedialog.askopenfilename(filetypes=filetypes)
    
    # 截取文件名
    file_name = os.path.basename(file_path)
    # 判断file_name是否满足条件
    if file_path and file_name.startswith(startswith) and file_name.endswith(endswith):
        return file_path
    else:
        return None

def open_file():
    file_path = select_file([("Excel files", "*.xlsx *.xls")], "上下班打卡", ".xlsx")
    if file_path:
        # 计算加班时长 和 总工时
        overtime_hours = calculate_overtime_hours(file_path)
        total_hours = calculate_total_hours(file_path) # 计算总工时
        if total_hours is not None:
            # 将加班时长显示在文本框中
            overtime_text.delete(1.0, tk.END)
            overtime_text.insert(tk.END, f" {overtime_hours}小时")

            # 将总工时显示在文本框中
            totaltime_text.delete(1.0, tk.END)
            totaltime_text.insert(tk.END, f" {total_hours}小时")
    else:
        messagebox.showerror("错误", "请选择正确的文件")

def open_ums_file():
    global DIFF_UMS_FILE_PATH  # 声明为全局变量

    file_path = select_file([("Excel files", "*.xlsx *.xls")], "export", ".xls")
    if file_path:
        dif_text.delete(1.0, tk.END)
        DIFF_UMS_FILE_PATH = file_path
    else:
        messagebox.showerror("错误", "请选择正确的文件")


def open_ew_file():
    global DIFF_EW_FILE_PATH  # 声明为全局变量
 
    file_path = select_file([("Excel files", "*.xlsx *.xls")], "上下班打卡", ".xlsx")
    if file_path:
        dif_text.delete(1.0, tk.END)
        DIFF_EW_FILE_PATH = file_path
    else:
        messagebox.showerror("错误", "请选择正确的文件")

def process_ums_file():
    # 首先判断DIFF_UMS_FILE_PATH和DIFF_EW_FILE_PATH是否为空
    if DIFF_UMS_FILE_PATH == "" and DIFF_EW_FILE_PATH == "":
        messagebox.showerror("错误", "请先选择ums和企业微信的导出文件")
        return None
    elif DIFF_UMS_FILE_PATH == "":
        messagebox.showerror("错误", "请先选择ums的导出文件")
        return None
    elif DIFF_EW_FILE_PATH == "":
        messagebox.showerror("错误", "请先选择企业微信的导出文件")
        return None
    try:
        ums_data = pd.read_html(DIFF_UMS_FILE_PATH, encoding='utf-8')[0]  # [0] 表示读取第一个表格
        ew_data = pd.read_excel(DIFF_EW_FILE_PATH, header=None, sheet_name=0)  # 读取第一个sheet

        # 获取ums_data中列名为 工时日期 和 工时(h) 的列
        ums_temp_data = ums_data[['工时日期', '工时(h)']]
        # 将 “工时日期” 这一列格式为"2025-03-01"转换为"2025/03/01"
        ums_temp_data['工时日期'] = ums_temp_data['工时日期'].str.replace('-', '/')

        num_rows = ew_data.shape[0]
        ew_temp_data = ew_data.loc[4: num_rows - 1, [0, 12]]

        # 筛选出包含“小时”的行
        ew_filtered_data = ew_temp_data[ew_temp_data[12].astype(str).str.contains("小时")]

        # 便利ums_temp_data中的每一行
        for index, ums_row in ums_temp_data.iterrows(): # 这个index是从0开始的，符合常规逻辑
            ums_date = ums_row['工时日期']
            ums_hours = ums_row['工时(h)']

            ew_count = 0 # 计数器，用于记录遍历了ew_filtered_data中的多少行，因为ew_filtered_data中index是从4开始的，不符合常规逻辑
            # 遍历ew_filtered_data中的每一行
            for index, ew_row in ew_filtered_data.iterrows(): # 这个index是从4开始的，跟ew excel表中的行号一致，不符合常规逻辑
                ew_count += 1
                ew_date = ew_row[0]

                if "." in ew_row[12]:
                    ew_hours = int(ew_row[12].split(".")[0])
                else: # 如果没有小数点，那么格式就为"7小时"，需要截取"小时"前面的数字
                    ew_hours = int(ew_row[12].split("小时")[0])

                # 如果ew_date中包含date，并且ew_hours不等于hours，就打印出来
                if ums_date in ew_date and ums_hours != ew_hours:
                    # 将结果写入dif_text中
                    dif_text.insert(tk.END, f"工时不对!!! ums日期: {ums_date}, ums工时(h): {ums_hours}, 企业微信日期: {ew_date}, 企业微信工时: {ew_hours}\n")
                    break
                if ums_date in ew_date and ums_hours== ew_hours:
                    break
    
                # 如果遍历完，date一直没有在ew_date中出现，将结果写入dif_text中
                if ew_count == ew_filtered_data.shape[0]:
                    dif_text.insert(tk.END, f"企业微信 缺少 {ums_date} 的记录!!!\n")

        # 遍历ew_filtered_data中的每一行
        for index, ew_row in ew_filtered_data.iterrows(): # 这个index是从4开始的，跟ew excel表中的行号一致，不符合常规逻辑
            ew_date = ew_row[0]
            if "." in ew_row[12]:
                ew_hours = int(ew_row[12].split(".")[0])
            else: # 如果没有小数点，那么格式就为"7小时"，需要截取"小时"前面的数字
                ew_hours = int(ew_row[12].split("小时")[0])
            
            # 遍历ums_temp_data中的每一行
            for ums_index, ums_row in ums_temp_data.iterrows(): # 这个ums_index是从0开始的，符合常规逻辑
                ums_date = ums_row['工时日期']
                ums_hours = ums_row['工时(h)']

                # 如果ew_date中包含date，并且ew_hours不等于hours，就打印出来
                if ums_date in ew_date and ums_hours == ew_hours:
                    break
                if ums_date in ew_date and ums_hours != ew_hours:
                    # 将结果写入dif_text中
                    dif_text.insert(tk.END, f"工时不对!!! ums日期: {ums_date}, ums工时(h): {ums_hours}, 企业微信日期: {ew_date}, 企业微信工时: {ew_hours}\n")
                    break
                # 如果遍历完，date一直没有在ew_date中出现，将结果写入dif_text中
                if ums_index == ums_temp_data.shape[0] - 1:
                    dif_text.insert(tk.END, f"ums 缺少 {ew_date} 的记录!!!\n")
                    break
    except Exception as e:
        messagebox.showerror("错误", f"处理文件时发生错误: {e}")
        return None
    

# 创建主窗口
root = tk.Tk()
root.title("工时统计工具")
root.geometry("700x400")  # 设置窗口大小

open_ew_file_frame = tk.Frame(root)
open_ew_file_frame.pack(pady=10)

# 创建label并添加到Frame中
open_ew_file_label = tk.Label(open_ew_file_frame, text="选择企业微信导出的Excel文件\n 来查看 加班时间 和 总工时:")
open_ew_file_label.pack(side=tk.LEFT)

# 创建文件选择按钮
open_ew_file_button = tk.Button(open_ew_file_frame, text="选择文件", command=open_file)
open_ew_file_button.pack(side=tk.LEFT, padx=5)

### 加班时长
overtime_frame = tk.Frame(root)
overtime_frame.pack(pady=10)

# 创建label并添加到Frame中
overtime_label = tk.Label(overtime_frame, text="加班时长:")
overtime_label.pack(side=tk.LEFT)

overtime_text = tk.Text(overtime_frame, height=1, width=10)
overtime_text.pack(side=tk.LEFT, padx=5)

### 总工时
totaltime_frame = tk.Frame(root)
totaltime_frame.pack(pady=10)
# 创建label并添加到Frame中
totaltime_label = tk.Label(totaltime_frame, text="总工时:")
totaltime_label.pack(side=tk.LEFT)

# 创建文本框用于显示结果
totaltime_text = tk.Text(totaltime_frame, height=1, width=10)
totaltime_text.pack(pady=10)

# 画一条横线来分隔
separator = tk.Frame(root, height=2, bd=1, relief=tk.SUNKEN)
separator.pack(fill=tk.X, padx=5, pady=5)

diff_frame = tk.Frame(root)
diff_frame.pack(pady=10)
open_ums_file_button = tk.Button(diff_frame, text="选择UMS导出的Excel文件", command=open_ums_file)
open_ums_file_button.pack(side=tk.LEFT, padx=5)

open_ew_file_button2 = tk.Button(diff_frame, text="选择企业微信导出的Excel文件", command=open_ew_file)
open_ew_file_button2.pack(side=tk.LEFT, padx=5)

diff_frame2 = tk.Frame(root)
diff_frame2.pack(pady=10)
# 创建label并添加到Frame中
diff_label = tk.Label(diff_frame2, text="选中两个文件后，点击执行对比，查看差异:")
diff_label.pack(side=tk.LEFT)
execute_diff_button = tk.Button(diff_frame2, text="执行对比", command=process_ums_file)
execute_diff_button.pack(side=tk.LEFT, padx=5)


# 绘制一个文本框，带滚动条，用于显示结果
dif_text = tk.Text(root, height=10, width=320)
dif_text.pack(pady=10)

# 创建滚动条并与文本框关联
scrollbar = tk.Scrollbar(root, orient='vertical',command=dif_text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
dif_text.config(yscrollcommand=scrollbar.set)

# 运行主循环
root.mainloop()
