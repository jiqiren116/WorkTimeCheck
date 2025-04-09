import pandas as pd
import tkinter as tk
import os
from tkinter import filedialog
from tkinter import messagebox

DIFF_UMS_FILE_PATH = ""
DIFF_EW_FILE_PATH = ""
key_sequence = []  # 用于存储按键序列，彩蛋

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
    elif file_path == "": # 处理选择文件对话框中点击取消的情况
        return "file_path is empty" 
    else:
        return None

def open_file():
    file_path = select_file([("Excel files", "*.xlsx *.xls")], "上下班打卡", ".xlsx")
    if file_path:
        if file_path == "file_path is empty":
            return None
        
        # 计算加班时长 和 总工时
        overtime_hours = calculate_overtime_hours(file_path)
        total_hours = calculate_total_hours(file_path) # 计算总工时
        if total_hours is not None:
            # 将加班时长显示在文本框中
            overtime_text.config(state=tk.NORMAL)
            overtime_text.delete(1.0, tk.END)
            overtime_text.insert(tk.END, f" {overtime_hours}小时")
            overtime_text.config(state=tk.DISABLED)

            # 将总工时显示在文本框中
            totaltime_text.config(state=tk.NORMAL)
            totaltime_text.delete(1.0, tk.END)
            totaltime_text.insert(tk.END, f" {total_hours}小时")
            totaltime_text.config(state=tk.DISABLED)
    else:
        messagebox.showerror("文件选择错误", "请选择从 企业微信 导出的文件\n 格式为 “上下班打卡_日报_20250301-20250331.xlsx")

def open_ums_file():
    global DIFF_UMS_FILE_PATH  # 声明为全局变量

    file_path = select_file([("Excel files", "*.xlsx *.xls")], "export", ".xls")
    if file_path:
        if file_path == "file_path is empty":
            return None
        
        dif_text.delete(1.0, tk.END)
        DIFF_UMS_FILE_PATH = file_path
    else:
        messagebox.showerror("文件选择错误", "请选择从 UMS 导出的文件\n 格式为 “export.xls")


def open_ew_file():
    global DIFF_EW_FILE_PATH  # 声明为全局变量
 
    file_path = select_file([("Excel files", "*.xlsx *.xls")], "上下班打卡", ".xlsx")
    if file_path:
        if file_path == "file_path is empty":
            return None
        
        dif_text.delete(1.0, tk.END)
        DIFF_EW_FILE_PATH = file_path
    else:
        messagebox.showerror("文件选择错误", "请选择从 企业微信 导出的文件\n 格式为 “上下班打卡_日报_20250301-20250331.xlsx")

def process_ums_file():
    dif_text.config(state=tk.NORMAL)
    # 清空dif_text中的内容
    dif_text.delete(1.0, tk.END)
    
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

        # 获取ums_data中列名为 工时日期 和 工时(h) 、状态的列
        ums_temp_data = ums_data[['工时日期', '工时(h)', '状态']]
        # 挑选出状态为 “审批完成” 的行赋值给ums_temp_data
        ums_temp_data = ums_temp_data[ums_temp_data['状态'] == '审批完成']
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
            # 确保ums_hours是整数类型
            if isinstance(ums_hours, float): # 如果ums_hours是float类型，那么需要转换为int类型
                ums_hours = int(ums_hours)

            ew_count = 0 # 计数器，用于记录遍历了ew_filtered_data中的多少行，因为ew_filtered_data中index是从4开始的，不符合常规逻辑
            # 遍历ew_filtered_data中的每一行
            for index, ew_row in ew_filtered_data.iterrows(): # 这个index是从4开始的，跟ew excel表中的行号一致，不符合常规逻辑
                ew_count += 1
                ew_date = ew_row[0]

                if "." in ew_row[12]:
                    ew_hours = int(ew_row[12].split(".")[0])
                else: # 如果没有小数点，那么格式就为"7小时"，需要截取"小时"前面的数字
                    ew_hours = int(ew_row[12].split("小时")[0])

                # 如果ew_date中包含date，并且ew_hours不等于hours
                if ums_date in ew_date and ums_hours != ew_hours:
                    # 将结果写入dif_text中
                    dif_text.insert(tk.END, f"工时不对!!! 日期: {ums_date}, ums工时: {ums_hours}, 企业微信工时: {ew_hours}\n")
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
            
            ums_index = 0 # 计数器，用于记录遍历了ums_temp_data中的多少行
            # 遍历ums_temp_data中的每一行
            for index, ums_row in ums_temp_data.iterrows(): # 这个ums_index是从0开始的，符合常规逻辑
                ums_index += 1
                ums_date = ums_row['工时日期']
                ums_hours = ums_row['工时(h)']
                # 确保ums_hours是整数类型
                if isinstance(ums_hours, float): # 如果ums_hours是float类型，那么需要转换为int类型
                    ums_hours = int(ums_hours)

                # 如果ew_date中包含date
                if ums_date in ew_date:
                    break
                
                # 如果遍历完，date一直没有在ew_date中出现，将结果写入dif_text中
                if ums_index == ums_temp_data.shape[0]:
                    dif_text.insert(tk.END, f"ums 缺少 {ew_date} 的记录!!!\n")
                    break

        ###### 处理dif_text中ums相同日期的工时 #####
        # 确定dif_text中是否有内容
        if dif_text.get("1.0", tk.END) == "\n":
            messagebox.showinfo("提示", "ums和企业微信的导出文件一致，没有差异")
            return None
        # 提取dif_text中的内容，用换行符分割
        content = dif_text.get("1.0", tk.END)
        lines = content.split("\n")
        # 确保lines不为空
        if not lines:
            messagebox.showinfo("提示", "ums和企业微信的导出文件一致，没有差异")
            return None
        pre_date = ""
        same_date_hours = 0
        # 定义一个数组，存储日期
        dates_arr = []
        # 遍历lines中的每一行
        for line in lines:
            # 如果line中包含"工时不对"
            if "工时不对" in line:
                # 具体格式为 "工时不对!!! 日期: 2025/02/08, ums工时: 8, 企业微信工时: 11"，需要截取"日期: "后面的内容，和"ums工时: "后面的内容，和"企业微信工时: "后面的内容
                date = line.split("日期: ")[1].split(",")[0]
                ums_hours = int(line.split("ums工时: ")[1].split(",")[0])
                ew_hours = int(line.split("企业微信工时: ")[1])
                
                if pre_date !=date:
                    pre_date = date
                    same_date_hours = ums_hours
                    continue
                else:
                    same_date_hours += ums_hours
                    if same_date_hours == ew_hours:
                        dates_arr.append(date)

        to_remove = []
        for i, line in enumerate(lines):
            if "工时不对" in line:
                date = line.split("日期: ")[1].split(",")[0]
                if date in dates_arr:
                    to_remove.append(i)

        # 倒序删除
        for index in sorted(to_remove, reverse=True):
            del lines[index]
        
        # # 将lines中的内容写入dif_text中
        dif_text.delete("1.0", tk.END)
        for line in lines:
            dif_text.insert(tk.END, line + "\n")

        dif_text.tag_add("red", "1.0", "end")
        dif_text.tag_config("red", foreground="red")
        dif_text.config(state=tk.DISABLED)
    except Exception as e:
        dif_text.config(state=tk.DISABLED)
        messagebox.showerror("错误", f"处理文件时发生错误: {e}")
        return None
    
# 处理按键事件，彩蛋
def on_key_press(event):
    global key_sequence
    # 获取按键的字符
    key = event.char

    # 如果按键是数字键，添加到序列中
    if key.isdigit():
        key_sequence.append(key)

        # 检查是否匹配目标序列 "1024"
        if ''.join(key_sequence[-4:]) == "1024":
            # 弹出消息框
            messagebox.showinfo("彩蛋", "作者是 robot-x")
            key_sequence = []  # 重置按键序列

# 用于使窗口居中
def center_window(root):
    # 获取屏幕尺寸
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # 获取窗口大小
    window_width = root.winfo_width()
    window_height = root.winfo_height()

    # 计算窗口位置
    x_cordinate = int((screen_width/2) - (window_width/2))
    y_cordinate = int((screen_height/2) - (window_height/2))

    # 设置窗口位置
    root.geometry(f'+{x_cordinate - 300}+{y_cordinate - 300}')

# 创建主窗口
root = tk.Tk()
root.title("工时统计工具")
root.geometry("460x500")  # 设置窗口大小
# 调用函数，使窗口居中
center_window(root)

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
# 设置文本框不可以手动更改
overtime_text.config(state=tk.DISABLED)

### 总工时
totaltime_frame = tk.Frame(root)
totaltime_frame.pack(pady=10)
# 创建label并添加到Frame中
totaltime_label = tk.Label(totaltime_frame, text="总工时:")
totaltime_label.pack(side=tk.LEFT)

# 创建文本框用于显示结果
totaltime_text = tk.Text(totaltime_frame, height=1, width=10)
totaltime_text.pack(pady=10)
# 设置文本框不可以手动更改
totaltime_text.config(state=tk.DISABLED)

# 画一条横线来分隔
separator = tk.Frame(root, height=2, bd=1, relief=tk.SUNKEN)
separator.pack(fill=tk.X, padx=5, pady=5)

# 创建一个label，用于显示提示信息
info_label = tk.Label(root, text="如果 UMS 工时和 企业微信 工时 不一致，再点击下面的按钮")
# 将label字体设置为红色，并加粗放大
info_label.config(fg="red", font=("Arial", 12, "bold"))
info_label.pack(pady=10)

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
dif_text = tk.Text(root, height=100, width=320)
dif_text.pack(pady=10)
dif_text.config(state=tk.DISABLED)

# 绑定键盘事件
root.bind("<Key>", on_key_press)

# 运行主循环
root.mainloop()
