import pandas as pd
import tkinter as tk
import os
from tkinter import filedialog
from tkinter import messagebox
import re
import attendance_fetcher  # 导入attendance_fetcher模块
from datetime import datetime, timedelta  # 添加日期处理依赖
import json

# 添加配置文件路径常量
CONFIG_FILE = "config.json"
DIFF_UMS_FILE_PATH = ""
DIFF_EW_FILE_PATH = ""
key_sequence = []  # 用于存储按键序列，彩蛋

def calculate_hours(file_path, column_name, pattern):
    try:
        # 读取Excel文件，从第4行开始读取，第四行为列名
        data = pd.read_excel(file_path, header=3, sheet_name=0)
        
        # 直接通过列名获取数据
        temp_data = data[column_name]

        # 筛选包含特定关键词的数据
        filtered_data = temp_data[~temp_data.str.contains("--", na=False)]

        # 提取小时数并转换为数值类型，如果数据中有不满足正则表达式的情况会返回NaN
        hours_data = filtered_data.str.extract(pattern).astype(float)
        # 处理 NaN 值（例如删除或填充）
        hours_data = hours_data.dropna()  # 删除包含 NaN 的行

        # 保留整数部分
        hours_data = hours_data[0].astype(int)

        # 求和
        result_hours = hours_data.sum()

        return result_hours
    except Exception as e:
        messagebox.showerror("错误", f"处理文件时发生错误: {e}")
        return None

def calculate_base_hours(file_path):
    # r'(\d+)\.?\d*'表示如果原123.45，678，901.23，会提取123，678，901
    temp_base_hours = calculate_hours(file_path, "标准工作时长(小时)", r'(\d+)\.?\d*')
    base_hours = int(temp_base_hours + temp_base_hours / 7)
    return base_hours

def calculate_overtime_hours(file_path):
    return calculate_hours(file_path, "假勤申请", r'加班(\d+\.\d+)小时')

def calculate_absence_hours(file_path):
    return calculate_hours(file_path, "假勤申请", r'假(\d+\.\d+)小时')

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
        # 如果选择文件对话框中点击取消
        if file_path == "file_path is empty":
            return None
        
        # 计算加班时长 和 总工时
        overtime_hours = calculate_overtime_hours(file_path)
        absence_hours = calculate_absence_hours(file_path)
        total_hours = calculate_base_hours(file_path) + overtime_hours - absence_hours # 计算总工时
        if overtime_hours is not None:
            # 将加班时长显示在文本框中
            overtime_text.config(state=tk.NORMAL)
            overtime_text.delete(1.0, tk.END)
            overtime_text.insert(tk.END, f" {overtime_hours}小时")
            overtime_text.config(state=tk.DISABLED)
        if absence_hours is not None:
            # 将请假时长显示在文本框中
            absence_text.config(state=tk.NORMAL)
            absence_text.delete(1.0, tk.END)
            absence_text.insert(tk.END, f" {absence_hours}小时")
            absence_text.config(state=tk.DISABLED)
        if total_hours is not None:
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
   
    try:
        # 首先判断DIFF_UMS_FILE_PATH和DIFF_EW_FILE_PATH是否为空
        if not validate_file_paths():
            return None

        ums_temp_data, ew_filtered_data = read_and_process_data() # 读取数据

        compare_hours(ums_temp_data, ew_filtered_data) # 比较工时

        handle_same_date_hours() # 处理dif_text中ums相同日期的工时

        dif_text.tag_add("red", "1.0", "end")
        dif_text.tag_config("red", foreground="red")
        dif_text.config(state=tk.DISABLED)
    except Exception as e:
        dif_text.config(state=tk.DISABLED)
        messagebox.showerror("错误", f"处理文件时发生错误: {e}")
        return None

# 用于判断DIFF_UMS_FILE_PATH和DIFF_EW_FILE_PATH是否为空
def validate_file_paths():
    if DIFF_UMS_FILE_PATH == "" and DIFF_EW_FILE_PATH == "":
        messagebox.showerror("错误", "请先选择ums和企业微信的导出文件")
        dif_text.config(state=tk.DISABLED)
        return False
    elif DIFF_UMS_FILE_PATH == "":
        messagebox.showerror("错误", "请先选择ums的导出文件")
        dif_text.config(state=tk.DISABLED)
        return False
    elif DIFF_EW_FILE_PATH == "":
        messagebox.showerror("错误", "请先选择企业微信的导出文件")
        dif_text.config(state=tk.DISABLED)
        return False
    return True

def read_and_process_data():
    ums_data = pd.read_html(DIFF_UMS_FILE_PATH, encoding='utf-8')[0]  # [0] 表示读取第一个表格
    ew_data = pd.read_excel(DIFF_EW_FILE_PATH, header=None, sheet_name=0)  # 读取第一个sheet

    # 获取ums_data中列名为 工时日期 和 工时(h) 、状态的列
    ums_temp_data = ums_data[['工时日期', '工时(h)', '状态']]
    # 挑选出状态为 “审批完成”或者是“待审批” 的行赋值给ums_temp_data
    # ums_temp_data = ums_temp_data[ums_temp_data['状态'] == '审批完成']
    ums_temp_data = ums_temp_data[ums_temp_data['状态'].isin(['审批完成', '待审批'])]
    # 将 “工时日期” 这一列格式为"2025-03-01"转换为"2025/03/01"
    ums_temp_data['工时日期'] = ums_temp_data['工时日期'].str.replace('-', '/')

    num_rows = ew_data.shape[0]
    ew_temp_data = ew_data.loc[4: num_rows - 1, [0, 11,12,13]]

    # 筛选出包含“小时”的行
    ew_filtered_data = ew_temp_data[ew_temp_data[12].astype(str).str.contains("小时")]

    return ums_temp_data, ew_filtered_data

def compare_hours(ums_temp_data, ew_filtered_data):
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

            temp_hour = 0
            # 如果ew_row[11]中包含“小时”，说明这最起码是8h
            if "小时" in ew_row[11]:
                temp_hour = 8
            # 如果ew_row[13]中包含“加班”，说明这是加班或者病假等
            if "小时" in ew_row[13]:
                # 按照正则表达式提取出数字 "'加班(\d+\.\d+)小时'"
                if "加班" in ew_row[13]:
                    pattern = r'加班(\d+\.\d+)小时'
                    match = re.search(pattern, ew_row[13])
                    if match:
                        temp_overtime = float(match.group(1))
                        temp_hour = int(temp_hour + temp_overtime)
                if "病假" in ew_row[13]:
                    pattern = r'病假(\d+\.\d+)小时'
                    match = re.search(pattern, ew_row[13])
                    if match:
                        temp_sicktime = float(match.group(1))
                        temp_hour = int(temp_hour - temp_sicktime)
                        # 如果temp_hour小于0,弹窗报错
                        if temp_hour < 0:
                            messagebox.showerror("错误", f"时间计算错误，{ums_date} 的 企业微信 工时为 {temp_hour} 小时，小于0")
                            dif_text.config(state=tk.DISABLED)
                            return None
            ew_hours = int(temp_hour)

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

# 处理dif_text中ums相同日期的工时
def handle_same_date_hours():
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
            messagebox.showinfo("彩蛋", "作者是 robot-x\n 源码https://github.com/jiqiren116/WorkTimeCheck\n软件更新日期：2025年4月18日 17:08:27")
            key_sequence = []  # 重置按键序列

def calculate_ums_worktime_from_json(data):
    """从UMS返回的JSON数据中计算总工时"""
    try:
        # 根据实际JSON结构调整键名（这里假设数据在'data'列表中）
        records = data.get('rows', [])
        total_hours = 0
        for record in records:
            # 累加工时（假设工时字段为'workHours'）
            hours = record.get('workHours', 0)
            total_hours += int(float(hours))  # 处理可能的浮点数
        return total_hours
    except Exception as e:
        messagebox.showerror("错误", f"处理UMS数据失败: {str(e)}")
        return None

# 添加配置文件读写函数
def load_config():
    """加载保存的账号密码配置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                ums_account_entry.insert(0, config.get('account', ''))
                ums_password_entry.insert(0, config.get('password', ''))
                remember_var.set(True)
        except Exception as e:
            print(f"加载配置失败: {e}")

def save_config(account, password):
    """保存账号密码到配置文件"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({
                'account': account,
                'password': password
            }, f, ensure_ascii=False, indent=4)
    except Exception as e:
        messagebox.showerror("错误", f"保存配置失败: {e}")

def get_ums_worktime():
    # 从输入框获取账号密码
    ums_account = ums_account_entry.get().strip()
    ums_password = ums_password_entry.get().strip()
    
    # 验证输入
    if not ums_account or not ums_password:
        messagebox.showwarning("提示", "请输入账号和密码")
        return None
    
    # 如果勾选了记住账号密码，则保存
    if remember_var.get():
        save_config(ums_account, ums_password)
    
    # 计算当前月份的起止日期（示例：当前月1号到月底）
    today = datetime.today()
    first_day = today.replace(day=1).strftime('%Y-%m-%d')
    last_day = (today.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    last_day = last_day.strftime('%Y-%m-%d')
    
    # 调用attendance_fetcher获取数据
    try:
        ums_data = attendance_fetcher.fetch_attendance_data(
            username=ums_account,
            password=ums_password,
            begin_time=first_day,
            end_time=last_day
        )

        if ums_data:
            # 计算总工时并显示
            total_hours = calculate_ums_worktime_from_json(ums_data)
            if total_hours is not None:
                ums_worktime_text.config(state=tk.NORMAL)
                ums_worktime_text.delete(1.0, tk.END)
                ums_worktime_text.insert(tk.END, f" {total_hours}小时")
                ums_worktime_text.config(state=tk.DISABLED)
        else:
            messagebox.showerror("错误", "获取UMS数据失败，请检查账号密码或网络")
    except Exception as e:
        messagebox.showerror("错误", f"调用UMS接口失败: {str(e)}")

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

def create_gui():
    global overtime_text, totaltime_text, dif_text, absence_text  # 声明全局控件变量
    global ums_account_entry, ums_password_entry, ums_worktime_text, remember_var  

    def on_account_change(event):
        """当账号输入框内容发生变化时调用"""
        if remember_var.get():
            remember_var.set(False)

    def on_password_change(event):
        """当密码输入框内容发生变化时调用"""
        if remember_var.get():
            remember_var.set(False)
    
    # 创建主窗口
    root = tk.Tk()
    root.title("工时统计工具")
    root.geometry("460x500")  # 设置窗口大小
    center_window(root)

    # 创建主水平框架
    main_frame = tk.Frame(root)
    main_frame.pack(pady=10, fill=tk.X, padx=10)

    # 左侧信息区域
    left_frame = tk.Frame(main_frame)
    left_frame.pack(side=tk.LEFT)

    # 企业微信文件选择区域
    open_ew_file_frame = tk.Frame(left_frame)
    open_ew_file_frame.pack(pady=10)
    open_ew_file_label = tk.Label(open_ew_file_frame, text="选择企业微信导出的Excel文件\n 来查看 加班时间 和 总工时:")
    open_ew_file_label.pack(side=tk.LEFT)
    open_ew_file_button = tk.Button(open_ew_file_frame, text="选择文件", command=open_file)
    open_ew_file_button.pack(side=tk.LEFT, padx=5)

    # 加班时长显示区域
    overtime_frame = tk.Frame(left_frame)
    overtime_frame.pack(pady=10)
    overtime_label = tk.Label(overtime_frame, text="加班时长:")
    overtime_label.pack(side=tk.LEFT)
    overtime_text = tk.Text(overtime_frame, height=1, width=10)
    overtime_text.pack(side=tk.LEFT, padx=5)
    overtime_text.config(state=tk.DISABLED)

    # 请假时长显示区域
    absence_frame = tk.Frame(left_frame)
    absence_frame.pack(pady=10)
    absence_label = tk.Label(absence_frame, text="请假时长:")
    absence_label.pack(side=tk.LEFT)
    absence_text = tk.Text(absence_frame, height=1, width=10)
    absence_text.pack(side=tk.LEFT, padx=5)
    absence_text.config(state=tk.DISABLED)

    # 总工时显示区域
    totaltime_frame = tk.Frame(left_frame)
    totaltime_frame.pack(pady=10)
    totaltime_label = tk.Label(totaltime_frame, text="企业微信总工时:")
    totaltime_label.pack(side=tk.LEFT)
    totaltime_text = tk.Text(totaltime_frame, height=1, width=10)
    totaltime_text.pack(pady=10)
    totaltime_text.config(state=tk.DISABLED)

    # 中间分割线
    separator = tk.Frame(main_frame, width=2, bg="gray")
    separator.pack(side=tk.LEFT, fill=tk.Y, padx=10)

    # 右侧按钮区域
    right_frame = tk.Frame(main_frame)
    right_frame.pack(side=tk.LEFT)

    # 右侧标签
    right_label = tk.Label(right_frame, text="输入ums账号,密码获取UMS工时", fg="blue")
    right_label.pack(pady=20)

    # ums账号输入框
    ums_account_frame = tk.Frame(right_frame)
    ums_account_frame.pack(pady=10)
    ums_account_label = tk.Label(ums_account_frame, text="UMS账号:")
    ums_account_label.pack(side=tk.LEFT)
    ums_account_entry = tk.Entry(ums_account_frame)
    ums_account_entry.pack(side=tk.LEFT, padx=5)
    ums_account_entry.bind('<KeyRelease>', on_account_change)
    # ums密码输入框
    ums_password_frame = tk.Frame(right_frame)
    ums_password_frame.pack(pady=10)
    ums_password_label = tk.Label(ums_password_frame, text="UMS密码:") 
    ums_password_label.pack(side=tk.LEFT)
    ums_password_entry = tk.Entry(ums_password_frame)
    ums_password_entry.pack(side=tk.LEFT, padx=5)
    ums_password_entry.bind('<KeyRelease>', on_password_change)
    
    # 添加记住账号密码选项
    remember_var = tk.BooleanVar()
    remember_check = tk.Checkbutton(right_frame, text="记住账号密码", variable=remember_var)
    remember_check.pack(pady=5)

    # 加载保存的配置
    load_config()
    
    # 获取ums工时按钮
    get_ums_worktime_button = tk.Button(right_frame, text="获取UMS工时", command=get_ums_worktime)
    get_ums_worktime_button.pack(pady=20)

    # ums工时显示区域
    ums_worktime_frame = tk.Frame(right_frame)
    ums_worktime_frame.pack(pady=10)
    ums_worktime_label = tk.Label(ums_worktime_frame, text="UMS工时:")
    ums_worktime_label.pack(side=tk.LEFT)
    ums_worktime_text = tk.Text(ums_worktime_frame, height=1, width=10)
    ums_worktime_text.pack(side=tk.LEFT, padx=5)
    ums_worktime_text.config(state=tk.DISABLED)
    

    # 分隔线
    separator = tk.Frame(root, height=2, bd=1, relief=tk.SUNKEN)
    separator.pack(fill=tk.X, padx=5, pady=5)

    # 提示信息
    info_label = tk.Label(root, text="如果 UMS 工时和 企业微信 工时 不一致，再点击下面的按钮")
    info_label.config(fg="red", font=("Arial", 12, "bold"))
    info_label.pack(pady=10)

    # 文件对比选择区域
    diff_frame = tk.Frame(root)
    diff_frame.pack(pady=10)
    open_ums_file_button = tk.Button(diff_frame, text="选择UMS导出的Excel文件", command=open_ums_file)
    open_ums_file_button.pack(side=tk.LEFT, padx=5)
    open_ew_file_button2 = tk.Button(diff_frame, text="选择企业微信导出的Excel文件", command=open_ew_file)
    open_ew_file_button2.pack(side=tk.LEFT, padx=5)

    # 对比执行区域
    diff_frame2 = tk.Frame(root)
    diff_frame2.pack(pady=10)
    diff_label = tk.Label(diff_frame2, text="选中两个文件后，点击执行对比，查看差异:")
    diff_label.pack(side=tk.LEFT)
    execute_diff_button = tk.Button(diff_frame2, text="执行对比", command=process_ums_file)
    execute_diff_button.pack(side=tk.LEFT, padx=5)

    # 差异结果显示区域
    dif_text = tk.Text(root, height=100, width=320)
    dif_text.pack(pady=10)
    dif_text.config(state=tk.DISABLED)

    # 绑定键盘事件
    root.bind("<Key>", on_key_press)
    
    return root

if __name__ == "__main__":
    root = create_gui()
    root.mainloop()
