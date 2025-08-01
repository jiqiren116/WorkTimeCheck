import pandas as pd
import tkinter as tk
import os
from tkinter import filedialog
from tkinter import messagebox
import re
import attendance_fetcher  # 导入attendance_fetcher模块
import json

# 添加配置文件路径常量
CONFIG_FILE = "config.json"
key_sequence = []  # 用于存储按键序列，彩蛋
UMS_DATA = {} # 用于存储ums数据
EW_JSON_DATA = {} # 用于存储企业微信导出的json数据
QUERY_START_TIME = ""
QUERY_END_TIME = ""
OVERTIME_REGEX = r'加班(\d+\.\d+)小时'
ABSENCE_REGEX = r'假(\d+\.\d+)小时'
BASE_HOURS = 0
OVERTIME_HOURS = 0
ABSENCE_HOURS = 0

def calculate_hours(file_path):
    global EW_JSON_DATA, BASE_HOURS, OVERTIME_HOURS, ABSENCE_HOURS
    try:
        # 读取Excel文件，从第4行开始读取，第四行为列名
        excel_data = pd.read_excel(file_path, header=3, sheet_name=0)

        # 初始化EW_JSON_DATA
        EW_JSON_DATA = {}
        for index, row in excel_data.iterrows():
            reportTime = row.iloc[0] # 日报日期
            reportTime = reportTime.split(" ")[0].replace("/", "-")
            EW_JSON_DATA[reportTime] = 0

            base_hour = row["标准工作时长(小时)"]
            if base_hour != "--":
                EW_JSON_DATA[reportTime] = 8
                BASE_HOURS += 8
            overtime_and_absence_hour = row["假勤申请"]
            if overtime_and_absence_hour != "--":
                if "加班" in overtime_and_absence_hour:
                    # 用pattern提取加班小时数
                    overtime_hours = re.search(OVERTIME_REGEX, overtime_and_absence_hour)
                    if overtime_hours:
                        # 提取出的overtime_hours格式为“加班3.0小时”，我只想要其中的整数部分
                        overtime_hours = overtime_hours.group(1)
                        # overtime_hours格式为“3.0”，是字符串，转换为整数
                        overtime_hours = int(float(overtime_hours))
                        EW_JSON_DATA[reportTime] = EW_JSON_DATA[reportTime] + overtime_hours
                        OVERTIME_HOURS += overtime_hours
                if "假" in overtime_and_absence_hour:
                    # 用pattern提取假勤小时数
                    absence_hours = re.search(ABSENCE_REGEX, overtime_and_absence_hour)
                    if absence_hours:
                        # 提取出的absence_hours格式为“假3.0小时”，我只想要其中的整数部分
                        absence_hours = absence_hours.group(1)
                        # 提取出的absence_hours格式为“3.0”，我只想要其中的整数部分
                        absence_hours = absence_hours.split(".")[0]
                        if absence_hours:
                            absence_hours = int(float(absence_hours))
                            EW_JSON_DATA[reportTime] = EW_JSON_DATA[reportTime] - absence_hours    
                            ABSENCE_HOURS += absence_hours
    except Exception as e:
        messagebox.showerror("错误", f"处理文件时发生错误: {e}")

def calculate_base_overtime_absence_hours(file_path):
    calculate_hours(file_path)



def select_file(filetypes, startswith, endswith):
    global QUERY_START_TIME, QUERY_END_TIME
    # 打开文件选择对话框
    file_path = filedialog.askopenfilename(filetypes=filetypes)
    
    # 截取文件名
    file_name = os.path.basename(file_path)
    # 判断file_name是否满足条件
    if file_path and file_name.startswith(startswith) and file_name.endswith(endswith):
        # 获取file_path文件名中的开始和结束日期，格式为“上下班打卡_日报_20250701-20250731.xlsx”，最后将日记转换为'2025-07-01'
        file_name = os.path.basename(file_path)
        file_name = file_name.split(".")[0]
        file_name = file_name.split("_")[2]
        QUERY_START_TIME = file_name.split("-")[0]
        QUERY_END_TIME = file_name.split("-")[1]
        # 转换为'2025-07-01'格式
        QUERY_START_TIME = QUERY_START_TIME[:4] + "-" + QUERY_START_TIME[4:6] + "-" + QUERY_START_TIME[6:]
        QUERY_END_TIME = QUERY_END_TIME[:4] + "-" + QUERY_END_TIME[4:6] + "-" + QUERY_END_TIME[6:]
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
        calculate_base_overtime_absence_hours(file_path)
        base_hours = BASE_HOURS
        overtime_hours = OVERTIME_HOURS
        absence_hours = ABSENCE_HOURS
        total_hours = base_hours + overtime_hours - absence_hours # 计算总工时
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

def open_ew_file():
    file_path = select_file([("Excel files", "*.xlsx *.xls")], "上下班打卡", ".xlsx")
    if file_path:
        if file_path == "file_path is empty":
            return None
        
        dif_text.delete(1.0, tk.END)
    else:
        messagebox.showerror("文件选择错误", "请选择从 企业微信 导出的文件\n 格式为 “上下班打卡_日报_20250301-20250331.xlsx")

def work_hours_compare():
    dif_text.config(state=tk.NORMAL)
    # 清空dif_text中的内容
    dif_text.delete(1.0, tk.END)
   
    try:
        if not validate_file_paths_and_account():
            return None

        # ums_temp_data, ew_filtered_data = read_and_process_data() # 读取数据

        compare_hours() # 比较工时

        # handle_same_date_hours() # 处理dif_text中ums相同日期的工时

        dif_text.tag_add("red", "1.0", "end")
        dif_text.tag_config("red", foreground="red")
        dif_text.config(state=tk.DISABLED)
    except Exception as e:
        dif_text.config(state=tk.DISABLED)
        messagebox.showerror("错误", f"处理文件时发生错误: {e}")
        return None

# 判断文件路径和账号是否为空
def validate_file_paths_and_account():
    global EW_JSON_DATA, UMS_DATA
    if EW_JSON_DATA == {} and UMS_DATA == {}:
        messagebox.showerror("错误", "请先选择企业微信的导出文件和输入账号，密码获取UMS工时")
        return False
    elif EW_JSON_DATA == {}:
        messagebox.showerror("错误", "请先选择企业微信的导出文件")
        return False
    elif UMS_DATA == {}:
        messagebox.showerror("错误", "请先输入账号，密码获取UMS工时")
        return False
    return True
        

def compare_hours():
    global UMS_DATA, EW_JSON_DATA
    # 其中EW_JSON_DATA的键为日期，格式为"2025-07-30",值为整数，而UMS_DATA键为日期，格式为"2025-07-30 00:00:00",值为小数，
    # 遍历EW_JSON_DATA,看EW_JSON_DATA的键是否在UMS_DATA中出现过，
    # 如果出现过，就将EW_JSON_DATA的值和UMS_DATA的值进行比较，相同就pass，不同就打印
    # 遍历EW_JSON_DATA,比较两个数据源的工时

    # 将UMS_DATA转换为字典格式便于比较
    ums_dict = {}
    for record in UMS_DATA:
        # 从 "2025-07-31 00:00:00" 提取日期部分 "2025-07-31"
        date = record['reportTime'].split(' ')[0]
        hours = int(float(record['workHours']))
        ums_dict[date] = hours

    # print(UMS_DATA)
    # 遍历EW_JSON_DATA,比较两个数据源的工时
    flag = True
    for ew_date in EW_JSON_DATA:
        ew_hours = EW_JSON_DATA[ew_date]
        if ew_date in ums_dict:
            ums_hours = ums_dict[ew_date]
            if ew_hours != ums_hours:
                flag = False
                dif_text.insert(tk.END, f"工时不对!!! 日期: {ew_date}, 企业微信工时: {ew_hours}, ums工时: {ums_hours}\n")
        elif ew_hours != 0:
            flag = False
            dif_text.insert(tk.END, f"UMS中无此日期数据!!! 日期: {ew_date}, 企业微信工时: {ew_hours}, ums工时: 0\n")
    if flag:
        messagebox.showinfo("OK", "工时一致，没有差异")

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
        # 如果data是字典且包含rows键
        if isinstance(data, dict) and 'rows' in data:
            records = data['rows']
        # 如果data本身就是列表
        elif isinstance(data, list):
            records = data
        else:
            records = []
            
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
    global UMS_DATA
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
    
    # 调用attendance_fetcher获取数据
    try:
        ums_data = attendance_fetcher.fetch_attendance_data(
            username=ums_account,
            password=ums_password,
            begin_time=QUERY_START_TIME,
            end_time=QUERY_END_TIME
        )

        # 只提取rows部分的数据
        if ums_data and 'rows' in ums_data:
            UMS_DATA = ums_data['rows']  # 只保存rows部分

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
    root.geometry("460x700")  # 设置窗口大小
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
    # open_ums_file_button = tk.Button(diff_frame, text="选择UMS导出的Excel文件", command=open_ums_file)
    # open_ums_file_button.pack(side=tk.LEFT, padx=5)
    open_ew_file_button2 = tk.Button(diff_frame, text="工时比对", command=work_hours_compare)
    open_ew_file_button2.pack(side=tk.LEFT, padx=5)

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
