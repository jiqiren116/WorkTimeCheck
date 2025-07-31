import requests
import json
from datetime import datetime

# 登录和数据URL（保持不变）
LOGIN_URL = 'http://192.168.117.70:8080/cover/login'
REPORT_URL = 'http://192.168.117.70:8080/workhours/approve/getList'

# 设置请求头（保持不变）
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
}

# 修改函数参数，接收用户名、密码和日期范围
def fetch_attendance_data(username, password, begin_time, end_time):
    try:
        # 创建会话
        session = requests.Session()
        
        # 动态生成登录数据（使用传入的用户名密码）
        login_data = {
            'username': username,
            'password': password,
            'rememberMe': 'false'
        }
        
        # 发送登录请求
        print("正在登录系统...")
        login_response = session.post(LOGIN_URL, data=login_data, headers=headers)
        
        # 检查登录是否成功
        if login_response.status_code == 200:
            print("登录成功")
            
            # 动态生成日报数据请求参数（使用传入的日期范围）
            report_data = {
                'projectId': '',
                'beginTime': begin_time,
                'endTime': end_time,
                'status': '2',
                'reportName': '',
                'workType': '',
                'pageSize': 50,
                'pageNum': 1,
                'orderByColumn': '',
                'isAsc': 'asc'
            }
            
            # 发送获取日报数据的请求（利用session自动处理cookie）
            print("正在获取日报数据...")
            report_response = session.post(REPORT_URL, data=report_data, headers=headers)
            
            # 检查数据请求是否成功
            if report_response.status_code == 200:
                # 尝试解析JSON数据
                try:
                    data = report_response.json()
                    # # 生成输出文件名
                    # now = datetime.now()
                    # filename = f"日报数据_{now.strftime('%Y%m%d_%H%M%S')}.json"
                    # # 保存数据到文件
                    # with open(filename, 'w', encoding='utf-8') as f:
                    #     json.dump(data, f, ensure_ascii=False, indent=4)
                    # print(f"数据获取成功，已保存到 {filename}")
                    return data
                except json.JSONDecodeError:
                    # 如果返回的不是JSON格式，保存原始数据
                    now = datetime.now()
                    filename = f"日报原始数据_{now.strftime('%Y%m%d_%H%M%S')}.txt"
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write(report_response.text)
                    print(f"返回数据不是JSON格式，已保存原始数据到 {filename}")
                    return None
            else:
                print(f"获取日报数据失败，状态码: {report_response.status_code}")
                return None
        else:
            print(f"登录失败，状态码: {login_response.status_code}")
            return None
            
    except requests.RequestException as e:
        print(f"请求发生错误: {e}")
        return None
    except Exception as e:
        print(f"发生未知错误: {e}")
        return None

if __name__ == "__main__":
    # 测试用例（实际使用由main.py调用）
    fetch_attendance_data('caomingyuan', '12345678', '2025-07-01', '2025-07-31')