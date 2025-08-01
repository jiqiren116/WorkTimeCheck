# 1. python虚拟环境

## 1.1 创建虚拟环境（.venv是虚拟环境的文件夹名）

`python -m venv .venv`

## 1.2 进入虚拟环境

### 1.2.1 powershell中进入

`.venv\Scripts\Activate.ps1`

在vscode中使用powershell，在使用`.venv\Scripts\Activate.ps1`时可能会报错，这时你需要使用命令`Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`来设置powershell脚本的执行策略。`Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser` 是一条用于**设置当前用户 PowerShell 脚本执行策略**的命令，允许运行本地脚本，同时对远程脚本增加了签名验证的安全性。

### 1.2.2 cmd进入

`.venv\Scripts\activate`

## 1.3 退出虚拟环境

`deactivate`

## 使用清华镜像安装依赖

`pip install requests --index-url https://pypi.tuna.tsinghua.edu.cn/simple`