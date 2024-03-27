# DocToTable
表结构与Word文档之间互相转换工具

### 创建虚拟环境
##### venv目录通常是指一个虚拟环境（virtual environment）
python -m venv venv

#### 激活虚拟环境
##### 在Unix或MacOS上：
source venv/bin/activate
##### 在Windows上：
venv\Scripts\activate


### 安装依赖
##### 运行环境
pip install python-docx

pip install pyodbc

pip install pyinstaller


### 打包命令
pyinstaller --onefile .\DocToTable9.py
