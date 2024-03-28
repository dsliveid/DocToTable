import tkinter as tk
from tkinter import filedialog, scrolledtext
from docx import Document

def analyze_table(table):
    # 初始化SQL，并设置一个标志来发现我们正在进行的表的标题
    table_sql = ""
    table_name = ""
    newline = "\n"
    columns = []

    # 遍历表格的每一行
    for i, row in enumerate(table.rows):
        text_in_cells = [cell.text.strip() for cell in row.cells]

        if i == 0:
            table_name = text_in_cells[1]  # 假设表名位于第二列
        elif i == 2:  # 假设列描述从第三行开始
            # 初始化列定义，跳过行号和最后的自定义列
            columns = ["{} {} {}".format(
                col_name.replace(' ', '_').lower(),  # 假设字段名为列名的小写形式，空格替换为下划线
                col_type,
                "NOT NULL" if "是" in null_flag else "NULL",
            ) for col_name, col_type, null_flag in zip(text_in_cells[1::6], text_in_cells[2::6], text_in_cells[3::6])]
        elif i > 2:
            # 跳过空行
            if set(text_in_cells) == {''}:
                continue
            if "主键" in text_in_cells:
                primary_key = text_in_cells[1].replace(' ', '_').lower()
                columns = [col + " PRIMARY KEY" if primary_key in col else col for col in columns]

    table_sql = f"CREATE TABLE {table_name.lower()} ({newline}    {',{newline}    '.join(columns)}{newline});"
    return table_sql

def open_docx():
    # 使用文件对话框选择.docx文件
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx"), ("All files", "*.*")])
    if not file_path:
        return

    try:
        # 加载Word文档
        doc = Document(file_path)

        sql_output = ""
        for table in doc.tables:
            sql_output += analyze_table(table) + '\n\n'

        # 显示在文本框中
        text_area.delete('1.0', tk.END)
        text_area.insert(tk.INSERT, sql_output)

    except Exception as e:
        text_area.delete('1.0', tk.END)
        text_area.insert(tk.INSERT, f'Error: {e}')

# 创建主窗口
root = tk.Tk()
root.title("Word文档表解析器")

# 创建滚动文本区域以显示文档内容
text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD)
text_area.pack(expand=True, fill=tk.BOTH)

# 添加按钮以打开文件
open_button = tk.Button(root, text="打开Word文档", command=open_docx)
open_button.pack(side=tk.TOP, pady=10)

# 运行主循环
root.mainloop()
