import tkinter as tk
from tkinter import filedialog, scrolledtext
from docx import Document


def analyze_table(table, table_comment):
    table_name = ""
    db_name = ""
    columns_definitions = []
    newline = ",\n    "
    primary_keys = []

    # 解析SQL表和列定义
    for i, row in enumerate(table.rows):
        cells = [cell.text.strip() for cell in row.cells]

        # 假定第一行包含列的注释，在第二列
        if i == 0:
            table_name = cells[1]
            db_name = cells[6]

        # 假定列定义从第三行开始
        elif i > 1:
            column_name = cells[1]
            if not column_name:  # 如果列名为空，则跳过该行
                continue

            # 建立列定义字符串
            data_type = cells[2]
            is_null = "NULL" if cells[3] == "" else "NOT NULL"
            is_key = cells[4]
            default = f"DEFAULT {cells[5]}" if cells[5] else ""

            column_definition = f"{column_name} {data_type} {is_null} {default}".strip()

            if is_key == "主键":
                primary_keys.append(column_name)
                # 在MySQL中，设置了PRIMARY KEY的列不能被定义为NULL
                # column_definition = column_definition.replace("NULL", "")
            elif is_key == "外键":
                # 在此示例中，外键的处理略过了，因为要建立外键，还需要知道外键引用了哪个表和列
                pass

            columns_definitions.append(column_definition)

    primary_key_definition = f"PRIMARY KEY ({', '.join(primary_keys)})" if primary_keys else ""
    columns_definitions.append(primary_key_definition)

    # 组装CREATE TABLE语句
    table_sql = f"if not exits (select * from sys.sysobjects where name='{table_name}') \nbegin \n"
    table_sql += f"CREATE TABLE {table_name} (\n    {''.join({newline}).join(col for col in columns_definitions if col)}\n);"
    # table_sql = f"CREATE TABLE {table_name.lower()} ({newline}    {',{newline}    '.join(col for col in columns_definitions if col)}{newline});"

    # 添加表的注释（在MySQL中的语法示例，根据你的数据库类型可能有所不同）
    comment_sql = f"EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'{table_comment}', @level0type=N'Schema', @level0name=N'dbo', @level1type=N'Table', @level1name=N'{table_name} '; "
    number_of_columns = len([col for col in columns_definitions if col and not col.startswith("PRIMARY KEY")])

    column_comments = []
    for i in range(2, number_of_columns + 2):
        column_name = table.cell(i, 1).text.strip()
        column_comment = table.cell(i, 6).text.strip()
        if column_comment:
            comment_statement = f"EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'{column_comment}', @level0type=N'Schema', @level0name=N'dbo', @level1type=N'Table', @level1name=N'{table_name}', @level2type=N'Column', @level2name=N'{column_name}'; "
            column_comments.append(comment_statement)

    return table_sql + "\n" + comment_sql + "\n" + "\n".join(column_comments) + "\nend"

def get_table_preceding_paragraph(table):
    tbl_element = table._element
    prev_element = tbl_element.getprevious()

    if prev_element is not None and prev_element.tag.endswith('p'):
        return prev_element.text
    else:
        # 找到真正的文本段落
        while prev_element is not None and not prev_element.tag.endswith('p'):
            prev_element = prev_element.getprevious()
        if prev_element is not None:
            return prev_element.text
    return ""  # 如果之前没有段落，则返回空字符串

def open_docx():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx"), ("All files", "*.*")])
    if not file_path:
        return

    try:
        doc = Document(file_path)
        sql_output = ""  # Store all the SQL statements here
        for table in doc.tables:
            table_comment = get_table_preceding_paragraph(table)
            sql_output += analyze_table(table, table_comment) + "\n\n"

        text_area.delete('1.0', tk.END)
        text_area.insert(tk.INSERT, sql_output)
    except Exception as e:
        text_area.delete('1.0', tk.END)
        text_area.insert(tk.INSERT, f'Error: {e}')


root = tk.Tk()
root.title("Word文档与表结构转换工具")

text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD)
text_area.pack(expand=True, fill=tk.BOTH)

open_button = tk.Button(root, text="打开Word文档", command=open_docx)
open_button.pack(side=tk.TOP, pady=10)

root.mainloop()
