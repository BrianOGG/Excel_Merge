import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def merge_excel_files():
    try:
        # 获取并验证文件类型
        file_type = file_type_entry.get().strip()
        if not file_type:
            raise ValueError("文件类型不能为空")
        if not file_type.startswith('.'):
            file_type = '.' + file_type

        # 获取并验证表头行数
        header_num_str = header_num_entry.get().strip()
        if not header_num_str:
            raise ValueError("表头行数不能为空")
        header_num = int(header_num_str)
        if header_num < 1:
            raise ValueError("表头行数必须至少为1")

        # 选择目录
        directory = filedialog.askdirectory()
        if not directory:
            messagebox.showinfo("取消", "操作已取消：未选择目录")
            return

        # 验证目录存在
        if not os.path.isdir(directory):
            raise FileNotFoundError(f"目录不存在：{directory}")

        # 获取文件列表
        excel_files = [f for f in os.listdir(directory) if f.endswith(file_type)]
        if not excel_files:
            raise ValueError(f"目录中没有找到{file_type}类型的文件")

        # 设置读取引擎
        engine_map = {
            '.xls': 'openpyxl',
            '.xlsx': 'openpyxl',
            '.xlsm': 'openpyxl'
        }
        read_engine = engine_map.get(file_type, None)

        # 读取文件
        dataframes = []
        for file in excel_files:
            try:
                file_path = os.path.join(directory, file)
                df = pd.read_excel(
                    file_path, 
                    header=header_num-1,
                    engine=read_engine
                )
                dataframes.append(df)
                print(f"成功读取：{file}")
            except Exception as e:
                messagebox.showwarning("读取警告", f"文件 {file} 读取失败：{str(e)}")

        # 检查有效数据
        if not dataframes:
            raise ValueError("没有成功读取任何有效数据")

        # 合并数据
        merged_df = pd.concat(dataframes, ignore_index=True)

        # 处理输出文件名
        result_name = result_name_entry.get().strip()
        if not result_name:
            result_name = "合并结果"  # 默认文件名
        
        base_name = os.path.splitext(result_name)[0]
        output_filename = f"{base_name}{file_type}"
        output_path = os.path.join(directory, output_filename)

        # 设置写入引擎
        write_engine_map = {
            '.xls': 'xlwt',
            '.xlsx': 'openpyxl',
            '.xlsm': 'openpyxl'
        }
        write_engine = write_engine_map.get(file_type, None)

        # 保存文件
        try:
            merged_df.to_excel(
                output_path,
                index=False,
                engine=write_engine
            )
            messagebox.showinfo("成功", f"文件已成功保存至：\n{output_path}")
        except Exception as e:
            raise RuntimeError(f"保存文件失败：{str(e)}")

    except ValueError as ve:
        messagebox.showerror("输入错误", str(ve))
    except Exception as e:
        messagebox.showerror("程序错误", f"发生未预期错误：\n{str(e)}")

# 创建主窗口
root = tk.Tk()
root.title('Excel文件合并工具')

# 文件类型输入
tk.Label(root, text='文件后缀（如 .xls）:').pack(pady=5)
file_type_entry = tk.Entry(root)
file_type_entry.pack(pady=5)
file_type_entry.insert(0, ".xlsx")  # 默认值

# 表头行数输入
tk.Label(root, text='表头行数（从1开始计数）:').pack(pady=5)
header_num_entry = tk.Entry(root)
header_num_entry.pack(pady=5)
header_num_entry.insert(0, "1")  # 默认值

# 输出文件名
tk.Label(root, text='输出文件名（无需后缀）:').pack(pady=5)
result_name_entry = tk.Entry(root)
result_name_entry.pack(pady=5)
result_name_entry.insert(0, "合并结果")  # 默认值

# 合并按钮
tk.Button(root, text='选择目录并合并', command=merge_excel_files, bg='#4CAF50', fg='white').pack(pady=20)

root.mainloop()
