import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Excel文件合并工具')
        self.root.resizable(False, False)
        self.setup_ui()
        
        # 初始化文件类型映射
        self.engine_map = {
            '.xls': {'read': 'openpyxl', 'write': 'openpyxl'},
            '.xlsx': {'read': 'openpyxl', 'write': 'openpyxl'},
            '.xlsm': {'read': 'openpyxl', 'write': 'openpyxl'}
        }

    def setup_ui(self):
        # 创建主容器
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill='both', expand=True)

        # 文件类型输入
        ttk.Label(main_frame, text="文件类型：").grid(row=0, column=0, sticky='w', pady=2)
        self.file_type_entry = ttk.Combobox(main_frame, values=('.xls', '.xlsx'), width=10)
        self.file_type_entry.set('.xlsx')
        self.file_type_entry.grid(row=0, column=1, sticky='ew', padx=5)

        # 表头行数输入
        ttk.Label(main_frame, text="表头行数：").grid(row=1, column=0, sticky='w', pady=2)
        self.header_num_entry = ttk.Spinbox(main_frame, from_=1, to=10, width=5)
        self.header_num_entry.set(1)
        self.header_num_entry.grid(row=1, column=1, sticky='w', padx=5)

        # 输出文件名
        ttk.Label(main_frame, text="输出名称：").grid(row=2, column=0, sticky='w', pady=2)
        self.result_name_entry = ttk.Entry(main_frame)
        self.result_name_entry.insert(0, "合并结果")
        self.result_name_entry.grid(row=2, column=1, sticky='ew', padx=5)

        # 文件名标记选项
        self.add_filename_var = tk.BooleanVar(value=False)
        self.filename_check = ttk.Checkbutton(
            main_frame, 
            text="添加文件名列", 
            variable=self.add_filename_var,
            command=self.toggle_filename_column
        )
        self.filename_check.grid(row=3, column=0, columnspan=2, pady=5, sticky='w')

        # 操作按钮
        self.merge_btn = ttk.Button(
            main_frame, 
            text="选择目录并合并", 
            command=self.start_merge_process
        )
        self.merge_btn.grid(row=4, column=0, columnspan=2, pady=10)

        # 进度条
        self.progress = ttk.Progressbar(
            main_frame, 
            orient='horizontal', 
            mode='determinate'
        )
        self.progress.grid(row=5, column=0, columnspan=2, sticky='ew')

    def toggle_filename_column(self):
        """更新提示信息"""
        if self.add_filename_var.get():
            messagebox.showinfo(
                "提示", 
                "将在合并文件左侧添加【源文件名】列", 
                parent=self.root
            )

    def start_merge_process(self):
        """启动合并流程"""
        try:
            # 禁用按钮防止重复点击
            self.merge_btn.config(state='disabled')
            self.progress['value'] = 0
            self.root.update_idletasks()
            
            self.merge_excel_files()
            
        finally:
            self.merge_btn.config(state='normal')
            self.progress['value'] = 100

    def merge_excel_files(self):
        """核心合并逻辑"""
        try:
            # 获取输入参数
            file_type = self.validate_file_type()
            header_num = self.validate_header_num()
            add_filename = self.add_filename_var.get()
            
            # 选择目录
            directory = filedialog.askdirectory(title="请选择包含Excel文件的目录")
            if not directory:
                messagebox.showinfo("取消", "操作已取消：未选择目录", parent=self.root)
                return

            # 获取文件列表
            excel_files = self.get_excel_files(directory, file_type)
            total_files = len(excel_files)
            
            # 读取并合并数据
            dataframes = []
            for idx, file in enumerate(excel_files, 1):
                df = self.read_excel_file(
                    directory, file, 
                    file_type, header_num, 
                    add_filename
                )
                if df is not None:
                    dataframes.append(df)
                
                # 更新进度
                progress = (idx / total_files) * 90  # 保留10%给保存操作
                self.progress['value'] = progress
                self.root.update_idletasks()

            # 合并数据
            if not dataframes:
                raise ValueError("没有成功读取任何有效数据")
                
            merged_df = pd.concat(dataframes, ignore_index=True)
            
            # 保存结果
            output_path = self.save_result(merged_df, directory, file_type)
            messagebox.showinfo(
                "成功", 
                f"文件已成功保存至：\n{output_path}", 
                parent=self.root
            )

        except Exception as e:
            messagebox.showerror("错误", str(e), parent=self.root)

    def validate_file_type(self):
        """验证文件类型输入"""
        file_type = self.file_type_entry.get().strip()
        if not file_type:
            raise ValueError("文件类型不能为空")
        if not file_type.startswith('.'):
            file_type = '.' + file_type
        if file_type not in self.engine_map:
            raise ValueError(f"不支持的文件类型：{file_type}")
        return file_type

    def validate_header_num(self):
        """验证表头行数输入"""
        header_num = self.header_num_entry.get().strip()
        if not header_num.isdigit() or int(header_num) < 1:
            raise ValueError("表头行数必须是大于0的整数")
        return int(header_num)

    def get_excel_files(self, directory, file_type):
        """获取符合条件的Excel文件列表"""
        if not os.path.isdir(directory):
            raise FileNotFoundError(f"目录不存在：{directory}")
            
        excel_files = [
            f for f in os.listdir(directory) 
            if f.lower().endswith(file_type.lower())
        ]
        
        if not excel_files:
            raise ValueError(f"目录中没有找到{file_type}类型的文件")
        return excel_files

    def read_excel_file(self, directory, filename, file_type, header_num, add_filename):
        """读取单个Excel文件"""
        try:
            file_path = os.path.join(directory, filename)
            engine = self.engine_map[file_type]['read']
            
            df = pd.read_excel(
                file_path,
                header=header_num-1,
                engine=engine
            )
            
            # 添加文件名列
            if add_filename:
                df.insert(0, '源文件名', filename)
                
            return df
            
        except Exception as e:
            messagebox.showwarning(
                "读取警告", 
                f"文件 {filename} 读取失败：{str(e)}", 
                parent=self.root
            )
            return None

    def save_result(self, df, directory, file_type):
        """保存合并结果"""
        result_name = self.result_name_entry.get().strip() or "合并结果"
        base_name = os.path.splitext(result_name)[0]
        output_filename = f"{base_name}{file_type}"
        output_path = os.path.join(directory, output_filename)
        
        engine = self.engine_map[file_type]['write']
        df.to_excel(output_path, index=False, engine=engine)
        
        return output_path

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()
