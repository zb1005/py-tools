import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class ExcelSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel文件选择器")
        
        # 存储文件路径和列名
        self.file_paths = {"主表": "", "人员表": "", "假期表": "", "特殊节点表": ""}
        self.column_options = {"主表": [], "人员表": [], "假期表": [], "特殊节点表": []}
        self.selected_columns = {"主表": [], "人员表": [], "假期表": [], "特殊节点表": []}
        
        # 创建界面
        self.create_widgets()
    
    def create_widgets(self):
        # 文件选择部分
        tk.Label(self.root, text="请选择4个Excel文件:").grid(row=0, column=0, columnspan=3, pady=5)
        
        # 主表
        self.create_file_selector("主表", 1)
        # 人员表
        self.create_file_selector("人员表", 2)
        # 假期表
        self.create_file_selector("假期表", 3)
        # 特殊节点表
        self.create_file_selector("特殊节点表", 4)
        
        # 预览按钮
        tk.Button(self.root, text="预览列名", command=self.preview_columns).grid(row=5, column=0, pady=10)
        
        # 确认按钮
        tk.Button(self.root, text="确认选择", command=self.confirm_selection).grid(row=5, column=1, pady=10)
        
        # 退出按钮
        tk.Button(self.root, text="退出", command=self.root.quit).grid(row=5, column=2, pady=10)
    
    def create_file_selector(self, file_type, row):
        tk.Label(self.root, text=f"{file_type}:").grid(row=row, column=0, sticky="e")
        tk.Entry(self.root, width=50).grid(row=row, column=1)
        tk.Button(self.root, text="浏览...", 
                 command=lambda: self.select_file(file_type)).grid(row=row, column=2)
    
    def select_file(self, file_type):
        file_path = filedialog.askopenfilename(
            title=f"选择{file_type}文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        if file_path:
            self.file_paths[file_type] = file_path
            self.root.grid_slaves(row=self.get_row(file_type), column=1)[0].delete(0, tk.END)
            self.root.grid_slaves(row=self.get_row(file_type), column=1)[0].insert(0, file_path)
            
            # 读取列名
            try:
                df = pd.read_excel(file_path)
                self.column_options[file_type] = list(df.columns)
            except Exception as e:
                messagebox.showerror("错误", f"读取文件失败: {str(e)}")
    
    def get_row(self, file_type):
        mapping = {"主表": 1, "人员表": 2, "假期表": 3, "特殊节点表": 4}
        return mapping[file_type]
    
    def preview_columns(self):
        preview_window = tk.Toplevel(self.root)
        preview_window.title("列名预览")
        
        notebook = ttk.Notebook(preview_window)
        notebook.pack(fill="both", expand=True)
        
        for file_type in self.file_paths:
            if not self.file_paths[file_type]:
                continue
                
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=file_type)
            
            # 显示所有列名
            tk.Label(frame, text=f"{file_type}的所有列名:").pack(pady=5)
            
            listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, height=10)
            for col in self.column_options[file_type]:
                listbox.insert(tk.END, col)
            listbox.pack(pady=5)
            
            # 选择按钮
            tk.Button(frame, text="选择这些列", 
                     command=lambda lb=listbox, ft=file_type: self.select_columns(lb, ft)).pack(pady=5)
    
    def select_columns(self, listbox, file_type):
        selected = [listbox.get(i) for i in listbox.curselection()]
        self.selected_columns[file_type] = selected
        messagebox.showinfo("选择成功", f"已为{file_type}选择列: {', '.join(selected)}")
    
    def confirm_selection(self):
        # 检查是否所有文件都已选择
        for file_type, path in self.file_paths.items():
            if not path:
                messagebox.showerror("错误", f"请先选择{file_type}文件")
                return
        
        # 检查是否所有表都选择了列
        for file_type, cols in self.selected_columns.items():
            if not cols:
                messagebox.showerror("错误", f"请为{file_type}选择至少一列")
                return
        
        messagebox.showinfo("成功", "所有选择已完成，可以开始处理数据")
        
        # 这里可以添加数据处理逻辑
        print("文件路径:", self.file_paths)
        print("选择的列:", self.selected_columns)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSelectorApp(root)
    root.mainloop()