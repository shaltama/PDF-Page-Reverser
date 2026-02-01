import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import PyPDF2
import os
from pathlib import Path

class PDFPageReverser:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF页面颠倒工具")
        self.root.geometry("600x450")
        self.root.resizable(False, False)
        
        # 设置窗口图标（如果有的话）
        try:
            self.root.iconbitmap('pdf_icon.ico')
        except:
            pass
        
        # 初始化变量
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.status_text = tk.StringVar()
        self.status_text.set("准备就绪")
        
        self.setup_ui()
        
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 标题
        title_label = ttk.Label(main_frame, text="PDF页面颠倒工具", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 输入文件选择
        input_frame = ttk.LabelFrame(main_frame, text="选择输入PDF文件", padding="10")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        input_entry = ttk.Entry(input_frame, textvariable=self.input_file_path, 
                               width=50, state="readonly")
        input_entry.grid(row=0, column=0, padx=(0, 10))
        
        input_button = ttk.Button(input_frame, text="浏览...", 
                                 command=self.select_input_file)
        input_button.grid(row=0, column=1)
        
        # 输出文件选择
        output_frame = ttk.LabelFrame(main_frame, text="选择输出PDF文件", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        output_entry = ttk.Entry(output_frame, textvariable=self.output_file_path, 
                                width=50, state="readonly")
        output_entry.grid(row=0, column=0, padx=(0, 10))
        
        output_button = ttk.Button(output_frame, text="浏览...", 
                                  command=self.select_output_file)
        output_button.grid(row=0, column=1)
        
        # 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(20, 10))
        
        self.reverse_button = ttk.Button(button_frame, text="颠倒页面顺序", 
                                        command=self.reverse_pages, state="disabled")
        self.reverse_button.grid(row=0, column=0, padx=(0, 10))
        
        clear_button = ttk.Button(button_frame, text="清除", 
                                   command=self.clear_fields)
        clear_button.grid(row=0, column=1, padx=(0, 10))
        
        exit_button = ttk.Button(button_frame, text="退出", 
                                command=self.root.quit)
        exit_button.grid(row=0, column=2)
        
        # 状态栏
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        status_label = ttk.Label(status_frame, text="状态:")
        status_label.grid(row=0, column=0, sticky=tk.W)
        
        self.status_display = ttk.Label(status_frame, textvariable=self.status_text, 
                                       font=("Arial", 10, "italic"))
        self.status_display.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        # 说明文本
        info_frame = ttk.LabelFrame(main_frame, text="使用说明", padding="10")
        info_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        info_text = tk.Text(info_frame, height=4, width=60, wrap=tk.WORD, 
                             bg='#f0f0f0', relief='flat')
        info_text.grid(row=0, column=0)
        
        instructions = """1. 选择要处理的PDF文件
2. 选择输出文件的位置和名称
3. 点击"颠倒页面顺序"按钮
4. 程序会将PDF的所有页面完全颠倒（第一页变为最后一页）
5. 所有PDF信息都会完整保留"""
        
        info_text.insert(1.0, instructions)
        info_text.config(state='disabled')
        
        # 配置网格权重
        main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def select_input_file(self):
        filename = filedialog.askopenfilename(
            title="选择输入PDF文件",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.input_file_path.set(filename)
            self.update_output_path(filename)
            self.update_button_state()
            self.status_text.set(f"已选择输入文件: {os.path.basename(filename)}")
            
    def select_output_file(self):
        if not self.input_file_path.get():
            messagebox.showwarning("警告", "请先选择输入文件！")
            return
            
        input_path = Path(self.input_file_path.get())
        default_name = f"{input_path.stem}_reversed{input_path.suffix}"
        
        filename = filedialog.asksaveasfilename(
            title="选择输出PDF文件",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=default_name
        )
        if filename:
            self.output_file_path.set(filename)
            self.update_button_state()
            
    def update_output_path(self, input_path):
        """根据输入文件自动生成输出文件路径"""
        input_path_obj = Path(input_path)
        parent_dir = input_path_obj.parent
        file_name = f"{input_path_obj.stem}_reversed{input_path_obj.suffix}"
        output_path = parent_dir / file_name
        self.output_file_path.set(str(output_path))
        
    def update_button_state(self):
        """根据文件路径更新按钮状态"""
        if self.input_file_path.get() and self.output_file_path.get():
            self.reverse_button.config(state="normal")
        else:
            self.reverse_button.config(state="disabled")
            
    def clear_fields(self):
        """清除所有字段"""
        self.input_file_path.set("")
        self.output_file_path.set("")
        self.status_text.set("准备就绪")
        self.reverse_button.config(state="disabled")
        
    def reverse_pages(self):
        """执行页面颠倒操作"""
        input_path = self.input_file_path.get()
        output_path = self.output_file_path.get()
        
        if not input_path or not output_path:
            messagebox.showerror("错误", "请选择输入和输出文件！")
            return
            
        if input_path == output_path:
            messagebox.showerror("错误", "输入和输出文件不能相同！")
            return
            
        try:
            self.status_text.set("正在处理PDF文件...")
            self.root.update()
            
            # 读取PDF文件
            with open(input_path, 'rb') as input_file:
                pdf_reader = PyPDF2.PdfReader(input_file)
                pdf_writer = PyPDF2.PdfWriter()
                
                # 获取总页数
                total_pages = len(pdf_reader.pages)
                self.status_text.set(f"PDF共有 {total_pages} 页，正在颠倒顺序...")
                self.root.update()
                
                # 颠倒页面顺序
                for page_num in range(total_pages - 1, -1, -1):
                    page = pdf_reader.pages[page_num]
                    pdf_writer.add_page(page)
                    
                    # 更新进度（每10页更新一次）
                    if page_num % 10 == 0 or page_num == 0:
                        self.status_text.set(f"正在处理第 {total_pages - page_num} 页...")
                        self.root.update()
                
                # 写入输出文件
                with open(output_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
                    
            self.status_text.set("处理完成！")
            messagebox.showinfo("成功", f"PDF页面颠倒完成！\n输出文件: {os.path.basename(output_path)}")
            
        except FileNotFoundError:
            messagebox.showerror("错误", f"找不到输入文件: {input_path}")
            self.status_text.set("错误：文件未找到")
        except PermissionError:
            messagebox.showerror("错误", f"没有权限访问文件。请检查文件权限或关闭其他程序。")
            self.status_text.set("错误：权限不足")
        except PyPDF2.errors.PdfReadError:
            messagebox.showerror("错误", "无法读取PDF文件。文件可能已损坏或不是有效的PDF格式。")
            self.status_text.set("错误：PDF文件损坏")
        except Exception as e:
            messagebox.showerror("错误", f"处理PDF时发生错误:\n{str(e)}")
            self.status_text.set(f"错误: {str(e)}")

def main():
    root = tk.Tk()
    app = PDFPageReverser(root)
    root.mainloop()

if __name__ == "__main__":
    main()