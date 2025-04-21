import os
import re
import argparse
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

def txt_to_marp(input_file, output_file=None, theme="default"):
    """
    將TXT文件轉換為Marp格式的Markdown文件
    
    Args:
        input_file: 輸入的TXT文件路徑
        output_file: 輸出的Markdown文件路徑，如果為None則自動生成
        theme: 使用的主題名稱
    """
    if output_file is None:
        output_file = os.path.splitext(input_file)[0] + ".md"
    
    # 讀取TXT文件
    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 分割內容為幻燈片
    # 使用兩個空白行作為幻燈片分隔
    slides = [slide.strip() for slide in re.split(r'\n[ \t]*\n[ \t]*\n', content)]
    
    # 創建Marp頭部
    marp_header = f"""---
marp: true
theme: {theme}
paginate: true
backgroundColor: #fff
---

"""
    
    # 創建CSS樣式
    css_styles = """
<style>
/* 全局樣式 */
section {
    font-family: 'Arial', sans-serif;
    padding: 40px;
}

/* 標題樣式 */
h1 {
    color: #2c3e50;
    font-size: 2.5em;
    margin-bottom: 0.5em;
}

h2 {
    color: #3498db;
    font-size: 2em;
    margin-bottom: 0.5em;
}

/* 列表樣式 */
ul, ol {
    margin-left: 1.5em;
    line-height: 1.6;
}

li {
    margin-bottom: 0.5em;
}

/* 強調文本 */
strong {
    color: #e74c3c;
}

em {
    color: #27ae60;
}

/* 代碼塊 */
code {
    background-color: #f8f8f8;
    border-radius: 3px;
    padding: 0.2em 0.4em;
    font-family: 'Courier New', monospace;
}

/* 引用塊 */
blockquote {
    border-left: 5px solid #3498db;
    padding-left: 1em;
    color: #7f8c8d;
    font-style: italic;
}

/* 表格樣式 */
table {
    border-collapse: collapse;
    width: 100%;
    margin: 1em 0;
}

th, td {
    border: 1px solid #ddd;
    padding: 8px 12px;
    text-align: left;
}

th {
    background-color: #f2f2f2;
    font-weight: bold;
}

/* 圖片樣式 */
img {
    max-width: 100%;
    height: auto;
    display: block;
    margin: 0 auto;
}

/* 頁腳樣式 */
footer {
    position: absolute;
    bottom: 20px;
    right: 20px;
    font-size: 0.8em;
    color: #95a5a6;
}
</style>

"""
    
    # 將幻燈片轉換為Marp格式
    marp_content = marp_header + css_styles
    
    for i, slide in enumerate(slides):
        slide = slide.strip()
        if not slide:
            continue
            
        # 處理幻燈片內容
        lines = slide.split('\n')
        
        # 假設第一行是標題
        if lines and lines[0]:
            title = lines[0]
            content_lines = lines[1:]
            
            # 將標題轉換為Markdown標題
            if not title.startswith('#'):
                title = f"# {title}"
                
            # 處理內容
            # 保留單個空白行作為換行
            content = "\n\n".join(line.strip() for line in content_lines if line.strip() or line.isspace())
            
            # 處理<br>換行標記
            content = content.replace('<br>', '\n')
            
            # 自動檢測列表項
            content = re.sub(r'^(\d+)\.\s', r'\1. ', content, flags=re.MULTILINE)  # 有序列表
            content = re.sub(r'^[-*]\s', '- ', content, flags=re.MULTILINE)  # 無序列表
            
            # 添加幻燈片分隔符
            if i > 0:
                marp_content += "\n---\n\n"
                
            marp_content += f"{title}\n\n{content}\n"
    
    # 寫入輸出文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(marp_content)
        
    return output_file

def check_marp_cli():
    """檢查Marp CLI是否已安裝"""
    import subprocess
    try:
        result = subprocess.run(['marp', '--version'], capture_output=True, text=True)
        return result.returncode == 0
    except FileNotFoundError:
        return False

def convert_to_pdf(md_file, output_dir=None):
    """
    使用Marp CLI將Markdown文件轉換為PDF
    
    需要先安裝Marp CLI: npm install -g @marp-team/marp-cli
    """
    if not check_marp_cli():
        raise RuntimeError("Marp CLI未安裝，請先執行：npm install -g @marp-team/marp-cli")
        
    if output_dir is None:
        output_dir = os.path.dirname(md_file)
    
    output_pdf = os.path.join(output_dir, os.path.splitext(os.path.basename(md_file))[0] + ".pdf")
    
    # 使用Marp CLI轉換
    result = subprocess.run(['marp', md_file, '--pdf', '--output', output_pdf], capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"PDF轉換失敗：{result.stderr}")
    
    return output_pdf

def convert_to_pptx(md_file, output_dir=None):
    """
    使用Marp CLI將Markdown文件轉換為PPTX
    
    需要先安裝Marp CLI: npm install -g @marp-team/marp-cli
    """
    if not check_marp_cli():
        raise RuntimeError("Marp CLI未安裝，請先執行：npm install -g @marp-team/marp-cli")
        
    if output_dir is None:
        output_dir = os.path.dirname(md_file)
    
    output_pptx = os.path.join(output_dir, os.path.splitext(os.path.basename(md_file))[0] + ".pptx")
    
    # 使用Marp CLI轉換
    result = subprocess.run(['marp', md_file, '--pptx', '--output', output_pptx], capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"PPTX轉換失敗：{result.stderr}")
    
    return output_pptx

class MarpConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Marp 簡報轉換器")
        self.root.geometry("600x400")
        
        # 創建主框架
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 輸入文件選擇
        ttk.Label(main_frame, text="輸入文件:").grid(row=0, column=0, sticky=tk.W)
        self.input_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.input_path, width=50).grid(row=0, column=1, padx=10)
        ttk.Button(main_frame, text="瀏覽", command=self.browse_input).grid(row=0, column=2, padx=10)
        
        # 輸出目錄選擇
        ttk.Label(main_frame, text="輸出目錄:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.output_path = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=1, column=1, padx=10)
        ttk.Button(main_frame, text="瀏覽", command=self.browse_output).grid(row=1, column=2, padx=10)
        
        # 主題選擇
        ttk.Label(main_frame, text="主題:").grid(row=2, column=0, sticky=tk.W)
        self.theme = tk.StringVar(value="default")
        themes = ["default", "gaia", "uncover"]
        ttk.Combobox(main_frame, textvariable=self.theme, values=themes).grid(row=2, column=1, sticky=tk.W)
        
        # 輸出格式選擇
        format_frame = ttk.LabelFrame(main_frame, text="輸出格式", padding="5")
        format_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky=tk.W)
        
        self.md_var = tk.BooleanVar(value=True)
        self.pdf_var = tk.BooleanVar(value=False)
        self.pptx_var = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(format_frame, text="Markdown", variable=self.md_var, state="disabled").pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(format_frame, text="PDF", variable=self.pdf_var).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(format_frame, text="PPTX", variable=self.pptx_var).pack(side=tk.LEFT, padx=5)
        
        # 轉換按鈕
        ttk.Button(main_frame, text="開始轉換", command=self.convert).grid(row=4, column=0, columnspan=3, pady=20)
        
        # 狀態標籤
        self.status_var = tk.StringVar()
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=5, column=0, columnspan=3)

    def browse_input(self):
        filename = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if filename:
            self.input_path.set(filename)
            # 自動設置輸出目錄
            self.output_path.set(os.path.dirname(filename))

    def browse_output(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_path.set(directory)

    def convert(self):
        input_file = self.input_path.get()
        output_dir = self.output_path.get()
        
        if not input_file or not output_dir:
            messagebox.showerror("錯誤", "請選擇輸入文件和輸出目錄")
            return
            
        try:
            # 生成輸出文件路徑
            base_name = os.path.splitext(os.path.basename(input_file))[0]
            output_md = os.path.join(output_dir, base_name + ".md")
            
            # 轉換為Markdown
            md_file = txt_to_marp(input_file, output_md, self.theme.get())
            self.status_var.set(f"已生成Markdown文件: {md_file}")
            
            # 轉換為PDF
            if self.pdf_var.get():
                pdf_file = convert_to_pdf(md_file, output_dir)
                self.status_var.set(self.status_var.get() + f"\n已生成PDF文件: {pdf_file}")
            
            # 轉換為PPTX
            if self.pptx_var.get():
                pptx_file = convert_to_pptx(md_file, output_dir)
                self.status_var.set(self.status_var.get() + f"\n已生成PPTX文件: {pptx_file}")
                
            messagebox.showinfo("成功", "轉換完成！")
            
        except Exception as e:
            messagebox.showerror("錯誤", f"轉換過程中發生錯誤：{str(e)}")

def main():
    # 檢查是否有命令行參數
    if len(sys.argv) > 1:
        # 使用原有的命令行介面
        parser = argparse.ArgumentParser(description='將TXT文件轉換為Marp格式的簡報')
        parser.add_argument('input_file', help='輸入的TXT文件路徑')
        parser.add_argument('-o', '--output', help='輸出的Markdown文件路徑')
        parser.add_argument('-t', '--theme', default='default', help='使用的主題名稱')
        parser.add_argument('--pdf', action='store_true', help='同時生成PDF文件')
        parser.add_argument('--pptx', action='store_true', help='同時生成PPTX文件')
        
        args = parser.parse_args()
        
        # 轉換為Marp格式
        md_file = txt_to_marp(args.input_file, args.output, args.theme)
        print(f"已生成Marp格式文件: {md_file}")
        
        # 轉換為PDF
        if args.pdf:
            pdf_file = convert_to_pdf(md_file)
            print(f"已生成PDF文件: {pdf_file}")
        
        # 轉換為PPTX
        if args.pptx:
            pptx_file = convert_to_pptx(md_file)
            print(f"已生成PPTX文件: {pptx_file}")
    else:
        # 啟動GUI介面
        root = tk.Tk()
        app = MarpConverterGUI(root)
        root.mainloop()

if __name__ == "__main__":
    import sys
    main()