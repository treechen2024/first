import os
import re
import argparse
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
# 移除 subprocess 導入，因為不再需要

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
    # 假設每個幻燈片由至少兩個連續的換行符分隔
    slides = re.split(r'\n{2,}', content)
    
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
            content = "\n".join(content_lines)
            
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

def convert_to_pdf(md_file, output_dir=None, node_path=None):
    """
    使用Marp CLI將Markdown文件轉換為PDF
    
    需要先安裝Marp CLI: npm install -g @marp-team/marp-cli
    或使用npx直接運行
    
    Args:
        md_file: Markdown文件路徑
        output_dir: 輸出目錄，如果為None則使用Markdown文件所在目錄
        node_path: Node.js可執行文件路徑，如果為None則使用系統PATH中的node
    """
    if output_dir is None:
        output_dir = os.path.dirname(md_file)
    
    output_pdf = os.path.join(output_dir, os.path.splitext(os.path.basename(md_file))[0] + ".pdf")
    
    # 構建命令
    if node_path and os.path.exists(node_path):
        npx_path = os.path.join(os.path.dirname(node_path), "npx")
        cmd = f'"{npx_path}" @marp-team/marp-cli {md_file} --pdf --output {output_pdf}'
    else:
        cmd = f'npx @marp-team/marp-cli {md_file} --pdf --output {output_pdf}'
    
    # 使用npx運行Marp CLI轉換
    result = os.system(cmd)
    
    # 檢查命令是否成功執行
    if result != 0:
        raise Exception("轉換PDF失敗，請確保已安裝Node.js和npm")
    
    return output_pdf

def convert_to_pptx(md_file, output_dir=None, node_path=None):
    """
    使用Marp CLI將Markdown文件轉換為PPTX
    
    需要先安裝Marp CLI: npm install -g @marp-team/marp-cli
    或使用npx直接運行
    
    Args:
        md_file: Markdown文件路徑
        output_dir: 輸出目錄，如果為None則使用Markdown文件所在目錄
        node_path: Node.js可執行文件路徑，如果為None則使用系統PATH中的node
    """
    if output_dir is None:
        output_dir = os.path.dirname(md_file)
    
    output_pptx = os.path.join(output_dir, os.path.splitext(os.path.basename(md_file))[0] + ".pptx")
    
    # 構建命令
    if node_path and os.path.exists(node_path):
        npx_path = os.path.join(os.path.dirname(node_path), "npx")
        cmd = f'"{npx_path}" @marp-team/marp-cli {md_file} --pptx --output {output_pptx}'
    else:
        cmd = f'npx @marp-team/marp-cli {md_file} --pptx --output {output_pptx}'
    
    # 使用npx運行Marp CLI轉換
    result = os.system(cmd)
    
    # 檢查命令是否成功執行
    if result != 0:
        raise Exception("轉換PPTX失敗，請確保已安裝Node.js和npm")
    
    return output_pptx

def create_gui():
    """創建GUI介面"""
    root = tk.Tk()
    root.title("TXT 轉 Marp 簡報轉換器")
    root.geometry("600x450")  # 增加高度以容納新的設置
    root.resizable(True, True)
    
    # 設置樣式
    style = ttk.Style()
    style.configure('TButton', font=('Arial', 10))
    style.configure('TLabel', font=('Arial', 10))
    style.configure('TCheckbutton', font=('Arial', 10))
    
    # 創建主框架
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 輸入文件選擇
    input_frame = ttk.Frame(main_frame)
    input_frame.pack(fill=tk.X, pady=5)
    
    ttk.Label(input_frame, text="輸入文件:").pack(side=tk.LEFT)
    input_var = tk.StringVar()
    input_entry = ttk.Entry(input_frame, textvariable=input_var, width=50)
    input_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    
    def browse_input():
        filename = filedialog.askopenfilename(
            title="選擇TXT文件",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            input_var.set(filename)
            # 自動設置輸出文件名
            if not output_var.get():
                output_path = os.path.splitext(filename)[0] + ".md"
                output_var.set(output_path)
    
    ttk.Button(input_frame, text="瀏覽...", command=browse_input).pack(side=tk.RIGHT)
    
    # 輸出文件選擇
    output_frame = ttk.Frame(main_frame)
    output_frame.pack(fill=tk.X, pady=5)
    
    ttk.Label(output_frame, text="輸出文件:").pack(side=tk.LEFT)
    output_var = tk.StringVar()
    output_entry = ttk.Entry(output_frame, textvariable=output_var, width=50)
    output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    
    def browse_output():
        filename = filedialog.asksaveasfilename(
            title="保存Markdown文件",
            filetypes=[("Markdown files", "*.md"), ("All files", "*.*")],
            defaultextension=".md"
        )
        if filename:
            output_var.set(filename)
    
    ttk.Button(output_frame, text="瀏覽...", command=browse_output).pack(side=tk.RIGHT)
    
    # 主題選擇
    theme_frame = ttk.Frame(main_frame)
    theme_frame.pack(fill=tk.X, pady=5)
    
    ttk.Label(theme_frame, text="主題:").pack(side=tk.LEFT)
    theme_var = tk.StringVar(value="default")
    theme_combo = ttk.Combobox(theme_frame, textvariable=theme_var, 
                              values=["default", "gaia", "uncover"])
    theme_combo.pack(side=tk.LEFT, padx=5)
    
    # 輸出格式選項
    format_frame = ttk.Frame(main_frame)
    format_frame.pack(fill=tk.X, pady=10)
    
    ttk.Label(format_frame, text="輸出格式:").pack(side=tk.LEFT, padx=(0, 10))
    
    md_var = tk.BooleanVar(value=True)
    md_check = ttk.Checkbutton(format_frame, text="Markdown", variable=md_var, state="disabled")
    md_check.pack(side=tk.LEFT, padx=5)
    
    pdf_var = tk.BooleanVar(value=False)
    pdf_check = ttk.Checkbutton(format_frame, text="PDF", variable=pdf_var)
    pdf_check.pack(side=tk.LEFT, padx=5)
    
    pptx_var = tk.BooleanVar(value=False)
    pptx_check = ttk.Checkbutton(format_frame, text="PPTX", variable=pptx_var)
    pptx_check.pack(side=tk.LEFT, padx=5)
    
    # 轉換按鈕
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=20)
    
    def convert():
        input_file = input_var.get()
        output_file = output_var.get() if output_var.get() else None
        theme = theme_var.get()
        node_path = node_path_var.get() if node_path_var.get() else None
        
        if not input_file:
            messagebox.showerror("錯誤", "請選擇輸入文件")
            return
        
        try:
            # 轉換為Marp格式
            status_var.set("正在生成Markdown文件...")
            root.update()
            md_file = txt_to_marp(input_file, output_file, theme)
            messagebox.showinfo("成功", f"已生成Marp格式文件: {md_file}")
            
            # 轉換為PDF
            if pdf_var.get():
                status_var.set("正在生成PDF文件...")
                root.update()
                try:
                    pdf_file = convert_to_pdf(md_file, node_path=node_path)
                    messagebox.showinfo("成功", f"已生成PDF文件: {pdf_file}")
                except Exception as e:
                    messagebox.showerror("錯誤", f"PDF轉換失敗: {str(e)}")
            
            # 轉換為PPTX
            if pptx_var.get():
                status_var.set("正在生成PPTX文件...")
                root.update()
                try:
                    pptx_file = convert_to_pptx(md_file, node_path=node_path)
                    messagebox.showinfo("成功", f"已生成PPTX文件: {pptx_file}")
                except Exception as e:
                    messagebox.showerror("錯誤", f"PPTX轉換失敗: {str(e)}")
                    
            status_var.set("就绪")
                
        except Exception as e:
            messagebox.showerror("錯誤", f"轉換過程中發生錯誤: {str(e)}")
            status_var.set("轉換失敗")
    
    ttk.Button(button_frame, text="轉換", command=convert, width=20).pack(side=tk.RIGHT)
    
    # 狀態欄
    status_var = tk.StringVar(value="就绪")
    status_bar = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor=tk.W)
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    return root

def main():
    parser = argparse.ArgumentParser(description='將TXT文件轉換為Marp格式的簡報')
    parser.add_argument('input_file', nargs='?', help='輸入的TXT文件路徑')
    parser.add_argument('-o', '--output', help='輸出的Markdown文件路徑')
    parser.add_argument('-t', '--theme', default='default', help='使用的主題名稱')
    parser.add_argument('--pdf', action='store_true', help='同時生成PDF文件')
    parser.add_argument('--pptx', action='store_true', help='同時生成PPTX文件')
    parser.add_argument('--gui', action='store_true', help='啟動GUI介面')
    parser.add_argument('--node-path', help='指定Node.js可執行文件路徑')
    
    args = parser.parse_args()
    
    # 如果指定了--gui參數或沒有提供輸入文件，則啟動GUI
    if args.gui or not args.input_file:
        root = create_gui()
        root.mainloop()
    else:
        # 命令行模式
        md_file = txt_to_marp(args.input_file, args.output, args.theme)
        print(f"已生成Marp格式文件: {md_file}")
        
        # 轉換為PDF
        if args.pdf:
            pdf_file = convert_to_pdf(md_file, node_path=args.node_path)
            print(f"已生成PDF文件: {pdf_file}")
        
        # 轉換為PPTX
        if args.pptx:
            pptx_file = convert_to_pptx(md_file, node_path=args.node_path)
            print(f"已生成PPTX文件: {pptx_file}")

if __name__ == "__main__":
    main()