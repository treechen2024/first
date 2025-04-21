import os
import re
import argparse
from pathlib import Path

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

def convert_to_pdf(md_file, output_dir=None):
    """
    使用Marp CLI將Markdown文件轉換為PDF
    
    需要先安裝Marp CLI: npm install -g @marp-team/marp-cli
    """
    if output_dir is None:
        output_dir = os.path.dirname(md_file)
    
    output_pdf = os.path.join(output_dir, os.path.splitext(os.path.basename(md_file))[0] + ".pdf")
    
    # 使用Marp CLI轉換
    os.system(f'marp {md_file} --pdf --output {output_pdf}')
    
    return output_pdf

def convert_to_pptx(md_file, output_dir=None):
    """
    使用Marp CLI將Markdown文件轉換為PPTX
    
    需要先安裝Marp CLI: npm install -g @marp-team/marp-cli
    """
    if output_dir is None:
        output_dir = os.path.dirname(md_file)
    
    output_pptx = os.path.join(output_dir, os.path.splitext(os.path.basename(md_file))[0] + ".pptx")
    
    # 使用Marp CLI轉換
    os.system(f'marp {md_file} --pptx --output {output_pptx}')
    
    return output_pptx

def main():
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

if __name__ == "__main__":
    main()