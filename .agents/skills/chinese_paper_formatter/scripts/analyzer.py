import win32com.client
import os
import sys

def analyze_document(doc_path):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    
    try:
        doc = word.Documents.Open(os.path.abspath(doc_path))
        
        print(f"=== 文档布局深入分析报告: {os.path.basename(doc_path)} ===\n")
        
        # 1. 页面设置分析
        print("--- 1. 页面边距与纸张 ---")
        section = doc.Sections(1)
        print(f"纸张宽度: {section.PageSetup.PageWidth / 28.35:.2f} cm")
        print(f"纸张高度: {section.PageSetup.PageHeight / 28.35:.2f} cm")
        print(f"上边距: {section.PageSetup.TopMargin / 28.35:.2f} cm")
        print(f"下边距: {section.PageSetup.BottomMargin / 28.35:.2f} cm")
        print(f"左边距: {section.PageSetup.LeftMargin / 28.35:.2f} cm")
        print(f"右边距: {section.PageSetup.RightMargin / 28.35:.2f} cm")
        
        # 2. 段落和字符级样式抽样分析
        print("\n--- 2. 各级段落排版抽样 (前 50 段) ---")
        abnormal_indent_count = 0
        soft_return_count = 0
        list_items_count = 0
        bold_normal_count = 0
        
        for i, para in enumerate(doc.Paragraphs):
            if i < 45:
                continue
            if i >= 80:
                break
                
            text = para.Range.Text.replace('\r', '').replace('\n', '').strip()
            if not text:
                continue
                
            style_name = para.Style.NameLocal
            outline_level = para.OutlineLevel
            
            # 统计特殊符号和格式
            if "\x0b" in para.Range.Text: # Soft Return (^l)
                soft_return_count += 1
                
            # 检测正文加粗
            if outline_level == 10 and para.Range.Font.Bold:
                bold_normal_count += 1
                
            # 检测列表
            if para.Range.ListFormat.ListType != 0:
                list_items_count += 1
                
            # 检测缩进异常 (大于常规首行缩进)
            left_indent = para.Format.LeftIndent / 28.35
            first_line_indent = para.Format.FirstLineIndent / 28.35
            if left_indent > 0.5:
                abnormal_indent_count += 1
                
            print(f"[段落 {i+1}] 样式: {style_name}, 大纲级: {outline_level}")
            print(f"         内容摘录: {text[:30]}...")
            print(f"         缩进(cm) -> 左: {left_indent:.2f}, 首行: {first_line_indent:.2f}")
            print(f"         行距类型: {para.Format.LineSpacingRule}, 值: {para.Format.LineSpacing}")
            
        print("\n--- 3. 宏观版面问题统计 ---")
        print(f"- 发现含有软回车 (^l) 的段落阻断版式: {soft_return_count}处 (前 50 段内)")
        print(f"- 发现异常左缩进的段落: {abnormal_indent_count}处 (前 50 段内)")
        print(f"- 发现携带自动列表格式的段落: {list_items_count}处 (前 50 段内)")
        print(f"- 发现正文(大纲10级)被违规加粗: {bold_normal_count}处 (前 50 段内)")
        
        # 4. 图片统计
        print("\n--- 4. 图片及周围元素 ---")
        shapes_count = doc.InlineShapes.Count
        print(f"总计图片数量: {shapes_count}")
        
    except Exception as e:
        print(f"分析出错: {e}")
    finally:
        try:
            doc.Close(0)
            word.Quit()
        except:
            pass

if __name__ == '__main__':
    doc_path = sys.argv[1]
    analyze_document(doc_path)
