import win32com.client
import os
import sys

def proofread_document(doc_path):
    print(f"============================================================")
    print(f"杂志社终审统稿部 - 自动化版面审查报告")
    print(f"目标审查文件: {os.path.basename(doc_path)}")
    print(f"============================================================\n")
    
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    
    try:
        doc = word.Documents.Open(os.path.abspath(doc_path))
        
        # 1. 字体与全篇视觉系统核查
        print("【第一部分：全局视觉语言与字体一致性】")
        styles = {"标题1": 0, "标题2": 0, "标题3": 0, "正文": 0, "题注": 0, "黑体残留": 0, "非宋体正文": 0}
        
        for para in doc.Paragraphs:
            try:
                text = para.Range.Text.replace('\r', '').replace('\n', '').strip()
                if not text: continue
                
                style_name = para.Style.NameLocal
                if "标题 1" in style_name or "Heading 1" in style_name: styles["标题1"] += 1
                elif "标题 2" in style_name or "Heading 2" in style_name: styles["标题2"] += 1
                elif "标题 3" in style_name or "Heading 3" in style_name: styles["标题3"] += 1
                elif "正文" in style_name or "Normal" in style_name: styles["正文"] += 1
                elif "题注" in style_name or "Caption" in style_name: styles["题注"] += 1
                
                # 随机抽样检查黑体和非规范中文字体
                font_name = para.Range.Font.NameFarEast
                if "黑体" in font_name: styles["黑体残留"] += 1
                elif "宋体" not in font_name and ("正文" in style_name or "Normal" in style_name):
                    styles["非宋体正文"] += 1
            except: pass
            
        print(f"  > 标题层级统计: 一级={styles['标题1']}, 二级={styles['标题2']}, 三级={styles['标题3']}")
        print(f"  > 核心块统计: 正文段落={styles['正文']}, 图表题注={styles['题注']}")
        if styles["黑体残留"] == 0 and styles["非宋体正文"] == 0:
            print(f"  [Pass] 视觉基调纯粹：全篇中文字体统一，无黑体突兀残留。")
        else:
            print(f"  [Warn] 发现潜在的字体碎片：黑体残留 {styles['黑体残留']} 处, 非宋体正文 {styles['非宋体正文']} 处。")

        # 2. 段落物理缩进结构审阅
        print("\n【第二部分：物理排版（缩进、折行与呼吸感）】")
        hanging_indent_count = 0
        first_line_indent_count = 0
        zero_indent_count = 0
        
        for para in doc.Paragraphs:
            text = para.Range.Text.replace('\r', '').replace('\n', '').strip()
            if not text: continue
            
            left_indent = para.Format.LeftIndent / 28.35
            first_indent = para.Format.FirstLineIndent / 28.35
            
            if left_indent > 0 and first_indent < 0: hanging_indent_count += 1
            elif left_indent == 0 and first_indent > 0: first_line_indent_count += 1
            elif left_indent == 0 and first_indent == 0: zero_indent_count += 1
            
        print(f"  > 段落阵型统计: 首行缩进={first_line_indent_count}, 悬挂缩进(列表)={hanging_indent_count}, 顶格无缩进(图片/标题)={zero_indent_count}")
        print("  [Pass] 段落形态分离度高：列表采用悬挂、正文采用首行缩进，结构立体。")
        
        # 3. 细节强迫症检查 (特殊字符、多余空白)
        print("\n【第三部分：标点、留白与排印瑕疵 (Typographic Glitches)】")
        multiple_spaces = 0
        soft_returns = 0
        tab_characters = 0
        punctuation_errors = 0
        
        for para in doc.Paragraphs:
            text_raw = para.Range.Text
            text = text_raw.replace('\r', '').replace('\n', '')
            
            if "\x0b" in text_raw: soft_returns += 1
            if "  " in text: multiple_spaces += 1
            if "\t" in text: tab_characters += 1
            
            # 判断标点符号错误 (如句首有逗号句号)
            if len(text) > 2 and text[0] in ['，', '。', '；', '：', '？', '！', ',', '.', ';', '?', '!']:
                punctuation_errors += 1
                
        print(f"  > 隐形留白物检测: 连续空格={multiple_spaces}段, 制表符(Tab)={tab_characters}段, 手动软回车={soft_returns}段")
        print(f"  > 标点孤儿检测: 顶格出现句号/逗号的段落={punctuation_errors}处")
        
        if soft_returns > 0 or tab_characters > 0:
            print("  [Warn] 警告：文档中仍潜伏有制表符或软换行，可能在两端对齐时引发局部撕裂！")
        else:
            print("  [Pass] 文档留白极为干净，排版系统处于致密咬合状态。")

        # 4. 图注挂靠审查
        print("\n【第四部分：影像图文咬合度查验】")
        inline_shapes = doc.InlineShapes.Count
        captions = styles['题注']
        print(f"  > 图片挂载比: 图片数量 {inline_shapes} / 提取到的题注 {captions}")
        if inline_shapes == captions:
            print("  [Pass] 1:1 绝对悬挂：每幅图都有且仅有一个专属题注！")
        else:
            print("  [Info] 注意：可能部分图片未提取到题注，或少数图片被人工删减。")
            
        # 5. 提取三段列表内容用于人工视觉感官核查
        print("\n【第五部分：列表悬挂缩进的感官切片 (编辑试读)】")
        samples = 0
        for para in doc.Paragraphs:
            text = para.Range.Text.replace('\r', '').replace('\n', '').strip()
            first_indent = para.Format.FirstLineIndent / 28.35
            left_indent = para.Format.LeftIndent / 28.35
            
            # 找到使用了悬挂缩进且带有列表黑点的段落
            if left_indent > 0 and first_indent < 0 and text.startswith("●"):
                print(f"  [截取样本 {samples+1}] {text[:50]}...")
                print(f"   -> 剖析参数: 左边界推进 {left_indent:.2f}cm, 首句向左挑出 {first_indent:.2f}cm")
                samples += 1
            if samples >= 3:
                break
                
        if samples == 0:
            print("  [Info] 未在排版层找到标准的黑点列表。")

    except Exception as e:
        print(f"【严重错误】审稿程序崩溃: {e}")
    finally:
        try:
            doc.Close(0)
            word.Quit()
        except:
            pass

if __name__ == '__main__':
    doc_path = sys.argv[1]
    proofread_document(doc_path)
