import win32com.client
import os
import sys
import traceback

def cm_to_points(cm):
    """厘米转磅 (1 cm = 28.35 points)"""
    return cm * 28.35

def format_word_document(input_path, output_path=None):
    if not os.path.exists(input_path):
        print(f"Error: 找不到输入文件 - {input_path}")
        return False
        
    input_abspath = os.path.abspath(input_path)
    # 如果未指定输出路径，覆盖原文件
    output_abspath = os.path.abspath(output_path) if output_path else input_abspath

    word = None
    doc = None
    
    try:
        # 启动 Word 应用程序的 COM 接口
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False # 在后台执行，不显示界面
        word.DisplayAlerts = 0 
    except Exception as e:
        print(f"无法启动 Word。请确认本机是否安装了 Microsoft Word 以及 pywin32。\n详细错误: {e}")
        return False

    try:
        print(f"正在打开文档: {input_abspath}")
        doc = word.Documents.Open(input_abspath)
        
        print("0.1 将所有自动编号和项目符号转换为纯文本以统一格式并根除缩排错乱...")
        try:
            doc.ConvertNumbersToText()
        except Exception as e:
            print(f"转换编号格式时出错: {e}")
            
        print("0. 自动符号清理 (移除 Markdown 残留符号及连续空行)...")
        
        # 定义需要精确清理的纯字符串符号
        exact_symbols = ["**", "---", "--"]
        
        for symbol in exact_symbols:
            find_obj = doc.Content.Find
            find_obj.ClearFormatting()
            find_obj.Replacement.ClearFormatting()
            find_obj.Execute(FindText=symbol, ReplaceWith="", Replace=2, Forward=True, Wrap=1, MatchWildcards=False, MatchCase=False)
            
        print("执行通配符正则清理 (清除带空格和各种空白符的符号如 '  - **' 和 '  ###')...")
        wildcard_patterns = [
            # Word 专有通配符: ^w 代表空白区（连续空格/制表符）
            r"^w- \*\*",
            r"^w-\*\*",
            r"^w-",
            r"^w#",
            r"^w##",
            r"^w###",
            r"^w####",
            r"\*\*", # 处理可能夹在中间的星号
            r"##",   # 处理残留的井号
            r"---",  # 清除无意义手工横线
            r"___"   # 清除无意义手工下划线
        ]
        
        for pattern in wildcard_patterns:
            find_obj = doc.Content.Find
            find_obj.ClearFormatting()
            find_obj.Replacement.ClearFormatting()
            find_obj.Execute(FindText=pattern, ReplaceWith="", Replace=2, Forward=True, Wrap=1, MatchWildcards=True)
            
        print("0.5 统一项目符号与编号的间隙 (将各类缩进原点强制收缩至统一距离)...")
        # --- 1. 使用通配符修正带空格/制表符的符号 ---
        bullet_spacing_fixes = [
            (r"([0-9]{1,}\.)^w", r"\1 "),      # 缩去数字编号的巨大间隙 1. \t -> 1. 
            (r"([A-Za-z]{1,}\.)^w", r"\1 "),  # 字母编号
            (r"○^w", r"● "),                # 统一空心圆点为实心并收紧空白
            (r"○", r"●"),                   # 兜底转换圆点
            (r"●^w", r"● "),                # 收紧实心圆点后的夸张缩进间隙
            (r"·^w", r"● "),                # 转换各种奇怪的标点符号为标准圆点
            (r"·", r"●")
        ]
        
        for find_text, replace_text in bullet_spacing_fixes:
            find_obj = doc.Content.Find
            find_obj.ClearFormatting()
            find_obj.Replacement.ClearFormatting()
            find_obj.Execute(FindText=find_text, ReplaceWith=replace_text, Replace=2, Forward=True, Wrap=1, MatchWildcards=True)
            
        # --- 2. 使用普通替换修正特定的 ASCII 字符和不可见系统乱码 ---
        # 避免受诸如 MatchWholeWord 或字体等潜在缓存状态的干扰，使用原生的段落遍历替换
        print("0.5.1 手动遍历查杀特殊或隐形乱码列表字符...")
        for p in doc.Paragraphs:
            try:
                t = p.Range.Text
                if t.startswith(chr(0xf0b7) + "\t") or t.startswith(chr(0xf0b7) + " "):
                    p.Range.Characters(1).Text = ""
                    p.Range.Characters(1).Text = "● "
                elif t.startswith(chr(0xf0b7)):
                    p.Range.Characters(1).Text = "●"
                elif t.startswith("o\t") or t.startswith("o "):
                    p.Range.Characters(1).Text = ""
                    p.Range.Characters(1).Text = "● "
            except Exception:
                pass
        
        print("0.6 全局正文去加粗 (确保正文没有加粗格式)...")
        # 强制清理正文格式中的加粗 (wdStyleNormal = -1)
        find_obj = doc.Content.Find
        find_obj.ClearFormatting()
        find_obj.Style = doc.Styles(-1)
        find_obj.Font.Bold = True
        find_obj.Replacement.ClearFormatting()
        find_obj.Replacement.Font.Bold = False
        find_obj.Execute(Replace=2, Forward=True, Wrap=1, MatchWildcards=False)
        
        print("清理手动换行符(Soft Return)和连续空行...")
        # 将所有的软回车 (Shift+Enter, ^l) 转换为标准的段落标记 (^p)，因为软回车会破坏段落级别的首行缩进等格式
        find_obj = doc.Content.Find
        find_obj.ClearFormatting()
        find_obj.Replacement.ClearFormatting()
        find_obj.Execute(FindText="^l", ReplaceWith="^p", Replace=2, Forward=True, Wrap=1, MatchWildcards=False)
        
        print(f"空行清理 (循环清理 ^p^p)")
        find_obj = doc.Content.Find
        find_obj.ClearFormatting()
        find_obj.Replacement.ClearFormatting()
        while find_obj.Execute(FindText="^p^p", ReplaceWith="^p", Replace=2, Forward=True, Wrap=1, MatchWildcards=False):
            pass
            
        print("0.5 清理强制换页符以防止出现空白页...")
        find_obj = doc.Content.Find
        find_obj.ClearFormatting()
        find_obj.Replacement.ClearFormatting()
        find_obj.Execute(
            FindText="^m",  # 手动换页符 manual page break
            ReplaceWith="",
            Replace=2,
            Forward=True,
            Wrap=1,
            MatchWildcards=False
        )

        # 1. 页面设置 (A4, 上3.0, 下2.5, 左2.6, 右2.6)
        print("1. 设置页面边距 (国标)...")
        for section in doc.Sections:
            section.PageSetup.TopMargin = cm_to_points(3.0)
            section.PageSetup.BottomMargin = cm_to_points(2.5)
            section.PageSetup.LeftMargin = cm_to_points(2.6)
            section.PageSetup.RightMargin = cm_to_points(2.6)
            section.PageSetup.PageWidth = cm_to_points(21.0)
            section.PageSetup.PageHeight = cm_to_points(29.7)
            
        # 开启孤行控制 (全局开启，避免逐段落遍历导致 COM 通信极慢)
        doc.Content.ParagraphFormat.WidowControl = True

        # 2. 标题和正文样式调整
        print("2. 调整正文和各级标题段落格式 (字体、缩进、行距)...")
        try:
            # 正文样式 (wdStyleNormal = -1)
            # 小四，宋体，英文字体 Times New Roman，行距固定值20磅，首行缩进2字符
            normal_style = doc.Styles(-1)
            normal_style.Font.NameFarEast = "宋体"
            normal_style.Font.NameAscii = "Times New Roman"
            normal_style.Font.NameOther = "Times New Roman"
            normal_style.Font.Size = 12 # 小四

            normal_style.ParagraphFormat.LineSpacingRule = 1 # 1.5 倍行距
            normal_style.ParagraphFormat.CharacterUnitFirstLineIndent = 2 # 首行缩进2字符
            normal_style.ParagraphFormat.SpaceBefore = 0
            normal_style.ParagraphFormat.SpaceAfter = 0
            normal_style.ParagraphFormat.Alignment = 3 # 两端对齐
            
            # 一级标题样式 (wdStyleHeading1 = -2)
            # 三号，宋体 (原要求为黑体，现统一所有字不得为黑体)，居中，段前0.5行 (12磅)，段后0.5行 (6磅)
            h1_style = doc.Styles(-2)
            h1_style.Font.NameFarEast = "宋体"
            h1_style.Font.NameAscii = "Times New Roman"
            h1_style.Font.Size = 16 # 三号
            h1_style.ParagraphFormat.Alignment = 1 # 1=居中
            h1_style.ParagraphFormat.SpaceBefore = 12
            h1_style.ParagraphFormat.SpaceAfter = 6
            h1_style.ParagraphFormat.CharacterUnitFirstLineIndent = 0
            h1_style.ParagraphFormat.LineSpacingRule = 4
            h1_style.ParagraphFormat.LineSpacing = 20
            
            # 二级标题样式 (wdStyleHeading2 = -3)
            # 四号，宋体，靠左，段前0.5行，段后0.5行
            h2_style = doc.Styles(-3)
            h2_style.Font.NameFarEast = "宋体"
            h2_style.Font.NameAscii = "Times New Roman"
            h2_style.Font.Size = 14 # 四号
            h2_style.ParagraphFormat.Alignment = 0 # 0=靠左
            h2_style.ParagraphFormat.SpaceBefore = 6
            h2_style.ParagraphFormat.SpaceAfter = 6
            h2_style.ParagraphFormat.CharacterUnitFirstLineIndent = 0
            h2_style.ParagraphFormat.LineSpacingRule = 4
            h2_style.ParagraphFormat.LineSpacing = 20
            
            # 三级标题样式 (wdStyleHeading3 = -4)
            # 小四，宋体，靠左
            h3_style = doc.Styles(-4)
            h3_style.Font.NameFarEast = "宋体"
            h3_style.Font.NameAscii = "Times New Roman"
            h3_style.Font.Size = 12 # 小四
            h3_style.ParagraphFormat.Alignment = 0
            h3_style.ParagraphFormat.SpaceBefore = 6
            h3_style.ParagraphFormat.SpaceAfter = 6
            h3_style.ParagraphFormat.CharacterUnitFirstLineIndent = 0
            h3_style.ParagraphFormat.LineSpacingRule = 4
            h3_style.ParagraphFormat.LineSpacing = 20

            # 图表题注样式 (wdStyleCaption = -35)
            # 小五号 (9磅)，居中
            caption_style = doc.Styles(-35)
            caption_style.Font.NameFarEast = "宋体"
            caption_style.Font.NameAscii = "Times New Roman"
            caption_style.Font.Size = 9
            caption_style.ParagraphFormat.Alignment = 1
            caption_style.ParagraphFormat.CharacterUnitFirstLineIndent = 0
            
        except Exception as e:
            print(f"设置样式时遇到部分错误 (可能原文档缺少此样式): {e}")

        print("2.5 强制清理所有段落的错误缩进 (标题 + 正文 + 列表)...")
        try:
            # 批量清理正文格式 (使用原生 Find.Execute 极速处理)
            find_obj = doc.Content.Find
            find_obj.ClearFormatting()
            find_obj.Replacement.ClearFormatting()
            
            # 正文 (-1), 列表段落 (-51), 正文缩进 (-29), 正文文本 (-67)
            body_styles = [-1, -51, -29, -67]
            for style_id in body_styles:
                try:
                    find_obj.Format = True
                    find_obj.Style = doc.Styles(style_id)
                    # 1. 清理杂乱的段落边框线 (Borders, 彻底消除文字背后的横线)
                    find_obj.Replacement.ParagraphFormat.Borders(1).LineStyle = 0 # wdLineStyleNone
                    find_obj.Replacement.ParagraphFormat.Borders(2).LineStyle = 0
                    find_obj.Replacement.ParagraphFormat.Borders(3).LineStyle = 0
                    find_obj.Replacement.ParagraphFormat.Borders(4).LineStyle = 0
                    
                    # 2. 通用正文格式设定 (剥夺所有私有格式以强制应用1.5倍行距和字体)
                    find_obj.Replacement.ParagraphFormat.LeftIndent = 0
                    find_obj.Replacement.ParagraphFormat.RightIndent = 0
                    find_obj.Replacement.ParagraphFormat.CharacterUnitFirstLineIndent = 2
                    find_obj.Replacement.ParagraphFormat.Alignment = 3 # 两端对齐
                    find_obj.Replacement.ParagraphFormat.LineSpacingRule = 1 # 1 = wdLineSpace1pt5 (1.5倍行距)
                    find_obj.Replacement.Font.NameFarEast = "宋体"
                    find_obj.Replacement.Font.NameAscii = "Times New Roman"
                    find_obj.Replacement.Font.Size = 12
                    
                    find_obj.Execute(FindText="", ReplaceWith="", Replace=2) # ReplaceAll
                except Exception:
                    pass
            
            # 暴力穿透旧版 .doc 文档自带的覆盖式“段前属性”锁定机制
            print("2.5.1 遍历普通段落执行高压格式覆盖 (应对顽固的直接格式)...")
            body_style_names = [doc.Styles(sid).NameLocal for sid in body_styles]
            for p in doc.Paragraphs:
                try:
                    sn = p.Style.NameLocal
                    if sn in body_style_names or "正文" in sn:
                        p.Range.Font.NameFarEast = "宋体"
                        p.Range.Font.NameAscii = "Times New Roman"
                        p.Range.Font.Size = 12
                        p.Format.LineSpacingRule = 1
                except Exception:
                    pass
            
            # 2.5 特殊情况：智能分离列表段的缩进模式 (悬挂缩进)
            print("2.6 剥离列表与正文：为带有子弹头(●)或序号(1.)的列表施加悬挂缩进...")
            for p in doc.Paragraphs:
                try:
                    t = p.Range.Text.lstrip()
                    # 识别：实心圆点，数字编号 (如 1. )，或单字母编号 (如 A. )
                    is_list = False
                    if t.startswith("● "):
                        is_list = True
                    elif len(t) > 2 and t[0].isdigit() and t[1:3] == ". ":
                        is_list = True
                    elif len(t) > 2 and t[0].isalpha() and t[1:3] == ". ":
                        is_list = True
                        
                    if is_list:
                        # 先给左边垫上两格空白
                        p.Format.CharacterUnitLeftIndent = 2
                        # 然后让首行的第一句话向左“凸出/悬挂” 2 个字符
                        p.Format.CharacterUnitFirstLineIndent = -2
                except Exception:
                    pass
            
            # H1 (-2)
            find_obj.ClearFormatting()
            find_obj.Replacement.ClearFormatting()
            find_obj.Style = doc.Styles(-2)
            find_obj.Replacement.ParagraphFormat.LeftIndent = 0
            find_obj.Replacement.ParagraphFormat.RightIndent = 0
            find_obj.Replacement.ParagraphFormat.CharacterUnitFirstLineIndent = 0
            find_obj.Replacement.ParagraphFormat.Alignment = 1 # 居中
            find_obj.Execute(FindText="", ReplaceWith="", Replace=2)
            
            # H2 (-3), H3 (-4) 及其他各级小标题
            for style_id in range(-9, -2): # -3 到 -9
                try:
                    find_obj.ClearFormatting()
                    find_obj.Replacement.ClearFormatting()
                    find_obj.Style = doc.Styles(style_id)
                    find_obj.Replacement.ParagraphFormat.LeftIndent = 0
                    find_obj.Replacement.ParagraphFormat.RightIndent = 0
                    find_obj.Replacement.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    find_obj.Replacement.ParagraphFormat.Alignment = 0 # 靠左
                    find_obj.Execute(FindText="", ReplaceWith="", Replace=2)
                except Exception:
                    pass
        except Exception as e:
            print(f"批量清理缩进时出错: {e}")

        # 3. 遍历图片，使其直接居中对齐，取消其因正文影响而产生的首行缩进，并提取标题作为题注
        shapes = list(doc.InlineShapes) # 转为 list 防止插入新段落时使迭代器错乱
        print(f"3. 设置所有内嵌图片及其动态图注 (共 {len(shapes)} 张图片)...")
        for i, inline_shape in enumerate(shapes):
            print(f"正在处理第 {i+1}/{len(shapes)} 张图片...")
            # Type 3 = wdInlineShapePicture
            try:
                img_para = inline_shape.Range.Paragraphs(1)
                img_para.Format.Alignment = 1 # 居中
                img_para.Format.CharacterUnitFirstLineIndent = 0
                img_para.Format.FirstLineIndent = 0
                img_para.Format.LeftIndent = 0
                img_para.Format.RightIndent = 0
                img_para.Format.LineSpacingRule = 0 # 0=wdLineSpaceSingle 单倍行距
                img_para.Format.SpaceBefore = 0
                img_para.Format.SpaceAfter = 0

                # 向前寻找最近包含有效文字内容的段落
                curr_para = img_para.Previous()
                caption_text = ""
                search_count = 0
                while curr_para and search_count < 30:
                    text_content = curr_para.Range.Text.replace('\r', '').replace('\n', '').replace('\x07', '').replace('\x0b', '').strip()
                    # 必须要有有效文字，且不能带有图片（InlineShapes count == 0 代表不是图片段）
                    if curr_para.Range.InlineShapes.Count == 0 and len(text_content) > 0:
                        # 截短如果太长
                        if len(text_content) > 60:
                            caption_text = text_content[:50] + "..."
                        else:
                            caption_text = text_content
                        break
                    
                    try:
                        next_para = curr_para.Previous()
                        if next_para is None:
                            break
                        curr_para = next_para
                    except Exception:
                        break
                        
                    search_count += 1
                
                if not caption_text:
                    caption_text = "自动提取图注"

                # --- 强化题注清理逻辑：删除图片下方所有可能是旧题注的段落 ---
                # 检查图片下方接下来的 3 个段落
                for _ in range(3):
                    next_para = img_para.Next()
                    if not next_para:
                        break
                    next_text = next_para.Range.Text.replace('\r', '').replace('\n', '').replace('\x07', '').replace('\x0b', '').replace('\x1e', '').replace('\x1f', '').strip()
                    # 如果该段落为空，或者是之前的题注，或者包含题注文字，则删除
                    is_old_caption = False
                    try:
                        # 检查样式名是否包含 "Caption" 或 "题注"
                        style_name = next_para.Style.NameLocal
                        if "Caption" in style_name or "题注" in style_name:
                            is_old_caption = True
                    except:
                        pass
                    
                    if is_old_caption or len(next_text) == 0 or (len(next_text) > 0 and (next_text in caption_text or caption_text in next_text)):
                        next_para.Range.Delete()
                    else:
                        # 遇到正文文字（且不是题注），停止清理
                        break

                # 在图片底部插入这一段题注
                img_para.Range.InsertParagraphAfter()
                inserted_para = img_para.Next()
                if inserted_para:
                    # 将刚才提取出来的文字作为 Text 内容写入新段落（InsertBefore可以防止删掉回车符）
                    inserted_para.Range.InsertBefore(caption_text)
                    try:
                        inserted_para.Style = doc.Styles(-35) # 应用 Caption 样式
                    except Exception:
                        inserted_para.Format.Alignment = 1
                        inserted_para.Format.CharacterUnitFirstLineIndent = 0
                        inserted_para.Format.FirstLineIndent = 0
            except Exception as e:
                print(f"处理图片及自动提取图注时发生错误: {e}")
                
        # 3.5 页眉页脚与页码
        print("3.5 格式化页眉页脚与页码...")
        try:
            for section in doc.Sections:
                # 页眉: 奇偶页不同往往需要整体设定，这里简化处理，统一设居中
                # 如果文档内没有页脚/页码，我们添加一个居中的阿伯数字页码
                footer = section.Footers(1) # wdHeaderFooterPrimary = 1
                if footer.PageNumbers.Count == 0:
                    # alignment = 1 (居中)
                    footer.PageNumbers.Add(PageNumberAlignment=1, FirstPage=True)
                
                # 设置页眉下划线和字体格式
                header = section.Headers(1)
                header.Range.Font.NameFarEast = "宋体"
                header.Range.Font.Size = 10.5 # 五号
                header.Range.ParagraphFormat.Alignment = 1 # 居中
                header.Range.ParagraphFormat.Borders(-3).LineStyle = 1 # 底部添加边框单实线
        except Exception as e:
            print(f"处理页眉页脚时产生部分错误: {e}")

        # 4. 更新全篇目录 (如果文档内包含自动目录)
        print("4. 更新自动目录...")
        if doc.TablesOfContents.Count > 0:
            for toc in doc.TablesOfContents:
                toc.Update()
                
        # 另存为/保存
        print("正在保存文档...")
        doc.SaveAs(output_abspath)
        print(f"成功完成格式排版，已输出至: {output_abspath}")
        return True

    except Exception as e:
        print(f"处理文档时发生严重错误:\n{traceback.format_exc()}")
        return False
    finally:
        if doc:
            try:
                doc.Close(0) # 0 = wdDoNotSaveChanges 
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python formatter.py <input.docx> [output.docx]")
        print("说明: 若不指定输出文件，则默认覆盖原输入文件。")
        sys.exit(1)
        
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    format_word_document(input_file, output_file)
