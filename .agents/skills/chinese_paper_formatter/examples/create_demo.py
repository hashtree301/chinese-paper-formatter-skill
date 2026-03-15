import win32com.client
import os
import sys

def create_demo_doc(output_path):
    abspath = os.path.abspath(output_path)
    if os.path.exists(abspath):
        try:
            os.remove(abspath)
        except Exception:
            pass
            
    print("Starting Word to create demo doc...")
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Add()
        
        selection = word.Selection
        
        # 一级标题
        selection.Style = doc.Styles(-2) # wdStyleHeading1
        selection.Font.Name = "微软雅黑"
        selection.Font.Size = 20
        selection.TypeText("第一章 绪论\n")
        
        # 正文，格式错乱
        selection.Style = doc.Styles(-1) # Normal
        selection.Font.Name = "Arial"
        selection.Font.Size = 10
        selection.TypeText("这是第一章的测试正文部分，用来检验格式整理脚本的功能。This sentence is used for testing English fonts (Times New Roman).\n")
        
        # 二级标题
        selection.Style = doc.Styles(-3) # wdStyleHeading2
        selection.Font.Name = "黑体"
        selection.Font.Size = 12
        # 故意增加两个字符的缩进错误
        selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2
        selection.TypeText("1.1 研究背景\n")
        
        # 正文
        selection.Style = doc.Styles(-1)
        selection.Font.Name = "宋体"
        selection.Font.Size = 14
        selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0 # 故意取消正文首行缩进
        selection.TypeText("测试二级标题下的正文段落。此段落故意打乱了行距，并且没有首行缩进。我们将通过脚本让其变回两字符的缩进。\n")

        # 三级标题 - 极端的左侧缩进和首行缩进混合
        selection.Style = doc.Styles(-4) # wdStyleHeading3
        selection.Font.Name = "宋体"
        selection.Font.Size = 10
        selection.ParagraphFormat.LeftIndent = 20 # 磅
        selection.ParagraphFormat.CharacterUnitFirstLineIndent = 4 # 离谱的4字符缩进
        selection.TypeText("   1.1.1 国内外研究现状 (前面还带有手动空格)\n")

        # 四级标题 - 居中对齐错误加上首行缩进
        selection.Style = doc.Styles(-5) # wdStyleHeading4
        selection.ParagraphFormat.Alignment = 1 # 错误地设为了居中
        selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2
        selection.TypeText("1.1.1.1 详细现状分析\n")
        
        # 保存
        doc.SaveAs(abspath)
        print(f"Demo generated securely at {abspath}")
        
    except Exception as e:
        print(f"Failed to create demo document: {e}")
    finally:
        try:
            doc.Close(0)
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass

if __name__ == "__main__":
    # Get dir of script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output = os.path.join(script_dir, "demo_input.docx")
    create_demo_doc(output)
