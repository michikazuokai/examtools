import copy
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import grey
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from pathlib import Path
import re
import sys
import json

# 現在の場所
# curdir = Path(__file__).parent.parent

# #subject = sys.argv[1]
# subject='1020801'
# fname=f'answers_{subject}.json'

def make_pdf(kaito,outfile, max_rows_per_page, version=None):

    if version:
        titleversion=f" ({version})"
    else:
        titleversion=""

# with open(curdir / "work" / fname, "r", encoding="utf-8") as f:
#     kaito = json.load(f)
# print(f"{fname}からjsonファイルを読み込みました")

    ehash=kaito[0]["metainfo"]["hash"][:7]
    ever=str(kaito[0]["metainfo"]["verno"])

    fsyear=kaito[0]['fsyear']
    meta_text = f"{ehash}{fsyear[2:4]}-{ever.zfill(2)}"

    # ✅ フォント登録（.ttfのパスをあなたの環境に合わせて修正）
    pdfmetrics.registerFont(TTFont('IPAexGothic', '/Library/Fonts/ipaexg.ttf'))

    styles = getSampleStyleSheet()
    #
    style = styles["Normal"]

    def safe_paragraph(text, style, context_label=""):
        try:
            para = Paragraph(text, style)
            # wrap() を試してレイアウトに関する不備を事前チェック（サイズは仮に指定）
            para.wrap(500, 800)
            return para
        except Exception as e:
            print(f"❌ エラー: Paragraph の生成または wrap に失敗しました（{context_label}）")
            print(f"　🔍 text: {repr(text)}")
            print(f"　⚠️ エラー内容: {e}")
            raise

    #
    title_style = ParagraphStyle(
        'title', parent=styles['Title'], fontName='IPAexGothic',
        fontSize=18, spaceAfter=6
    )

    subtitle_style = ParagraphStyle(
        'subtitle', parent=styles['Normal'], fontName='IPAexGothic',
        fontSize=12, textColor=colors.black
    )

    small_style = ParagraphStyle(
        'smaiistyle', parent=styles['Normal'], fontName='IPAexGothic',
        fontSize=9,alignment=TA_CENTER  # ← ここで中央揃えを指定
    )

    body_style = ParagraphStyle(
        'body', parent=styles['BodyText'], fontName='IPAexGothic'
    )

    my_style = ParagraphStyle(
        name="MyCode",
        fontName="Courier",
        fontSize=11,
        leading=12,alignment=TA_CENTER  # ← ここで中央揃えを指定
    )
    my_code = ParagraphStyle(
        name="MyCode",
        fontName="Courier",
        fontSize=11,
        leading=18,
        leftIndent=40  # ← ここで右にずらす（単位はポイント）
    )

    # def contains_japanese(text):
    #     return re.search(r'[\u3040-\u30FF\u4E00-\u9FFF]', text) is not None

    def contains_japanese(text):
        # ※ や 全角記号など「ASCII以外」が入っていたら日本語フォント側へ
        return re.search(r'[^\x00-\x7F]', str(text)) is not None


    # 表のスタイル
    tblstyle=[
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('ROWHEIGHTS', (0,0), (-1,-1), 20),
    ]

    story = []

    def add_title(story):
        wtitle=kaito[0]['title']+titleversion  #試験タイトル
        story.append(safe_paragraph(f"{wtitle} 履修判定試験", title_style))
        story.append(Spacer(1, 10))
        nenji=kaito[0]['nenji']+'年'
        # 表のデータと幅を指定
        l=[[ safe_paragraph(v, subtitle_style) for v in [nenji, "学籍（下２桁）", "氏名", "点"] ]]
        table1 = Table(l, colWidths=[40, 120, 220,100])
        # 行数に応じた下線スタイルだけを先に作る
        underline_commands = [('LINEBELOW', (0, i), (-1, i), 0.5, colors.black) for i in range(4)]
        # 基本スタイルと結合して全体のスタイルにする
        style = TableStyle(underline_commands + [
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 5),
        ])

        table1.setStyle(style)
        table1.hAlign = 'LEFT'
        story.append(table1)
        story.append(Spacer(1, 20))

    command=copy.deepcopy(tblstyle)  # ← ここが重要

    def create_table_from_text(data_list, haba_list,  style_command, row_height):
        data = []
        row = []
        for w in data_list:
            if "nbsp" in str(w):
                fcode=f"<pre>{w}</pre>"
                row.append(safe_paragraph(fcode, my_code))
            elif contains_japanese(str(w)):
                row.append(safe_paragraph(w, small_style))
            elif "※" in str(w) or contains_japanese(str(w)):
                row.append(safe_paragraph(str(w), small_style))
            else:
                row.append(safe_paragraph(str(w), my_style))
        data.append(row)
        haba = [w * kaito[0]['width'] for w in haba_list]
      #  print(data)
      #  print(haba)
        table = Table(data, colWidths=haba, rowHeights=[row_height])
        table.setStyle(TableStyle(style_command))
        table.hAlign = 'LEFT'
        return table

    # --- メイン処理部分 ---
    for k in range(2):
        # 1ページ目：問題ページ
        add_title(story)

        for i,v in enumerate(kaito[1:]):
            lwidth=[w * 50 for w in v['width']]
            table1 = create_table_from_text(v['label'], v['width'],  command, row_height=20)
            story.append(table1)
            if k==0:
                # 解答欄（空白）を表示
                table2 = create_table_from_text(["" for _ in v['label']], v['width'],  command, row_height=v['height'][0])
            else:
                # 解答を表示
                table2 = create_table_from_text(v['answer'], v['width'],  command, row_height=v['height'][0])
            story.append(table2)

            ##if i == 6:  #改ページ（１枚の解答用紙で裏面を使う時の処理（７行を超える解答欄）
            ##    if len(kaito) - 1 > 7 :
            if i == max_rows_per_page - 1:
                if len(kaito) - 1 > max_rows_per_page:
                    story.append(Spacer(1, 12))
                    story.append(safe_paragraph(f"裏面につづく", body_style))
                    story.append(PageBreak())
            else:
                story.append(Spacer(1, 8))
            
        story.append(Spacer(1, 12))
        story.append(safe_paragraph(kaito[0]['kaito_message'], body_style))

        #バージョン情報を表示
        # alignment=2 は右寄せを意味します（0=左, 1=中央, 2=右）。
        # fontSize=6 は非常に小さい文字です。必要に応じて 7〜8 に調整可能です。
        # textColor=grey で文字色を薄くしています。lightgrey にしてもさらに淡くなります。
        # rightIndent=0 は右端ピッタリに寄せる調整用です。
        meta_style = ParagraphStyle(
            name='MetaStyle',
            fontSize=6,
            textColor=grey,
            alignment=2,  # right-align
            rightIndent=0
        )
        story.append(Paragraph(meta_text, meta_style))

        story.append(PageBreak())

    # ✅ 文書生成
    # outfile=str(curdir / "output" / subject / (f"{subject}_{kaito[0]['title']}解答用紙.pdf"))
    # outfile=str(curdir / "output" / subject / (f"{subject}_{kaito[0]['title']}解答用紙.pdf"))
    doc = SimpleDocTemplate(
        str(outfile), 
        leftMargin=45, 
        pagesize=A4, 
        topMargin=10*mm, 
        bottomMargin=20*mm
        )

    #doc.build(story) 
    try:
        doc.build(story)
    except Exception as e:
        print("build中にエラー:", e)
        for i, item in enumerate(story):
            try:
                item.wrap(400, 800)  # 幅と高さは仮の値
            except Exception as e2:
                print(f"→ story[{i}] でエラー: {e2}")
                print(item)
                break
    