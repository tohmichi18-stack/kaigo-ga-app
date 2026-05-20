"""介護DXスライド資料 生成スクリプト（現場職員向け）"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ===== 配色（介護業界らしい温かみのあるトーン） =====
COLOR_PRIMARY = RGBColor(0x2C, 0x6E, 0x91)     # 落ち着いたブルー
COLOR_ACCENT = RGBColor(0xF2, 0x8C, 0x28)      # 温かみのあるオレンジ
COLOR_BG = RGBColor(0xF7, 0xF7, 0xF2)          # オフホワイト
COLOR_DARK = RGBColor(0x1F, 0x2D, 0x3D)        # 濃紺グレー
COLOR_LIGHT = RGBColor(0xE9, 0xEF, 0xF4)       # 淡いブルー
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_SUB = RGBColor(0x55, 0x6B, 0x7D)         # サブテキスト用

# 16:9 サイズ
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H

BLANK = prs.slide_layouts[6]


def add_bg(slide, color=COLOR_BG):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    return bg


def add_side_bar(slide):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.18), SLIDE_H)
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_PRIMARY
    bar.line.fill.background()


def add_text(slide, left, top, width, height, text, size=18, bold=False,
             color=COLOR_DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             font="Yu Gothic UI"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def add_header(slide, title, subtitle=None, page=None, total=None):
    # ヘッダーバー
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.9))
    band.fill.solid()
    band.fill.fore_color.rgb = COLOR_PRIMARY
    band.line.fill.background()
    # アクセント
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.9),
                                    SLIDE_W, Inches(0.06))
    accent.fill.solid()
    accent.fill.fore_color.rgb = COLOR_ACCENT
    accent.line.fill.background()
    # タイトル
    add_text(slide, Inches(0.4), Inches(0.18), Inches(10), Inches(0.6),
             title, size=24, bold=True, color=COLOR_WHITE)
    if subtitle:
        add_text(slide, Inches(0.4), Inches(0.55), Inches(10), Inches(0.35),
                 subtitle, size=12, color=COLOR_LIGHT)
    if page and total:
        add_text(slide, Inches(11.8), Inches(0.3), Inches(1.2), Inches(0.4),
                 f"{page} / {total}", size=12, color=COLOR_WHITE,
                 align=PP_ALIGN.RIGHT)


def add_footer(slide, page=None, total=None):
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(7.15),
                                  Inches(12.5), Emu(9525))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_LIGHT
    line.line.fill.background()
    add_text(slide, Inches(0.4), Inches(7.2), Inches(8), Inches(0.25),
             "介護DXについて  ｜  現場で活きるデジタル活用", size=10, color=COLOR_SUB)
    if page and total:
        add_text(slide, Inches(11.8), Inches(7.2), Inches(1.2), Inches(0.25),
                 f"Page {page} / {total}", size=10, color=COLOR_SUB,
                 align=PP_ALIGN.RIGHT)


def add_card(slide, left, top, width, height, title, body,
             title_color=COLOR_PRIMARY, icon=None):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_WHITE
    card.line.color.rgb = COLOR_LIGHT
    card.line.width = Pt(1)
    # アクセント帯
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = title_color
    bar.line.fill.background()
    # アイコン（絵文字代用）
    title_offset = Inches(0.25)
    if icon:
        add_text(slide, left + Inches(0.25), top + Inches(0.25),
                 Inches(0.6), Inches(0.6), icon, size=24, bold=True,
                 color=title_color)
        title_offset = Inches(0.9)
    add_text(slide, left + title_offset, top + Inches(0.3),
             width - title_offset - Inches(0.2), Inches(0.5),
             title, size=15, bold=True, color=title_color)
    add_text(slide, left + Inches(0.25), top + Inches(0.95),
             width - Inches(0.5), height - Inches(1.1),
             body, size=11, color=COLOR_DARK)


TOTAL_PAGES = 12

# ============================================================
# Slide 1: タイトル
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s, COLOR_WHITE)
# 背景装飾
deco1 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-2), Inches(-2),
                           Inches(6), Inches(6))
deco1.fill.solid()
deco1.fill.fore_color.rgb = COLOR_LIGHT
deco1.line.fill.background()

deco2 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10), Inches(4.5),
                           Inches(5), Inches(5))
deco2.fill.solid()
deco2.fill.fore_color.rgb = COLOR_PRIMARY
deco2.line.fill.background()

deco3 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(11.5), Inches(-1),
                           Inches(3), Inches(3))
deco3.fill.solid()
deco3.fill.fore_color.rgb = COLOR_ACCENT
deco3.line.fill.background()

# ラベル
label = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                           Inches(1.0), Inches(2.2),
                           Inches(2.4), Inches(0.5))
label.fill.solid()
label.fill.fore_color.rgb = COLOR_ACCENT
label.line.fill.background()
add_text(s, Inches(1.0), Inches(2.25), Inches(2.4), Inches(0.4),
         "FOR FRONTLINE STAFF", size=12, bold=True,
         color=COLOR_WHITE, align=PP_ALIGN.CENTER)

add_text(s, Inches(1.0), Inches(2.9), Inches(11), Inches(1.5),
         "介護DXについて", size=60, bold=True, color=COLOR_DARK)
add_text(s, Inches(1.0), Inches(4.3), Inches(11), Inches(0.8),
         "現場で活きるデジタル活用と、これからの介護のかたち",
         size=22, color=COLOR_PRIMARY)

# 区切り線
sep = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(5.4),
                         Inches(2), Inches(0.05))
sep.fill.solid()
sep.fill.fore_color.rgb = COLOR_ACCENT
sep.line.fill.background()

add_text(s, Inches(1.0), Inches(5.6), Inches(8), Inches(0.4),
         "2026年5月  ｜  社内勉強会資料", size=14, color=COLOR_SUB)

# ============================================================
# Slide 2: 目次
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "目次", "本日お話しする内容", page=2, total=TOTAL_PAGES)

agenda = [
    ("01", "介護現場の今", "人手不足・記録業務・多忙化"),
    ("02", "介護DXとは何か", "DXの定義と「IT化」との違い"),
    ("03", "なぜ今、DXなのか", "国の動きと業界の必要性"),
    ("04", "主なデジタル技術", "ICT記録／見守り／ロボット／AI"),
    ("05", "現場での活用事例", "明日から使えるイメージ"),
    ("06", "導入のメリット", "職員・利用者・事業所の3視点"),
    ("07", "うまく進めるコツ", "失敗しないための5原則"),
    ("08", "これからの介護", "私たちの役割の変化"),
    ("09", "まとめ", "まず一歩、何から始めるか"),
]

start_y = 1.3
for i, (no, title, desc) in enumerate(agenda):
    row = i % 5
    col = i // 5
    x = Inches(0.5 + col * 6.4)
    y = Inches(start_y + row * 1.05)
    # No.バッジ
    badge = s.shapes.add_shape(MSO_SHAPE.OVAL, x, y, Inches(0.7), Inches(0.7))
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_PRIMARY
    badge.line.fill.background()
    add_text(s, x, y, Inches(0.7), Inches(0.7), no, size=18, bold=True,
             color=COLOR_WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.9), y + Inches(0.05),
             Inches(5.3), Inches(0.4), title, size=16, bold=True, color=COLOR_DARK)
    add_text(s, x + Inches(0.9), y + Inches(0.45),
             Inches(5.3), Inches(0.35), desc, size=11, color=COLOR_SUB)

add_footer(s, page=2, total=TOTAL_PAGES)

# ============================================================
# Slide 3: 介護現場の今
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "01. 介護現場の今", "皆さんが毎日感じている「あの大変さ」",
           page=3, total=TOTAL_PAGES)

add_text(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.5),
         "現場の声から見える3つの課題", size=18, bold=True, color=COLOR_PRIMARY)

cards = [
    ("人手不足", "👥",
     "・採用しても定着しない\n・夜勤体制を組むのもギリギリ\n"
     "・2040年には約57万人の介護人材が不足\n  （厚労省試算）",
     COLOR_PRIMARY),
    ("記録・書類業務", "📝",
     "・1日の業務時間のうち約3割が記録\n・紙とパソコンで二重入力\n"
     "・申し送りに時間がかかる\n・記録の質にばらつきがある",
     COLOR_ACCENT),
    ("身体的・精神的負担", "💪",
     "・移乗介助での腰痛\n・常に気を張った見守り\n"
     "・突発対応で休憩が取れない\n・利用者と向き合う時間が削られる",
     RGBColor(0x6A, 0x9C, 0x4E)),
]

for i, (title, icon, body, color) in enumerate(cards):
    add_card(s, Inches(0.5 + i * 4.25), Inches(1.9),
             Inches(4.0), Inches(4.5), title, body,
             title_color=color, icon=icon)

# 強調メッセージ
msg = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                        Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.5))
msg.fill.solid()
msg.fill.fore_color.rgb = COLOR_DARK
msg.line.fill.background()
add_text(s, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.5),
         "👉 「これ以上、人だけでがんばる」のは限界。仕組みで支える発想が必要です。",
         size=14, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)

add_footer(s, page=3, total=TOTAL_PAGES)

# ============================================================
# Slide 4: 介護DXとは何か
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "02. 介護DXとは何か", "「ICT化」と「DX」は何が違うのか",
           page=4, total=TOTAL_PAGES)

# 定義ボックス
defbox = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                            Inches(0.5), Inches(1.2),
                            Inches(12.3), Inches(1.3))
defbox.fill.solid()
defbox.fill.fore_color.rgb = COLOR_LIGHT
defbox.line.color.rgb = COLOR_PRIMARY
defbox.line.width = Pt(1.5)

add_text(s, Inches(0.8), Inches(1.35), Inches(12), Inches(0.4),
         "📘 介護DXとは", size=14, bold=True, color=COLOR_PRIMARY)
add_text(s, Inches(0.8), Inches(1.75), Inches(11.7), Inches(0.8),
         "デジタル技術を使って、介護サービスの提供方法・働き方・組織のあり方そのものを\n"
         "見直し、利用者にも職員にもよりよい介護を実現していく取り組み。",
         size=15, color=COLOR_DARK)

# 比較表
add_text(s, Inches(0.5), Inches(2.8), Inches(6), Inches(0.5),
         "比較してみると…", size=16, bold=True, color=COLOR_PRIMARY)

# 表
headers = ["観点", "ICT化（これまで）", "DX（これから）"]
rows = [
    ["目的", "業務をデジタルに置き換える", "ケアの質と働き方を変える"],
    ["主役", "ツール（道具）", "利用者と職員（人）"],
    ["範囲", "一部の業務", "業務全体・組織文化"],
    ["効果", "作業時間の削減", "ケア時間の創出・離職率改善"],
]

table_top = Inches(3.35)
col_widths = [Inches(2.0), Inches(5.0), Inches(5.3)]
col_x = [Inches(0.5), Inches(2.5), Inches(7.5)]
row_h = Inches(0.55)

# ヘッダー
for i, h in enumerate(headers):
    cell = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, col_x[i], table_top,
                              col_widths[i], row_h)
    cell.fill.solid()
    cell.fill.fore_color.rgb = COLOR_PRIMARY
    cell.line.color.rgb = COLOR_WHITE
    add_text(s, col_x[i], table_top, col_widths[i], row_h, h,
             size=13, bold=True, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

for r, row in enumerate(rows):
    for c, val in enumerate(row):
        y = table_top + row_h * (r + 1)
        cell = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, col_x[c], y,
                                  col_widths[c], row_h)
        cell.fill.solid()
        if c == 0:
            cell.fill.fore_color.rgb = COLOR_LIGHT
        else:
            cell.fill.fore_color.rgb = COLOR_WHITE
        cell.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
        bold = (c == 0)
        add_text(s, col_x[c], y, col_widths[c], row_h, val,
                 size=12, bold=bold, color=COLOR_DARK,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 一言
add_text(s, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.4),
         "💡 ポイント：道具を入れるだけがゴールではなく、入れたことで何が変わるか、が大切",
         size=13, bold=True, color=COLOR_ACCENT)

add_footer(s, page=4, total=TOTAL_PAGES)

# ============================================================
# Slide 5: なぜ今、DXなのか
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "03. なぜ今、DXなのか", "国も業界も、待ったなしの状況",
           page=5, total=TOTAL_PAGES)

# 左：数字で見る現状
add_text(s, Inches(0.5), Inches(1.2), Inches(6), Inches(0.5),
         "数字で見る、いまの介護", size=18, bold=True, color=COLOR_PRIMARY)

stats = [
    ("2040年", "高齢者人口がピークに", COLOR_PRIMARY),
    ("約57万人", "不足する介護人材（2040年）", COLOR_ACCENT),
    ("約3割", "記録業務が占める就業時間", RGBColor(0x6A, 0x9C, 0x4E)),
    ("14%超", "介護職の年間離職率", RGBColor(0xC0, 0x4F, 0x4F)),
]

for i, (num, desc, color) in enumerate(stats):
    y = Inches(1.85 + i * 1.15)
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                             Inches(0.5), y, Inches(6.0), Inches(1.0))
    box.fill.solid()
    box.fill.fore_color.rgb = COLOR_WHITE
    box.line.color.rgb = color
    box.line.width = Pt(1.5)
    add_text(s, Inches(0.7), y + Inches(0.1), Inches(2.5), Inches(0.8),
             num, size=28, bold=True, color=color, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(3.2), y + Inches(0.1), Inches(3.2), Inches(0.8),
             desc, size=13, color=COLOR_DARK, anchor=MSO_ANCHOR.MIDDLE)

# 右：国の動き
add_text(s, Inches(7.0), Inches(1.2), Inches(6), Inches(0.5),
         "国・業界の動き", size=18, bold=True, color=COLOR_PRIMARY)

policies = [
    ("介護報酬改定 (2024)",
     "ICT・介護ロボット導入により人員配置基準が柔軟化"),
    ("LIFE（科学的介護情報システム）",
     "ケアの記録データを国が分析し、エビデンスベースのケアへ"),
    ("ICT導入支援補助金",
     "都道府県ごとに、機器・ソフトの導入に補助制度あり"),
    ("BCP（業務継続計画）の義務化",
     "災害・感染症対応にもデジタル基盤が不可欠に"),
]

for i, (title, body) in enumerate(policies):
    y = Inches(1.85 + i * 1.15)
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(7.0), y, Inches(5.8), Inches(1.0))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_WHITE
    card.line.color.rgb = COLOR_LIGHT
    # 左の帯
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.0), y,
                             Inches(0.12), Inches(1.0))
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_ACCENT
    bar.line.fill.background()
    add_text(s, Inches(7.3), y + Inches(0.12), Inches(5.4), Inches(0.4),
             title, size=13, bold=True, color=COLOR_DARK)
    add_text(s, Inches(7.3), y + Inches(0.5), Inches(5.4), Inches(0.5),
             body, size=11, color=COLOR_SUB)

add_footer(s, page=5, total=TOTAL_PAGES)

# ============================================================
# Slide 6: 主なデジタル技術
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "04. 主なデジタル技術", "現場で広がっている4つの領域",
           page=6, total=TOTAL_PAGES)

techs = [
    ("ICT記録ソフト", "📱",
     "スマホ・タブレットで\n記録をその場で入力。\n紙の連絡帳や手書きの\n申し送りを置き換え。",
     "例：ほのぼの、CareViewer、\nワイズマン、カナミック など"),
    ("見守りセンサー", "🛏",
     "ベッドや居室の状態を\nセンサーで把握。\n離床・呼吸・心拍を\n夜勤者へ通知。",
     "例：眠りSCAN、aams、\nNeOSなど"),
    ("介護ロボット", "🤖",
     "移乗・移動・入浴を\nサポートする機器。\n腰痛予防と一人介助化に\nつながる。",
     "例：ハグL1、リショーネ、\nROBEAR、Hug など"),
    ("AI・データ活用", "📊",
     "蓄積した記録を分析し、\nケアプラン作成や\n転倒予測、シフト最適化\nなどに応用。",
     "例：LIFE連携ソフト、\nAIケアプラン提案 など"),
]

card_w = Inches(3.0)
card_h = Inches(5.2)
gap = Inches(0.16)
total_w = card_w * 4 + gap * 3
start_x = (SLIDE_W - total_w) / 2

colors = [COLOR_PRIMARY, COLOR_ACCENT,
          RGBColor(0x6A, 0x9C, 0x4E), RGBColor(0x8E, 0x5A, 0xB7)]

for i, (title, icon, body, ex) in enumerate(techs):
    x = start_x + (card_w + gap) * i
    y = Inches(1.3)
    # カード本体
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h)
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_WHITE
    card.line.color.rgb = COLOR_LIGHT
    # ヘッダー帯
    head = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_w, Inches(1.3))
    head.fill.solid()
    head.fill.fore_color.rgb = colors[i]
    head.line.fill.background()
    add_text(s, x, y + Inches(0.15), card_w, Inches(0.6),
             icon, size=36, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x, y + Inches(0.8), card_w, Inches(0.4),
             title, size=15, bold=True, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER)
    # 本文
    add_text(s, x + Inches(0.25), y + Inches(1.55),
             card_w - Inches(0.5), Inches(2.0),
             body, size=12, color=COLOR_DARK)
    # 例
    add_text(s, x + Inches(0.25), y + Inches(3.7),
             card_w - Inches(0.5), Inches(0.3),
             "▼ 代表例", size=10, bold=True, color=COLOR_ACCENT)
    add_text(s, x + Inches(0.25), y + Inches(4.0),
             card_w - Inches(0.5), Inches(1.0),
             ex, size=10, color=COLOR_SUB)

add_footer(s, page=6, total=TOTAL_PAGES)

# ============================================================
# Slide 7: 現場での活用事例
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "05. 現場での活用事例", "「明日の業務」がこう変わる",
           page=7, total=TOTAL_PAGES)

cases = [
    ("👵", "夜勤帯", "見守りセンサーで先回り",
     "Before：30分ごとの巡視で寝ている方を起こしてしまう。\n"
     "After：センサー通知で必要な方の所だけ訪室。職員の負担も利用者の睡眠も改善。"),
    ("📝", "日中の記録", "音声入力＆スマホ記録",
     "Before：休憩時間や勤務後に記録のために事務所に戻る。\n"
     "After：歩きながら音声で記録。1日あたり30〜60分の業務時間短縮。"),
    ("🍽", "食事・水分管理", "タブレットで一元管理",
     "Before：食事量・水分量を紙に書き、後で別の表に転記。\n"
     "After：その場で入力すれば家族や看護師にも自動共有。ヒヤリハットの早期発見に。"),
    ("🧓", "ケアプラン作成", "AIによる提案＆LIFE連携",
     "Before：経験頼みで時間がかかる。\n"
     "After：過去データやエビデンスをAIが提示。ケアマネと多職種で議論する時間に充てられる。"),
]

for i, (icon, scene, title, body) in enumerate(cases):
    y = Inches(1.25 + i * 1.4)
    # 背景カード
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(0.5), y, Inches(12.3), Inches(1.25))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_WHITE
    card.line.color.rgb = COLOR_LIGHT
    # アイコン円
    circle = s.shapes.add_shape(MSO_SHAPE.OVAL,
                                Inches(0.75), y + Inches(0.18),
                                Inches(0.9), Inches(0.9))
    circle.fill.solid()
    circle.fill.fore_color.rgb = COLOR_LIGHT
    circle.line.fill.background()
    add_text(s, Inches(0.75), y + Inches(0.18), Inches(0.9), Inches(0.9),
             icon, size=24, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # シーンタグ
    tag = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                             Inches(1.85), y + Inches(0.2),
                             Inches(1.3), Inches(0.35))
    tag.fill.solid()
    tag.fill.fore_color.rgb = COLOR_ACCENT
    tag.line.fill.background()
    add_text(s, Inches(1.85), y + Inches(0.2), Inches(1.3), Inches(0.35),
             scene, size=11, bold=True, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # タイトル
    add_text(s, Inches(3.3), y + Inches(0.2), Inches(9.0), Inches(0.4),
             title, size=15, bold=True, color=COLOR_PRIMARY)
    # 本文
    add_text(s, Inches(3.3), y + Inches(0.6), Inches(9.0), Inches(0.7),
             body, size=11, color=COLOR_DARK)

add_footer(s, page=7, total=TOTAL_PAGES)

# ============================================================
# Slide 8: 導入のメリット
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "06. 導入のメリット", "3つの立場から見たプラス効果",
           page=8, total=TOTAL_PAGES)

groups = [
    ("👨‍⚕️ 職員にとって", COLOR_PRIMARY, [
        "記録・転記の時間が減る",
        "夜勤の精神的負担が軽減",
        "腰痛などの身体的負担が減る",
        "申し送りがスムーズになる",
        "ケアに集中できる時間が増える",
    ]),
    ("👵 利用者にとって", COLOR_ACCENT, [
        "見守られている安心感",
        "睡眠を妨げない巡視",
        "より個別性の高いケア",
        "状態変化への早い対応",
        "家族との情報共有が円滑に",
    ]),
    ("🏢 事業所にとって", RGBColor(0x6A, 0x9C, 0x4E), [
        "離職率の低下・採用力アップ",
        "サービスの質の標準化",
        "事故・ヒヤリハットの低減",
        "加算取得・人員基準の柔軟化",
        "データに基づく経営判断",
    ]),
]

card_w = Inches(4.05)
card_h = Inches(5.3)
gap = Inches(0.2)
start_x = (SLIDE_W - (card_w * 3 + gap * 2)) / 2

for i, (title, color, items) in enumerate(groups):
    x = start_x + (card_w + gap) * i
    y = Inches(1.3)
    # カード
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h)
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_WHITE
    card.line.color.rgb = color
    card.line.width = Pt(1.5)
    # タイトル帯
    head = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_w, Inches(0.8))
    head.fill.solid()
    head.fill.fore_color.rgb = color
    head.line.fill.background()
    add_text(s, x, y, card_w, Inches(0.8), title, size=17, bold=True,
             color=COLOR_WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 項目
    for j, item in enumerate(items):
        iy = y + Inches(1.0 + j * 0.8)
        # チェックマーク
        chk = s.shapes.add_shape(MSO_SHAPE.OVAL,
                                 x + Inches(0.3), iy + Inches(0.1),
                                 Inches(0.35), Inches(0.35))
        chk.fill.solid()
        chk.fill.fore_color.rgb = color
        chk.line.fill.background()
        add_text(s, x + Inches(0.3), iy + Inches(0.1),
                 Inches(0.35), Inches(0.35), "✓", size=12, bold=True,
                 color=COLOR_WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + Inches(0.8), iy + Inches(0.08),
                 card_w - Inches(1.0), Inches(0.5),
                 item, size=12, color=COLOR_DARK, anchor=MSO_ANCHOR.TOP)

add_footer(s, page=8, total=TOTAL_PAGES)

# ============================================================
# Slide 9: うまく進めるコツ
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "07. うまく進めるコツ", "失敗しないための5原則",
           page=9, total=TOTAL_PAGES)

tips = [
    ("①", "「困りごと」から始める",
     "ツールありきではなく、現場の困りごとを起点に。"),
    ("②", "小さく始めて、広げる",
     "まずは1フロア・1業務から。成功体験を積み重ねる。"),
    ("③", "推進役と現場の声をつなぐ",
     "「DX担当」を孤立させない。現場の声を吸い上げる仕組みを。"),
    ("④", "苦手な人を置き去りにしない",
     "操作研修・マニュアル整備・聞ける相手を用意する。"),
    ("⑤", "効果を見える化する",
     "時間・件数・声をデータで残し、次の改善につなげる。"),
]

for i, (no, title, desc) in enumerate(tips):
    y = Inches(1.3 + i * 1.05)
    # 番号バッジ
    badge = s.shapes.add_shape(MSO_SHAPE.OVAL,
                               Inches(0.7), y, Inches(0.85), Inches(0.85))
    badge.fill.solid()
    badge.fill.fore_color.rgb = COLOR_ACCENT
    badge.line.fill.background()
    add_text(s, Inches(0.7), y, Inches(0.85), Inches(0.85), no,
             size=26, bold=True, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # カード
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(1.8), y, Inches(11.0), Inches(0.85))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_WHITE
    card.line.color.rgb = COLOR_LIGHT
    add_text(s, Inches(2.0), y + Inches(0.05),
             Inches(10.8), Inches(0.4), title, size=15, bold=True,
             color=COLOR_PRIMARY)
    add_text(s, Inches(2.0), y + Inches(0.45),
             Inches(10.8), Inches(0.4), desc, size=12, color=COLOR_DARK)

add_footer(s, page=9, total=TOTAL_PAGES)

# ============================================================
# Slide 10: これからの介護
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "08. これからの介護", "私たちの仕事はどう変わる？",
           page=10, total=TOTAL_PAGES)

# 中央矢印
arrow_y = Inches(3.5)
arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                           Inches(6.0), arrow_y, Inches(1.3), Inches(0.7))
arrow.fill.solid()
arrow.fill.fore_color.rgb = COLOR_ACCENT
arrow.line.fill.background()

# Before
before = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                            Inches(0.5), Inches(1.3),
                            Inches(5.4), Inches(5.4))
before.fill.solid()
before.fill.fore_color.rgb = COLOR_WHITE
before.line.color.rgb = COLOR_SUB
add_text(s, Inches(0.5), Inches(1.5), Inches(5.4), Inches(0.5),
         "これまで", size=18, bold=True, color=COLOR_SUB,
         align=PP_ALIGN.CENTER)
add_text(s, Inches(0.5), Inches(1.9), Inches(5.4), Inches(0.4),
         "── 人が頑張る介護 ──", size=12, color=COLOR_SUB,
         align=PP_ALIGN.CENTER)

before_items = [
    "📋 記録は紙とPCで二重入力",
    "🚶 30分おきの定時巡視",
    "💪 二人介助が前提の移乗",
    "🗣 口頭中心の申し送り",
    "👀 経験と勘によるケア判断",
    "🕰 残業前提の業務スタイル",
]
for i, t in enumerate(before_items):
    add_text(s, Inches(0.8), Inches(2.5 + i * 0.6), Inches(4.8), Inches(0.5),
             t, size=12, color=COLOR_DARK)

# After
after = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                           Inches(7.4), Inches(1.3),
                           Inches(5.4), Inches(5.4))
after.fill.solid()
after.fill.fore_color.rgb = COLOR_LIGHT
after.line.color.rgb = COLOR_PRIMARY
after.line.width = Pt(1.5)
add_text(s, Inches(7.4), Inches(1.5), Inches(5.4), Inches(0.5),
         "これから", size=18, bold=True, color=COLOR_PRIMARY,
         align=PP_ALIGN.CENTER)
add_text(s, Inches(7.4), Inches(1.9), Inches(5.4), Inches(0.4),
         "── 人と技術で支える介護 ──", size=12, color=COLOR_PRIMARY,
         align=PP_ALIGN.CENTER)

after_items = [
    "📱 その場でスマホ記録・自動共有",
    "🛏 センサーで必要な時だけ訪室",
    "🤖 ロボット補助で一人介助も可能",
    "💬 デジタル申し送りで漏れゼロへ",
    "📊 データに基づくケアの根拠化",
    "⏰ 定時で帰り、ケアに集中する時間",
]
for i, t in enumerate(after_items):
    add_text(s, Inches(7.7), Inches(2.5 + i * 0.6), Inches(4.8), Inches(0.5),
             t, size=12, bold=True, color=COLOR_DARK)

add_footer(s, page=10, total=TOTAL_PAGES)

# ============================================================
# Slide 11: まとめ
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s)
add_header(s, "09. まとめ", "まず一歩、明日からできること",
           page=11, total=TOTAL_PAGES)

# キーメッセージ
key = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                         Inches(0.5), Inches(1.3),
                         Inches(12.3), Inches(1.5))
key.fill.solid()
key.fill.fore_color.rgb = COLOR_PRIMARY
key.line.fill.background()
add_text(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(0.6),
         "🌱 介護DXのゴールは「人にしかできないケア」に集中できる現場をつくること。",
         size=18, bold=True, color=COLOR_WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Inches(0.5), Inches(2.0), Inches(12.3), Inches(0.7),
         "デジタルは目的ではなく、利用者と職員の笑顔を増やすための手段です。",
         size=14, color=COLOR_LIGHT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 3つの行動
add_text(s, Inches(0.5), Inches(3.1), Inches(12), Inches(0.5),
         "今日から、私たちにできる3つのこと", size=18, bold=True, color=COLOR_PRIMARY)

actions = [
    ("🔍", "知る",
     "今、自分の現場にはどんなICT機器・ソフトがあるかを\n"
     "あらためて確認してみる。"),
    ("💬", "話す",
     "「ここが大変」「もっとこうしたい」を同僚と共有する。\n"
     "現場の声がDXの第一歩。"),
    ("✋", "試す",
     "新しい仕組みが入ったとき、まず触ってみる。\n"
     "気づきはマニュアルにフィードバック。"),
]

card_w = Inches(4.0)
gap = Inches(0.15)
start_x = (SLIDE_W - (card_w * 3 + gap * 2)) / 2

for i, (icon, title, body) in enumerate(actions):
    x = start_x + (card_w + gap) * i
    y = Inches(3.75)
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y,
                              card_w, Inches(3.0))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_WHITE
    card.line.color.rgb = COLOR_ACCENT
    card.line.width = Pt(1.5)
    add_text(s, x, y + Inches(0.25), card_w, Inches(0.8),
             icon, size=42, color=COLOR_ACCENT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x, y + Inches(1.2), card_w, Inches(0.5),
             title, size=22, bold=True, color=COLOR_DARK,
             align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.25), y + Inches(1.85),
             card_w - Inches(0.5), Inches(1.0),
             body, size=12, color=COLOR_SUB, align=PP_ALIGN.CENTER)

add_footer(s, page=11, total=TOTAL_PAGES)

# ============================================================
# Slide 12: Thank you / Q&A
# ============================================================
s = prs.slides.add_slide(BLANK)
add_bg(s, COLOR_DARK)

# 装飾
circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-3), Inches(-3),
                            Inches(7), Inches(7))
circle.fill.solid()
circle.fill.fore_color.rgb = COLOR_PRIMARY
circle.line.fill.background()

circle2 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10), Inches(4),
                             Inches(6), Inches(6))
circle2.fill.solid()
circle2.fill.fore_color.rgb = COLOR_ACCENT
circle2.line.fill.background()

add_text(s, Inches(0.5), Inches(2.5), Inches(12.3), Inches(1.5),
         "Thank you.", size=72, bold=True, color=COLOR_WHITE,
         align=PP_ALIGN.CENTER)

# 区切り
sep = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                        Inches(5.5), Inches(4.2),
                        Inches(2.3), Inches(0.05))
sep.fill.solid()
sep.fill.fore_color.rgb = COLOR_ACCENT
sep.line.fill.background()

add_text(s, Inches(0.5), Inches(4.5), Inches(12.3), Inches(0.7),
         "ご質問・ご意見をお待ちしています",
         size=22, color=COLOR_LIGHT, align=PP_ALIGN.CENTER)

add_text(s, Inches(0.5), Inches(5.5), Inches(12.3), Inches(0.5),
         "「やってみたい」「ここが不安」── どんな声も歓迎です。",
         size=14, color=COLOR_LIGHT, align=PP_ALIGN.CENTER)

# 保存
out = "/home/user/kaigo-ga-app/slides/介護DXについて.pptx"
prs.save(out)
print(f"Saved: {out}")
