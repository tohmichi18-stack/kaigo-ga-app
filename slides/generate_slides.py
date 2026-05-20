"""主任ケアマネ研修スライド生成スクリプト

テーマ: 適切なケアマネジメント手法の実践活用
     ー厚生労働省が推進する基本ケア、疾患別ケアの考え方を
       実際のアセスメント、ケアプラン作成にどう落とし込むか
時間: 90分 / 認知症中心
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ===== カラーパレット =====
NAVY = RGBColor(0x1F, 0x3A, 0x5F)
TEAL = RGBColor(0x2E, 0x86, 0xAB)
ORANGE = RGBColor(0xE8, 0x7A, 0x41)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
DARK_GRAY = RGBColor(0x40, 0x40, 0x40)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT_BG = RGBColor(0xEA, 0xF3, 0xF9)
SOFT_ORANGE = RGBColor(0xFC, 0xEC, 0xDD)

FONT = "Yu Gothic"

prs = Presentation()
prs.slide_width = Cm(33.867)   # 16:9 ワイド
prs.slide_height = Cm(19.05)

SW = prs.slide_width
SH = prs.slide_height

blank_layout = prs.slide_layouts[6]


# ===== ヘルパー関数 =====
def add_slide():
    return prs.slides.add_slide(blank_layout)


def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
    shape.shadow.inherit = False
    return shape


def add_text(slide, left, top, width, height, text, *,
             size=14, bold=False, color=DARK_GRAY,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=FONT,
             line_spacing=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0.1)
    tf.margin_right = Cm(0.1)
    tf.margin_top = Cm(0.05)
    tf.margin_bottom = Cm(0.05)
    tf.vertical_anchor = anchor

    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing is not None:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
    return box


def add_header(slide, section_label, title):
    # 上部帯
    add_rect(slide, 0, 0, SW, Cm(1.6), NAVY)
    # セクションラベル
    add_text(slide, Cm(0.8), Cm(0.25), Cm(8), Cm(0.6),
             section_label, size=11, color=WHITE, bold=True)
    # タイトル
    add_text(slide, Cm(0.8), Cm(1.9), SW - Cm(1.6), Cm(1.3),
             title, size=24, bold=True, color=NAVY)
    # アクセントライン
    add_rect(slide, Cm(0.8), Cm(3.2), Cm(2.5), Cm(0.12), ORANGE)


def add_footer(slide, page_num, total=37):
    add_text(slide, Cm(0.8), SH - Cm(0.8), Cm(20), Cm(0.5),
             "主任介護支援専門員研修 / 適切なケアマネジメント手法の実践活用",
             size=9, color=DARK_GRAY)
    add_text(slide, SW - Cm(3), SH - Cm(0.8), Cm(2.2), Cm(0.5),
             f"{page_num} / {total}", size=9, color=DARK_GRAY, align=PP_ALIGN.RIGHT)


def add_bullets(slide, left, top, width, height, items, *,
                size=14, color=DARK_GRAY, bullet_color=TEAL, line_spacing=1.3):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Cm(0.1)

    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        # bullet
        r1 = p.add_run()
        r1.text = "■  "
        r1.font.name = FONT
        r1.font.size = Pt(size)
        r1.font.bold = True
        r1.font.color.rgb = bullet_color
        # text
        r2 = p.add_run()
        r2.text = item
        r2.font.name = FONT
        r2.font.size = Pt(size)
        r2.font.color.rgb = color
    return box


def add_box(slide, left, top, width, height, title, body, *,
            title_bg=TEAL, body_bg=ACCENT_BG, title_color=WHITE,
            body_color=DARK_GRAY, title_size=14, body_size=12):
    # title bar
    add_rect(slide, left, top, width, Cm(1.0), title_bg)
    add_text(slide, left + Cm(0.2), top + Cm(0.1), width - Cm(0.4), Cm(0.8),
             title, size=title_size, bold=True, color=title_color,
             anchor=MSO_ANCHOR.MIDDLE)
    # body
    add_rect(slide, left, top + Cm(1.0), width, height - Cm(1.0), body_bg)
    add_text(slide, left + Cm(0.3), top + Cm(1.1), width - Cm(0.6),
             height - Cm(1.1), body, size=body_size, color=body_color)


# =============================================================
#  SLIDE 1: 表紙
# =============================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, NAVY)
# 装飾バー
add_rect(s, 0, Cm(8), SW, Cm(0.15), ORANGE)
add_rect(s, 0, Cm(12.5), SW, Cm(0.05), TEAL)

# サブタイトル(上)
add_text(s, Cm(2), Cm(4.5), SW - Cm(4), Cm(1),
         "主任介護支援専門員 研修会",
         size=18, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)

# メインタイトル
add_text(s, Cm(1.5), Cm(6.0), SW - Cm(3), Cm(2.2),
         "適切なケアマネジメント手法の実践活用",
         size=36, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

# サブ
add_text(s, Cm(1.5), Cm(9), SW - Cm(3), Cm(2.5),
         "ー厚生労働省が推進する「基本ケア」「疾患別ケア」の考え方を\nアセスメント・ケアプラン作成にどう落とし込むかー",
         size=18, color=WHITE, align=PP_ALIGN.CENTER)

# 開催情報
add_text(s, Cm(2), Cm(14.5), SW - Cm(4), Cm(1),
         "主任ケアマネジャーの会  主催",
         size=16, color=WHITE, align=PP_ALIGN.CENTER)
add_text(s, Cm(2), Cm(15.5), SW - Cm(4), Cm(1),
         "研修時間 90 分",
         size=14, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)


# =============================================================
#  SLIDE 2: 本日の目標
# =============================================================
s = add_slide()
add_header(s, "Introduction", "本日の研修目標")
add_footer(s, 2)

goals = [
    "「適切なケアマネジメント手法」の全体構造(基本ケア・疾患別ケア)を理解する",
    "基本ケアの4つの柱と支援の必要性を判断する項目を実務に位置づける",
    "認知症の疾患別ケアの考え方を、進行段階に応じて整理できる",
    "課題分析(アセスメント)で「想定される支援内容」を仮説検証的に活用できる",
    "第1表〜第3表のケアプランに「基本ケア+認知症ケア」を落とし込める",
    "事業所内のケース検討で活用するためのポイントを持ち帰る",
]
add_bullets(s, Cm(1.2), Cm(4.5), SW - Cm(2.4), Cm(13),
            goals, size=17, line_spacing=1.45)


# =============================================================
#  SLIDE 3: 本日の流れ(アジェンダ)
# =============================================================
s = add_slide()
add_header(s, "Agenda", "本日の流れ(90分)")
add_footer(s, 3)

agenda = [
    ("1", "適切なケアマネジメント手法とは",            "10分"),
    ("2", "基本ケアの考え方(4つの柱)",                 "15分"),
    ("3", "認知症の疾患別ケアの考え方",                 "20分"),
    ("4", "アセスメントへの落とし込み",                 "15分"),
    ("5", "ケアプラン作成への落とし込み",               "15分"),
    ("6", "事例検討(認知症ケース)・グループ共有",     "10分"),
    ("7", "まとめ・質疑",                               " 5分"),
]
y = Cm(4.5)
for num, title, mins in agenda:
    # 番号丸
    circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(1.5), y, Cm(1.2), Cm(1.2))
    circle.fill.solid()
    circle.fill.fore_color.rgb = TEAL
    circle.line.fill.background()
    tf = circle.text_frame
    tf.margin_top = Cm(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = num
    r.font.name = FONT
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = WHITE

    add_text(s, Cm(3.2), y + Cm(0.15), Cm(20), Cm(1),
             title, size=18, bold=True, color=NAVY)
    add_text(s, SW - Cm(4), y + Cm(0.2), Cm(3), Cm(1),
             mins, size=14, color=ORANGE, bold=True, align=PP_ALIGN.RIGHT)
    y += Cm(1.6)


# =============================================================
#  SLIDE 4: 第1部 タイトル
# =============================================================
def section_title(num, jp, en):
    s = add_slide()
    add_rect(s, 0, 0, SW, SH, NAVY)
    add_rect(s, 0, Cm(8.5), SW, Cm(0.1), ORANGE)
    add_text(s, Cm(2), Cm(6.5), SW - Cm(4), Cm(1.5),
             f"第 {num} 部",
             size=22, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)
    add_text(s, Cm(1), Cm(9.5), SW - Cm(2), Cm(2),
             jp, size=40, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(s, Cm(1), Cm(13), SW - Cm(2), Cm(1),
             en, size=14, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
    return s


section_title(1, "適切なケアマネジメント手法とは", "What is the Appropriate Care Management Method?")


# =============================================================
#  SLIDE 5: 制度的背景
# =============================================================
s = add_slide()
add_header(s, "1. 適切なケアマネジメント手法とは", "なぜ今、適切なケアマネジメント手法か")
add_footer(s, 5)

add_box(s, Cm(1.2), Cm(4.5), Cm(15), Cm(6),
        "背景:制度・社会的要請",
        "・高齢化の進展と認知症高齢者の増加(2025年問題・2040年問題)\n"
        "・ケアマネジメントの質のばらつきへの指摘\n"
        "・「自立支援」「重度化防止」「尊厳の保持」の法的位置づけ\n"
        "  (介護保険法第1条・第2条)\n"
        "・科学的介護(LIFE)・データに基づくケアの推進\n"
        "・主任ケアマネへの「指導・育成」機能の期待",
        title_size=15, body_size=13)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6),
        "目的:手法が目指すもの",
        "・経験や勘に依存しない、科学的根拠に基づくケアマネジメント\n"
        "・要介護高齢者の生活の継続を支える共通言語\n"
        "・想定される支援内容を「仮説」として持ち、\n"
        "  アセスメントで検証する思考プロセスの定着\n"
        "・利用者像の見落とし・支援漏れの防止\n"
        "・多職種協働を促進する土台づくり",
        title_bg=ORANGE, body_bg=SOFT_ORANGE,
        title_size=15, body_size=13)

add_box(s, Cm(1.2), Cm(11), Cm(31.4), Cm(6),
        "ポイント",
        "■ 手法は「マニュアル」ではなく「思考の枠組み」\n"
        "■ 個別性を否定するものではない。むしろ標準を持つことで個別性が際立つ\n"
        "■ 主任ケアマネは、後輩ケアマネに「なぜその支援を入れたのか」を説明できる根拠として活用する\n"
        "■ R3年度から研修カリキュラム・法定研修にも段階的に組み込まれている",
        title_bg=NAVY, body_bg=ACCENT_BG,
        title_size=15, body_size=14)


# =============================================================
#  SLIDE 6: 全体構造
# =============================================================
s = add_slide()
add_header(s, "1. 適切なケアマネジメント手法とは", "適切なケアマネジメント手法の全体構造")
add_footer(s, 6)

# 上段:基本ケア
add_rect(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(0.9), TEAL)
add_text(s, Cm(1.5), Cm(4.4), Cm(31), Cm(0.8),
         "【基本ケア】 すべての高齢者に共通する支援の考え方",
         size=16, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, Cm(1.2), Cm(5.2), Cm(31.4), Cm(2.2), ACCENT_BG)
add_text(s, Cm(1.6), Cm(5.4), Cm(31), Cm(2),
         "①尊厳の保持と自立支援  ②全体像の把握(疾病・生活・家族・住環境)\n"
         "③意思決定の支援(本人の意向を中心に)  ④リスク管理(予防的視点)",
         size=14, color=DARK_GRAY)

# 下段:疾患別ケア
add_rect(s, Cm(1.2), Cm(8.0), Cm(31.4), Cm(0.9), ORANGE)
add_text(s, Cm(1.5), Cm(8.1), Cm(31), Cm(0.8),
         "【疾患別ケア】 主要7疾患群ごとの支援の考え方",
         size=16, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

# 疾患カード x 7
diseases = [
    ("脳血管疾患",       "再発予防・麻痺"),
    ("大腿骨頸部骨折",   "リハ・転倒予防"),
    ("心疾患(心不全)", "増悪予防・服薬"),
    ("認知症",           "進行・BPSD・家族"),
    ("誤嚥性肺炎",       "口腔・嚥下"),
    ("パーキンソン病",   "ON/OFF・転倒"),
    ("糖尿病",           "血糖・合併症"),
]
card_w = Cm(4.3)
card_h = Cm(3.3)
gap = Cm(0.18)
total_w = card_w * 7 + gap * 6
start_x = (SW - total_w) // 2
y = Cm(9.1)
for i, (name, kw) in enumerate(diseases):
    x = start_x + (card_w + gap) * i
    if name == "認知症":
        add_rect(s, x, y, card_w, card_h, ORANGE)
        add_text(s, x + Cm(0.2), y + Cm(0.4), card_w - Cm(0.4), Cm(1.5),
                 name, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(s, x + Cm(0.2), y + Cm(1.8), card_w - Cm(0.4), Cm(1.5),
                 kw, size=10, color=WHITE, align=PP_ALIGN.CENTER)
        # 「本日の中心」マーク
        add_text(s, x, y + card_h + Cm(0.1), card_w, Cm(0.6),
                 "★ 本日の中心",
                 size=10, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)
    else:
        add_rect(s, x, y, card_w, card_h, ACCENT_BG, line_color=TEAL)
        add_text(s, x + Cm(0.2), y + Cm(0.4), card_w - Cm(0.4), Cm(1.5),
                 name, size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_text(s, x + Cm(0.2), y + Cm(1.8), card_w - Cm(0.4), Cm(1.5),
                 kw, size=10, color=DARK_GRAY, align=PP_ALIGN.CENTER)

# 下部:統合メッセージ
add_text(s, Cm(1.5), Cm(15.5), SW - Cm(3), Cm(2),
         "▶ 実務では「基本ケア」を土台に、利用者の疾患に応じた「疾患別ケア」を重ねる。\n"
         "▶ どの疾患でも、基本ケアは必ず通底する。",
         size=14, color=NAVY, bold=True)


# =============================================================
#  SLIDE 7: 思考プロセス
# =============================================================
s = add_slide()
add_header(s, "1. 適切なケアマネジメント手法とは", "「想定される支援内容」を仮説として持つ")
add_footer(s, 7)

# フロー図
boxes = [
    ("①想定される支援内容\n(手引きから仮説立て)", TEAL),
    ("②アセスメントで\n必要性を判断", ORANGE),
    ("③ケアプランへ\n落とし込み", TEAL),
    ("④モニタリングで\n検証・修正", ORANGE),
]
box_w = Cm(6.5)
box_h = Cm(3.5)
gap_x = Cm(1.4)
total_w = box_w * 4 + gap_x * 3
start_x = (SW - total_w) // 2
y = Cm(5.5)
for i, (txt, color) in enumerate(boxes):
    x = start_x + (box_w + gap_x) * i
    add_rect(s, x, y, box_w, box_h, color)
    add_text(s, x + Cm(0.2), y + Cm(0.3), box_w - Cm(0.4), box_h - Cm(0.6),
             txt, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 矢印
    if i < 3:
        arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                   x + box_w + Cm(0.1), y + Cm(1.3),
                                   gap_x - Cm(0.2), Cm(0.9))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = NAVY
        arrow.line.fill.background()

# 補足
add_box(s, Cm(1.2), Cm(11), Cm(31.4), Cm(6),
        "主任ケアマネが押さえるべき発想の転換",
        "BEFORE  「アセスメントしてから、支援内容を考える」\n\n"
        "AFTER     「想定される支援内容を 仮説 として持って、\n"
        "                アセスメントで本人にとっての必要性を 検証 する」\n\n"
        "→ 漏れ・抜けを減らし、根拠を持って「入れる/入れない」を判断できるようになる",
        title_bg=NAVY, body_bg=SOFT_ORANGE, body_size=14)


# =============================================================
#  SLIDE 8: 第2部
# =============================================================
section_title(2, "基本ケアの考え方", "Foundation: The Four Pillars of Basic Care")


# =============================================================
#  SLIDE 9: 基本ケアの4つの柱
# =============================================================
s = add_slide()
add_header(s, "2. 基本ケアの考え方", "基本ケアの4つの柱(全体像)")
add_footer(s, 9)

pillars = [
    ("Ⅰ", "尊厳の保持と\n自立支援",
     "・本人の「したい」「ありたい」を起点に\n・できることを奪わない\n・役割・参加の機会の確保"),
    ("Ⅱ", "全体像の\n把握",
     "・疾病管理 / 生活全般 / 家族・介護者 / 住環境\n・既往と現状の整合を確認\n・潜在ニーズの可視化"),
    ("Ⅲ", "意思決定の\n支援",
     "・ACP・本人の価値観の確認\n・代理意思決定者の整理\n・段階に応じた情報提供"),
    ("Ⅳ", "リスクの\n予測と管理",
     "・服薬・転倒・誤嚥・栄養・脱水\n・感染症・予期せぬ入院\n・予防的アプローチ"),
]
col_w = Cm(7.7)
col_h = Cm(11.5)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
y = Cm(4.5)
for i, (rom, title, body) in enumerate(pillars):
    x = start_x + (col_w + gap) * i
    # header
    add_rect(s, x, y, col_w, Cm(2.8), NAVY)
    add_text(s, x, y + Cm(0.2), col_w, Cm(0.8),
             rom, size=20, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)
    add_text(s, x + Cm(0.2), y + Cm(1.1), col_w - Cm(0.4), Cm(1.6),
             title, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # body
    add_rect(s, x, y + Cm(2.8), col_w, col_h - Cm(2.8), ACCENT_BG)
    add_text(s, x + Cm(0.3), y + Cm(3.0), col_w - Cm(0.6), col_h - Cm(3),
             body, size=11, color=DARK_GRAY)

add_text(s, Cm(1.2), Cm(16.5), SW - Cm(2.4), Cm(1),
         "▶ 4つの柱は独立ではなく相互に作用する。実務では「Ⅱ全体像」を土台に、Ⅰ・Ⅲ・Ⅳを組み合わせる。",
         size=13, color=NAVY, bold=True)


# =============================================================
#  SLIDE 10: 柱Ⅰ 尊厳と自立支援
# =============================================================
s = add_slide()
add_header(s, "2. 基本ケアの考え方", "Ⅰ. 尊厳の保持と自立支援 ― 実務での落とし込み")
add_footer(s, 10)

add_box(s, Cm(1.2), Cm(4.5), Cm(15), Cm(6),
        "想定される支援内容(仮説)",
        "・本人の意向・好みの把握(食・整容・生活リズム)\n"
        "・残存機能を活かす活動の場の確保\n"
        "・「できる/しているADL」の差を埋める働きかけ\n"
        "・社会的役割・地域参加の機会の確保\n"
        "・過剰介助の回避(家族・サービス事業所へのアプローチ)",
        body_size=13)
add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6),
        "アセスメントで確認する項目",
        "・本人の希望(短期/長期) / 価値観 / 人生史\n"
        "・現在のADL・IADL(できる/している)\n"
        "・趣味・社会参加の頻度\n"
        "・家族・介護者の介護観(代行/見守りのバランス)\n"
        "・本人が「諦めていること」「我慢していること」",
        title_bg=ORANGE, body_bg=SOFT_ORANGE, body_size=13)

add_box(s, Cm(1.2), Cm(11), Cm(31.4), Cm(6),
        "主任ケアマネの視点 ― 後輩への問いかけ例",
        "■ 「このプラン、本人の言葉で目標が書けている?家族の希望にすり替わっていない?」\n"
        "■ 「この支援は『代わりにやる』になっていない?『一緒にやる』『見守る』に置き換えられない?」\n"
        "■ 「本人が、5年後も同じ生活を続けるための『止めない』ことは何だろう?」",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 11: 柱Ⅱ 全体像
# =============================================================
s = add_slide()
add_header(s, "2. 基本ケアの考え方", "Ⅱ. 全体像の把握 ― 4つの領域で漏れなく見る")
add_footer(s, 11)

# 中央に「本人」、周囲に4領域(2x2グリッド)
center_x = SW // 2
center_y = Cm(10.5)

# 本人(中央)
oval = s.shapes.add_shape(MSO_SHAPE.OVAL,
                          center_x - Cm(2), center_y - Cm(1.5),
                          Cm(4), Cm(3))
oval.fill.solid()
oval.fill.fore_color.rgb = ORANGE
oval.line.fill.background()
tf = oval.text_frame
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "本人"
r.font.name = FONT
r.font.size = Pt(20)
r.font.bold = True
r.font.color.rgb = WHITE

# 4領域
regions = [
    ("① 疾病管理",     "既往歴・現病歴・服薬\n医療連携・服薬コンプライアンス",     Cm(2),   Cm(5.5)),
    ("② 生活全般",     "ADL/IADL・栄養・睡眠\n生活リズム・趣味・社会参加",          Cm(21),  Cm(5.5)),
    ("③ 家族・介護者", "介護力・健康・就労\n介護観・キーパーソン・経済",              Cm(2),   Cm(13)),
    ("④ 住環境",       "段差・手すり・トイレ動線\n地域資源・近隣・交通",              Cm(21),  Cm(13)),
]
for title, body, x, y in regions:
    add_rect(s, x, y, Cm(10), Cm(4), ACCENT_BG, line_color=TEAL)
    add_rect(s, x, y, Cm(10), Cm(1), TEAL)
    add_text(s, x + Cm(0.2), y + Cm(0.1), Cm(9.6), Cm(0.8),
             title, size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Cm(0.3), y + Cm(1.2), Cm(9.4), Cm(2.6),
             body, size=12, color=DARK_GRAY)

add_text(s, Cm(1.2), Cm(17.3), SW - Cm(2.4), Cm(1),
         "▶ 4領域はチェックリストではなく「視点」。一度で完結せず、関わりの中で更新し続ける。",
         size=13, color=NAVY, bold=True)


# =============================================================
#  SLIDE 12: 柱Ⅲ 意思決定支援
# =============================================================
s = add_slide()
add_header(s, "2. 基本ケアの考え方", "Ⅲ. 意思決定支援 ― ACPと日常の選好")
add_footer(s, 12)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(6),
        "2つのレベルで考える",
        "【日常の意思決定】\n"
        "・食事・入浴・外出・服装などの選好\n"
        "・「選べる」場面の確保\n\n"
        "【人生の重要な意思決定】\n"
        "・在宅継続/施設入所/医療処置の選択\n"
        "・ACP(人生会議)・看取りの場所",
        body_size=13)
add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6),
        "認知症がある場合の留意点",
        "・能力を「ある/ない」の二分で判断しない\n"
        "・場面・タイミング・支援の有無で揺らぐ\n"
        "・本人の表情・反応も意思表示として汲み取る\n"
        "・家族の代理判断と本人意思を区別する\n"
        "・意思決定支援ガイドライン(厚労省)を活用",
        title_bg=ORANGE, body_bg=SOFT_ORANGE, body_size=13)

add_box(s, Cm(1.2), Cm(11), Cm(31.4), Cm(6),
        "ケアプランへの落とし込みポイント",
        "■ 第1表「利用者及び家族の生活に対する意向」に、本人と家族の意向を分けて記載\n"
        "■ 第2表のニーズに、本人が了解・納得しているプロセスを残す(支援経過の活用)\n"
        "■ 「意思決定支援」自体をサービス内容に明示する例:本人参加の担当者会議、診察同席",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 13: 柱Ⅳ リスク管理
# =============================================================
s = add_slide()
add_header(s, "2. 基本ケアの考え方", "Ⅳ. リスクの予測と管理 ― 予防的視点")
add_footer(s, 13)

# リスクマトリクス
risks = [
    ("服薬",   "残薬・重複処方・自己中断", "服薬一元化・お薬カレンダー"),
    ("転倒",   "下肢筋力低下・薬剤性ふらつき", "環境整備・リハ・履物"),
    ("誤嚥",   "食事中のむせ・体重減少",       "口腔ケア・食形態調整"),
    ("栄養",   "食事量低下・買い物困難",       "配食・栄養補助・受診"),
    ("脱水",   "夏季・利尿剤・口渇低下",       "水分摂取支援・室温管理"),
    ("感染症", "肺炎・尿路・皮膚",             "ワクチン・口腔・清潔保持"),
    ("家族",   "介護負担・虐待リスク",         "レスパイト・相談機関連携"),
    ("社会的", "孤立・経済困窮",               "地域包括・成年後見・生保"),
]
col_widths = [Cm(4.5), Cm(11), Cm(11)]
row_h = Cm(1.2)
start_x = Cm(3.4)
start_y = Cm(4.5)

# ヘッダー
headers = ["領域", "リスクの兆候", "想定される予防的支援"]
x = start_x
add_rect(s, start_x, start_y, sum(col_widths, Cm(0)), row_h, NAVY)
for i, h in enumerate(headers):
    add_text(s, x + Cm(0.2), start_y + Cm(0.15), col_widths[i] - Cm(0.4), Cm(0.9),
             h, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    x += col_widths[i]

# 行
for r_i, (area, sign, prev) in enumerate(risks):
    y = start_y + row_h * (r_i + 1)
    bg = ACCENT_BG if r_i % 2 == 0 else WHITE
    add_rect(s, start_x, y, sum(col_widths, Cm(0)), row_h, bg, line_color=LIGHT_GRAY)
    cells = [area, sign, prev]
    x = start_x
    for i, c in enumerate(cells):
        align = PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT
        col = NAVY if i == 0 else DARK_GRAY
        bold = (i == 0)
        add_text(s, x + Cm(0.2), y + Cm(0.15), col_widths[i] - Cm(0.4), Cm(0.9),
                 c, size=12, bold=bold, color=col,
                 align=align, anchor=MSO_ANCHOR.MIDDLE)
        x += col_widths[i]


# =============================================================
#  SLIDE 14: 第3部 タイトル
# =============================================================
section_title(3, "認知症の疾患別ケアの考え方",
              "Dementia-Specific Care: From Diagnosis to End-of-Life")


# =============================================================
#  SLIDE 15: 認知症ケアの全体像
# =============================================================
s = add_slide()
add_header(s, "3. 認知症の疾患別ケア", "認知症ケアの全体像 ― 4つの観点")
add_footer(s, 15)

points = [
    ("A", "病態と進行の理解",
     "・原因疾患の違いと特徴\n  (AD/血管性/レビー/FTD)\n・進行段階の見立て(FAST等)\n・併存疾患・身体合併症"),
    ("B", "BPSDへの対応",
     "・行動・心理症状の意味を読む\n・誘因の把握(身体・環境・心理)\n・非薬物的対応を優先\n・薬物療法との関係整理"),
    ("C", "意思決定支援",
     "・早期からの本人の意向把握\n・ACP・看取り場所の話し合い\n・成年後見・任意後見の検討\n・身寄りなしケースへの対応"),
    ("D", "家族介護者支援",
     "・介護負担・抑うつ\n・介護観のすり合わせ\n・就労との両立支援\n・家族の理解と教育"),
]
col_w = Cm(7.7)
col_h = Cm(11.5)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
y = Cm(4.5)
for i, (label, title, body) in enumerate(points):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(2.8), ORANGE)
    add_text(s, x, y + Cm(0.2), col_w, Cm(0.8),
             label, size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, x + Cm(0.2), y + Cm(1.1), col_w - Cm(0.4), Cm(1.6),
             title, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(2.8), col_w, col_h - Cm(2.8), SOFT_ORANGE)
    add_text(s, x + Cm(0.3), y + Cm(3.0), col_w - Cm(0.6), col_h - Cm(3),
             body, size=11, color=DARK_GRAY)

add_text(s, Cm(1.2), Cm(16.5), SW - Cm(2.4), Cm(1),
         "▶ 4観点は同時並行で考える。特に「B BPSDへの対応」は「A 進行」と「D 家族」とセットで見立てる。",
         size=13, color=NAVY, bold=True)


# =============================================================
#  SLIDE 16: 原因疾患の特徴
# =============================================================
s = add_slide()
add_header(s, "3. 認知症の疾患別ケア", "A. 原因疾患による特徴の違い")
add_footer(s, 16)

rows = [
    ("アルツハイマー型", "緩徐進行・近時記憶障害先行", "見当識低下・徘徊・物盗られ妄想"),
    ("血管性",           "段階的進行・基礎疾患合併", "感情失禁・意欲低下・運動麻痺"),
    ("レビー小体型",     "認知の変動・パーキンソン症状", "幻視・REM睡眠行動・転倒"),
    ("前頭側頭型",       "人格変化・常同行動が初期", "脱抑制・社会的逸脱・食行動異常"),
]
col_widths = [Cm(6.5), Cm(11), Cm(13)]
headers = ["原因疾患", "進行・特徴", "ケアで注意するポイント"]
start_x = (SW - sum(col_widths, Cm(0))) // 2
start_y = Cm(5)
row_h = Cm(1.6)

x = start_x
add_rect(s, start_x, start_y, sum(col_widths, Cm(0)), Cm(1.2), NAVY)
for i, h in enumerate(headers):
    add_text(s, x + Cm(0.2), start_y + Cm(0.15), col_widths[i] - Cm(0.4), Cm(0.9),
             h, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    x += col_widths[i]

for r_i, row in enumerate(rows):
    y = start_y + Cm(1.2) + row_h * r_i
    bg = ACCENT_BG if r_i % 2 == 0 else WHITE
    add_rect(s, start_x, y, sum(col_widths, Cm(0)), row_h, bg, line_color=LIGHT_GRAY)
    x = start_x
    for i, c in enumerate(row):
        align = PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT
        col = NAVY if i == 0 else DARK_GRAY
        bold = (i == 0)
        add_text(s, x + Cm(0.2), y + Cm(0.2), col_widths[i] - Cm(0.4), Cm(1.2),
                 c, size=12, bold=bold, color=col,
                 align=align, anchor=MSO_ANCHOR.MIDDLE)
        x += col_widths[i]

add_text(s, Cm(1.2), Cm(14.5), SW - Cm(2.4), Cm(2),
         "▶ 「認知症」とひとくくりにせず、原因疾患を主治医と共有することが、想定される支援内容の起点。\n"
         "▶ 例:レビー小体型でハロペリドール系投与は禁忌に近い ⇒ 服薬情報の共有はケアマネの役割。",
         size=14, color=NAVY, bold=True, line_spacing=1.3)


# =============================================================
#  SLIDE 17: 進行段階と支援
# =============================================================
s = add_slide()
add_header(s, "3. 認知症の疾患別ケア", "A. 進行段階別の「想定される支援内容」")
add_footer(s, 17)

stages = [
    ("MCI〜軽度",
     "・本人の自覚を尊重\n・運転・金銭管理の見直し\n・就労継続支援\n・社会参加の維持"),
    ("中等度",
     "・BPSD対応(中心課題)\n・服薬管理のサポート\n・通所サービス・レスパイト\n・徘徊・火気の事故予防"),
    ("重度",
     "・身体合併症の予防\n・誤嚥・脱水・褥瘡対策\n・コミュニケーション維持\n・在宅/施設の意思決定"),
    ("看取り期",
     "・本人の安楽・尊厳\n・看取り場所の選択\n・家族の予期悲嘆ケア\n・多職種看取りカンファ"),
]
col_w = Cm(7.7)
col_h = Cm(11.5)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
y = Cm(4.5)
colors = [TEAL, ORANGE, NAVY, RGBColor(0x6A, 0x4C, 0x93)]
for i, (label, body) in enumerate(stages):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1.6), colors[i])
    add_text(s, x + Cm(0.2), y + Cm(0.2), col_w - Cm(0.4), Cm(1.2),
             label, size=16, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1.6), col_w, col_h - Cm(1.6), ACCENT_BG)
    add_text(s, x + Cm(0.3), y + Cm(1.8), col_w - Cm(0.6), col_h - Cm(1.8),
             body, size=12, color=DARK_GRAY)

add_text(s, Cm(1.2), Cm(16.5), SW - Cm(2.4), Cm(1),
         "▶ 「今、本人はどの段階か」を多職種で共有することが、支援の過不足を防ぐ最大のポイント。",
         size=13, color=NAVY, bold=True)


# =============================================================
#  SLIDE 18: BPSD対応
# =============================================================
s = add_slide()
add_header(s, "3. 認知症の疾患別ケア", "B. BPSDの捉え方 ― 「行動」ではなく「意味」を見る")
add_footer(s, 18)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(6.5),
        "BPSDを引き起こす誘因(3つの層)",
        "【身体的要因】\n"
        "・疼痛・便秘・脱水・発熱・薬剤性・睡眠障害\n\n"
        "【環境的要因】\n"
        "・過剰な刺激・場所の変化・人間関係・不適切ケア\n\n"
        "【心理的要因】\n"
        "・不安・孤独・自尊心の傷つき・喪失体験",
        body_size=13)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6.5),
        "対応の優先順位",
        "① 非薬物的対応を最優先\n"
        "   ・誘因の除去、安心できる関わり\n"
        "   ・センター方式・ひもときシート活用\n\n"
        "② 主治医と相談の上での薬物療法\n"
        "   ・副作用(転倒・嚥下低下)に注意\n\n"
        "③ 家族・事業所への「対応の意味」共有\n"
        "   ・なぜその対応か、根拠を伝える",
        title_bg=ORANGE, body_bg=SOFT_ORANGE, body_size=13)

add_box(s, Cm(1.2), Cm(11.5), Cm(31.4), Cm(5.5),
        "主任ケアマネの腕の見せどころ",
        "■ 後輩から「対応に困っている」相談が来たら、まず「3つの誘因」を一緒に整理する習慣をつける\n"
        "■ 事業所内のケース会議で「困った行動」ではなく「困っている本人の状態」として言い換えさせる\n"
        "■ 短期間で改善しなくても、誘因仮説の検証プロセスを記録に残す(支援経過)",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 19: 意思決定支援(認知症)
# =============================================================
s = add_slide()
add_header(s, "3. 認知症の疾患別ケア", "C. 認知症と意思決定支援 ― ガイドラインの活用")
add_footer(s, 19)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(6),
        "厚労省ガイドラインの3つの原則",
        "① 本人の意思の尊重(意思決定能力を前提に)\n\n"
        "② 本人が自ら意思決定できるよう支援\n"
        "   ・支援を尽くしたか?(情報提示・環境調整)\n\n"
        "③ 本人の意思の推定・最善の利益\n"
        "   ・推定意思 → 最善の利益の検討の順",
        body_size=13)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6),
        "現場でのチェックポイント",
        "・本人にとってわかる言葉・速度で説明したか\n"
        "・体調がよい時間帯を選んだか\n"
        "・選択肢を絞りすぎていないか/多すぎないか\n"
        "・家族の意向と本人意思を混同していないか\n"
        "・「決められない」を「拒否」と捉えていないか",
        title_bg=ORANGE, body_bg=SOFT_ORANGE, body_size=13)

add_box(s, Cm(1.2), Cm(11), Cm(31.4), Cm(6),
        "ケアプランに残すべきプロセス",
        "■ 第1表 意向欄:「本人:〜と希望(○月○日 自宅にて確認)」と日時・場面を明記\n"
        "■ 担当者会議録:本人が同席した発言・うなずき・表情も記録\n"
        "■ 支援経過:意思の揺らぎ・変化、家族との合意プロセス\n"
        "■ ACPに関する話し合いは「結論」より「プロセス」を残す",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 20: 家族介護者支援
# =============================================================
s = add_slide()
add_header(s, "3. 認知症の疾患別ケア", "D. 家族介護者支援 ― 「もう一人の利用者」として")
add_footer(s, 20)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(6),
        "家族のアセスメント項目",
        "・主介護者の健康状態・年齢・就労状況\n"
        "・介護経験・介護観(代行型/見守り型)\n"
        "・家族内の役割分担・関係性\n"
        "・経済状況・住環境\n"
        "・抑うつ傾向・燃え尽き・虐待リスク\n"
        "・社会的孤立・支援者の有無",
        body_size=13)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6),
        "想定される支援内容",
        "・介護教室・家族会(ピアサポート)\n"
        "・レスパイト(ショートステイ・通所)\n"
        "・介護休業制度・両立支援の情報提供\n"
        "・経済的支援(医療費・介護費・年金)\n"
        "・本人の状態説明・進行の見通しの共有\n"
        "・看取り期に向けた予期的ガイダンス",
        title_bg=ORANGE, body_bg=SOFT_ORANGE, body_size=13)

add_box(s, Cm(1.2), Cm(11), Cm(31.4), Cm(6),
        "虐待・不適切ケアのサインを見逃さない",
        "■ 不自然なあざ・体重減少・脱水の繰り返し → 身体的虐待・ネグレクトの可能性\n"
        "■ 本人の前で家族が大声で叱責する/本人が萎縮 → 心理的虐待の可能性\n"
        "■ 通帳・年金を家族が管理し、本人に必要な物が買えていない → 経済的虐待の可能性\n"
        "■ 通報義務(高齢者虐待防止法)を主任ケアマネは正しく理解し、後輩を支える",
        title_bg=RGBColor(0xB0, 0x3A, 0x2E), body_bg=SOFT_ORANGE, body_size=14)


# =============================================================
#  SLIDE 21: 第4部 タイトル
# =============================================================
section_title(4, "アセスメントへの落とし込み",
              "Embedding the Framework into Assessment")


# =============================================================
#  SLIDE 22: 課題分析標準項目との関係
# =============================================================
s = add_slide()
add_header(s, "4. アセスメントへの落とし込み", "課題分析標準項目(23項目)との対応")
add_footer(s, 22)

add_text(s, Cm(1.2), Cm(4.3), SW - Cm(2.4), Cm(1),
         "「基本ケア+認知症ケア」の視点を、すでにある23項目のどこで拾うかを意識する。",
         size=14, color=NAVY, bold=True)

rows = [
    ("基本情報(1〜9)",   "①基本情報 ②生活状況 ③利用サービス ④障害高齢者 ⑤認知症高齢者\n⑥主訴 ⑦認定情報 ⑧課題分析理由 ⑨健康状態", "■ ⑤認知症高齢者の日常生活自立度 → 進行段階の起点\n■ ⑥主訴 → 本人/家族を分けて記載"),
    ("ADL/IADL(10〜11)","ADL・IADL", "■ 「できる/している」の差を必ず確認\n■ 認知症では声かけ・見守りの有無で大きく変動"),
    ("認知・コミュ(12〜14)","認知・コミュニケーション・社会との関わり", "■ HDS-R/MMSE等の参考値\n■ 表情・うなずき・拒否のサインも記録"),
    ("身体(15〜18)",       "排尿・排便・褥瘡・口腔衛生", "■ 失禁→BPSD誘因/口腔→誤嚥予防に直結"),
    ("生活全般(19〜23)",   "食事・問題行動・介護力・居住環境・特別な状況", "■ ⑳問題行動 → BPSDの誘因を一緒に記載\n■ ㉑介護力・㉒居住環境 → 家族・住環境(基本ケアⅡ)"),
]
col_widths = [Cm(6), Cm(13), Cm(12)]
headers = ["領域(項目番号)", "課題分析標準項目", "認知症ケア視点での着眼点"]
start_x = (SW - sum(col_widths, Cm(0))) // 2
start_y = Cm(5.7)

x = start_x
add_rect(s, start_x, start_y, sum(col_widths, Cm(0)), Cm(1.0), NAVY)
for i, h in enumerate(headers):
    add_text(s, x + Cm(0.2), start_y + Cm(0.1), col_widths[i] - Cm(0.4), Cm(0.8),
             h, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    x += col_widths[i]

y = start_y + Cm(1.0)
for r_i, row in enumerate(rows):
    row_h = Cm(2.2)
    bg = ACCENT_BG if r_i % 2 == 0 else WHITE
    add_rect(s, start_x, y, sum(col_widths, Cm(0)), row_h, bg, line_color=LIGHT_GRAY)
    x = start_x
    for i, c in enumerate(row):
        align = PP_ALIGN.LEFT
        col = NAVY if i == 0 else DARK_GRAY
        bold = (i == 0)
        size = 11 if i != 0 else 12
        add_text(s, x + Cm(0.2), y + Cm(0.15), col_widths[i] - Cm(0.4), row_h - Cm(0.3),
                 c, size=size, bold=bold, color=col, align=align,
                 anchor=MSO_ANCHOR.MIDDLE)
        x += col_widths[i]
    y += row_h


# =============================================================
#  SLIDE 23: ICFと統合
# =============================================================
s = add_slide()
add_header(s, "4. アセスメントへの落とし込み", "ICFの視点で「全体像」を見える化")
add_footer(s, 23)

# ICF図
center_x = SW // 2

# 健康状態(上)
add_rect(s, center_x - Cm(5), Cm(4.5), Cm(10), Cm(1.5), NAVY)
add_text(s, center_x - Cm(5), Cm(4.7), Cm(10), Cm(1.1),
         "健康状態(認知症 中等度・高血圧)",
         size=14, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 心身機能・身体構造 / 活動 / 参加(3列)
mid_y = Cm(7)
items = [
    ("心身機能・身体構造", "・近時記憶低下\n・見当識低下\n・下肢筋力低下\n・聴力低下"),
    ("活動",               "・更衣・整容 自立\n・調理は支援要\n・服薬 声かけ要\n・外出 同伴必要"),
    ("参加",               "・週1の通所\n・近所付き合い継続\n・畑作業(週2回)\n・自治会 不参加"),
]
col_w = Cm(9)
gap = Cm(1)
total_w = col_w * 3 + gap * 2
start_x = (SW - total_w) // 2
for i, (t, b) in enumerate(items):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, mid_y, col_w, Cm(1.0), TEAL)
    add_text(s, x + Cm(0.2), mid_y + Cm(0.1), col_w - Cm(0.4), Cm(0.8),
             t, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, mid_y + Cm(1.0), col_w, Cm(4), ACCENT_BG)
    add_text(s, x + Cm(0.3), mid_y + Cm(1.1), col_w - Cm(0.6), Cm(3.8),
             b, size=11, color=DARK_GRAY)

# 環境因子・個人因子(下2列)
bot_y = Cm(12.7)
items2 = [
    ("環境因子", "・長男夫婦と同居(共働き)\n・段差多い旧家屋\n・かかりつけ医近隣\n・地域に親しい友人"),
    ("個人因子", "・78歳 女性\n・元教員(現役感あり)\n・几帳面・社交的\n・「人に迷惑をかけたくない」"),
]
col_w2 = Cm(14)
gap2 = Cm(1)
total_w2 = col_w2 * 2 + gap2
start_x2 = (SW - total_w2) // 2
for i, (t, b) in enumerate(items2):
    x = start_x2 + (col_w2 + gap2) * i
    add_rect(s, x, bot_y, col_w2, Cm(1.0), ORANGE)
    add_text(s, x + Cm(0.2), bot_y + Cm(0.1), col_w2 - Cm(0.4), Cm(0.8),
             t, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, bot_y + Cm(1.0), col_w2, Cm(3.5), SOFT_ORANGE)
    add_text(s, x + Cm(0.3), bot_y + Cm(1.1), col_w2 - Cm(0.6), Cm(3.3),
             b, size=11, color=DARK_GRAY)


# =============================================================
#  SLIDE 24: 事例紹介
# =============================================================
s = add_slide()
add_header(s, "4. アセスメントへの落とし込み", "事例:Aさん(78歳・女性・要介護2)")
add_footer(s, 24)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(6.5),
        "プロフィール",
        "・78歳 女性 要介護2(認知症日常生活自立度 Ⅱb)\n"
        "・アルツハイマー型認知症(2年前 診断)\n"
        "・既往:高血圧・変形性膝関節症\n"
        "・長男夫婦と同居(長男夫婦は共働き)\n"
        "・元小学校教員、几帳面で社交的\n"
        "・趣味は家庭菜園(週2回 30分程度)\n"
        "・希望:「できるだけ自宅で、迷惑をかけずに暮らしたい」",
        body_size=13)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6.5),
        "現在の状況・家族の困りごと",
        "・調理時に火を消し忘れることが2回(直近1ヶ月)\n"
        "・「お金が無い」と通帳を何度も確認\n"
        "・夕方に落ち着かなくなり、外出しようとする\n"
        "・服薬は嫁が声かけしているが、拒否することも\n\n"
        "【嫁の訴え】\n"
        "「目が離せず仕事を辞めるか悩んでいる。\n"
        " このまま在宅介護を続けられるか不安」",
        title_bg=ORANGE, body_bg=SOFT_ORANGE, body_size=13)

add_box(s, Cm(1.2), Cm(11.5), Cm(31.4), Cm(5.5),
        "想定される支援内容(仮説立て)",
        "■ 基本ケア:残存機能の活用(菜園・家事参加)、家族の介護負担軽減、火気・服薬のリスク管理\n"
        "■ 認知症ケア:夕方症候群への対応(誘因仮説:疲労・不安)、本人意思の確認(早期から)、嫁の支援\n"
        "▶ 次のスライドで、これら仮説をアセスメントで「検証」する",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 25: アセスメント:見立てと根拠
# =============================================================
s = add_slide()
add_header(s, "4. アセスメントへの落とし込み", "Aさんのアセスメント ― 見立てと根拠")
add_footer(s, 25)

rows = [
    ("尊厳・自立",
     "教員としての自負・几帳面さ。\n菜園を続けたい意向。",
     "「迷惑をかけたくない」を尊重しつつ、\n役割を奪わない関わりが必要。"),
    ("全体像",
     "AD中等度・高血圧・膝OA。\n長男夫婦共働きで日中独居。\n旧家屋で段差多い。",
     "日中独居時間のリスク管理。\n服薬は嫁の出勤後〜帰宅前の空白。"),
    ("意思決定",
     "現時点では本人の意思表示可能。\nACPは未実施。",
     "比較的早期にACP着手すべき段階。\n本人主体で進行を見通した話し合い。"),
    ("リスク",
     "火気の消し忘れ・服薬拒否・\n夕方の不穏・転倒(膝OA)。",
     "BPSD(夕方症候群)の誘因仮説:\n疲労+夕方の家族不在による不安。"),
    ("家族",
     "嫁の介護負担増・離職検討。\n夫婦間で介護分担曖昧。",
     "嫁=もう一人の利用者として支援。\nレスパイトと役割分担調整が急務。"),
]
col_widths = [Cm(5), Cm(13), Cm(13)]
headers = ["観点", "アセスメント(事実)", "見立て・解釈"]
start_x = (SW - sum(col_widths, Cm(0))) // 2
start_y = Cm(4.5)

x = start_x
add_rect(s, start_x, start_y, sum(col_widths, Cm(0)), Cm(1.0), NAVY)
for i, h in enumerate(headers):
    add_text(s, x + Cm(0.2), start_y + Cm(0.1), col_widths[i] - Cm(0.4), Cm(0.8),
             h, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    x += col_widths[i]

y = start_y + Cm(1.0)
for r_i, row in enumerate(rows):
    row_h = Cm(2.2)
    bg = ACCENT_BG if r_i % 2 == 0 else WHITE
    add_rect(s, start_x, y, sum(col_widths, Cm(0)), row_h, bg, line_color=LIGHT_GRAY)
    x = start_x
    for i, c in enumerate(row):
        align = PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT
        col = NAVY if i == 0 else DARK_GRAY
        bold = (i == 0)
        size = 12 if i != 0 else 13
        add_text(s, x + Cm(0.2), y + Cm(0.15), col_widths[i] - Cm(0.4), row_h - Cm(0.3),
                 c, size=size, bold=bold, color=col, align=align,
                 anchor=MSO_ANCHOR.MIDDLE)
        x += col_widths[i]
    y += row_h


# =============================================================
#  SLIDE 26: 第5部 タイトル
# =============================================================
section_title(5, "ケアプラン作成への落とし込み",
              "From Assessment to Care Plan")


# =============================================================
#  SLIDE 27: ケアプランへの転換
# =============================================================
s = add_slide()
add_header(s, "5. ケアプランへの落とし込み", "アセスメント → ニーズ → 目標 → 支援内容")
add_footer(s, 27)

# 流れ図
boxes = [
    ("アセスメント\n(事実と解釈)", TEAL),
    ("ニーズ\n(本人の言葉で)", ORANGE),
    ("長期/短期目標\n(評価可能な形で)", TEAL),
    ("支援内容\n(誰が何をいつ)", ORANGE),
]
box_w = Cm(6.5)
box_h = Cm(3.0)
gap_x = Cm(1.4)
total_w = box_w * 4 + gap_x * 3
start_x = (SW - total_w) // 2
y = Cm(4.5)
for i, (txt, color) in enumerate(boxes):
    x = start_x + (box_w + gap_x) * i
    add_rect(s, x, y, box_w, box_h, color)
    add_text(s, x + Cm(0.2), y + Cm(0.3), box_w - Cm(0.4), box_h - Cm(0.6),
             txt, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if i < 3:
        arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                   x + box_w + Cm(0.1), y + Cm(1.0),
                                   gap_x - Cm(0.2), Cm(0.9))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = NAVY
        arrow.line.fill.background()

add_box(s, Cm(1.2), Cm(9), Cm(31.4), Cm(8),
        "落とし込みのポイント",
        "■ ニーズは「〜したい」「〜できるようになりたい」と本人の言葉で書く\n"
        "    × 「認知症のため見守りが必要」 → 〇 「自分で食事の準備を続けたい」\n\n"
        "■ 目標は「誰が」「何を」「いつまでに」「どの程度」を具体的に\n"
        "    × 「ADLを維持する」 → 〇 「6ヶ月後も、台所で野菜を洗う作業を継続できる」\n\n"
        "■ 支援内容は「想定される支援内容(基本ケア+認知症ケア)」を再点検し、漏れがないか確認\n\n"
        "■ 「入れない」と判断したサービスも、なぜ入れないかを支援経過に残す(主任CMの腕の見せどころ)",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=13)


# =============================================================
#  SLIDE 28: Aさん 第1表
# =============================================================
s = add_slide()
add_header(s, "5. ケアプランへの落とし込み", "Aさんのケアプラン:第1表(抜粋)")
add_footer(s, 28)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(3.5),
        "利用者及び家族の生活に対する意向",
        "本人:「家族に迷惑をかけたくないけれど、できるだけ自宅で、自分でできることは続けたい。\n"
        "        畑の野菜を作って、孫に食べさせたい。」(○月○日 自宅にて確認)\n\n"
        "家族(嫁):「日中の安全と、私の仕事の継続を両立させたい。本人の意向は尊重したい。」",
        title_bg=TEAL, body_size=12)

add_box(s, Cm(1.2), Cm(8.2), Cm(15.5), Cm(4),
        "総合的な援助の方針",
        "■ 残存機能を活かし、本人が「役割」を持ち続けられる生活を支える\n"
        "■ 家族の介護負担を軽減し、在宅生活継続を可能とする\n"
        "■ 進行を見通したACPを早期から始め、本人の意思を尊重する",
        title_bg=NAVY, body_size=12)

add_box(s, Cm(17.5), Cm(8.2), Cm(15), Cm(4),
        "認定情報・課題分析の理由",
        "・要介護2(R6.○月認定)\n"
        "・主治医意見書:AD進行、抗認知症薬服用中\n"
        "・更新申請に伴う再アセスメント、夕方症候群\n"
        "  および家族介護負担の顕在化により再評価",
        title_bg=ORANGE, body_bg=SOFT_ORANGE, body_size=12)

add_box(s, Cm(1.2), Cm(12.5), Cm(31.4), Cm(4.5),
        "第1表 作成のポイント(主任CMチェック)",
        "■ 本人と家族の意向を分けて記載できているか/本人の言葉に近づいているか\n"
        "■ 援助方針に「基本ケアの4つの柱」が反映されているか\n"
        "■ 認知症の進行を見通した内容になっているか(数年単位)\n"
        "■ 家族介護者支援が方針に明示されているか",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=13)


# =============================================================
#  SLIDE 29: Aさん 第2表
# =============================================================
s = add_slide()
add_header(s, "5. ケアプランへの落とし込み", "Aさんのケアプラン:第2表(抜粋)")
add_footer(s, 29)

needs_data = [
    ("生活全般の解決すべき課題(ニーズ)", "長期目標", "短期目標", "サービス内容"),
    ("自分で家事の一部を続けながら、\n畑作業を楽しみたい",
     "野菜を育て、収穫物を孫に\n渡せる(12ヶ月)",
     "週2回、台所で野菜を洗う\n作業を続けられる(3ヶ月)",
     "通所介護(週2)/家族による\n声かけ・見守り/作業療法士助言"),
    ("夕方も落ち着いて過ごし、\n家族との時間を持ちたい",
     "夕方の不穏が週1回以下に\nなり、家族の負担軽減(12ヶ月)",
     "夕方の活動メニューが定着\nし、不穏の頻度が減る(3ヶ月)",
     "通所介護(17時まで延長)/\n家族への対応方法の助言"),
    ("安全に薬を飲み、\n体調を崩さない",
     "服薬を継続し、入院せず\n在宅生活を続ける(12ヶ月)",
     "毎日の服薬が確実に行わ\nれる(3ヶ月)",
     "訪問看護(週1)/服薬カレンダー\n/主治医・薬剤師と連携"),
    ("嫁が安心して仕事を続け、\n介護を続けられる",
     "嫁の就労継続、介護負担感\nの軽減(12ヶ月)",
     "月1回のレスパイト利用\n家族会に参加(3ヶ月)",
     "短期入所(月1回 2泊)\n地域包括/家族会の紹介"),
]
col_widths = [Cm(8), Cm(7.5), Cm(7.5), Cm(8.4)]
start_x = (SW - sum(col_widths, Cm(0))) // 2
start_y = Cm(4.5)

# ヘッダー
add_rect(s, start_x, start_y, sum(col_widths, Cm(0)), Cm(1.2), NAVY)
x = start_x
for i, h in enumerate(needs_data[0]):
    add_text(s, x + Cm(0.1), start_y + Cm(0.1), col_widths[i] - Cm(0.2), Cm(1),
             h, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    x += col_widths[i]

y = start_y + Cm(1.2)
for r_i, row in enumerate(needs_data[1:]):
    row_h = Cm(2.6)
    bg = ACCENT_BG if r_i % 2 == 0 else WHITE
    add_rect(s, start_x, y, sum(col_widths, Cm(0)), row_h, bg, line_color=LIGHT_GRAY)
    x = start_x
    for i, c in enumerate(row):
        add_text(s, x + Cm(0.15), y + Cm(0.2), col_widths[i] - Cm(0.3), row_h - Cm(0.4),
                 c, size=10.5, color=DARK_GRAY,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        x += col_widths[i]
    y += row_h

add_text(s, Cm(1.2), y + Cm(0.2), SW - Cm(2.4), Cm(1),
         "▶ ニーズは本人の言葉で。目標は評価可能に。サービス内容は基本ケア+認知症ケアの両輪を意識。",
         size=12, color=NAVY, bold=True)


# =============================================================
#  SLIDE 30: モニタリングと再アセスメント
# =============================================================
s = add_slide()
add_header(s, "5. ケアプランへの落とし込み", "モニタリングで「仮説」を検証し続ける")
add_footer(s, 30)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(6.5),
        "認知症ケースのモニタリング着眼点",
        "・進行段階の変化(FAST・自立度の見直し)\n"
        "・BPSDの出現/消失と誘因の変化\n"
        "・服薬状況・身体合併症の有無\n"
        "・本人の表情・発語の質的変化\n"
        "・家族の負担感(介護負担尺度の活用)\n"
        "・ACPの内容に変化はないか",
        body_size=13)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6.5),
        "再アセスメントのタイミング",
        "・退院後・入院後\n"
        "・BPSDの新規出現/増悪\n"
        "・家族の介護環境の変化\n"
        "  (主介護者の体調/就労変化)\n"
        "・住環境の変化(転居・改修)\n"
        "・本人の意思表示が変わったとき",
        title_bg=ORANGE, body_bg=SOFT_ORANGE, body_size=13)

add_box(s, Cm(1.2), Cm(11.5), Cm(31.4), Cm(5.5),
        "主任ケアマネへの問いかけ",
        "■ あなたの担当ケースで、ケアプランは「3ヶ月前」と「今」で内容が動いていますか?\n"
        "■ サービスを増やすだけでなく、「卒業」「変更」「終了」も検討できていますか?\n"
        "■ モニタリング記録に、「想定支援内容のうち提供しなかったもの」の根拠は残っていますか?",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 31: 第6部 タイトル
# =============================================================
section_title(6, "事例検討・グループワーク",
              "Group Work: Case Discussion")


# =============================================================
#  SLIDE 32: グループワーク課題
# =============================================================
s = add_slide()
add_header(s, "6. 事例検討", "グループワーク(10分)")
add_footer(s, 32)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(5),
        "事例:Bさん(82歳・男性・要介護3)",
        "・レビー小体型認知症(3年前 診断)、パーキンソン症状あり\n"
        "・妻(80歳)と二人暮らし、長女が車で30分の所在\n"
        "・幻視を訴えることがあり、夜間に大声を出す\n"
        "・最近、食事中のむせが増え、3ヶ月で2kg体重減\n"
        "・「家で死にたい」と繰り返し発言、妻も同意\n"
        "・妻が腰痛悪化、夜間の介護が辛い",
        title_bg=ORANGE, body_size=13)

add_box(s, Cm(1.2), Cm(9.7), Cm(31.4), Cm(7.3),
        "ワーク内容(3〜4人組)",
        "Q1. Bさんの「基本ケア4つの柱」の視点で、想定される支援内容を挙げてみよう(3分)\n\n"
        "Q2. レビー小体型認知症ならではの「疾患別ケア」の留意点は何か?\n"
        "       (薬剤・転倒・嚥下・幻視への対応 など)(3分)\n\n"
        "Q3. このケースのケアプラン第1表「総合的な援助の方針」をグループで1文書いてみよう(3分)\n\n"
        "Q4. 共有・全体発表(1分×グループ)",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=13)


# =============================================================
#  SLIDE 33: ワークシート(参考)
# =============================================================
s = add_slide()
add_header(s, "6. 事例検討", "ワークシート ― 整理のフレーム")
add_footer(s, 33)

# 4つのワークボックス
items = [
    ("基本ケアⅠ 尊厳・自立",   "本人の意向・残存機能・役割は?"),
    ("基本ケアⅡ 全体像",       "疾病・生活・家族・住環境のリスクは?"),
    ("基本ケアⅢ 意思決定",     "「家で死にたい」をどう支える?ACPは?"),
    ("基本ケアⅣ リスク管理",   "嚥下・転倒・薬剤・夜間対応・家族の健康"),
    ("認知症ケアA 病態理解",   "レビー小体型ならではの特徴は?"),
    ("認知症ケアB BPSD",       "幻視・夜間の大声、誘因仮説は?"),
    ("認知症ケアC 意思決定",   "本人意思の確認方法・ACPの進め方"),
    ("認知症ケアD 家族支援",   "妻の腰痛・夜間負担・長女との連携"),
]
col_w = Cm(15.5)
row_h = Cm(2.5)
gap_x = Cm(0.4)
gap_y = Cm(0.3)
start_x = Cm(1.2)
start_y = Cm(4.5)
for i, (t, q) in enumerate(items):
    col = i % 2
    row = i // 2
    x = start_x + (col_w + gap_x) * col
    y = start_y + (row_h + gap_y) * row
    color = TEAL if i < 4 else ORANGE
    add_rect(s, x, y, col_w, Cm(0.9), color)
    add_text(s, x + Cm(0.2), y + Cm(0.05), col_w - Cm(0.4), Cm(0.8),
             t, size=12, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(0.9), col_w, row_h - Cm(0.9), ACCENT_BG)
    add_text(s, x + Cm(0.3), y + Cm(1.1), col_w - Cm(0.6), row_h - Cm(1.2),
             q, size=11, color=DARK_GRAY)


# =============================================================
#  SLIDE 34: まとめ1
# =============================================================
s = add_slide()
add_header(s, "Summary", "本日の学びの整理(1/2)")
add_footer(s, 34)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(12.5),
        "「適切なケアマネジメント手法」の本質",
        "■ マニュアルではなく、思考の枠組み\n\n"
        "■ 「想定される支援内容」を仮説として持ち、\n"
        "    アセスメントで本人にとっての必要性を検証する\n\n"
        "■ 基本ケアは すべての利用者に共通する土台\n\n"
        "■ 疾患別ケアは 基本ケアの上に重ねるもの\n\n"
        "■ 個別性を否定するものではない\n"
        "    むしろ標準があるからこそ「この人らしさ」が際立つ\n\n"
        "■ 説明責任(なぜそのケアか)を果たす道具",
        body_size=14)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(12.5),
        "認知症ケアの実践ポイント",
        "■ 原因疾患による違いを押さえる\n   (AD/血管性/レビー/前頭側頭)\n\n"
        "■ 進行段階に応じた支援を見立てる\n   (MCI〜看取りまで連続したケア)\n\n"
        "■ BPSDは「行動」ではなく「意味」を見る\n   (身体・環境・心理の3層で誘因把握)\n\n"
        "■ 意思決定支援は早期から、プロセスを残す\n\n"
        "■ 家族介護者を「もう一人の利用者」として\n   アセスメント・支援する",
        title_bg=ORANGE, body_bg=SOFT_ORANGE, body_size=14)


# =============================================================
#  SLIDE 35: まとめ2 主任CMへ
# =============================================================
s = add_slide()
add_header(s, "Summary", "本日の学びの整理(2/2) ― 主任CMへの期待")
add_footer(s, 35)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(12.5),
        "主任ケアマネジャーが事業所・地域で果たす役割",
        "\n"
        "【①  個別ケースの質を担保する】\n"
        "  ・後輩CMのアセスメント・ケアプランを「基本ケア+認知症ケア」の視点でレビューする\n"
        "  ・「なぜその支援か」「なぜその支援を入れないか」を一緒に言語化する\n\n"
        "【②  事業所の文化として根付かせる】\n"
        "  ・事例検討会で「想定される支援内容(仮説)→ 検証」の構造を共通言語にする\n"
        "  ・記録(支援経過)に「検証プロセス」が残る習慣をつくる\n\n"
        "【③  地域の質を底上げする】\n"
        "  ・地域ケア会議・主任CM連絡会で他事業所と知見を共有する\n"
        "  ・新人〜中堅CMの相談先となる(「一人で抱えさせない」)\n\n"
        "【④  自分自身も学び続ける】\n"
        "  ・手引き・ガイドライン・最新の研究知見を取りに行く姿勢を持ち続ける",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 36: 参考資料
# =============================================================
s = add_slide()
add_header(s, "Reference", "参考資料・関連ガイドライン")
add_footer(s, 36)

refs = [
    "■ 厚生労働省「適切なケアマネジメント手法の手引き」",
    "■ 厚生労働省「認知症の人の日常生活・社会生活における意思決定支援ガイドライン」",
    "■ 厚生労働省「人生の最終段階における医療・ケアの決定プロセスに関するガイドライン」",
    "■ 厚生労働省「高齢者虐待防止の手引き」",
    "■ 日本ケアマネジメント学会「課題分析標準項目に関する解説」",
    "■ 認知症介護研究・研修センター「センター方式」「ひもときシート」",
    "■ 国際生活機能分類(ICF) ― WHO",
    "■ 介護保険最新情報(課題分析標準項目 R5改訂)",
]
y = Cm(4.5)
for r in refs:
    add_text(s, Cm(2), y, SW - Cm(4), Cm(1),
             r, size=15, color=DARK_GRAY)
    y += Cm(1.3)


# =============================================================
#  SLIDE 37: クロージング
# =============================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, NAVY)
add_rect(s, 0, Cm(8), SW, Cm(0.15), ORANGE)

add_text(s, Cm(2), Cm(5), SW - Cm(4), Cm(2),
         "Thank you",
         size=54, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

add_text(s, Cm(2), Cm(9.5), SW - Cm(4), Cm(3),
         "「適切なケアマネジメント手法」は、\n"
         "私たちが利用者の人生に責任を持つための共通言語です。",
         size=18, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

add_text(s, Cm(2), Cm(13.5), SW - Cm(4), Cm(2),
         "まずは「次の1ケース」から、\n仮説を持って関わってみましょう。",
         size=20, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)

add_text(s, Cm(2), Cm(17), SW - Cm(4), Cm(1),
         "ご質問・ご意見をお願いいたします",
         size=14, color=WHITE, align=PP_ALIGN.CENTER)


# ===== 保存 =====
output = "/home/user/kaigo-ga-app/slides/適切なケアマネジメント手法_認知症ケア_研修スライド.pptx"
prs.save(output)
print(f"Saved: {output}")
print(f"Total slides: {len(prs.slides)}")
