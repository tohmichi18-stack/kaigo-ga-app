"""主任ケアマネ研修スライド生成スクリプト 第6弾

テーマ: 身体拘束最小化と高齢者虐待防止
重点: 居宅サービス事業所への指導・助言の視点 / 不適切ケアの早期発見と対応
トーン: 実務型(視点・手順・場面対応)
時間: 90分
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ===== カラーパレット(尊厳・人権テーマ:ワイン基調) =====
WINE        = RGBColor(0x5E, 0x2A, 0x3C)   # メイン
WINE_MID    = RGBColor(0x84, 0x44, 0x57)
WINE_LT     = RGBColor(0xA8, 0x73, 0x80)
GOLD        = RGBColor(0xBE, 0x8B, 0x3C)   # アクセント
SAGE        = RGBColor(0x4F, 0x7A, 0x5E)   # 適切なケア・OK
TEAL        = RGBColor(0x2E, 0x6E, 0x73)   # 連携・協働
AMBER       = RGBColor(0xCF, 0x96, 0x2E)   # 注意・グレーゾーン
RED         = RGBColor(0xB5, 0x3A, 0x33)   # 虐待・NG
LIGHT_BG    = RGBColor(0xF4, 0xF1, 0xF2)
MID_GRAY    = RGBColor(0xDC, 0xD4, 0xD7)
DARK_GRAY   = RGBColor(0x33, 0x2E, 0x30)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_BG    = RGBColor(0xEE, 0xE7, 0xE9)
RED_BG      = RGBColor(0xF8, 0xE5, 0xE3)
AMBER_BG    = RGBColor(0xFA, 0xF0, 0xDA)
SAGE_BG     = RGBColor(0xE5, 0xED, 0xE7)
TEAL_BG     = RGBColor(0xDF, 0xEA, 0xEA)

FONT = "Yu Gothic"

prs = Presentation()
prs.slide_width = Cm(33.867)
prs.slide_height = Cm(19.05)
SW = prs.slide_width
SH = prs.slide_height
blank_layout = prs.slide_layouts[6]


# ===== ヘルパー =====
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
             line_spacing=None, italic=False):
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
        run.font.italic = italic
        run.font.color.rgb = color
    return box


def add_header(slide, section_label, title):
    add_rect(slide, 0, 0, SW, Cm(1.6), WINE)
    add_text(slide, Cm(0.8), Cm(0.25), Cm(22), Cm(0.6),
             section_label, size=11, color=WHITE, bold=True)
    add_text(slide, Cm(0.8), Cm(1.9), SW - Cm(1.6), Cm(1.3),
             title, size=23, bold=True, color=WINE)
    add_rect(slide, Cm(0.8), Cm(3.2), Cm(2.5), Cm(0.12), GOLD)


def add_footer(slide, page_num, total=34):
    add_text(slide, Cm(0.8), SH - Cm(0.8), Cm(26), Cm(0.5),
             "主任CM研修 / 身体拘束最小化と高齢者虐待防止",
             size=9, color=WINE_LT)
    add_text(slide, SW - Cm(3), SH - Cm(0.8), Cm(2.2), Cm(0.5),
             f"{page_num} / {total}", size=9, color=WINE_LT, align=PP_ALIGN.RIGHT)


def add_box(slide, left, top, width, height, title, body, *,
            title_bg=WINE_MID, body_bg=TABLE_BG, title_color=WHITE,
            body_color=DARK_GRAY, title_size=14, body_size=12,
            body_line_spacing=None):
    add_rect(slide, left, top, width, Cm(1.0), title_bg)
    add_text(slide, left + Cm(0.2), top + Cm(0.1), width - Cm(0.4), Cm(0.8),
             title, size=title_size, bold=True, color=title_color,
             anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, left, top + Cm(1.0), width, height - Cm(1.0), body_bg)
    add_text(slide, left + Cm(0.3), top + Cm(1.1), width - Cm(0.6),
             height - Cm(1.1), body, size=body_size, color=body_color,
             line_spacing=body_line_spacing)


def add_table(slide, left, top, col_widths, headers, rows, *,
              header_bg=WINE, header_color=WHITE,
              row_bg_a=TABLE_BG, row_bg_b=WHITE,
              header_size=12, body_size=11, row_height=Cm(1.4),
              first_col_color=None, first_col_bold=True):
    total_w = sum(col_widths, Cm(0))
    add_rect(slide, left, top, total_w, Cm(1.0), header_bg)
    x = left
    for i, h in enumerate(headers):
        add_text(slide, x + Cm(0.15), top + Cm(0.1), col_widths[i] - Cm(0.3), Cm(0.8),
                 h, size=header_size, bold=True, color=header_color,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        x += col_widths[i]

    y = top + Cm(1.0)
    for r_i, row in enumerate(rows):
        bg = row_bg_a if r_i % 2 == 0 else row_bg_b
        add_rect(slide, left, y, total_w, row_height, bg, line_color=MID_GRAY)
        x = left
        for i, c in enumerate(row):
            color = first_col_color if (i == 0 and first_col_color) else DARK_GRAY
            bold = first_col_bold and (i == 0)
            align = PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT
            add_text(slide, x + Cm(0.15), y + Cm(0.15),
                     col_widths[i] - Cm(0.3), row_height - Cm(0.3),
                     c, size=body_size, bold=bold, color=color,
                     align=align, anchor=MSO_ANCHOR.MIDDLE)
            x += col_widths[i]
        y += row_height
    return y


def section_title(num, jp, en, *, accent_color=GOLD):
    s = add_slide()
    add_rect(s, 0, 0, SW, SH, WINE)
    add_rect(s, 0, Cm(8.5), SW, Cm(0.1), accent_color)
    add_text(s, Cm(2), Cm(6.5), SW - Cm(4), Cm(1.5),
             f"PART {num}", size=22, color=accent_color, bold=True,
             align=PP_ALIGN.CENTER)
    add_text(s, Cm(1), Cm(9.5), SW - Cm(2), Cm(2),
             jp, size=33, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(s, Cm(1), Cm(13), SW - Cm(2), Cm(1),
             en, size=14, color=MID_GRAY, align=PP_ALIGN.CENTER)
    return s


def add_numbered_circle(slide, left, top, diameter, num, color, fontsize=15):
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, diameter, diameter)
    circle.fill.solid()
    circle.fill.fore_color.rgb = color
    circle.line.fill.background()
    circle.shadow.inherit = False
    tf = circle.text_frame
    tf.margin_top = Cm(0)
    tf.margin_bottom = Cm(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(num)
    r.font.name = FONT
    r.font.size = Pt(fontsize)
    r.font.bold = True
    r.font.color.rgb = WHITE
    return circle


# =============================================================
#  SLIDE 1: 表紙
# =============================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, WINE)
add_rect(s, 0, Cm(8), SW, Cm(0.15), GOLD)
add_rect(s, 0, Cm(12.3), SW, Cm(0.05), WINE_LT)

add_text(s, Cm(2), Cm(4.0), SW - Cm(4), Cm(1),
         "主任介護支援専門員 研修会",
         size=18, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
add_text(s, Cm(1), Cm(5.4), SW - Cm(2), Cm(2.5),
         "身体拘束最小化と\n高齢者虐待防止",
         size=33, color=WHITE, bold=True, align=PP_ALIGN.CENTER, line_spacing=1.25)
add_text(s, Cm(1.5), Cm(9.7), SW - Cm(3), Cm(2),
         "ー 不適切ケアを早期に発見し、利用者の尊厳を守る ー",
         size=16, color=WHITE, align=PP_ALIGN.CENTER)

add_text(s, Cm(1.5), Cm(13), SW - Cm(3), Cm(2),
         "■ 身体拘束・高齢者虐待をめぐる制度と基本知識\n"
         "■ 不適切ケアの早期発見 ― グレーゾーンへの感度\n"
         "■ 居宅サービス事業所への指導・助言の視点と対応",
         size=14, color=AMBER, align=PP_ALIGN.CENTER, line_spacing=1.5)

add_text(s, Cm(2), Cm(16.3), SW - Cm(4), Cm(1),
         "主任ケアマネジャーの会  主催  ・  研修時間 90 分",
         size=14, color=MID_GRAY, align=PP_ALIGN.CENTER)


# =============================================================
#  SLIDE 2: 本日の目標
# =============================================================
s = add_slide()
add_header(s, "Introduction", "本日の研修目標")
add_footer(s, 2)

goals = [
    "身体拘束最小化・高齢者虐待防止をめぐる制度的背景と基本知識を整理する",
    "身体拘束の定義・弊害・「緊急やむを得ない場合」の3要件を正しく理解する",
    "高齢者虐待の5類型と、不適切ケアとの「連続性(グレーゾーン)」を理解する",
    "不適切ケアのサインを、利用者・環境・職員・記録から早期に発見する視点を持つ",
    "居宅サービス事業所への指導・助言の「立場」と「やり方」を学ぶ",
    "気づいたときの対応・通報・連携の流れを、後輩CMにも伝えられるようになる",
]
y = Cm(4.5)
for g in goals:
    add_rect(s, Cm(1.2), y, Cm(0.5), Cm(1.5), GOLD)
    add_text(s, Cm(2), y + Cm(0.15), SW - Cm(3.5), Cm(1.2),
             g, size=15, color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.3)
    y += Cm(1.8)


# =============================================================
#  SLIDE 3: アジェンダ
# =============================================================
s = add_slide()
add_header(s, "Agenda", "本日の流れ(90分)")
add_footer(s, 3)

agenda = [
    ("1", "なぜ今、拘束最小化と虐待防止か",          "10分"),
    ("2", "身体拘束の理解 ― 定義・弊害・3要件",      "15分"),
    ("3", "高齢者虐待の理解 ― 5類型と連続性",        "15分"),
    ("4", "不適切ケアの早期発見 ― サインを読む",     "20分"),
    ("5", "居宅サービス事業所への指導・助言",        "20分"),
    ("6", "対応と連携 ・ まとめ",                    "10分"),
]
y = Cm(4.7)
for num, title, mins in agenda:
    add_numbered_circle(s, Cm(1.5), y, Cm(1.3), num, WINE_MID)
    add_text(s, Cm(3.3), y + Cm(0.15), Cm(22), Cm(1),
             title, size=17, bold=True, color=WINE)
    add_text(s, SW - Cm(4), y + Cm(0.2), Cm(3), Cm(1),
             mins, size=14, color=GOLD, bold=True, align=PP_ALIGN.RIGHT)
    y += Cm(1.85)


# =============================================================
#  SLIDE 4: 第1部
# =============================================================
section_title(1, "なぜ今、拘束最小化と\n虐待防止か",
              "Why It Matters Now")


# =============================================================
#  SLIDE 5: 制度的背景
# =============================================================
s = add_slide()
add_header(s, "1. なぜ今か", "制度的背景 ― 「体制整備」が義務になった")
add_footer(s, 5)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.5),
        "すべての介護サービス事業所が対象",
        "■ 高齢者虐待防止・身体拘束等の適正化のための「体制整備」が、\n"
        "    居宅介護支援事業所を含むすべての事業所に義務づけられている(経過措置は終了)",
        title_bg=WINE, body_size=13)

col_widths = [Cm(10), Cm(11.7), Cm(9.7)]
headers = ["求められる措置", "具体的な内容", "未実施の場合"]
rows = [
    ["虐待防止の措置",
     "委員会の設置 / 指針の整備 /\n研修の実施 / 担当者の選定",
     "虐待防止措置\n未実施減算"],
    ["身体拘束等の\n適正化(該当сサービス)",
     "適正化検討委員会 / 指針 /\n研修 ※サービス類型により異なる",
     "身体拘束廃止\n未実施減算 等"],
    ["記録・説明",
     "やむを得ず拘束した場合の\n記録、利用者・家族への説明",
     "運営基準違反\n指導・監査の対象"],
]
add_table(s, Cm(1.2), Cm(7.2), col_widths, headers, rows,
          header_bg=WINE, first_col_color=WINE,
          row_height=Cm(2.2), body_size=11)

add_box(s, Cm(1.2), Cm(14.5), Cm(31.4), Cm(2.5),
        "「減算」は本質ではない",
        "■ 減算は「最低限のライン」 ― 罰則を避けるための取り組みでは、利用者は守れない\n"
        "■ 本質は、利用者一人ひとりの「尊厳」と「権利」を守ること。制度はその下支えにすぎない",
        title_bg=GOLD, body_bg=AMBER_BG, body_size=12)


# =============================================================
#  SLIDE 6: 高齢者虐待防止法
# =============================================================
s = add_slide()
add_header(s, "1. なぜ今か", "高齢者虐待防止法 ― 基本の枠組み")
add_footer(s, 6)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.7),
        "高齢者虐待防止法とは",
        "正式名称「高齢者虐待の防止、高齢者の養護者に対する支援等に関する法律」(平成18年施行)。\n"
        "高齢者の虐待を防ぐとともに、虐待してしまう養護者(家族等)への支援も定めた法律。",
        title_bg=WINE, body_size=13)

add_box(s, Cm(1.2), Cm(7.3), Cm(15.5), Cm(5),
        "2つの虐待の区分",
        "■ 養介護施設従事者等による虐待\n   施設・事業所の職員等によるもの\n\n"
        "■ 養護者による虐待\n   家族・親族・同居人等によるもの\n\n"
        "▶ 居宅CMは、どちらにも出会いうる立場",
        title_bg=WINE_MID, body_size=12)

add_box(s, Cm(17.5), Cm(7.3), Cm(15), Cm(5),
        "通報の義務",
        "■ 虐待を受けたと思われる高齢者を発見した者は、\n   市町村に 通報する義務 がある\n\n"
        "■ 特に、生命・身体に重大な危険が生じている\n   場合は、速やかな通報が必須\n\n"
        "■ 「確証」がなくても「思われる」段階で通報",
        title_bg=RED, body_bg=RED_BG, body_size=12)

add_box(s, Cm(1.2), Cm(12.8), Cm(31.4), Cm(4.2),
        "通報をためらわせるものと、その誤解",
        "■ 「確証がない」→ 確証は不要。「虐待を受けたと思われる」段階で通報してよい\n"
        "■ 「事業所との関係が悪くなる」→ 通報者は法律で保護される。利用者の安全が最優先\n"
        "■ 「家族を追い詰めたくない」→ 通報は「罰」ではなく、養護者支援につなぐ入口でもある",
        title_bg=GOLD, body_bg=AMBER_BG, body_size=13)


# =============================================================
#  SLIDE 7: 居宅CM・主任CMの立場
# =============================================================
s = add_slide()
add_header(s, "1. なぜ今か", "居宅CM・主任CMの立場 ― できること/できないこと")
add_footer(s, 7)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.7),
        "居宅CMは「直接介護をしない」立場",
        "■ 居宅介護支援事業所のCMは、自ら身体介護を行う場面は通常ない\n"
        "■ しかし、訪問・通所・施設系サービスの現場で起きる不適切ケアに「気づける」立場にある",
        title_bg=WINE, body_size=13)

add_box(s, Cm(1.2), Cm(7.3), Cm(15.5), Cm(6.5),
        "居宅CMにできること",
        "■ モニタリング訪問で、利用者の様子・\n   サービスの場の異変に気づく\n\n"
        "■ サービス担当者会議で、ケアのあり方を\n   議題にし、多職種で検討する\n\n"
        "■ 気づきを記録し、関係機関につなぐ\n\n"
        "■ 虐待を疑ったら、市町村・包括に通報する\n\n"
        "■ 後輩CMの「気づく力」を育てる",
        title_bg=SAGE, body_bg=SAGE_BG, body_size=12)

add_box(s, Cm(17.5), Cm(7.3), Cm(15), Cm(6.5),
        "CMにできないこと/限界",
        "■ 行政のような「指導・処分」の権限はない\n\n"
        "■ サービス事業所の内部運営を直接\n   変えることはできない\n\n"
        "■ 一人で抱え込んで解決することはできない\n\n"
        "▶ だからこそ「気づく → つなぐ → 連携する」\n   が居宅CMの役割の中心になる",
        title_bg=AMBER, body_bg=AMBER_BG, body_size=12)

add_text(s, Cm(1.2), Cm(14.3), SW - Cm(2.4), Cm(1.2),
         "▶ 「指導する」より「気づき、議論の場をつくり、つなぐ」 ― これが主任CMの現実的な貢献。",
         size=13, color=WINE, bold=True)


# =============================================================
#  SLIDE 8: 第2部
# =============================================================
section_title(2, "身体拘束の理解",
              "Understanding Physical Restraint")


# =============================================================
#  SLIDE 9: 身体拘束とは
# =============================================================
s = add_slide()
add_header(s, "2. 身体拘束の理解", "身体拘束とは何か ― 具体的な行為")
add_footer(s, 9)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.3),
        "身体拘束 = 利用者の行動を制限する行為",
        "ひも・抑制帯・ミトン・柵・つなぎ服などで身体の自由を奪う行為のほか、\n"
        "言葉や薬による行動の制限も含まれる。下記は厚労省が例示する代表的な行為。",
        title_bg=WINE, body_size=12)

restraints = [
    "徘徊しないよう、車椅子・椅子・ベッドに体幹や四肢をひも等で縛る",
    "転落しないよう、ベッドに体幹や四肢をひも等で縛る",
    "自分で降りられないよう、ベッドを柵(サイドレール)で囲む",
    "チューブを抜かないよう、四肢をひも等で縛る",
    "チューブ抜去や皮膚かきむしりを防ぐため、ミトン型の手袋等をつける",
    "ずり落ち・立ち上がりを防ぐため、Y字型抑制帯・腰ベルト・車椅子テーブルをつける",
    "立ち上がる能力のある人の立ち上がりを妨げる椅子を使用する",
    "脱衣やおむつはずしを制限するため、介護衣(つなぎ服)を着せる",
    "他人への迷惑行為を防ぐため、ベッド等に体幹や四肢をひも等で縛る",
    "行動を落ち着かせるため、向精神薬を過剰に服用させる",
    "自分の意思で開けられない居室等に隔離する",
]
y = Cm(7)
col_w = Cm(15.7)
for i, r in enumerate(restraints):
    col = i % 2
    row = i // 2
    x = Cm(1.2) + col * (col_w + Cm(0.4))
    yy = y + row * Cm(1.55)
    add_rect(s, x, yy, col_w, Cm(1.4), TABLE_BG if (i % 2 == 0) else WHITE,
             line_color=MID_GRAY)
    add_numbered_circle(s, x + Cm(0.2), yy + Cm(0.3), Cm(0.8), i + 1, WINE_MID, fontsize=11)
    add_text(s, x + Cm(1.3), yy + Cm(0.1), col_w - Cm(1.5), Cm(1.2),
             r, size=10, color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.1)


# =============================================================
#  SLIDE 10: 身体拘束の弊害
# =============================================================
s = add_slide()
add_header(s, "2. 身体拘束の理解", "身体拘束がもたらす弊害 ― 「安全」のはずが…")
add_footer(s, 10)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.3),
        "拘束は「悪循環」を生む",
        "「安全のため」と始めた拘束が、心身の機能を奪い、かえって事故やBPSDを招き、\n"
        "さらに拘束が必要に見えてしまう ― この悪循環を断ち切る視点が必要。",
        title_bg=WINE, body_size=12)

harms = [
    ("身体的弊害", RED,
     "・関節拘縮・筋力低下・廃用症候群\n・褥瘡・食欲低下・心肺機能の低下\n・拘束具による事故(窒息等)の危険"),
    ("精神的弊害", WINE_MID,
     "・不安・怒り・屈辱・あきらめ\n・認知症の進行・せん妄の悪化\n・「人として扱われない」尊厳の傷つき"),
    ("社会的弊害", AMBER,
     "・家族の罪悪感・後悔\n・職員の士気・誇りの低下\n・事業所への社会的信頼の失墜"),
]
col_w = Cm(10.2)
col_h = Cm(6.5)
gap = Cm(0.4)
start_x = (SW - (col_w * 3 + gap * 2)) // 2
y = Cm(7.3)
for i, (title, color, body) in enumerate(harms):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1.3), color)
    add_text(s, x, y, col_w, Cm(1.3),
             title, size=16, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1.3), col_w, col_h - Cm(1.3), TABLE_BG)
    add_text(s, x + Cm(0.4), y + Cm(1.7), col_w - Cm(0.8), col_h - Cm(2.1),
             body, size=12, color=DARK_GRAY, line_spacing=1.5)

add_box(s, Cm(1.2), Cm(14.5), Cm(31.4), Cm(2.5),
        "視点の転換",
        "■ 「拘束しないと危ない」ではなく「なぜその行動が起きるのか」を考える\n"
        "■ 行動の背景(痛み・不安・環境・コミュニケーション不足)に手を当てれば、拘束は減らせる",
        title_bg=SAGE, body_bg=SAGE_BG, body_size=13)


# =============================================================
#  SLIDE 11: 3要件
# =============================================================
s = add_slide()
add_header(s, "2. 身体拘束の理解", "「緊急やむを得ない場合」の3要件")
add_footer(s, 11)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.3),
        "原則は「拘束しない」 ― 例外は極めて限定的",
        "身体拘束は原則禁止。例外的に認められるのは、下記3要件を「すべて」満たし、\n"
        "かつ組織として手続きを踏んだ場合のみ。1つでも欠ければ認められない。",
        title_bg=WINE, body_size=12)

reqs = [
    ("切迫性",
     "利用者本人または他の人の生命・\n身体が危険にさらされる可能性が\n著しく高いこと"),
    ("非代替性",
     "身体拘束以外に代わる介護方法が\nないこと(他の手段を尽くした\n上での最終手段であること)"),
    ("一時性",
     "身体拘束が一時的なものである\nこと(必要最小限の時間に\nとどめること)"),
]
col_w = Cm(10.2)
col_h = Cm(6)
gap = Cm(0.4)
start_x = (SW - (col_w * 3 + gap * 2)) // 2
y = Cm(7.3)
for i, (title, body) in enumerate(reqs):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1.5), WINE_MID)
    add_text(s, x, y, col_w, Cm(1.5),
             f"要件{i+1}  {title}", size=16, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1.5), col_w, col_h - Cm(1.5), TABLE_BG)
    add_text(s, x + Cm(0.4), y + Cm(1.9), col_w - Cm(0.8), col_h - Cm(2.3),
             body, size=12, color=DARK_GRAY, line_spacing=1.5,
             anchor=MSO_ANCHOR.MIDDLE)

add_box(s, Cm(1.2), Cm(14.3), Cm(31.4), Cm(2.7),
        "3要件は「拘束を正当化する道具」ではない",
        "■ 3要件は「拘束してよい理由」を探すためのものではなく、「本当に他に手段がないか」を\n"
        "    組織で吟味するためのチェックポイント ■ 「3要件を満たした」で思考を止めないこと",
        title_bg=GOLD, body_bg=AMBER_BG, body_size=13)


# =============================================================
#  SLIDE 12: 手続きと適正化の取り組み
# =============================================================
s = add_slide()
add_header(s, "2. 身体拘束の理解", "やむを得ず拘束する場合の手続きと、適正化の体制")
add_footer(s, 12)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(6.5),
        "やむを得ず拘束する場合の手続き",
        "■ 個人ではなく「組織」として判断する\n   (拘束廃止委員会等での検討)\n\n"
        "■ 利用者本人・家族に、目的・理由・\n   時間帯・期間等を説明し、同意を得る\n\n"
        "■ 「態様・時間・心身の状況・やむを得ない\n   理由」を記録する\n\n"
        "■ 必要性を常に再検討し、速やかな解除に向けて動く",
        title_bg=WINE_MID, body_size=12)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6.5),
        "適正化のための体制(事業所の義務)",
        "■ 身体拘束等適正化検討委員会の設置\n   (定期開催、結果を職員に周知)\n\n"
        "■ 適正化のための指針の整備\n\n"
        "■ 職員への定期的な研修の実施\n\n"
        "▶ 「委員会・指針・研修」の3点セットは\n   虐待防止の体制整備とも共通する",
        title_bg=SAGE, body_bg=SAGE_BG, body_size=12)

add_box(s, Cm(1.2), Cm(11.5), Cm(31.4), Cm(5.5),
        "居宅CMがモニタリングで確認したいこと",
        "■ 通所・施設系サービスを利用している場合、身体拘束が行われていないか\n\n"
        "■ 行われている場合、3要件・組織決定・記録・本人家族への説明 が適切になされているか\n\n"
        "■ サービス担当者会議で、拘束の解除・代替策を多職種で検討できているか\n\n"
        "■ 「ずっと続いている拘束」がないか ― 一時性が形骸化していないか",
        title_bg=WINE, body_size=13)


# =============================================================
#  SLIDE 13: 第3部
# =============================================================
section_title(3, "高齢者虐待の理解",
              "Understanding Elder Abuse")


# =============================================================
#  SLIDE 14: 虐待の5類型
# =============================================================
s = add_slide()
add_header(s, "3. 高齢者虐待の理解", "高齢者虐待の5類型")
add_footer(s, 14)

types = [
    ("身体的虐待",
     "暴力的行為で身体に傷・痛み\nを与える、または与えるおそれ\nのある行為。不適切な身体拘束も含む",
     "・たたく、つねる、蹴る\n・無理に食べさせる\n・縛る、閉じ込める"),
    ("心理的虐待",
     "脅し・侮辱・無視などにより\n精神的な苦痛を与える行為",
     "・どなる、ののしる\n・子ども扱いする\n・無視する、嫌がらせ"),
    ("性的虐待",
     "本人が同意していない、\nあらゆる性的な行為・その強要",
     "・性的な行為の強要\n・排泄の失敗等への\n  辱め"),
    ("経済的虐待",
     "本人の財産を不当に処分する、\n本人から不当に財産上の利益を得る",
     "・年金・預貯金の無断使用\n・財産の無断処分\n・金銭を渡さない"),
    ("介護等放棄\n(ネグレクト)",
     "必要な介護・世話を放棄し、\n心身を衰弱させる行為",
     "・食事・水分を与えない\n・入浴・排泄の世話をしない\n・必要な受診をさせない"),
]
col_w = Cm(6.2)
col_h = Cm(11)
gap = Cm(0.2)
start_x = (SW - (col_w * 5 + gap * 4)) // 2
y = Cm(4.5)
colors = [RED, WINE_MID, WINE, AMBER, TEAL]
for i, (title, desc, ex) in enumerate(types):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1.8), colors[i])
    add_text(s, x + Cm(0.1), y, col_w - Cm(0.2), Cm(1.8),
             title, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1.8), col_w, Cm(4), TABLE_BG)
    add_text(s, x + Cm(0.25), y + Cm(2), col_w - Cm(0.5), Cm(3.7),
             desc, size=10, color=DARK_GRAY, line_spacing=1.3)
    add_rect(s, x, y + Cm(5.8), col_w, col_h - Cm(5.8), LIGHT_BG)
    add_text(s, x + Cm(0.25), y + Cm(6), col_w - Cm(0.5), col_h - Cm(6.2),
             "【例】\n" + ex, size=10, color=WINE, line_spacing=1.3)

add_box(s, Cm(1.2), Cm(15.8), Cm(31.4), Cm(1.5),
        "",
        "▶ 1つの事例で複数の類型が重なることが多い。「どれか1つ」と決めつけず、全体像で捉える。",
        title_bg=WHITE, body_bg=AMBER_BG, title_size=1, body_size=12)


# =============================================================
#  SLIDE 15: 虐待の気づきにくさ
# =============================================================
s = add_slide()
add_header(s, "3. 高齢者虐待の理解", "なぜ虐待は「気づきにくい」のか")
add_footer(s, 15)

reasons = [
    ("本人が訴えない",
     "・認知症で説明できない\n・家族をかばう、あきらめている\n・「迷惑をかけたくない」"),
    ("加害側に自覚がない",
     "・「しつけ」「介護のため」と思っている\n・介護疲れで余裕を失っている\n・不適切ケアが日常化している"),
    ("密室性",
     "・在宅も施設も「閉じた空間」\n・第三者の目が入りにくい\n・サービスの場は見えにくい"),
    ("発見側の心理",
     "・「考えすぎかも」と思いたい\n・関係が壊れることへの不安\n・確証がないと動けないという誤解"),
]
col_w = Cm(15.5)
gap_x = Cm(0.4)
y_start = Cm(4.5)
row_h = Cm(3.7)
colors = [WINE_MID, RED, AMBER, TEAL]
for i, (title, body) in enumerate(reasons):
    col = i % 2
    row = i // 2
    x = Cm(1.2) + col * (col_w + gap_x)
    y = y_start + row * (row_h + Cm(0.3))
    add_rect(s, x, y, col_w, Cm(1), colors[i])
    add_text(s, x + Cm(0.5), y, col_w - Cm(1), Cm(1),
             title, size=15, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1), col_w, row_h - Cm(1), TABLE_BG)
    add_text(s, x + Cm(0.5), y + Cm(1.2), col_w - Cm(1), row_h - Cm(1.4),
             body, size=12, color=DARK_GRAY, line_spacing=1.4)

add_box(s, Cm(1.2), Cm(13), Cm(31.4), Cm(4),
        "だからこそ、定期的に関わるCMの「気づき」が重要",
        "■ 居宅CMは、利用者・家族・サービスの場に定期的に関わる数少ない専門職\n\n"
        "■ 「気づきにくい」ことを知っているからこそ、意識して見る・聴く・記録する\n\n"
        "■ 「考えすぎかも」と思ったその違和感こそ、最初の大切なサイン",
        title_bg=WINE, body_size=13)


# =============================================================
#  SLIDE 16: 不適切ケアと虐待の連続性
# =============================================================
s = add_slide()
add_header(s, "3. 高齢者虐待の理解", "不適切ケアと虐待の「連続性」")
add_footer(s, 16)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.3),
        "虐待は「ある日突然」起きるのではない",
        "明らかな虐待の手前に、広大な「不適切ケア(グレーゾーン)」がある。\n"
        "不適切ケアを見逃して放置すると、それが日常化し、やがて虐待へとつながっていく。",
        title_bg=WINE, body_size=12)

# 連続性の図(三角形/段階)
levels = [
    ("適切なケア", SAGE, "利用者の尊厳・意思を尊重したケア"),
    ("不適切なケア(グレーゾーン)", AMBER,
     "悪意はないが、利用者の尊厳を損なう関わり\n例:スピーチロック、子ども扱い、流れ作業的な介助"),
    ("高齢者虐待", RED, "5類型に該当する行為 ― 通報・対応が必要"),
]
y = Cm(7.2)
widths = [Cm(31.4), Cm(25), Cm(16)]
for i, (label, color, desc) in enumerate(levels):
    w = widths[i]
    x = (SW - w) // 2
    h = Cm(2.4)
    add_rect(s, x, y, w, h, color)
    add_text(s, x + Cm(0.3), y + Cm(0.2), w - Cm(0.6), Cm(0.9),
             label, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, x + Cm(0.3), y + Cm(1.05), w - Cm(0.6), Cm(1.2),
             desc, size=10, color=WHITE, align=PP_ALIGN.CENTER, line_spacing=1.15)
    y += Cm(2.7)

add_box(s, Cm(1.2), Cm(15.3), Cm(31.4), Cm(2.2),
        "早期発見の本質",
        "■ 「虐待かどうか」を見極めようとするより、「不適切ケアの芽」の段階で気づき、関わる\n"
        "■ グレーゾーンへの感度こそが、虐待を未然に防ぐ最大の力になる",
        title_bg=GOLD, body_bg=AMBER_BG, body_size=12)


# =============================================================
#  SLIDE 17: 第4部
# =============================================================
section_title(4, "不適切ケアの早期発見",
              "Early Detection of Inappropriate Care")


# =============================================================
#  SLIDE 18: 3つのロック
# =============================================================
s = add_slide()
add_header(s, "4. 早期発見", "見えにくい拘束 ― 「3つのロック」")
add_footer(s, 18)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.3),
        "ひもや柵だけが拘束ではない",
        "身体を縛らなくても、言葉や薬で人の行動を縛ることができる。\n"
        "これらは「見えにくい拘束」として、不適切ケアの中でも特に見過ごされやすい。",
        title_bg=WINE, body_size=12)

locks = [
    ("フィジカルロック",
     "身体的な拘束",
     "ひも・抑制帯・ミトン・柵・\nつなぎ服など、物理的に\n身体の自由を奪う",
     RED),
    ("スピーチロック",
     "言葉による拘束",
     "「動かないで」「ちょっと待って」\n「ダメ」等の言葉で\n行動を抑え込む",
     AMBER),
    ("ドラッグロック",
     "薬による拘束",
     "向精神薬等を過剰に用いて\n行動を抑制する。\n医師との連携が不可欠",
     WINE_MID),
]
col_w = Cm(10.2)
col_h = Cm(7)
gap = Cm(0.4)
start_x = (SW - (col_w * 3 + gap * 2)) // 2
y = Cm(7.3)
for i, (title, sub, body, color) in enumerate(locks):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1.8), color)
    add_text(s, x + Cm(0.2), y + Cm(0.2), col_w - Cm(0.4), Cm(0.9),
             title, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, x + Cm(0.2), y + Cm(1.05), col_w - Cm(0.4), Cm(0.6),
             sub, size=11, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(s, x, y + Cm(1.8), col_w, col_h - Cm(1.8), TABLE_BG)
    add_text(s, x + Cm(0.4), y + Cm(2.2), col_w - Cm(0.8), col_h - Cm(2.6),
             body, size=12, color=DARK_GRAY, line_spacing=1.5,
             anchor=MSO_ANCHOR.MIDDLE)

add_box(s, Cm(1.2), Cm(14.7), Cm(31.4), Cm(2.3),
        "スピーチロックは「最も身近で、最も気づかれにくい拘束」",
        "■ 忙しさの中で、つい出てしまう言葉。悪意がないからこそ日常化しやすい\n"
        "■ 「待っててね」の一言を、「○○しますね、もう少しだけ」に変える ― 小さな意識が現場を変える",
        title_bg=GOLD, body_bg=AMBER_BG, body_size=12)


# =============================================================
#  SLIDE 19: サインの見つけ方
# =============================================================
s = add_slide()
add_header(s, "4. 早期発見", "サインを読む ― 4つの視点で観る")
add_footer(s, 19)

signs = [
    ("利用者本人から",
     "・説明のつかない あざ・傷・やせ\n・おびえ・無表情・急な変化\n・「家に帰りたくない」等の言葉\n・身なり・清潔の乱れ"),
    ("環境・生活から",
     "・居室の不衛生・異臭\n・必要な物(眼鏡・杖・薬)の不足\n・食事・水分の不足のサイン\n・お金・通帳が管理されすぎ"),
    ("職員・介護者から",
     "・利用者への乱暴な言葉・態度\n・「困った人」という語り方\n・余裕のなさ・疲弊\n・質問への防衛的な反応"),
    ("記録・情報から",
     "・サービス記録の不自然な空白\n・受診・服薬の中断\n・拘束記録の形骸化\n・他職種からの違和感の声"),
]
col_w = Cm(15.5)
gap_x = Cm(0.4)
y_start = Cm(4.5)
row_h = Cm(4)
colors = [RED, AMBER, WINE_MID, TEAL]
for i, (title, body) in enumerate(signs):
    col = i % 2
    row = i // 2
    x = Cm(1.2) + col * (col_w + gap_x)
    y = y_start + row * (row_h + Cm(0.3))
    add_rect(s, x, y, col_w, Cm(1), colors[i])
    add_text(s, x + Cm(0.5), y, col_w - Cm(1), Cm(1),
             title, size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1), col_w, row_h - Cm(1), TABLE_BG)
    add_text(s, x + Cm(0.5), y + Cm(1.2), col_w - Cm(1), row_h - Cm(1.4),
             body, size=12, color=DARK_GRAY, line_spacing=1.4)

add_text(s, Cm(1.2), Cm(16.7), SW - Cm(2.4), Cm(0.8),
         "▶ 1つのサインで断定しない。複数のサインが「重なる」とき、危険度は高まる。",
         size=12, color=WINE, bold=True)


# =============================================================
#  SLIDE 20: モニタリング訪問での着眼点
# =============================================================
s = add_slide()
add_header(s, "4. 早期発見", "モニタリング訪問での着眼点")
add_footer(s, 20)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.3),
        "モニタリングは「気づきの最前線」",
        "居宅CMが利用者宅・サービスの場に定期的に足を運ぶモニタリングは、\n"
        "不適切ケアに気づける貴重な機会。「サービスの実施状況」だけを見て終わらせない。",
        title_bg=WINE, body_size=12)

points = [
    ("観る",
     "・利用者の表情・様子の変化\n・身体の状態(あざ・やせ等)\n・生活環境・清潔さ"),
    ("聴く",
     "・利用者本人の言葉(さりげなく)\n・家族・介護者の語り方\n・サービス職員の利用者の語り方"),
    ("会う",
     "・できるだけ本人と二人で話す機会を\n・家族同席だと言えないことがある\n・複数回・時間帯を変えて訪問"),
    ("記録する",
     "・違和感は「事実」として具体的に記録\n・日時・状況・誰の発言かを明確に\n・自分の解釈と事実を分けて書く"),
]
col_w = Cm(7.7)
col_h = Cm(6)
gap = Cm(0.3)
start_x = (SW - (col_w * 4 + gap * 3)) // 2
y = Cm(7.3)
colors = [WINE_MID, TEAL, GOLD, SAGE]
for i, (title, body) in enumerate(points):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1.3), colors[i])
    add_text(s, x, y, col_w, Cm(1.3),
             title, size=16, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1.3), col_w, col_h - Cm(1.3), TABLE_BG)
    add_text(s, x + Cm(0.3), y + Cm(1.6), col_w - Cm(0.6), col_h - Cm(1.9),
             body, size=11, color=DARK_GRAY, line_spacing=1.4)

add_box(s, Cm(1.2), Cm(14.7), Cm(31.4), Cm(2.3),
        "「本人と二人で話す」を意識的につくる",
        "■ 家族や職員が同席していると、利用者は本当のことを言えないことがある\n"
        "■ さりげなく場をつくり、「困っていることはないですか」と本人の声を聴く機会を持つ",
        title_bg=GOLD, body_bg=AMBER_BG, body_size=12)


# =============================================================
#  SLIDE 21: グレーゾーンへの感度
# =============================================================
s = add_slide()
add_header(s, "4. 早期発見", "グレーゾーンへの感度を磨く ― 事例で考える")
add_footer(s, 21)

col_widths = [Cm(13), Cm(8), Cm(10.4)]
headers = ["こんな場面、どう感じますか?", "見え方", "考えたいこと"]
rows = [
    ["「危ないから」と日中ずっと\n車椅子にテーブルをつけている",
     "不適切ケア\n〜身体的虐待",
     "立ち上がりの理由は?\n代替策は検討されたか"],
    ["職員が利用者を\nあだ名・呼び捨てで呼ぶ",
     "不適切ケア\n(心理面)",
     "親しみか、子ども扱いか\n本人はどう感じているか"],
    ["「待っててね」「動かないで」が\n一日中飛び交っている",
     "スピーチロック",
     "なぜ待たせるのか\n人員・段取りの問題では"],
    ["家族が利用者の年金を管理し\n本人に小遣いを渡さない",
     "経済的虐待の\n疑い",
     "本人の意思は?\n生活に支障は出ていないか"],
    ["「この人は手がかかる」と\n職員がため息まじりに話す",
     "不適切ケアの\n前兆・職員の疲弊",
     "職員の負担・体制は?\nチームで支えられているか"],
]
add_table(s, Cm(1.2), Cm(4.5), col_widths, headers, rows,
          header_bg=WINE, first_col_color=DARK_GRAY,
          row_height=Cm(2.0), body_size=11, header_size=12,
          first_col_bold=False)

add_box(s, Cm(1.2), Cm(16), Cm(31.4), Cm(1.5),
        "",
        "▶ 「白か黒か」で判断しない。グレーの段階で「気にかける」「話題にする」ことが予防になる。",
        title_bg=WHITE, body_bg=AMBER_BG, title_size=1, body_size=12)


# =============================================================
#  SLIDE 22: 第5部
# =============================================================
section_title(5, "居宅サービス事業所への\n指導・助言",
              "Guidance to Service Providers")


# =============================================================
#  SLIDE 23: 指導・助言のスタンス
# =============================================================
s = add_slide()
add_header(s, "5. 指導・助言", "指導・助言の基本スタンス")
add_footer(s, 23)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.5),
        "CMの「指導・助言」は、行政の指導とは違う",
        "■ CMには処分権限はない。だからこそ「上から正す」のではなく、\n"
        "    「同じ目標に向かう専門職どうし」として、対等に問題を共有する姿勢が基本になる",
        title_bg=WINE, body_size=13)

add_box(s, Cm(1.2), Cm(7.3), Cm(15.5), Cm(6.5),
        "避けたい関わり方",
        "■ 「虐待では?」と決めつけて責める\n\n"
        "■ 事業所を一方的に批判する\n\n"
        "■ 関係悪化を恐れて何も言わない\n\n"
        "■ 一人のCMだけで抱え込む\n\n"
        "■ 逆に、見て見ぬふりをして流す\n\n"
        "▶ 「対立」も「黙認」も利用者を守れない",
        title_bg=RED, body_bg=RED_BG, body_size=12)

add_box(s, Cm(17.5), Cm(7.3), Cm(15), Cm(6.5),
        "目指したい関わり方",
        "■ 事実を具体的に、冷静に共有する\n\n"
        "■ 「利用者のために一緒に考えたい」と伝える\n\n"
        "■ 事業所の事情・背景にも耳を傾ける\n\n"
        "■ サービス担当者会議など「場」を活用する\n\n"
        "■ 緊急性が高い場合は迷わず通報する\n\n"
        "▶ 関係を保ちつつ、見過ごさない",
        title_bg=SAGE, body_bg=SAGE_BG, body_size=12)

add_text(s, Cm(1.2), Cm(14.3), SW - Cm(2.4), Cm(1.2),
         "▶ ゴールは「事業所を正すこと」ではなく「利用者の尊厳が守られる状態をつくること」。",
         size=13, color=WINE, bold=True)


# =============================================================
#  SLIDE 24: 対応フロー
# =============================================================
s = add_slide()
add_header(s, "5. 指導・助言", "気づいたときの対応フロー")
add_footer(s, 24)

steps = [
    ("1", "事実を確認・記録する",
     "・見聞きしたことを「事実」として具体的に記録\n・推測と事実を分けて整理する"),
    ("2", "緊急性を判断する",
     "・生命・身体に重大な危険があるか\n・あれば即・市町村へ通報(ステップ5へ)"),
    ("3", "一人で抱えず相談する",
     "・主任CM・管理者・地域包括に相談\n・複数の目で状況を見立てる"),
    ("4", "事業所と事実を共有する",
     "・サービス担当者会議等の「場」を活用\n・ケアの見直しを多職種で検討する"),
    ("5", "通報・連携する",
     "・虐待が疑われれば市町村へ通報\n・以降は行政・地域包括と連携して対応"),
]
y = Cm(4.4)
colors = [WINE_MID, AMBER, TEAL, SAGE, RED]
for i, (num, title, body) in enumerate(steps):
    add_numbered_circle(s, Cm(1.5), y + Cm(0.35), Cm(1.3), num, colors[i])
    add_rect(s, Cm(3.4), y, Cm(8), Cm(2), colors[i])
    add_text(s, Cm(3.5), y, Cm(7.8), Cm(2),
             title, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Cm(11.4), y, SW - Cm(12.6), Cm(2), TABLE_BG)
    add_text(s, Cm(11.8), y + Cm(0.2), SW - Cm(13.2), Cm(1.7),
             body, size=12, color=DARK_GRAY,
             anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.35)
    y += Cm(2.35)

add_box(s, Cm(1.2), Cm(16.2), Cm(31.4), Cm(1.3),
        "",
        "▶ 緊急性が高ければ、ステップ2から一気に5へ。「順番どおり」より「利用者の安全」が優先。",
        title_bg=WHITE, body_bg=RED_BG, title_size=1, body_size=12)


# =============================================================
#  SLIDE 25: 事業所への伝え方
# =============================================================
s = add_slide()
add_header(s, "5. 指導・助言", "事業所への伝え方 ― 言葉の選び方")
add_footer(s, 25)

col_widths = [Cm(15.7), Cm(15.7)]
headers = ["関係を壊しやすい言い方", "一緒に考える言い方"]
rows = [
    ["「それ、虐待じゃないですか」",
     "「気になる場面があったので、\n一緒に状況を整理させてください」"],
    ["「拘束はダメですよね」",
     "「この拘束、解除に向けて何か\n一緒に工夫できないでしょうか」"],
    ["「ちゃんと見てるんですか」",
     "「ご本人の様子が少し気になって。\n最近のご様子はいかがですか」"],
    ["「記録が不十分です」",
     "「記録があると、皆で経過を\n共有しやすいので助かります」"],
    ["「家族が虐待しています」",
     "「ご家族もかなりお疲れのようです。\n支える方法を一緒に考えませんか」"],
]
add_table(s, Cm(1.2), Cm(4.5), col_widths, headers, rows,
          header_bg=WINE, first_col_color=RED,
          row_height=Cm(1.85), body_size=12, first_col_bold=False)

add_box(s, Cm(1.2), Cm(15.3), Cm(31.4), Cm(2.2),
        "言い方の原則",
        "■ 「事実」を共有し、「人」を責めない ■ 「問いかけ」の形にして、一緒に考える姿勢を示す\n"
        "■ ただし、緊急性が高い虐待は、言葉を選んでいる場合ではない ― 速やかな通報を優先する",
        title_bg=GOLD, body_bg=AMBER_BG, body_size=12)


# =============================================================
#  SLIDE 26: サ担会議の活用
# =============================================================
s = add_slide()
add_header(s, "5. 指導・助言", "サービス担当者会議を「拘束を減らす場」にする")
add_footer(s, 26)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.3),
        "一人で言うより、多職種で考える",
        "CMが個別に事業所へ言うより、サービス担当者会議で多職種が一緒に検討するほうが、\n"
        "建設的で、関係も壊れにくく、実効性のある解決につながる。",
        title_bg=WINE, body_size=12)

add_box(s, Cm(1.2), Cm(7), Cm(15.5), Cm(5.3),
        "会議で扱いたい論点",
        "■ その行動(立ち上がり・徘徊等)は\n   なぜ起きているのか\n\n"
        "■ 拘束・抑制以外の代替策はないか\n\n"
        "■ 環境・関わり方・薬の見直しの余地は\n\n"
        "■ 拘束を「いつ・どう解除していくか」",
        title_bg=TEAL, body_bg=TEAL_BG, body_size=12)

add_box(s, Cm(17.5), Cm(7), Cm(15), Cm(5.3),
        "ファシリテーターとしてのCM",
        "■ 事業所を「責める場」にしない\n\n"
        "■ 各職種(医師・看護・リハ・介護)の\n   視点を引き出す\n\n"
        "■ 「誰が・何を・いつまでに」を決める\n\n"
        "■ 次回会議で必ず経過を振り返る",
        title_bg=SAGE, body_bg=SAGE_BG, body_size=12)

add_box(s, Cm(1.2), Cm(12.7), Cm(31.4), Cm(4.3),
        "主任CMの役割 ― 第4弾の学びをここで活かす",
        "■ ファシリテーションの技術(中立を保つ・全員参加・論点の可視化・結論を出す)を発揮する\n\n"
        "■ 「拘束ゼロ」は一度の会議では実現しない ― 小さな代替策から試し、経過を追い続ける\n\n"
        "■ 後輩CMには、サ担会議をこうした「ケアの質を問い直す場」として使う姿を見せて伝える",
        title_bg=WINE, body_size=13)


# =============================================================
#  SLIDE 27: 第6部
# =============================================================
section_title(6, "対応と連携 ・ まとめ",
              "Response, Collaboration & Wrap Up")


# =============================================================
#  SLIDE 28: 通報・連携の流れ
# =============================================================
s = add_slide()
add_header(s, "6. 対応と連携", "通報・連携の流れ ― 一人で抱えない")
add_footer(s, 28)

# フロー図
add_rect(s, Cm(2), Cm(4.7), Cm(8.5), Cm(2.3), WINE_MID)
add_text(s, Cm(2), Cm(4.7), Cm(8.5), Cm(2.3),
         "CMが気づく\n(モニタリング等)",
         size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
arr1 = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(10.7), Cm(5.4), Cm(2), Cm(0.9))
arr1.fill.solid(); arr1.fill.fore_color.rgb = GOLD; arr1.line.fill.background()
arr1.shadow.inherit = False
add_rect(s, Cm(12.9), Cm(4.7), Cm(8.5), Cm(2.3), TEAL)
add_text(s, Cm(12.9), Cm(4.7), Cm(8.5), Cm(2.3),
         "事業所内で相談\n(主任CM・管理者)",
         size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
arr2 = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(21.6), Cm(5.4), Cm(2), Cm(0.9))
arr2.fill.solid(); arr2.fill.fore_color.rgb = GOLD; arr2.line.fill.background()
arr2.shadow.inherit = False
add_rect(s, Cm(23.8), Cm(4.7), Cm(8), Cm(2.3), RED)
add_text(s, Cm(23.8), Cm(4.7), Cm(8), Cm(2.3),
         "市町村へ通報\n地域包括へ相談",
         size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_box(s, Cm(1.2), Cm(7.7), Cm(15.5), Cm(5),
        "通報・相談先",
        "■ 市町村(高齢者虐待の通報窓口)\n\n"
        "■ 地域包括支援センター\n   (総合相談・権利擁護の窓口)\n\n"
        "■ 緊急時は迷わず ― 確証より「疑い」で動く\n\n"
        "▶ 通報後の対応は、市町村が中心となって\n   進める(事実確認・立入調査等)",
        title_bg=WINE_MID, body_size=12)

add_box(s, Cm(17.5), Cm(7.7), Cm(15), Cm(5),
        "通報後もCMの関わりは続く",
        "■ 通報して「終わり」ではない\n\n"
        "■ 行政・地域包括と連携しながら、\n   利用者の生活を支え続ける\n\n"
        "■ ケアプランの見直し、サービス調整\n\n"
        "■ 養護者(家族)への支援にも関わる",
        title_bg=SAGE, body_bg=SAGE_BG, body_size=12)

add_box(s, Cm(1.2), Cm(13.2), Cm(31.4), Cm(3.8),
        "通報は「告発」ではなく「支援の入口」",
        "■ 通報の目的は、誰かを罰することではなく、利用者の安全と、養護者を含めた支援につなぐこと\n"
        "■ 通報者の情報は保護される。「関係が壊れる」ことを恐れて、利用者の安全を後回しにしない\n"
        "■ 一人で判断せず、組織として・地域として対応する ― それが主任CMの示すべき姿",
        title_bg=GOLD, body_bg=AMBER_BG, body_size=12)


# =============================================================
#  SLIDE 29: 養護者による虐待への対応
# =============================================================
s = add_slide()
add_header(s, "6. 対応と連携", "養護者(家族)による虐待 ― 「支援」の視点")
add_footer(s, 29)

add_box(s, Cm(1.2), Cm(4.3), Cm(31.4), Cm(2.5),
        "家族を「加害者」とだけ見ない",
        "■ 養護者による虐待の背景には、介護疲れ・孤立・経済的困窮・知識不足・本人との関係史がある\n"
        "■ 高齢者虐待防止法は、養護者への「支援」も定めている ― 罰するだけでは解決しない",
        title_bg=WINE, body_size=12)

add_box(s, Cm(1.2), Cm(7.3), Cm(15.5), Cm(5.5),
        "養護者の虐待の背景にあるもの",
        "■ 介護による心身の疲弊・追い詰められ\n\n"
        "■ 社会的な孤立・相談相手の不在\n\n"
        "■ 認知症等への知識・対応力の不足\n\n"
        "■ 経済的な困窮\n\n"
        "■ 本人と養護者の長年の関係・葛藤",
        title_bg=WINE_MID, body_size=12)

add_box(s, Cm(17.5), Cm(7.3), Cm(15), Cm(5.5),
        "CMができる支援的関わり",
        "■ 養護者の労をねぎらい、話を聴く\n\n"
        "■ サービス導入で介護負担を軽減する\n\n"
        "■ 孤立を防ぐ ― 相談先・社会資源につなぐ\n\n"
        "■ 分離が必要な場合は、行政と連携\n\n"
        "■ 「責める」前に「支える」道を探す",
        title_bg=SAGE, body_bg=SAGE_BG, body_size=12)

add_text(s, Cm(1.2), Cm(13.3), SW - Cm(2.4), Cm(1.2),
         "▶ ただし、生命・身体に重大な危険がある場合は、支援の視点と同時に「速やかな通報」が必要。",
         size=13, color=RED, bold=True)


# =============================================================
#  SLIDE 30: 後輩CMへの指導
# =============================================================
s = add_slide()
add_header(s, "6. 対応と連携", "後輩CMに、何をどう伝えるか")
add_footer(s, 30)

points = [
    ("「気づく目」を育てる",
     "・グレーゾーンへの感度を、事例を通じて伝える\n・「これって変かも」という違和感を大切にさせる\n・モニタリングの着眼点を一緒に確認する"),
    ("「抱え込ませない」",
     "・「迷ったら相談していい」を文化にする\n・後輩が気づきを話せる場を定期的に持つ\n・通報をためらわせる空気をつくらない"),
    ("対応をモデルで示す",
     "・事業所への伝え方、サ担会議の進め方を見せる\n・同行訪問で「観る・聴く」の実際を伝える\n・通報・連携の手順を一緒に確認しておく"),
    ("CM自身を支える",
     "・虐待対応はCM自身も傷つき、消耗する\n・対応後の振り返り・感情のケアを行う\n・「一人の責任にしない」と明確に伝える"),
]
col_w = Cm(15.5)
gap_x = Cm(0.4)
y_start = Cm(4.5)
row_h = Cm(3.8)
colors = [WINE_MID, TEAL, GOLD, SAGE]
for i, (title, body) in enumerate(points):
    col = i % 2
    row = i // 2
    x = Cm(1.2) + col * (col_w + gap_x)
    y = y_start + row * (row_h + Cm(0.3))
    add_rect(s, x, y, col_w, Cm(1), colors[i])
    add_text(s, x + Cm(0.5), y, col_w - Cm(1), Cm(1),
             title, size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1), col_w, row_h - Cm(1), TABLE_BG)
    add_text(s, x + Cm(0.5), y + Cm(1.2), col_w - Cm(1), row_h - Cm(1.4),
             body, size=11, color=DARK_GRAY, line_spacing=1.4)

add_box(s, Cm(1.2), Cm(13.3), Cm(31.4), Cm(3.7),
        "虐待対応はチームで ― 主任CMがその要に",
        "■ 虐待対応は、判断も対応も重く、一人のCMが背負うには大きすぎる\n"
        "■ 主任CMは、後輩が「気づき・相談し・対応できる」チームの仕組みと空気をつくる\n"
        "■ スーパービジョンの3機能(管理的・教育的・支持的)を、ここでこそ発揮する",
        title_bg=WINE, body_size=13)


# =============================================================
#  SLIDE 31: 持ち帰りアクション
# =============================================================
s = add_slide()
add_header(s, "Closing", "明日からの 3 つのアクション")
add_footer(s, 31)

actions = [
    ("Action 1", "モニタリングの「観る目」を変える",
     "・次の訪問から「観る・聴く・会う・記録する」の4視点を意識する\n"
     "・本人と二人で話せる場面を、意識的につくる"),
    ("Action 2", "グレーゾーンを事業所内で話題にする",
     "・スピーチロックなど「身近な不適切ケア」を、事例検討や朝礼で取り上げる\n"
     "・「白か黒か」でなく、グレーの段階で気にかける文化をつくる"),
    ("Action 3", "対応・通報の流れを確認しておく",
     "・市町村・地域包括の通報窓口、事業所内の相談ルートを再確認する\n"
     "・後輩CMと「気づいたらどう動くか」を一度話し合っておく"),
]
y = Cm(4.5)
colors = [WINE_MID, GOLD, TEAL]
for i, (label, title, body) in enumerate(actions):
    color = colors[i]
    add_rect(s, Cm(1.2), y, Cm(4.5), Cm(3.7), color)
    add_text(s, Cm(1.2), y + Cm(0.4), Cm(4.5), Cm(1),
             label, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Cm(1.2), y + Cm(1.5), Cm(4.5), Cm(1.8),
             label.split(" ")[1], size=24, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Cm(5.7), y, SW - Cm(6.9), Cm(3.7), TABLE_BG)
    add_text(s, Cm(6), y + Cm(0.3), SW - Cm(7.5), Cm(1),
             title, size=16, bold=True, color=WINE)
    add_text(s, Cm(6), y + Cm(1.4), SW - Cm(7.5), Cm(2.2),
             body, size=12, color=DARK_GRAY, line_spacing=1.4)
    y += Cm(4)


# =============================================================
#  SLIDE 32: クロージング
# =============================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, WINE)
add_rect(s, 0, Cm(7), SW, Cm(0.1), GOLD)

add_text(s, Cm(2), Cm(3), SW - Cm(4), Cm(1),
         "Closing Message", size=14, color=GOLD, bold=True,
         align=PP_ALIGN.CENTER)
add_text(s, Cm(2), Cm(8), SW - Cm(4), Cm(3),
         "拘束しないことは、手間ではなく、\n「その人を人として扱う」という約束。",
         size=25, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
         line_spacing=1.4)
add_text(s, Cm(3), Cm(12.5), SW - Cm(6), Cm(3),
         "あなたが訪問先で覚えた小さな違和感、本人と交わした二人きりの会話、\n"
         "サービス担当者会議で投げかけた一つの問い。\n"
         "その積み重ねが、声をあげられない誰かの尊厳を、確かに守っています。",
         size=14, color=MID_GRAY, align=PP_ALIGN.CENTER, line_spacing=1.6)


# =============================================================
#  SLIDE 33: 参考資料
# =============================================================
s = add_slide()
add_header(s, "Reference", "参考資料 ・ 情報源")
add_footer(s, 33)

refs = [
    "■ 高齢者虐待の防止、高齢者の養護者に対する支援等に関する法律(高齢者虐待防止法)",
    "■ 厚生労働省「市町村・都道府県における高齢者虐待への対応と養護者支援について」(対応マニュアル)",
    "■ 厚生労働省「身体拘束ゼロへの手引き」",
    "■ 厚生労働省「介護報酬改定」関係通知(虐待防止・身体拘束等の適正化の体制整備)",
    "■ 各事業所の虐待防止指針・身体拘束等適正化指針",
    "■ 各都道府県・市町村の高齢者虐待対応の手引き・通報窓口情報",
    "■ 日本介護支援専門員協会等による権利擁護・虐待防止関連の研修資料",
]
y = Cm(4.5)
for r in refs:
    add_text(s, Cm(2), y, SW - Cm(4), Cm(1),
             r, size=13, color=DARK_GRAY)
    y += Cm(1.25)

add_box(s, Cm(1.2), Cm(13.3), Cm(31.4), Cm(3.7),
        "学び続けるために",
        "■ 制度・通知は改定されます。最新の法令・厚労省マニュアル・自治体情報を必ずご確認ください\n"
        "■ 身体拘束・虐待への向き合い方は、知識だけでなく「事例を通じた振り返り」で深まります\n"
        "■ 本資料は 2026年5月時点の一般的な情報を基に構成しています",
        title_bg=GOLD, body_bg=AMBER_BG, body_size=12)


# =============================================================
#  SLIDE 34: 質疑応答
# =============================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, WINE)
add_rect(s, Cm(11), Cm(7.5), Cm(11.8), Cm(0.12), GOLD)
add_text(s, Cm(2), Cm(7.8), SW - Cm(4), Cm(2),
         "ご清聴ありがとうございました",
         size=30, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
add_text(s, Cm(2), Cm(11), SW - Cm(4), Cm(1.5),
         "質疑応答 ・ 事例検討",
         size=18, color=GOLD, align=PP_ALIGN.CENTER)
add_text(s, Cm(2), Cm(13), SW - Cm(4), Cm(1.5),
         "現場で出会った「気になる場面」を持ち寄り、一緒に考えましょう。",
         size=14, color=MID_GRAY, align=PP_ALIGN.CENTER)


# ===== 保存 =====
output = "/home/user/kaigo-ga-app/slides/身体拘束最小化と高齢者虐待防止_研修スライド.pptx"
prs.save(output)
print(f"Saved: {output}")
print(f"Total slides: {len(prs.slides)}")
