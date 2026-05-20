"""主任ケアマネ研修スライド生成スクリプト 第3弾

テーマ: 令和6年度介護報酬改定の振り返りと令和9年度改定に向けた論点整理
重点: 特定事業所加算 / ターミナルCM加算 + その他R6改定論点
トーン: 実務マニュアル型 / R9は制度設計視点
時間: 90分
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ===== カラーパレット(政策・経営トーン) =====
NAVY        = RGBColor(0x14, 0x2B, 0x4A)
STEEL       = RGBColor(0x37, 0x5A, 0x7F)
SLATE       = RGBColor(0x5F, 0x7A, 0x96)
ACCENT_RED  = RGBColor(0xC0, 0x3A, 0x2E)   # 改定・規制強化
ACCENT_AMB  = RGBColor(0xE0, 0xA6, 0x36)   # 注意・要確認
ACCENT_GREEN = RGBColor(0x2E, 0x8B, 0x57)  # 加算・チャンス
LIGHT_GRAY  = RGBColor(0xF1, 0xF3, 0xF6)
MID_GRAY    = RGBColor(0xD6, 0xDC, 0xE3)
DARK_GRAY   = RGBColor(0x33, 0x38, 0x3F)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_BG    = RGBColor(0xE8, 0xED, 0xF3)
RED_BG      = RGBColor(0xFB, 0xE7, 0xE5)
AMB_BG      = RGBColor(0xFB, 0xF1, 0xDA)
GREEN_BG    = RGBColor(0xE3, 0xEE, 0xE6)

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
    add_rect(slide, 0, 0, SW, Cm(1.6), NAVY)
    add_text(slide, Cm(0.8), Cm(0.25), Cm(22), Cm(0.6),
             section_label, size=11, color=WHITE, bold=True)
    add_text(slide, Cm(0.8), Cm(1.9), SW - Cm(1.6), Cm(1.3),
             title, size=23, bold=True, color=NAVY)
    add_rect(slide, Cm(0.8), Cm(3.2), Cm(2.5), Cm(0.12), ACCENT_RED)


def add_footer(slide, page_num, total=38):
    add_text(slide, Cm(0.8), SH - Cm(0.8), Cm(24), Cm(0.5),
             "主任CM研修 / R6改定の振り返りとR9改定への論点整理  ※具体的単位数は最新の告示・通知を必ずご確認ください",
             size=9, color=SLATE)
    add_text(slide, SW - Cm(3), SH - Cm(0.8), Cm(2.2), Cm(0.5),
             f"{page_num} / {total}", size=9, color=SLATE, align=PP_ALIGN.RIGHT)


def add_box(slide, left, top, width, height, title, body, *,
            title_bg=STEEL, body_bg=TABLE_BG, title_color=WHITE,
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
              header_bg=NAVY, header_color=WHITE,
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


def section_title(num, jp, en, *, accent_color=ACCENT_RED):
    s = add_slide()
    add_rect(s, 0, 0, SW, SH, NAVY)
    add_rect(s, 0, Cm(8.5), SW, Cm(0.1), accent_color)
    add_text(s, Cm(2), Cm(6.5), SW - Cm(4), Cm(1.5),
             f"PART {num}", size=22, color=accent_color, bold=True,
             align=PP_ALIGN.CENTER)
    add_text(s, Cm(1), Cm(9.5), SW - Cm(2), Cm(2),
             jp, size=36, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(s, Cm(1), Cm(13), SW - Cm(2), Cm(1),
             en, size=14, color=MID_GRAY, align=PP_ALIGN.CENTER)
    return s


# =============================================================
#  SLIDE 1: 表紙
# =============================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, NAVY)
add_rect(s, 0, Cm(8), SW, Cm(0.15), ACCENT_RED)
add_rect(s, 0, Cm(12.5), SW, Cm(0.05), STEEL)

add_text(s, Cm(2), Cm(4.0), SW - Cm(4), Cm(1),
         "主任介護支援専門員 研修会",
         size=18, color=ACCENT_RED, bold=True, align=PP_ALIGN.CENTER)
add_text(s, Cm(1), Cm(5.5), SW - Cm(2), Cm(2.5),
         "令和6年度介護報酬改定の振り返りと\n令和9年度改定に向けた論点整理",
         size=30, color=WHITE, bold=True, align=PP_ALIGN.CENTER, line_spacing=1.3)
add_text(s, Cm(1.5), Cm(10), SW - Cm(3), Cm(2.5),
         "ー 特定事業所加算 / ターミナルCM加算など主要加算の運用実態と、\n"
         "  次期改定に向けた制度設計の動向を読み解く ー",
         size=16, color=WHITE, align=PP_ALIGN.CENTER, line_spacing=1.4)

add_text(s, Cm(1.5), Cm(13.5), SW - Cm(3), Cm(1),
         "■ R6 改定施行から 2 年経過時点での運用棚卸し\n"
         "■ R9 (2027.4) 同時改定の論点を制度設計視点で俯瞰",
         size=14, color=ACCENT_AMB, align=PP_ALIGN.CENTER, line_spacing=1.5)

add_text(s, Cm(2), Cm(16), SW - Cm(4), Cm(1),
         "主任ケアマネジャーの会  主催  ・  研修時間 90 分",
         size=14, color=MID_GRAY, align=PP_ALIGN.CENTER)


# =============================================================
#  SLIDE 2: 本日の目標
# =============================================================
s = add_slide()
add_header(s, "Introduction", "本日の研修目標")
add_footer(s, 2)

goals = [
    "令和6年度改定の居宅介護支援関連の主要変更点を、施行2年経過の視点で棚卸しできる",
    "特定事業所加算(Ⅰ〜Ⅳ)の要件・算定実態・運用上の落とし穴を体系的に整理する",
    "ターミナルケアマネジメント加算の算定漏れ・記録不備のリスクを洗い出せる",
    "同一建物減算・逓減制・BCP・虐待防止など、改定で論点となった項目を再確認する",
    "令和9年度改定(同時改定)に向けた制度設計の方向性を理解する",
    "主任CMとして、事業所内・地域でどう情報を共有し議論を導くかを持ち帰る",
]
y = Cm(4.5)
for g in goals:
    add_rect(s, Cm(1.2), y, Cm(0.5), Cm(1.5), ACCENT_RED)
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
    ("1", "令和6年度改定の全体像 ― 何が変わったか",            "10分"),
    ("2", "特定事業所加算 ― 要件・実態・運用上の論点",         "20分"),
    ("3", "ターミナルケアマネジメント加算 ― 算定漏れを防ぐ",   "15分"),
    ("4", "その他 R6 改定の論点 ― 減算・逓減制・BCP・LIFE",   "15分"),
    ("5", "R6 改定の運用棚卸し ― 自事業所点検チェック",        "10分"),
    ("6", "令和9年度改定に向けた論点(制度設計視点)",         "15分"),
    ("7", "主任CMが事業所・地域に持ち帰るもの",                " 5分"),
]
y = Cm(4.5)
for num, title, mins in agenda:
    circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(1.5), y, Cm(1.2), Cm(1.2))
    circle.fill.solid()
    circle.fill.fore_color.rgb = STEEL
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

    add_text(s, Cm(3.2), y + Cm(0.15), Cm(22), Cm(1),
             title, size=17, bold=True, color=NAVY)
    add_text(s, SW - Cm(4), y + Cm(0.2), Cm(3), Cm(1),
             mins, size=14, color=ACCENT_RED, bold=True, align=PP_ALIGN.RIGHT)
    y += Cm(1.55)


# =============================================================
#  SLIDE 4: 第1部
# =============================================================
section_title(1, "令和6年度改定の全体像",
              "Overview of the FY2024 Revision")


# =============================================================
#  SLIDE 5: 改定の位置づけ
# =============================================================
s = add_slide()
add_header(s, "1. R6改定の全体像", "R6改定の位置づけ ― 「3 年に 1 度」「同時改定」の意味")
add_footer(s, 5)

# 改定スケジュール
years = [
    ("R3 (2021)", "介護のみ",  "コロナ対応・LIFE 導入", STEEL),
    ("R6 (2024)", "同時改定",  "医療・介護・障害\n3制度同時",     ACCENT_RED),
    ("R9 (2027)", "同時改定",  "次回・診療報酬と同時",          ACCENT_AMB),
    ("R12(2030)","介護のみ",  "予測:抜本的見直し",            STEEL),
]
col_w = Cm(7.7)
col_h = Cm(5)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
y = Cm(4.5)
for i, (label, kind, note, color) in enumerate(years):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1.2), color)
    add_text(s, x + Cm(0.2), y + Cm(0.15), col_w - Cm(0.4), Cm(1),
             label, size=15, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1.2), col_w, col_h - Cm(1.2), TABLE_BG)
    add_text(s, x + Cm(0.2), y + Cm(1.3), col_w - Cm(0.4), Cm(1),
             kind, size=13, bold=True, color=color,
             align=PP_ALIGN.CENTER)
    add_text(s, x + Cm(0.2), y + Cm(2.4), col_w - Cm(0.4), Cm(2),
             note, size=11, color=DARK_GRAY,
             align=PP_ALIGN.CENTER)

add_box(s, Cm(1.2), Cm(10.5), Cm(31.4), Cm(6.5),
        "「同時改定」の意味するもの",
        "■ 医療(診療報酬)・介護・障害(障害福祉サービス等)が同じ年度に同時改定される\n"
        "    → 制度間の整合性・連携が強く意識される\n\n"
        "■ R6 改定の柱(社会保障審議会・介護給付費分科会の答申より)\n"
        "    ① 地域包括ケアシステムの深化・推進\n"
        "    ② 自立支援・重度化防止の質の向上\n"
        "    ③ 良質な介護サービスの効率的な提供への取組\n"
        "    ④ 制度の安定性・持続可能性の確保\n\n"
        "■ R9 改定 (2027.4) も同時改定。R6 で残された課題が次の俎上に乗る。",
        title_bg=NAVY, body_size=13)


# =============================================================
#  SLIDE 6: 居宅介護支援の主要変更点
# =============================================================
s = add_slide()
add_header(s, "1. R6改定の全体像", "居宅介護支援に関わる主要変更点 ― 一覧マップ")
add_footer(s, 6)

# 4カテゴリ
cats = [
    ("加算の見直し",  STEEL, [
        "・特定事業所加算 (要件強化)",
        "・ターミナルCM加算 (算定要件明確化)",
        "・入院時情報連携加算 (区分見直し)",
        "・通院時情報連携加算 (継続)",
    ]),
    ("減算・基準改定", ACCENT_RED, [
        "・同一建物減算 (居宅CMに新設)",
        "・逓減制 (ICT活用要件の見直し)",
        "・運営基準減算 (継続)",
    ]),
    ("基準・体制整備", ACCENT_AMB, [
        "・BCP策定の本格運用 (経過措置終了)",
        "・高齢者虐待防止 委員会・指針・研修",
        "・身体的拘束等の適正化",
        "・ハラスメント対策の明示",
    ]),
    ("ICT・DX",       ACCENT_GREEN, [
        "・LIFE の活用方向性",
        "・テレワーク・オンラインモニタリング",
        "・ケアプランデータ連携システム",
    ]),
]
col_w = Cm(7.7)
col_h = Cm(11.5)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
y = Cm(4.5)
for i, (cat, color, items) in enumerate(cats):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1.2), color)
    add_text(s, x + Cm(0.2), y + Cm(0.15), col_w - Cm(0.4), Cm(1),
             cat, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1.2), col_w, col_h - Cm(1.2), TABLE_BG)
    body = "\n\n".join(items)
    add_text(s, x + Cm(0.3), y + Cm(1.4), col_w - Cm(0.6), col_h - Cm(1.6),
             body, size=11, color=DARK_GRAY, line_spacing=1.25)

add_text(s, Cm(1.2), Cm(16.5), SW - Cm(2.4), Cm(1),
         "▶ 本日は左 2 列(加算 / 減算)を中心に、右 2 列(基準 / ICT)を R9 論点として扱う。",
         size=13, color=NAVY, bold=True)


# =============================================================
#  SLIDE 7: 第2部
# =============================================================
section_title(2, "特定事業所加算",
              "Designated Office Bonus: Requirements and Reality")


# =============================================================
#  SLIDE 8: 特定事業所加算の位置づけ
# =============================================================
s = add_slide()
add_header(s, "2. 特定事業所加算", "特定事業所加算とは ― 「質の高いケアマネジメント」の評価")
add_footer(s, 8)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(6),
        "加算の趣旨",
        "■ 中重度者・支援困難ケースへの積極的対応\n"
        "■ 主任CMを中心とした計画的研修・指導体制\n"
        "■ 24時間連絡体制と緊急時の即応性\n"
        "■ 地域包括支援センター・他機関との連携\n"
        "■ 多職種・多サービス事業者との協働\n\n"
        "→ 個別ケースの質を担保する 体制 への評価",
        title_bg=STEEL, body_size=13)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6),
        "事業所への影響",
        "■ 経営的影響:大\n"
        "    収入の数% 〜 10%超に達することも\n\n"
        "■ 人材的影響:大\n"
        "    主任CMの配置・研修計画・採用戦略\n\n"
        "■ 業務的影響:大\n"
        "    24時間体制・困難事例の受入・記録整備\n\n"
        "▶ 「取る/取らない」の判断は事業所戦略そのもの",
        title_bg=ACCENT_GREEN, body_bg=GREEN_BG, body_size=13)

add_box(s, Cm(1.2), Cm(11), Cm(31.4), Cm(6),
        "主任CMの役割 ― 特定事業所加算における要(かなめ)",
        "■ 配置要件:加算Ⅰ・Ⅱで複数名、Ⅲ・Ⅳで1名以上(区分により異なる)\n"
        "■ 機能要件:CMの指導・教育、計画的な研修、事例検討の主催\n"
        "■ 困難事例の受け持ち:他機関との調整役、虐待・身寄りなしへの対応\n"
        "■ 地域貢献:地域ケア会議への参加、他事業所への助言\n\n"
        "▶ 主任CMが「配置」されるだけでは不十分。実質的に機能していることが要件。",
        title_bg=NAVY, body_size=13)


# =============================================================
#  SLIDE 9: 区分の整理
# =============================================================
s = add_slide()
add_header(s, "2. 特定事業所加算", "特定事業所加算 Ⅰ〜Ⅳ の区分整理")
add_footer(s, 9)

col_widths = [Cm(4.5), Cm(7), Cm(7), Cm(7), Cm(5.4)]
headers = ["区分", "主任CM 配置", "常勤CM 配置", "中重度者割合", "想定する事業所"]
rows = [
    ["Ⅰ",  "複数名",          "複数名",          "高水準",        "大規模・困難ケース\n主軸事業所"],
    ["Ⅱ",  "複数名",          "複数名",          "中水準",        "中規模・地域\n中核事業所"],
    ["Ⅲ",  "1 名以上",        "複数名",          "一定水準",      "中小規模・体制\n整備中事業所"],
    ["Ⅳ",  "1 名以上",        "1 名以上",        "要件あり",      "新規取得を目指す\n小規模事業所"],
]
add_table(s, Cm(0.7), Cm(4.5), col_widths, headers, rows,
          header_bg=NAVY, first_col_color=ACCENT_RED, row_height=Cm(1.8),
          body_size=12)

add_box(s, Cm(1.2), Cm(13.5), Cm(31.4), Cm(3.5),
        "区分共通の主要要件",
        "■ 24時間連絡可能な体制  ■ 計画的な研修(年度計画)  ■ 事例検討の実施(自・他事業所も含む)\n"
        "■ 運営基準減算がない  ■ 居宅サービス計画の届出(条文上の要件)\n"
        "■ 必要に応じた利用者宅訪問・他機関連携  ■ 地域包括支援センターからの相談受け入れ実績",
        title_bg=STEEL, body_size=12)


# =============================================================
#  SLIDE 10: 算定要件の詳細チェック
# =============================================================
s = add_slide()
add_header(s, "2. 特定事業所加算", "算定要件チェックリスト ― 共通要件")
add_footer(s, 10)

items = [
    ("人員配置",     "主任CM・常勤CMの配置基準を満たしているか(区分ごとに異なる)"),
    ("研修計画",     "個別CMの研修計画を策定し、実施・記録しているか"),
    ("会議の定例化", "利用者情報の伝達・取扱方針について定期的な会議を開催しているか"),
    ("事例検討",     "事業所内・他事業所も含めた事例検討会を定期的に実施しているか"),
    ("24時間体制",   "夜間・休日の連絡体制を文書で定め、利用者・家族に周知しているか"),
    ("運営基準減算", "現在算定していない(過去 6 月以内に減算を受けていない)"),
    ("中重度者割合", "区分ごとの基準(要介護3以上の利用者割合)を満たしているか"),
    ("計画届出",     "居宅サービス計画の総数のうち、地域包括等から依頼を受けた件数の把握"),
    ("実地指導対応", "上記の運用状況を書面で記録し、いつでも提示できるか"),
]
y = Cm(4.5)
for i, (k, v) in enumerate(items):
    bg = TABLE_BG if i % 2 == 0 else WHITE
    add_rect(s, Cm(1.2), y, Cm(31.4), Cm(1.2), bg, line_color=MID_GRAY)
    # チェック
    add_rect(s, Cm(1.5), y + Cm(0.25), Cm(0.7), Cm(0.7), WHITE, line_color=NAVY)
    add_text(s, Cm(2.4), y + Cm(0.2), Cm(5.5), Cm(0.8),
             k, size=12, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Cm(8), y + Cm(0.2), SW - Cm(9), Cm(0.8),
             v, size=11, color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE)
    y += Cm(1.25)

add_text(s, Cm(1.2), Cm(16.5), SW - Cm(2.4), Cm(1),
         "▶ R6 改定で「実質的に機能しているか」の確認が強化。形だけの会議・研修は監査リスク。",
         size=12, color=ACCENT_RED, bold=True)


# =============================================================
#  SLIDE 11: R6改定での変更点
# =============================================================
s = add_slide()
add_header(s, "2. 特定事業所加算", "R6 改定で変わった/明確化された点")
add_footer(s, 11)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(6.5),
        "明確化・強化された要件",
        "■ 多様な主体等が提供する生活支援サービスの活用\n"
        "    地域資源マップ作成・更新を求める運用が増加\n\n"
        "■ ヤングケアラー等を含む家族介護者への支援\n"
        "    アセスメント項目への明示\n\n"
        "■ 他法他制度の活用(障害・医療・成年後見・生活困窮)\n"
        "    連携の実績を記録に残す\n\n"
        "■ 必要に応じた地域包括ケア会議への参加",
        title_bg=ACCENT_RED, body_bg=RED_BG, body_size=12)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6.5),
        "運用上の留意事項",
        "■ BCP・虐待防止の体制整備が前提に\n"
        "   (未整備事業所は加算継続困難)\n\n"
        "■ 主任CMの「専門研修・更新研修」修了\n"
        "    が引き続き要件として運用される\n\n"
        "■ 「困難事例の受入」は数だけでなく\n"
        "    アセスメント・支援経過の質も問われる\n\n"
        "■ 地域包括との連携実績は\n"
        "    具体的件数・内容を記録",
        title_bg=ACCENT_AMB, body_bg=AMB_BG, body_size=12)

add_box(s, Cm(1.2), Cm(11.5), Cm(31.4), Cm(5.5),
        "主任CMが押さえるべき変化のシグナル",
        "■ 「加算 = 報酬」ではなく 「加算 = 質の証明」 への転換が R6 の通底テーマ\n"
        "■ 「やったか/やっていないか」より 「どう機能したか」 を問われる方向\n"
        "■ R9 ではさらに、 アウトカム評価 (利用者の生活機能維持・再入院抑制等)が論点化する見込み\n"
        "■ 主任CMは、「加算を取るための作業」を超えて「事業所の質を可視化する道具」として位置づけ直す",
        title_bg=NAVY, body_size=13)


# =============================================================
#  SLIDE 12: 算定実態と落とし穴
# =============================================================
s = add_slide()
add_header(s, "2. 特定事業所加算", "算定実態と「落とし穴」 ― 実地指導で指摘されやすい点")
add_footer(s, 12)

col_widths = [Cm(8), Cm(11), Cm(12.4)]
headers = ["落とし穴", "ありがちな実態", "監査・実地指導の指摘例"]
rows = [
    ["研修計画\nの形骸化",
     "年度初めに作って、\nそれっきり",
     "「個別CMの研修記録がない」\n「年度末に振り返り・修正がない」"],
    ["事例検討の\n形だけ実施",
     "短時間・参加者偏り\n司会も主任CM一人",
     "「議事録に学びが書かれていない」\n「参加CMが固定化している」"],
    ["24時間体制\nの不十分",
     "携帯転送のみ・\n利用者に周知不十分",
     "「連絡先が利用者宅にない」\n「夜間応答記録が残っていない」"],
    ["中重度者割合\nの記録不備",
     "毎月の確認なし\n年度途中で低下",
     "「割合計算の根拠書類がない」\n「割合低下時の対応記録がない」"],
    ["主任CMの\n機能不全",
     "配置はあるが\n指導実態がない",
     "「個別ケース指導記録がない」\n「同行訪問・スーパーバイズ記録がない」"],
]
add_table(s, Cm(1.2), Cm(4.5), col_widths, headers, rows,
          header_bg=ACCENT_RED, first_col_color=ACCENT_RED,
          row_height=Cm(2.0), body_size=11, header_size=12)

add_text(s, Cm(1.2), Cm(16), SW - Cm(2.4), Cm(1.5),
         "▶ 「書類が揃っている」は最低条件。「中身が伴っている」を主任CM自身が定期的に点検する。\n"
         "▶ 自主点検 → 内部監査 → 外部相談(行政・職能団体)の 3 段階チェックを推奨。",
         size=12, color=NAVY, bold=True, line_spacing=1.3)


# =============================================================
#  SLIDE 13: 取得戦略
# =============================================================
s = add_slide()
add_header(s, "2. 特定事業所加算", "未取得事業所への助言 ― 段階的取得戦略")
add_footer(s, 13)

steps = [
    ("Phase 1\n基盤整備",
     "0〜6ヶ月",
     "・運営基準減算ゼロの徹底\n・記録様式の統一・電子化\n・BCP/虐待防止の体制整備"),
    ("Phase 2\nⅣ 取得",
     "6〜12ヶ月",
     "・主任CM 1名配置\n・24時間体制構築\n・研修計画・事例検討の定例化"),
    ("Phase 3\nⅢ → Ⅱ",
     "12〜24ヶ月",
     "・常勤CM体制整備\n・中重度者割合の確保\n・地域連携実績の蓄積"),
    ("Phase 4\nⅡ → Ⅰ",
     "24ヶ月〜",
     "・困難事例の積極受入\n・地域貢献の言語化\n・主任CM 複数化"),
]
col_w = Cm(7.7)
col_h = Cm(11)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
y = Cm(4.5)
colors = [STEEL, ACCENT_GREEN, ACCENT_AMB, ACCENT_RED]
for i, (label, time, body) in enumerate(steps):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(2.5), colors[i])
    add_text(s, x + Cm(0.2), y + Cm(0.2), col_w - Cm(0.4), Cm(1.5),
             label, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Cm(0.2), y + Cm(1.7), col_w - Cm(0.4), Cm(0.8),
             time, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER)
    add_rect(s, x, y + Cm(2.5), col_w, col_h - Cm(2.5), TABLE_BG)
    add_text(s, x + Cm(0.3), y + Cm(2.7), col_w - Cm(0.6), col_h - Cm(2.7),
             body, size=11, color=DARK_GRAY, line_spacing=1.3)

add_text(s, Cm(1.2), Cm(16), SW - Cm(2.4), Cm(1.5),
         "▶ 一気にⅠを狙わず、段階を踏む。\n"
         "▶ 主任CM がいない事業所では「主任研修受講者の確保」が最優先課題。",
         size=13, color=NAVY, bold=True, line_spacing=1.3)


# =============================================================
#  SLIDE 14: 第3部
# =============================================================
section_title(3, "ターミナルケアマネジメント加算",
              "Terminal Care Management Bonus")


# =============================================================
#  SLIDE 15: 加算の概要
# =============================================================
s = add_slide()
add_header(s, "3. ターミナルCM加算", "加算の概要と算定要件の構造")
add_footer(s, 15)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(4),
        "加算の趣旨",
        "■ 末期がん等の利用者が、住み慣れた自宅で人生の最終段階を過ごせるよう、\n"
        "    CMが頻回の訪問・多職種連携・看取り後の対応まで一貫して支える業務を評価\n\n"
        "■ R6 改定では算定要件が明確化され、記録の整備がより重要に",
        title_bg=ACCENT_RED, body_bg=RED_BG, body_size=13)

# 4要件
items = [
    ("①", "対象利用者",
     "末期の悪性腫瘍その他、\n医師が一般的に認められて\nいる医学的知見に基づき\n回復の見込みがないと診断"),
    ("②", "死亡日 + 死亡前\n14 日以内の対応",
     "本人・家族の同意のもと、\n2回以上の利用者宅訪問\n(状態変化等の理由必須)"),
    ("③", "多職種連携",
     "主治医・看護師等との\n連携、サービス担当者会議\nまたは個別の情報共有"),
    ("④", "看取り後の対応",
     "本人の死亡日も含めた\n対応を行い、\n記録を整備"),
]
col_w = Cm(7.7)
col_h = Cm(7.5)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
y = Cm(9)
for i, (num, title, body) in enumerate(items):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(2.2), ACCENT_RED)
    add_text(s, x, y + Cm(0.2), col_w, Cm(0.8),
             num, size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, x + Cm(0.2), y + Cm(1), col_w - Cm(0.4), Cm(1.2),
             title, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(2.2), col_w, col_h - Cm(2.2), RED_BG)
    add_text(s, x + Cm(0.2), y + Cm(2.4), col_w - Cm(0.4), col_h - Cm(2.6),
             body, size=11, color=DARK_GRAY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.3)


# =============================================================
#  SLIDE 16: 実務フロー
# =============================================================
s = add_slide()
add_header(s, "3. ターミナルCM加算", "実務フロー ― 末期診断 〜 看取り後まで")
add_footer(s, 16)

# 縦のフロー
phases = [
    ("Step 1", "末期診断・在宅看取りの方針確認",
     "・主治医からの情報収集(末期の旨の確認)\n・本人・家族の意向確認(在宅看取り希望)\n・支援体制の見直し(医療系サービス導入検討)"),
    ("Step 2", "ターミナル期のケアプラン再作成",
     "・サービス担当者会議の開催\n・訪問看護・訪問診療との連携強化\n・看取り場所・救急時の方針を明文化"),
    ("Step 3", "頻回訪問と多職種共有",
     "・状態変化に応じた訪問(2回以上必須)\n・訪問の都度、訪問理由・状態変化を記録\n・主治医・看護師との情報共有(電話・FAX・ICT)"),
    ("Step 4", "死亡日対応 + 看取り後支援",
     "・死亡確認後の家族支援(死後事務・グリーフケア)\n・関係事業所への連絡・サービス終了処理\n・記録の最終整備(算定根拠の確定)"),
]
y = Cm(4.5)
for i, (step, title, body) in enumerate(phases):
    color = [STEEL, ACCENT_AMB, ACCENT_RED, ACCENT_GREEN][i]
    # ステップラベル
    add_rect(s, Cm(1.2), y, Cm(4), Cm(2.5), color)
    add_text(s, Cm(1.2), y + Cm(0.2), Cm(4), Cm(1),
             step, size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Cm(1.2), y + Cm(1.1), Cm(4), Cm(1.3),
             title, size=11, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 内容
    add_rect(s, Cm(5.4), y, SW - Cm(6.6), Cm(2.5), TABLE_BG)
    add_text(s, Cm(5.7), y + Cm(0.2), SW - Cm(7), Cm(2.1),
             body, size=12, color=DARK_GRAY, line_spacing=1.3)
    y += Cm(2.8)

add_text(s, Cm(1.2), Cm(16), SW - Cm(2.4), Cm(1),
         "▶ 算定要件は「死亡日 + 死亡前 14 日以内」。この期間の業務記録が決め手。",
         size=13, color=ACCENT_RED, bold=True)


# =============================================================
#  SLIDE 17: 算定漏れチェックリスト
# =============================================================
s = add_slide()
add_header(s, "3. ターミナルCM加算", "算定漏れを防ぐ 8 つのチェックポイント")
add_footer(s, 17)

checks = [
    "末期である旨が、主治医意見書または診療情報提供書で 文書 で確認できるか",
    "本人・家族の意向確認の記録があるか(在宅看取りに同意した旨)",
    "死亡日を含む 14 日以内に 2 回以上の訪問記録があるか",
    "各訪問記録に「訪問理由」「状態変化」「対応内容」が明記されているか",
    "サ担会議または個別連携の記録があるか(訪問看護・訪問診療と)",
    "ターミナル期のケアプランが再作成されているか(変更も含む)",
    "死亡日の対応記録があるか(電話対応・関係機関への連絡も含む)",
    "請求月の算定根拠が、上記書類で説明できる状態に整理されているか",
]
y = Cm(4.3)
for i, c in enumerate(checks):
    bg = TABLE_BG if i % 2 == 0 else WHITE
    add_rect(s, Cm(1.2), y, Cm(31.4), Cm(1.4), bg, line_color=MID_GRAY)
    # 番号
    circle = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(1.5), y + Cm(0.25),
                                Cm(0.9), Cm(0.9))
    circle.fill.solid()
    circle.fill.fore_color.rgb = ACCENT_RED
    circle.line.fill.background()
    tf = circle.text_frame
    tf.margin_top = Cm(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(i + 1)
    r.font.name = FONT
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = WHITE
    # チェック枠
    add_rect(s, Cm(2.8), y + Cm(0.3), Cm(0.7), Cm(0.7), WHITE, line_color=NAVY)
    # 本文
    add_text(s, Cm(3.8), y + Cm(0.2), SW - Cm(5), Cm(1),
             c, size=12, color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE)
    y += Cm(1.45)


# =============================================================
#  SLIDE 18: 連携先と記録
# =============================================================
s = add_slide()
add_header(s, "3. ターミナルCM加算", "連携先別 ― 共有すべき情報と記録の例")
add_footer(s, 18)

col_widths = [Cm(7), Cm(12), Cm(12.4)]
headers = ["連携先", "共有すべき情報", "記録として残す内容"]
rows = [
    ["主治医 / 訪問診療",
     "状態変化、疼痛コントロール、\n看取り場所の意向、急変時対応",
     "情報提供日、伝達手段(電話/FAX/ICT)、\n伝達内容、医師からの指示"],
    ["訪問看護",
     "バイタル変化、本人・家族の不安、\n医療処置の内容、緊急時連絡網",
     "共有日時、共有手段、共有内容、\n看護師からの助言"],
    ["薬剤師 (在宅)",
     "麻薬・疼痛薬の調整、\n服薬困難への対応",
     "情報提供と薬剤調整の経過、\n薬剤師訪問日との連動"],
    ["訪問介護",
     "本人の状態に応じた介護方法、\n家族のレスパイトニーズ",
     "サービス内容変更の記録、\n担当ヘルパー教育の記録"],
    ["家族 (キーパーソン)",
     "状態の見通し、看取りの心構え、\n医療処置の意思決定",
     "面談記録、意思決定支援の経過、\nグリーフケアの初期対応"],
]
add_table(s, Cm(1.2), Cm(4.5), col_widths, headers, rows,
          header_bg=NAVY, first_col_color=ACCENT_RED,
          row_height=Cm(2.1), body_size=11, header_size=12)


# =============================================================
#  SLIDE 19: ACPとの統合
# =============================================================
s = add_slide()
add_header(s, "3. ターミナルCM加算", "ACP ・ 人生会議との統合 ― 加算は「結果」")
add_footer(s, 19)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(3.5),
        "考え方の転換",
        "「加算を取りに行く」のではなく、\n"
        "「本人の意思を尊重した看取り支援を行った結果として、加算要件を満たす」が本来の姿。",
        title_bg=NAVY, body_size=14)

add_box(s, Cm(1.2), Cm(8.5), Cm(15.5), Cm(8.5),
        "ACP の早期着手が鍵",
        "■ 末期診断より前から、本人の価値観・希望を聴く\n\n"
        "■ 「もしもの時の話」を進行に応じて深める\n\n"
        "■ 看取りの場所・医療処置の希望を文書化\n\n"
        "■ 厚労省ガイドライン\n   「人生の最終段階における医療・ケアの\n    決定プロセスに関するガイドライン」を活用\n\n"
        "■ ACP の記録が、ターミナル期の判断根拠になる",
        title_bg=STEEL, body_size=12)

add_box(s, Cm(17.5), Cm(8.5), Cm(15), Cm(8.5),
        "事業所内で整備したい体制",
        "■ ターミナル対応 CM の経験不足を補う仕組み\n   (主任CMによる同行・スーパーバイズ)\n\n"
        "■ 24時間体制(特定事業所加算と連動)\n\n"
        "■ グリーフケア研修の実施\n\n"
        "■ 看取り後の振り返り会議(デスカンファ)\n\n"
        "■ 主治医・訪問看護との顔の見える関係づくり\n   (平時からの連携)",
        title_bg=ACCENT_GREEN, body_bg=GREEN_BG, body_size=12)


# =============================================================
#  SLIDE 20: 第4部
# =============================================================
section_title(4, "その他 R6 改定の論点",
              "Other Key Points of the FY2024 Revision",
              accent_color=ACCENT_AMB)


# =============================================================
#  SLIDE 21: 同一建物減算
# =============================================================
s = add_slide()
add_header(s, "4. その他 R6 改定", "同一建物減算 ― 居宅介護支援への新設")
add_footer(s, 21)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(4),
        "改定の趣旨",
        "■ サービス付き高齢者向け住宅(サ高住)・有料老人ホーム等に居住する利用者を、\n"
        "    同一建物に所在する事業所のCMが多数担当する事例への対応\n"
        "■ 公正中立性・移動時間の効率性を踏まえた評価",
        title_bg=ACCENT_RED, body_bg=RED_BG, body_size=13)

col_widths = [Cm(10), Cm(11), Cm(10.4)]
headers = ["対象", "減算の考え方", "実務上の留意点"]
rows = [
    ["同一建物に居住する\n利用者が一定割合\n以上",
     "居宅介護支援費を\n所定の割合で減算",
     "毎月の居住者状況の確認\n建物単位での集計が必要"],
    ["対象外:\n例外規定あり",
     "やむを得ない事情等が\n認められる場合の扱い",
     "個別判断のため\n指定権者に確認が必要"],
    ["集合住宅で\n複数事業所が分担",
     "他事業所と分担している\n場合の取扱いも整理",
     "事業所間の調整\n中立性確保の証跡"],
]
add_table(s, Cm(1.2), Cm(9), col_widths, headers, rows,
          header_bg=NAVY, first_col_color=ACCENT_RED,
          row_height=Cm(2), body_size=11)

add_text(s, Cm(1.2), Cm(15.7), SW - Cm(2.4), Cm(1.5),
         "▶ サ高住併設の居宅介護支援事業所では、経営インパクトが大きい。\n"
         "▶ R9 改定では更に強化される可能性 ― 公正中立性は継続論点。",
         size=12, color=NAVY, bold=True, line_spacing=1.3)


# =============================================================
#  SLIDE 22: 逓減制とICT
# =============================================================
s = add_slide()
add_header(s, "4. その他 R6 改定", "逓減制 ― ICT 活用 / 事務職員配置による緩和")
add_footer(s, 22)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(3.5),
        "逓減制の基本構造",
        "■ 1人のCMが担当する利用者数が増えると、所定単位数を逓減\n"
        "■ ただし、ICT活用 または 事務職員配置 を行う事業所は、逓減の開始件数を緩和\n"
        "■ R6 改定で、ICT 活用要件の具体的な内容が明確化された",
        title_bg=STEEL, body_size=13)

add_box(s, Cm(1.2), Cm(8.2), Cm(15.5), Cm(8.5),
        "ICT 活用要件 ― 何をすればよいか",
        "■ ケアプラン作成・記録・モニタリングの一連の業務\n   を支援するシステムの導入\n\n"
        "■ 例:\n   ・ケアプランデータ連携システム\n   ・記録のクラウド化(訪問先からの記録)\n   ・スケジュール管理ソフト\n   ・モバイル端末での情報共有\n\n"
        "■ システム導入だけでは不十分\n   実際に業務効率化に寄与している証跡が必要",
        title_bg=ACCENT_GREEN, body_bg=GREEN_BG, body_size=12)

add_box(s, Cm(17.5), Cm(8.2), Cm(15), Cm(8.5),
        "事務職員配置要件",
        "■ 常勤換算で一定割合の事務職員配置\n\n"
        "■ 事務職員が行う業務の例:\n   ・利用者ファイルの管理\n   ・請求業務\n   ・各種書類の作成補助\n   ・電話応対・予約調整\n\n"
        "■ CMが本来業務に集中できる体制を作る\n\n"
        "■ 業務分担表の整備が望ましい",
        title_bg=ACCENT_AMB, body_bg=AMB_BG, body_size=12)


# =============================================================
#  SLIDE 23: BCP・虐待防止
# =============================================================
s = add_slide()
add_header(s, "4. その他 R6 改定", "BCP / 虐待防止 ― 経過措置終了後の本格運用")
add_footer(s, 23)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(6),
        "BCP (業務継続計画)",
        "■ 感染症 + 自然災害の双方を網羅\n\n"
        "■ 必要な構成要素\n   ・基本方針 / 推進体制\n   ・リスクの把握 / 優先業務の選定\n   ・必要な人員・物資・設備の確保\n   ・職員の訓練・研修\n\n"
        "■ 年 1 回以上の研修・訓練 と 計画の見直し\n\n"
        "■ 未策定は減算対象(経過措置終了後)",
        title_bg=ACCENT_RED, body_bg=RED_BG, body_size=12)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6),
        "高齢者虐待防止",
        "■ 委員会の設置・指針の整備・研修の実施\n   ・担当者の選定が必須\n\n"
        "■ 利用者からの相談窓口の周知\n\n"
        "■ 養介護施設従事者等による虐待 + 養護者\n   による虐待、両方への対応体制\n\n"
        "■ 通報義務(法律)を職員全員が理解\n\n"
        "■ 未整備は減算対象",
        title_bg=ACCENT_AMB, body_bg=AMB_BG, body_size=12)

add_box(s, Cm(1.2), Cm(11), Cm(31.4), Cm(6),
        "主任CMが事業所内で取りまとめる役割",
        "■ BCP の「絵に描いた餅」化を防ぐ ― 訓練後の振り返り・改善のサイクルを回す\n\n"
        "■ 虐待防止委員会の事例検討で、後輩CMが「気になるサイン」を共有できる文化を作る\n\n"
        "■ 介護現場のハラスメント(利用者・家族から職員へ)も含めた相談窓口の機能化\n\n"
        "■ 経過措置終了後、減算事業所として指摘を受けないための内部点検を年度初めに実施",
        title_bg=NAVY, body_size=13)


# =============================================================
#  SLIDE 24: 入退院時連携加算
# =============================================================
s = add_slide()
add_header(s, "4. その他 R6 改定", "入退院時連携加算 ・ 通院時情報連携加算")
add_footer(s, 24)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(6.5),
        "入院時情報連携加算",
        "■ 利用者が入院した際、医療機関へ\n   利用者情報を提供することで算定\n\n"
        "■ R6 改定での区分整理\n   ・入院後 3 日以内に情報提供\n   ・上記より遅い情報提供\n   の区分整理(評価に差)\n\n"
        "■ 「面談」または「文書/電子的方法」での提供\n\n"
        "■ 情報提供日と方法を必ず記録\n   (相手の受領サインがあるとなお良い)",
        title_bg=STEEL, body_size=12)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(6.5),
        "通院時情報連携加算",
        "■ CM が利用者の通院に同行し、医師等\n   と情報共有することで算定\n\n"
        "■ 算定要件\n   ・利用者・家族の同意\n   ・通院同席または医師との情報共有\n   ・記録(医師に伝達した内容・指示の受領)\n\n"
        "■ 頻度の上限あり(月 1 回等)\n\n"
        "■ 認知症ケース・複数科受診ケースで\n   活用余地が大きい",
        title_bg=ACCENT_GREEN, body_bg=GREEN_BG, body_size=12)

add_box(s, Cm(1.2), Cm(11.5), Cm(31.4), Cm(5.5),
        "算定漏れを防ぐコツ",
        "■ 入院連絡を受けたら、その場で「3 日以内に情報提供」をタスク化 ― 期限管理が命\n"
        "■ 通院同行は、家族からの依頼ベースで日常化しているケースが多い ― 算定可能性を再点検\n"
        "■ 「やったけど算定していない」を年度末に棚卸し ― 主任CMの仕事として位置づける",
        title_bg=NAVY, body_size=13)


# =============================================================
#  SLIDE 25: LIFEと居宅
# =============================================================
s = add_slide()
add_header(s, "4. その他 R6 改定", "LIFE と居宅介護支援 ― 「次の波」への準備")
add_footer(s, 25)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(3.5),
        "現状(R6 時点)",
        "■ 居宅介護支援は LIFE 提出 を直接の加算要件としていない\n"
        "■ ただし、サービス事業所(訪問介護・通所介護等)では LIFE 連動の加算が拡大\n"
        "■ CMは、各サービス事業所からの LIFE フィードバックを ケアプランに反映 する責務",
        title_bg=STEEL, body_size=13)

add_box(s, Cm(1.2), Cm(8.2), Cm(15.5), Cm(8.5),
        "今、CMが行うべきこと",
        "■ サービス事業所が提出している LIFE\n   データの内容を把握する\n\n"
        "■ フィードバック票の確認\n   ・科学的介護推進体制加算等\n   ・ADL/IADL/認知/栄養の推移\n\n"
        "■ ケアプラン見直しに反映する仕組み\n   ・モニタリング項目への追加\n   ・サ担会議の議題化\n\n"
        "■ 「データに基づくケア」の入口に立つ",
        title_bg=ACCENT_GREEN, body_bg=GREEN_BG, body_size=12)

add_box(s, Cm(17.5), Cm(8.2), Cm(15), Cm(8.5),
        "R9 で予想される変化",
        "■ 居宅介護支援にも LIFE 連動の加算\n   が導入される可能性\n\n"
        "■ ケアプランデータ連携システムが\n   標準化されていく\n\n"
        "■ アウトカム評価(自立度の維持・改善等)\n   が報酬体系に組み込まれる議論\n\n"
        "■ CM の業務が「データを読む」「データで\n   説明する」方向にシフト\n\n"
        "■ ICT・データリテラシーが主任CMの\n   必須スキルに",
        title_bg=ACCENT_AMB, body_bg=AMB_BG, body_size=12)


# =============================================================
#  SLIDE 26: 第5部
# =============================================================
section_title(5, "R6改定の運用棚卸し",
              "Operational Self-Check",
              accent_color=ACCENT_GREEN)


# =============================================================
#  SLIDE 27: 自事業所点検チェックリスト
# =============================================================
s = add_slide()
add_header(s, "5. 運用棚卸し", "自事業所点検チェックリスト ― 5 領域 × 各項目")
add_footer(s, 27)

areas = [
    ("加算管理",
     ["特定事業所加算の区分維持要件を月次点検",
      "ターミナルCM加算の算定漏れチェック",
      "入退院時 / 通院時連携加算の算定実績把握"]),
    ("減算回避",
     ["運営基準減算の発生有無を毎月確認",
      "同一建物減算の対象者集計が毎月できているか",
      "BCP / 虐待防止 未整備による減算リスクなし"]),
    ("記録整備",
     ["居宅介護支援経過の記録が業務単位で残る",
      "サ担会議録の整備(本人参加・意向確認)",
      "アセスメント様式と23項目の対応の確認"]),
    ("人材育成",
     ["主任CM の計画的研修参加",
      "個別CM の研修計画と振り返り",
      "事例検討会の年間スケジュール"]),
    ("地域連携",
     ["地域包括との情報共有の定例化",
      "医療機関 / 居宅サービス事業所との連携記録",
      "地域ケア会議への参加実績"]),
]

col_w = Cm(6.2)
col_h = Cm(11)
gap = Cm(0.2)
total_w = col_w * 5 + gap * 4
start_x = (SW - total_w) // 2
y = Cm(4.5)
colors = [ACCENT_RED, ACCENT_AMB, STEEL, ACCENT_GREEN, NAVY]
for i, (area, items) in enumerate(areas):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1.2), colors[i])
    add_text(s, x + Cm(0.1), y + Cm(0.15), col_w - Cm(0.2), Cm(1),
             area, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1.2), col_w, col_h - Cm(1.2), TABLE_BG)
    for j, it in enumerate(items):
        item_y = y + Cm(1.5) + j * Cm(3)
        # チェック枠
        add_rect(s, x + Cm(0.2), item_y, Cm(0.5), Cm(0.5),
                 WHITE, line_color=NAVY)
        add_text(s, x + Cm(0.9), item_y - Cm(0.05), col_w - Cm(1), Cm(2.7),
                 it, size=10, color=DARK_GRAY, line_spacing=1.3)

add_text(s, Cm(1.2), Cm(16.5), SW - Cm(2.4), Cm(1),
         "▶ 主任CMが事業所内で月次・四半期・年次で点検を回す。",
         size=13, color=NAVY, bold=True)


# =============================================================
#  SLIDE 28: 第6部
# =============================================================
section_title(6, "令和9年度改定に向けた論点",
              "Toward the FY2027 Revision",
              accent_color=ACCENT_RED)


# =============================================================
#  SLIDE 29: R9 改定の俯瞰
# =============================================================
s = add_slide()
add_header(s, "6. R9 改定への論点", "R9 改定 ― 制度設計の大きな潮流")
add_footer(s, 29)

# 5論点
points = [
    ("論点 1", "ケアプラン作成料 の有料化議論",
     "・現在は 10 割給付。利用者負担導入の是非\n・公平性 vs 介護離れの懸念"),
    ("論点 2", "主任CM の位置づけ強化",
     "・管理者要件としての主任CM\n・配置加算・経験年数要件の見直し"),
    ("論点 3", "アウトカム評価 の本格導入",
     "・自立支援・重度化防止の数値化\n・LIFE データに基づく加算の拡大"),
    ("論点 4", "ICT / AI 活用 の評価",
     "・ケアプランデータ連携の標準化\n・記録 AI・予測アラートへの対応"),
    ("論点 5", "公正中立性 の更なる担保",
     "・同一建物減算の強化\n・特定事業者への偏り是正"),
]
y = Cm(4.5)
for i, (label, title, body) in enumerate(points):
    row_h = Cm(2.3)
    add_rect(s, Cm(1.2), y, Cm(3.5), row_h, ACCENT_RED)
    add_text(s, Cm(1.2), y + Cm(0.1), Cm(3.5), Cm(1),
             label, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER)
    add_text(s, Cm(1.2), y + Cm(1.0), Cm(3.5), Cm(1.2),
             "R9", size=22, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Cm(4.7), y, SW - Cm(5.9), row_h, TABLE_BG)
    add_text(s, Cm(5), y + Cm(0.2), SW - Cm(6.5), Cm(1),
             title, size=15, bold=True, color=NAVY)
    add_text(s, Cm(5), y + Cm(1.15), SW - Cm(6.5), row_h - Cm(1.2),
             body, size=11, color=DARK_GRAY, line_spacing=1.3)
    y += Cm(2.5)


# =============================================================
#  SLIDE 30: ケアプラン作成料の論点
# =============================================================
s = add_slide()
add_header(s, "6. R9 改定への論点", "論点 1 ― ケアプラン作成料の有料化議論")
add_footer(s, 30)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(8.5),
        "推進側の論拠",
        "■ 公平性\n   他の介護サービスは利用者負担あり\n   居宅介護支援だけ 10 割給付の是正\n\n"
        "■ 質の向上\n   利用者がCMを選ぶインセンティブ\n   競争を通じた質の底上げ\n\n"
        "■ 制度の持続可能性\n   介護給付費の伸びを抑制\n\n"
        "■ 海外比較\n   独・仏では有料が一般的",
        title_bg=STEEL, body_size=12)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(8.5),
        "慎重論",
        "■ 介護離れの懸念\n   負担を嫌い、CM を介さない自己作成\n   増加 → 質の低下リスク\n\n"
        "■ 公正中立性への影響\n   利用者の経済力により選別される\n   サービスの偏在\n\n"
        "■ 制度趣旨との整合\n   「いつでも誰でも相談できる窓口」\n   としての機能の毀損\n\n"
        "■ 自己作成の運用負荷\n   保険者・包括への負担増",
        title_bg=ACCENT_RED, body_bg=RED_BG, body_size=12)

add_box(s, Cm(1.2), Cm(13.5), Cm(31.4), Cm(3.5),
        "現実的なシナリオ予測",
        "■ R9 で 全面有料化 の可能性は低いが、 部分導入(高所得者・特定事業所利用者等) の議論は加速\n"
        "■ 段階導入・低額負担からの試行も視野\n"
        "■ いずれにしても、 CM の選ばれる存在 への変化が問われる",
        title_bg=NAVY, body_size=12)


# =============================================================
#  SLIDE 31: 主任CMの位置づけ
# =============================================================
s = add_slide()
add_header(s, "6. R9 改定への論点", "論点 2 ― 主任 CM の位置づけ強化の方向性")
add_footer(s, 31)

cards = [
    ("管理者要件",
     "現状:R9 以降、管理者は\n主任CM が原則(経過措置中)\n\n"
     "→ 経過措置の終了タイミング\n→ 移行猶予の延長議論"),
    ("配置加算",
     "特定事業所加算における\n主任CM 配置の要件強化\n\n"
     "→ 複数配置の評価拡大\n→ 機能要件(指導実態)の明確化"),
    ("育成 / 研修",
     "主任研修 ・ 更新研修の\n質と量の見直し\n\n"
     "→ ICT 活用研修の追加\n→ 法定研修体系の再編議論"),
    ("地域での役割",
     "地域包括ケアの中核\n地域ケア会議の運営支援\n\n"
     "→ 報酬体系での評価\n→ 兼任・派遣の柔軟化"),
]
col_w = Cm(7.7)
col_h = Cm(10)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
y = Cm(4.5)
colors = [ACCENT_RED, STEEL, ACCENT_AMB, ACCENT_GREEN]
for i, (title, body) in enumerate(cards):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1.2), colors[i])
    add_text(s, x + Cm(0.2), y + Cm(0.15), col_w - Cm(0.4), Cm(1),
             title, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1.2), col_w, col_h - Cm(1.2), TABLE_BG)
    add_text(s, x + Cm(0.3), y + Cm(1.4), col_w - Cm(0.6), col_h - Cm(1.6),
             body, size=11, color=DARK_GRAY, line_spacing=1.35)

add_box(s, Cm(1.2), Cm(15), Cm(31.4), Cm(2),
        "主任CMの「数」と「機能」の両立がR9 の焦点",
        "■ 全国的に主任CM が不足 → 養成スピード vs 質の維持のバランス\n"
        "■ 「機能していない主任CM」を生まない仕組みの議論",
        title_bg=NAVY, body_size=12)


# =============================================================
#  SLIDE 32: アウトカム評価
# =============================================================
s = add_slide()
add_header(s, "6. R9 改定への論点", "論点 3 ― アウトカム評価の本格導入")
add_footer(s, 32)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(4),
        "アウトカム評価とは",
        "■ プロセス評価(何を行ったか)から、アウトカム評価(どんな結果が出たか)への転換\n"
        "■ 例:ADL の維持 / 改善、認知機能の悪化抑制、入院抑制、自宅生活継続期間\n"
        "■ サービス事業所(訪問・通所等)で既に始まっており、居宅介護支援への波及が論点",
        title_bg=ACCENT_AMB, body_bg=AMB_BG, body_size=13)

add_box(s, Cm(1.2), Cm(9), Cm(15.5), Cm(8),
        "想定される指標例",
        "■ 利用者の自立度の維持・改善割合\n\n"
        "■ 計画的なサービス調整による\n   重度化抑制率\n\n"
        "■ 緊急入院・救急搬送の発生率\n\n"
        "■ 看取り対応の質(本人意向の尊重)\n\n"
        "■ 在宅生活継続日数",
        title_bg=STEEL, body_size=12)

add_box(s, Cm(17.5), Cm(9), Cm(15), Cm(8),
        "実装上の課題",
        "■ ケースミックスの調整\n   重度者の多い事業所が不利にならない仕組み\n\n"
        "■ 環境要因の切り分け\n   家族介護力・住環境の影響をどう扱うか\n\n"
        "■ 「卒業」と「維持」の評価の両立\n   要支援→自立も、要介護→現状維持も評価\n\n"
        "■ データ入力の負担",
        title_bg=ACCENT_RED, body_bg=RED_BG, body_size=12)


# =============================================================
#  SLIDE 33: ICT/AI
# =============================================================
s = add_slide()
add_header(s, "6. R9 改定への論点", "論点 4 ― ICT / AI 活用と業務の再定義")
add_footer(s, 33)

# ICTの3層
layers = [
    ("基盤層",
     "ケアプランデータ連携システム ・ クラウド記録 ・ 電子請求",
     "→ 多事業所間の情報共有 / 標準化",
     STEEL),
    ("活用層",
     "AI ケアプラン作成支援 ・ アセスメント補助 ・ リスク予測",
     "→ CM の判断を補助、業務時間短縮",
     ACCENT_GREEN),
    ("評価層",
     "LIFE 連動 ・ アウトカム分析 ・ ベンチマーク",
     "→ データに基づく PDCA、報酬連動",
     ACCENT_AMB),
]
y = Cm(4.5)
for i, (label, body, arrow, color) in enumerate(layers):
    add_rect(s, Cm(1.2), y, Cm(5), Cm(3), color)
    add_text(s, Cm(1.2), y + Cm(0.5), Cm(5), Cm(2),
             label, size=18, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Cm(6.2), y, SW - Cm(7.4), Cm(3), TABLE_BG)
    add_text(s, Cm(6.5), y + Cm(0.3), SW - Cm(8), Cm(1.2),
             body, size=14, bold=True, color=NAVY)
    add_text(s, Cm(6.5), y + Cm(1.5), SW - Cm(8), Cm(1.2),
             arrow, size=12, color=DARK_GRAY)
    y += Cm(3.3)

add_box(s, Cm(1.2), Cm(14.5), Cm(31.4), Cm(2.5),
        "R9 で予想される動き",
        "■ ICT 活用が「加算」ではなく「標準」になる ― 未対応の事業所は競争力を失う\n"
        "■ 主任CM には、ICT を活用した業務再設計能力 が求められる",
        title_bg=NAVY, body_size=12)


# =============================================================
#  SLIDE 34: 公正中立性
# =============================================================
s = add_slide()
add_header(s, "6. R9 改定への論点", "論点 5 ― 公正中立性 / 集中減算の深化")
add_footer(s, 34)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(3.5),
        "公正中立性が問われる構造",
        "■ 同一法人内のサービス事業所への利用集中\n"
        "■ サ高住併設の居宅介護支援事業所\n"
        "■ 営業活動の影響を受けたケアプラン",
        title_bg=ACCENT_RED, body_bg=RED_BG, body_size=13)

add_box(s, Cm(1.2), Cm(8.2), Cm(15.5), Cm(8.5),
        "R6 までに実装済みの規制",
        "■ 集中減算(特定事業所への集中時)\n\n"
        "■ 同一建物減算(R6 で居宅にも)\n\n"
        "■ 説明・同意の徹底\n   ・複数事業所の紹介\n   ・選定理由の説明と記録\n\n"
        "■ 居宅サービス計画の届出\n   (見直し制度)",
        title_bg=STEEL, body_size=12)

add_box(s, Cm(17.5), Cm(8.2), Cm(15), Cm(8.5),
        "R9 で議論される可能性",
        "■ 同一建物減算の強化(割合基準の更なる厳格化)\n\n"
        "■ 紹介加算 / 紹介減算の新設議論\n\n"
        "■ AI による営業バイアス検出\n   (実現性は議論段階)\n\n"
        "■ 苦情処理 ・ 第三者評価の活用\n\n"
        "■ 公正中立性に関する宣言 / 開示の義務化",
        title_bg=ACCENT_AMB, body_bg=AMB_BG, body_size=12)


# =============================================================
#  SLIDE 35: 同時改定の連動
# =============================================================
s = add_slide()
add_header(s, "6. R9 改定への論点", "同時改定 ― 医療 / 介護 / 障害 の連動視点")
add_footer(s, 35)

# 3制度の連動
add_rect(s, Cm(1.2), Cm(4.5), Cm(10), Cm(2), STEEL)
add_text(s, Cm(1.2), Cm(4.5), Cm(10), Cm(2),
         "医療(診療報酬)",
         size=18, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, Cm(11.6), Cm(4.5), Cm(10), Cm(2), ACCENT_RED)
add_text(s, Cm(11.6), Cm(4.5), Cm(10), Cm(2),
         "介護(介護報酬)",
         size=18, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, Cm(22), Cm(4.5), Cm(10), Cm(2), ACCENT_GREEN)
add_text(s, Cm(22), Cm(4.5), Cm(10), Cm(2),
         "障害(障害福祉)",
         size=18, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 連動論点
issues = [
    ("入退院連携",
     "退院時共同指導料(医療)と入院時情報連携加算(介護)の連動。\n"
     "病院 ・ CM ・ 在宅医療の三者がシームレスに動ける仕組み"),
    ("看取り",
     "在宅ターミナルケア加算(医療)とターミナルCM 加算(介護)の整合。\n"
     "ACP の早期実施が両制度で重視される方向"),
    ("精神 ・ 障害",
     "障害者の高齢化(65歳問題)。\n"
     "障害福祉サービス → 介護保険移行時のCMの役割"),
    ("認知症",
     "認知症の医療(専門医療)と介護(BPSD対応・意思決定支援)の連動。\n"
     "若年性認知症への両制度対応"),
]
y = Cm(7.5)
for i, (title, body) in enumerate(issues):
    bg = TABLE_BG if i % 2 == 0 else WHITE
    add_rect(s, Cm(1.2), y, Cm(31.4), Cm(2.2), bg, line_color=MID_GRAY)
    add_text(s, Cm(1.5), y + Cm(0.15), Cm(6), Cm(1.9),
             title, size=14, bold=True, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Cm(7.8), y + Cm(0.2), SW - Cm(9), Cm(1.9),
             body, size=11, color=DARK_GRAY,
             anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.3)
    y += Cm(2.3)


# =============================================================
#  SLIDE 36: 第7部
# =============================================================
section_title(7, "主任CMが持ち帰るもの",
              "Take Home")


# =============================================================
#  SLIDE 37: 持ち帰り3つ
# =============================================================
s = add_slide()
add_header(s, "Closing", "明日からの 3 つのアクション")
add_footer(s, 37)

actions = [
    ("Action 1", "事業所内 ― R6 改定の運用棚卸し",
     "・特定事業所加算 / ターミナルCM加算 の算定実態を月次で点検する\n"
     "・記録様式の整備と、後輩CM への教育機会を作る\n"
     "・BCP / 虐待防止 / 同一建物減算 の漏れがないか確認する"),
    ("Action 2", "地域 ― 制度動向の共有と議論",
     "・主任CM連絡会・地域ケア会議で R9 論点を共有する\n"
     "・他事業所の運用実態を聴き、ベンチマークする\n"
     "・行政・職能団体・地域包括 と顔の見える関係を維持する"),
    ("Action 3", "自分自身 ― 制度を読む力を磨く",
     "・社会保障審議会・介護給付費分科会の資料を定期的に確認する\n"
     "・職能団体(日本介護支援専門員協会等)の発信を追う\n"
     "・「制度を運用する」だけでなく「制度に意見する」立場を持つ"),
]
y = Cm(4.5)
colors = [STEEL, ACCENT_AMB, ACCENT_RED]
for i, (label, title, body) in enumerate(actions):
    color = colors[i]
    add_rect(s, Cm(1.2), y, Cm(4.5), Cm(3.7), color)
    add_text(s, Cm(1.2), y + Cm(0.3), Cm(4.5), Cm(1),
             label, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER)
    add_text(s, Cm(1.2), y + Cm(1.5), Cm(4.5), Cm(2),
             title.split(" ― ")[0] if " ― " not in title else title.split(" ― ")[0],
             size=20, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Cm(5.7), y, SW - Cm(6.9), Cm(3.7), TABLE_BG)
    add_text(s, Cm(6), y + Cm(0.2), SW - Cm(7.5), Cm(1),
             title, size=15, bold=True, color=NAVY)
    add_text(s, Cm(6), y + Cm(1.2), SW - Cm(7.5), Cm(2.5),
             body, size=12, color=DARK_GRAY, line_spacing=1.35)
    y += Cm(4)


# =============================================================
#  SLIDE 38: 参考資料 + クロージング
# =============================================================
s = add_slide()
add_header(s, "Reference", "参考資料 ・ 情報源")
add_footer(s, 38)

refs = [
    "■ 厚生労働省「令和6年度介護報酬改定について」(関係告示・通知・Q&A)",
    "■ 社会保障審議会 介護給付費分科会 議事録・配布資料",
    "■ 介護給付費単位数表 ・ 算定基準告示 ・ 解釈通知",
    "■ 日本介護支援専門員協会「主任介護支援専門員研修」関連資料",
    "■ 厚生労働省「適切なケアマネジメント手法の手引き」",
    "■ WAM NET「介護保険制度・報酬」情報",
    "■ 各都道府県・保険者の Q&A、集団指導資料",
    "■ R9 改定議論:介護給付費分科会(2025年〜2026年の審議)",
]
y = Cm(4.5)
for r in refs:
    add_text(s, Cm(2), y, SW - Cm(4), Cm(1),
             r, size=14, color=DARK_GRAY)
    y += Cm(1.2)

add_text(s, Cm(1.2), Cm(16), SW - Cm(2.4), Cm(1.5),
         "▶ 制度は動き続ける。主任CMは「最新の確からしい情報源」を常に複数持つこと。\n"
         "▶ 本日の資料は R6 施行2 年経過時点(2026年5月現在)の情報を基にしています。",
         size=12, color=ACCENT_RED, bold=True, line_spacing=1.3)


# ===== 保存 =====
output = "/home/user/kaigo-ga-app/slides/R6改定振り返り_R9改定論点_研修スライド.pptx"
prs.save(output)
print(f"Saved: {output}")
print(f"Total slides: {len(prs.slides)}")
