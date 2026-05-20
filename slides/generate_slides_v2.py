"""主任ケアマネ研修スライド生成スクリプト 第2弾(姉妹版)

テーマ: 適切なケアマネジメント手法の実践活用
重点疾患: 脳血管疾患 / 糖尿病
トーン: ストーリー型(2人の主人公の生活を追う構成)
時間: 90分
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ===== カラーパレット =====
NAVY     = RGBColor(0x1F, 0x3A, 0x5F)
TEAL     = RGBColor(0x2E, 0x86, 0xAB)
ORANGE   = RGBColor(0xE8, 0x7A, 0x41)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
DARK_GRAY = RGBColor(0x40, 0x40, 0x40)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT_BG = RGBColor(0xEA, 0xF3, 0xF9)
SOFT_ORANGE = RGBColor(0xFC, 0xEC, 0xDD)

# キャラクター固有色
STROKE_COLOR = RGBColor(0xC1, 0x4C, 0x4C)        # 脳血管(深紅)
STROKE_LIGHT = RGBColor(0xFB, 0xE9, 0xE9)
DM_COLOR     = RGBColor(0x4B, 0x86, 0x55)        # 糖尿病(緑)
DM_LIGHT     = RGBColor(0xE6, 0xF1, 0xE9)
CM_COLOR     = RGBColor(0x6A, 0x4C, 0x93)        # CM・対話(紫)
CM_LIGHT     = RGBColor(0xEC, 0xE5, 0xF5)

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


def add_round(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
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
    add_text(slide, Cm(0.8), Cm(0.25), Cm(20), Cm(0.6),
             section_label, size=11, color=WHITE, bold=True)
    add_text(slide, Cm(0.8), Cm(1.9), SW - Cm(1.6), Cm(1.3),
             title, size=24, bold=True, color=NAVY)
    add_rect(slide, Cm(0.8), Cm(3.2), Cm(2.5), Cm(0.12), ORANGE)


def add_footer(slide, page_num, total=38):
    add_text(slide, Cm(0.8), SH - Cm(0.8), Cm(22), Cm(0.5),
             "主任介護支援専門員研修 / 適切なケアマネジメント手法の実践活用 ― 脳血管疾患・糖尿病編",
             size=9, color=DARK_GRAY)
    add_text(slide, SW - Cm(3), SH - Cm(0.8), Cm(2.2), Cm(0.5),
             f"{page_num} / {total}", size=9, color=DARK_GRAY, align=PP_ALIGN.RIGHT)


def add_box(slide, left, top, width, height, title, body, *,
            title_bg=TEAL, body_bg=ACCENT_BG, title_color=WHITE,
            body_color=DARK_GRAY, title_size=14, body_size=12):
    add_rect(slide, left, top, width, Cm(1.0), title_bg)
    add_text(slide, left + Cm(0.2), top + Cm(0.1), width - Cm(0.4), Cm(0.8),
             title, size=title_size, bold=True, color=title_color,
             anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, left, top + Cm(1.0), width, height - Cm(1.0), body_bg)
    add_text(slide, left + Cm(0.3), top + Cm(1.1), width - Cm(0.6),
             height - Cm(1.1), body, size=body_size, color=body_color)


def add_dialog(slide, left, top, width, speaker, quote, *,
               speaker_color=NAVY, body_bg=ACCENT_BG, body_color=DARK_GRAY,
               size=13, height=Cm(2.5)):
    """セリフ吹き出し風ボックス"""
    add_text(slide, left, top, Cm(7), Cm(0.6),
             speaker, size=11, bold=True, color=speaker_color)
    add_round(slide, left, top + Cm(0.7), width, height - Cm(0.7), body_bg)
    add_text(slide, left + Cm(0.4), top + Cm(0.85), width - Cm(0.8), height - Cm(1),
             f"「{quote}」", size=size, color=body_color, italic=True,
             anchor=MSO_ANCHOR.MIDDLE)


def section_title(num, jp, en, *, accent_color=ORANGE):
    s = add_slide()
    add_rect(s, 0, 0, SW, SH, NAVY)
    add_rect(s, 0, Cm(8.5), SW, Cm(0.1), accent_color)
    add_text(s, Cm(2), Cm(6.5), SW - Cm(4), Cm(1.5),
             f"第 {num} 部", size=22, color=accent_color, bold=True,
             align=PP_ALIGN.CENTER)
    add_text(s, Cm(1), Cm(9.5), SW - Cm(2), Cm(2),
             jp, size=38, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(s, Cm(1), Cm(13), SW - Cm(2), Cm(1),
             en, size=14, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
    return s


def chapter_break(label, story_title, sub, *, accent=STROKE_COLOR):
    """主人公の章扉(物語の区切り)"""
    s = add_slide()
    add_rect(s, 0, 0, SW, SH, ACCENT_BG)
    add_rect(s, 0, 0, Cm(0.4), SH, accent)
    add_text(s, Cm(2), Cm(5), SW - Cm(4), Cm(1),
             label, size=14, bold=True, color=accent, align=PP_ALIGN.CENTER)
    add_text(s, Cm(1), Cm(7), SW - Cm(2), Cm(3),
             story_title, size=36, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(s, Cm(2), Cm(12), SW - Cm(4), Cm(3),
             sub, size=16, color=DARK_GRAY, align=PP_ALIGN.CENTER, line_spacing=1.4)
    return s


# =============================================================
#  SLIDE 1: 表紙
# =============================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, NAVY)
add_rect(s, 0, Cm(8), SW, Cm(0.15), ORANGE)
add_rect(s, 0, Cm(12.5), SW, Cm(0.05), TEAL)

add_text(s, Cm(2), Cm(4.5), SW - Cm(4), Cm(1),
         "主任介護支援専門員 研修会 [ 姉妹編 ]",
         size=18, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)
add_text(s, Cm(1.5), Cm(6.0), SW - Cm(3), Cm(2.2),
         "適切なケアマネジメント手法の実践活用",
         size=34, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
add_text(s, Cm(1.5), Cm(8.8), SW - Cm(3), Cm(2.5),
         "ー基本ケア・疾患別ケアを、\nアセスメント・ケアプラン作成にどう落とし込むかー",
         size=18, color=WHITE, align=PP_ALIGN.CENTER, line_spacing=1.3)
add_text(s, Cm(1.5), Cm(12.3), SW - Cm(3), Cm(1.2),
         "〜 二人の利用者の物語で読み解く ・ 脳血管疾患 / 糖尿病 〜",
         size=18, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)

add_text(s, Cm(2), Cm(15), SW - Cm(4), Cm(1),
         "主任ケアマネジャーの会  主催",
         size=16, color=WHITE, align=PP_ALIGN.CENTER)
add_text(s, Cm(2), Cm(16), SW - Cm(4), Cm(1),
         "研修時間 90 分",
         size=14, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)


# =============================================================
#  SLIDE 2: 物語の主人公たち
# =============================================================
s = add_slide()
add_header(s, "Prologue", "本日の物語 ― 二人の主人公に出会う")
add_footer(s, 2)

add_text(s, Cm(1.2), Cm(4.3), SW - Cm(2.4), Cm(1),
         "今日の研修は、二人の利用者の物語を通して進みます。",
         size=15, color=DARK_GRAY)

# Cさん(脳血管)
add_rect(s, Cm(1.2), Cm(5.7), Cm(15.5), Cm(1.0), STROKE_COLOR)
add_text(s, Cm(1.5), Cm(5.75), Cm(15), Cm(0.9),
         "C さん  72歳 男性  ・ 脳梗塞 退院直後",
         size=16, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, Cm(1.2), Cm(6.7), Cm(15.5), Cm(9.5), STROKE_LIGHT)
add_text(s, Cm(1.5), Cm(6.9), Cm(15), Cm(9),
         "・左片麻痺(BRSⅢ)、軽度構音障害\n"
         "・元タクシードライバー、無口で頑固\n"
         "・妻(70歳)と二人暮らし、長女は車で1時間\n"
         "・高血圧・心房細動の既往、ワーファリン服用\n"
         "・3週間の急性期 + 2ヶ月の回復期リハ\n"
         "・要介護2、回復期リハ病棟から自宅退院\n\n"
         "本人の口癖:\n「もう自分は終わりだ。タクシーに乗れない男に\n  何の意味がある」",
         size=13, color=DARK_GRAY, line_spacing=1.3)

# Dさん(糖尿病)
add_rect(s, Cm(17.5), Cm(5.7), Cm(15), Cm(1.0), DM_COLOR)
add_text(s, Cm(17.8), Cm(5.75), Cm(14.5), Cm(0.9),
         "D さん  68歳 女性  ・ 糖尿病 30年来",
         size=16, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, Cm(17.5), Cm(6.7), Cm(15), Cm(9.5), DM_LIGHT)
add_text(s, Cm(17.8), Cm(6.9), Cm(14.5), Cm(9),
         "・2型糖尿病、HbA1c 8.6%、インスリン4回\n"
         "・糖尿病性網膜症 (要矯正で歩行可)\n"
         "・糖尿病性末梢神経障害、軽度腎症\n"
         "・独居、息子家族は同市内に在住\n"
         "・夫を5年前に他界、料理が趣味\n"
         "・要支援2、訪問看護でインスリン確認\n\n"
         "本人の口癖:\n「甘いものくらい、楽しまないと\n  生きてる意味がないでしょう」",
         size=13, color=DARK_GRAY, line_spacing=1.3)

add_text(s, Cm(1.2), Cm(17), SW - Cm(2.4), Cm(1.2),
         "▶ この二人の生活を「基本ケア + 疾患別ケア」の視点で読み解いていきます。",
         size=14, color=NAVY, bold=True)


# =============================================================
#  SLIDE 3: 本日の流れ
# =============================================================
s = add_slide()
add_header(s, "Agenda", "本日の流れ(90分・物語をたどる構成)")
add_footer(s, 3)

agenda = [
    ("1", "プロローグ ― なぜ「手法」が必要だったのか",       "10分"),
    ("2", "基本ケアの4つの柱を物語で読み解く",                 "15分"),
    ("3", "C さんの物語 ― 脳血管疾患の疾患別ケア",            "15分"),
    ("4", "D さんの物語 ― 糖尿病の疾患別ケア",                "15分"),
    ("5", "アセスメントへの落とし込み(両ケース比較)",       "15分"),
    ("6", "ケアプラン作成への落とし込み(両ケース比較)",     "10分"),
    ("7", "事例検討・グループ共有・まとめ",                   "10分"),
]
y = Cm(4.5)
for num, title, mins in agenda:
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

    add_text(s, Cm(3.2), y + Cm(0.15), Cm(22), Cm(1),
             title, size=17, bold=True, color=NAVY)
    add_text(s, SW - Cm(4), y + Cm(0.2), Cm(3), Cm(1),
             mins, size=14, color=ORANGE, bold=True, align=PP_ALIGN.RIGHT)
    y += Cm(1.6)


# =============================================================
#  SLIDE 4: 第1部 タイトル
# =============================================================
section_title(1, "なぜ「手法」が必要だったのか",
              "Prologue: A young CM's struggle")


# =============================================================
#  SLIDE 5: 後輩CMの物語
# =============================================================
s = add_slide()
add_header(s, "1. プロローグ", "ある日、後輩 CM から相談を受けた")
add_footer(s, 5)

# シーン1
add_dialog(s, Cm(1.2), Cm(4.5), Cm(31.4),
           "後輩CM(2年目)",
           "Cさんのケアプラン、なんかしっくり来ないんです。\n"
           " 退院サマリーに書いてあるサービスは入れたんですけど、\n"
           " 本人はずっと黙ったままで…これで合ってるんでしょうか",
           speaker_color=CM_COLOR, body_bg=CM_LIGHT, height=Cm(3))

# シーン2
add_dialog(s, Cm(1.2), Cm(8.2), Cm(31.4),
           "あなた(主任CM)",
           "何を聞いて、何が分からなくて、しっくり来ないんだろうね。\n"
           " 一緒に整理してみようか。",
           speaker_color=NAVY, body_bg=ACCENT_BG, height=Cm(2.5))

# 問いかけ
add_box(s, Cm(1.2), Cm(11.5), Cm(31.4), Cm(5.5),
        "主任CMとして、あなたなら何を問いかけますか?",
        "■ 「Cさんがリハビリ病院でどんな顔をしていたか、見てきた?」\n"
        "■ 「奥さんは、これからの生活で何が一番心配だと言っている?」\n"
        "■ 「Cさんが『またタクシーに乗りたい』と言ったとして、私たちは何ができる?」\n\n"
        "▶ 漠然とした「しっくり来ない」を、構造化して言語化させる ― これが手法の出番。",
        title_bg=NAVY, body_size=14)


# =============================================================
#  SLIDE 6: 手法の全体像
# =============================================================
s = add_slide()
add_header(s, "1. プロローグ", "「適切なケアマネジメント手法」とは ― 一言で言うと")
add_footer(s, 6)

# 中央の大きなキーフレーズ
add_rect(s, Cm(2), Cm(4.5), SW - Cm(4), Cm(3.5), NAVY)
add_text(s, Cm(2), Cm(4.8), SW - Cm(4), Cm(1),
         "経験や勘 だけ に頼らない、",
         size=18, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
add_text(s, Cm(2), Cm(5.7), SW - Cm(4), Cm(1.5),
         "「想定される支援内容」 を 仮説 として持って、",
         size=22, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
add_text(s, Cm(2), Cm(6.8), SW - Cm(4), Cm(1),
         "アセスメントで 検証 する 思考の型。",
         size=22, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)

# 2層構造
add_rect(s, Cm(2), Cm(9), SW - Cm(4), Cm(0.9), TEAL)
add_text(s, Cm(2), Cm(9.05), SW - Cm(4), Cm(0.8),
         "【基本ケア】 すべての高齢者に共通する支援の考え方  (4つの柱)",
         size=15, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
add_rect(s, Cm(2), Cm(9.9), SW - Cm(4), Cm(1.8), ACCENT_BG)
add_text(s, Cm(2.3), Cm(10.1), SW - Cm(4.6), Cm(1.6),
         "①尊厳の保持と自立支援  ②全体像の把握  ③意思決定の支援  ④リスクの予測と管理",
         size=15, color=DARK_GRAY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_rect(s, Cm(2), Cm(12.2), SW - Cm(4), Cm(0.9), ORANGE)
add_text(s, Cm(2), Cm(12.25), SW - Cm(4), Cm(0.8),
         "【疾患別ケア】 主要 7 疾患群ごとの「想定される支援内容」",
         size=15, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
add_rect(s, Cm(2), Cm(13.1), SW - Cm(4), Cm(1.8), SOFT_ORANGE)
add_text(s, Cm(2.3), Cm(13.3), SW - Cm(4.6), Cm(1.6),
         "脳血管疾患  ・  大腿骨頸部骨折  ・  心疾患  ・  認知症  ・  誤嚥性肺炎  ・  パーキンソン病  ・  糖尿病",
         size=14, color=DARK_GRAY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Cm(1.2), Cm(15.7), SW - Cm(2.4), Cm(1.5),
         "▶ 今日は、★ 脳血管疾患(Cさん) と ★ 糖尿病(Dさん) を物語で追います。\n"
         "▶ 「手法」は答えを与えるものではなく、問いの順序を整える地図です。",
         size=14, color=NAVY, bold=True, line_spacing=1.3)


# =============================================================
#  SLIDE 7: 第2部 タイトル
# =============================================================
section_title(2, "基本ケアの4つの柱を物語で読み解く",
              "The Four Pillars, Through the Story")


# =============================================================
#  SLIDE 8: 柱Ⅰ 尊厳・自立(Cさんの怒り)
# =============================================================
s = add_slide()
add_header(s, "2. 基本ケアを物語で読み解く", "Ⅰ. 尊厳の保持と自立支援 ― Cさんが黙り込んだ朝")
add_footer(s, 8)

# 場面
add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(2.5),
        "【場面】退院翌週・初回モニタリング訪問",
        "ベッドサイドで、若いCMが「リハビリ頑張りましょうね!」と声をかけた瞬間、\n"
        "Cさんは天井を見たまま、何も答えなかった。",
        title_bg=STROKE_COLOR, body_bg=STROKE_LIGHT, body_size=13)

# セリフ
add_dialog(s, Cm(1.2), Cm(7.4), Cm(15.5),
           "C さんの心の声",
           "頑張る?何を頑張るんだ。\n この左手で何ができる。\n 客なんかもう乗せられない。",
           speaker_color=STROKE_COLOR, body_bg=STROKE_LIGHT, height=Cm(3))
add_dialog(s, Cm(17.5), Cm(7.4), Cm(15),
           "妻 ・ Cさんに 30 年連れ添う",
           "あの人は、車で家族を養ってきた。\n 仕事を取り上げられて、生きる気力をなくして。",
           speaker_color=NAVY, body_bg=ACCENT_BG, height=Cm(3))

add_box(s, Cm(1.2), Cm(11), Cm(31.4), Cm(6),
        "ここで主任CMが見るべきこと",
        "■ 「ADL自立度」ではなく、本人にとっての 役割の喪失感 が課題の核心\n\n"
        "■ 「リハビリを頑張る」は支援者の言葉。本人の言葉に翻訳されているか?\n"
        "    → 「もう一度、孫を助手席に乗せて運転したい」「妻と買い物に行きたい」\n\n"
        "■ 尊厳の保持とは、できないことを補うだけでなく、できることを奪わない こと\n"
        "    → 朝刊を取りに行く、自分でひげを剃る、こうした小さな「自分」を残す",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 9: 柱Ⅱ 全体像(Dさんの冷蔵庫)
# =============================================================
s = add_slide()
add_header(s, "2. 基本ケアを物語で読み解く", "Ⅱ. 全体像の把握 ― Dさんの冷蔵庫が教えてくれたこと")
add_footer(s, 9)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(2.5),
        "【場面】訪問看護師からの一報",
        "「Dさんの台所、見せてもらいました。\n 冷蔵庫に大福が3つ、流しに菓子パンの袋。\n これだとHbA1cが下がらないのも当然です」",
        title_bg=DM_COLOR, body_bg=DM_LIGHT, body_size=13)

# 4領域で全体像を把握
regions = [
    ("① 疾病管理",  "HbA1c 8.6%・インスリン4回\n眼科・透析リスクの将来像",  STROKE_COLOR),
    ("② 生活全般",  "間食習慣・趣味の和菓子\n料理が好き・1人の食卓",       ORANGE),
    ("③ 家族",       "息子は週末のみ訪問\n嫁との関係は良好だが遠慮",         CM_COLOR),
    ("④ 住環境",     "古い団地・夜間トイレ遠い\n近所のスーパーが甘味豊富",  TEAL),
]
y = Cm(7.3)
col_w = Cm(7.7)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
for i, (title, body, color) in enumerate(regions):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1), color)
    add_text(s, x + Cm(0.2), y + Cm(0.1), col_w - Cm(0.4), Cm(0.8),
             title, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1), col_w, Cm(3.5), ACCENT_BG)
    add_text(s, x + Cm(0.3), y + Cm(1.2), col_w - Cm(0.6), Cm(3.3),
             body, size=11, color=DARK_GRAY)

add_box(s, Cm(1.2), Cm(13.5), Cm(31.4), Cm(3.5),
        "全体像を見たCMの気づき",
        "■ 「血糖コントロール不良」を「服薬指導」だけで解こうとしていなかったか?\n"
        "■ Dさんの「楽しみ」を奪うのではなく、「楽しみと共存する血糖管理」が必要\n"
        "■ 「料理が好き」は強み。低糖質レシピを学ぶ場が、Dさんの社会参加にもなる",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 10: 柱Ⅲ 意思決定(Cさんの選択)
# =============================================================
s = add_slide()
add_header(s, "2. 基本ケアを物語で読み解く", "Ⅲ. 意思決定支援 ― Cさんは何を選びたいのか")
add_footer(s, 10)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(2.5),
        "【場面】退院後 3 週目・サービス担当者会議",
        "「通所リハと訪問リハ、どちらにしますか?」\n"
        "サ担で問われたCさんは、首を振って、目を伏せた。",
        title_bg=STROKE_COLOR, body_bg=STROKE_LIGHT, body_size=13)

add_box(s, Cm(1.2), Cm(7.5), Cm(15.5), Cm(5),
        "選択を阻む3つの壁",
        "① 情報の壁:選択肢の違いが理解できない\n"
        "② 体験の壁:そもそも体験したことがない\n"
        "③ 自己効力感の壁:「自分で決めてもどうせ」",
        title_bg=ORANGE, body_bg=SOFT_ORANGE, body_size=13)

add_box(s, Cm(17.5), Cm(7.5), Cm(15), Cm(5),
        "主任CMの工夫",
        "・体験利用を組み込む(両方やってみる)\n"
        "・写真・パンフレット・先輩利用者の話を見せる\n"
        "・本人が「不安に思っていること」を先に言語化\n"
        "・「今日は決めなくていい」と保留も認める",
        title_bg=TEAL, body_size=13)

add_box(s, Cm(1.2), Cm(13), Cm(31.4), Cm(4),
        "意思決定支援は「結論」ではなく「プロセス」を残す",
        "■ 第1表「本人の意向」に、確認した日時・場面を記載する\n"
        "■ サ担録に、Cさんがうなずいた瞬間・首を振った瞬間を記す\n"
        "■ 支援経過に「2週間体験後、訪問リハを選んだ理由:外に出るのが怖いと話されたため」",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=13)


# =============================================================
#  SLIDE 11: 柱Ⅳ リスク管理(Dさんの低血糖)
# =============================================================
s = add_slide()
add_header(s, "2. 基本ケアを物語で読み解く", "Ⅳ. リスクの予測と管理 ― Dさんが救急搬送された日")
add_footer(s, 11)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(2.5),
        "【場面】真冬の朝、息子からの電話",
        "「母が、トイレで倒れていました。意識がぼんやりして、汗をかいていて。\n"
        "  …低血糖だったみたいです。救急車を呼びました」",
        title_bg=DM_COLOR, body_bg=DM_LIGHT, body_size=13)

# 予防と発見のフロー
add_text(s, Cm(1.2), Cm(7.3), SW - Cm(2.4), Cm(0.8),
         "  ▼ 予測できたサイン  ―  振り返り",
         size=15, bold=True, color=NAVY)

signs = [
    ("食欲低下", "前週から夕食を残すように", STROKE_COLOR),
    ("睡眠変化", "夜中に何度もトイレ",       ORANGE),
    ("認知低下", "薬の名前を間違える",       CM_COLOR),
    ("孤立",     "息子の訪問が2週空いた",  TEAL),
]
y = Cm(8.4)
col_w = Cm(7.7)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
for i, (title, body, color) in enumerate(signs):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(1.0), color)
    add_text(s, x + Cm(0.2), y + Cm(0.1), col_w - Cm(0.4), Cm(0.8),
             title, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1), col_w, Cm(2.2), ACCENT_BG)
    add_text(s, x + Cm(0.3), y + Cm(1.1), col_w - Cm(0.6), Cm(2),
             body, size=11, color=DARK_GRAY, anchor=MSO_ANCHOR.MIDDLE)

add_box(s, Cm(1.2), Cm(12.5), Cm(31.4), Cm(4.5),
        "主任CMの問いかけ ― 「サインの拾い方」を後輩に教える",
        "■ モニタリング時、「変わりないですか?」だけで終わっていないか?\n"
        "■ 食欲・睡眠・薬・人との会話、4点セットを毎回確認する習慣\n"
        "■ 予測できたリスクは「想定される支援内容」に書いてある ― 手法を知っていれば防げた",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 12: 第3部 Cさん編 タイトル
# =============================================================
section_title(3, "Cさんの物語 ― 脳血管疾患の疾患別ケア",
              "Chapter Three: The Story of Mr. C",
              accent_color=STROKE_COLOR)


# =============================================================
#  SLIDE 13: 脳血管疾患の特徴
# =============================================================
s = add_slide()
add_header(s, "3. Cさんの物語 ― 脳血管疾患", "脳血管疾患のケアで押さえる4つの視点")
add_footer(s, 13)

views = [
    ("V1", "再発予防",   "・血圧・糖代謝・脂質の管理\n・抗血栓薬の確実な服用\n・出血リスクの観察"),
    ("V2", "機能回復",   "・廃用症候群の予防\n・回復期 → 維持期への移行\n・自主トレ・通所リハ"),
    ("V3", "ADL再構築", "・残存機能を活かす\n・福祉用具・住宅改修\n・代償手段の獲得"),
    ("V4", "心理・社会", "・うつ・意欲低下への対応\n・役割再獲得・社会参加\n・家族の受容"),
]
col_w = Cm(7.7)
col_h = Cm(11.5)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
y = Cm(4.5)
for i, (label, title, body) in enumerate(views):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(2.8), STROKE_COLOR)
    add_text(s, x, y + Cm(0.2), col_w, Cm(0.8),
             label, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, x + Cm(0.2), y + Cm(1.1), col_w - Cm(0.4), Cm(1.6),
             title, size=15, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(2.8), col_w, col_h - Cm(2.8), STROKE_LIGHT)
    add_text(s, x + Cm(0.3), y + Cm(3.0), col_w - Cm(0.6), col_h - Cm(3),
             body, size=12, color=DARK_GRAY)

add_text(s, Cm(1.2), Cm(16.5), SW - Cm(2.4), Cm(1),
         "▶ 「再発予防」と「ADL再構築」は両輪。どちらかに偏ると、生活の質も命も守れない。",
         size=13, color=NAVY, bold=True)


# =============================================================
#  SLIDE 14: Cさんのタイムライン
# =============================================================
s = add_slide()
add_header(s, "3. Cさんの物語 ― 脳血管疾患", "C さんの 1 年 ― 退院後の生活経過")
add_footer(s, 14)

# タイムライン
timeline = [
    ("退院直後\n0〜1ヶ月",  "・抑うつ・閉じこもり\n・自宅で何もしない時間\n・服薬は妻が管理",  STROKE_COLOR),
    ("導入期\n1〜3ヶ月",   "・訪問リハ開始(週2)\n・ADL拡大、洗面自立\n・近所散歩(妻同伴)", ORANGE),
    ("安定期\n3〜6ヶ月",   "・通所リハに移行\n・他の利用者と交流開始\n・運転リハ評価開始",      TEAL),
    ("拡大期\n6〜12ヶ月", "・地域サロン参加\n・限定運転再開\n・心房細動コントロール継続",     CM_COLOR),
]
y = Cm(4.5)
col_w = Cm(7.7)
col_h = Cm(8)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2

# 線
add_rect(s, start_x, y + Cm(2.5), total_w, Cm(0.1), NAVY)

for i, (label, body, color) in enumerate(timeline):
    x = start_x + (col_w + gap) * i
    # 円
    circ = s.shapes.add_shape(MSO_SHAPE.OVAL, x + col_w/2 - Cm(0.4),
                              y + Cm(2.2), Cm(0.8), Cm(0.8))
    circ.fill.solid()
    circ.fill.fore_color.rgb = color
    circ.line.fill.background()
    # ラベル
    add_text(s, x, y, col_w, Cm(2.0),
             label, size=14, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 本文
    add_rect(s, x, y + Cm(3.5), col_w, col_h - Cm(3.5), ACCENT_BG)
    add_text(s, x + Cm(0.3), y + Cm(3.7), col_w - Cm(0.6), col_h - Cm(3.7),
             body, size=11, color=DARK_GRAY)

add_box(s, Cm(1.2), Cm(13.5), Cm(31.4), Cm(3.5),
        "経過から読み取れること ― 「想定される支援内容」は時間軸を持つ",
        "■ 同じ「リハビリ」でも、各期で意味合いが違う(治療 → 維持 → 拡大)\n"
        "■ 主任CMは、3ヶ月先・1年先を見据えて、今のプランを設計しているか?\n"
        "■ サービスの「卒業」「変更」も計画的に ― ずっと同じプランは黄信号",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=13)


# =============================================================
#  SLIDE 15: Cさんの「想定される支援内容」
# =============================================================
s = add_slide()
add_header(s, "3. Cさんの物語 ― 脳血管疾患", "C さんに 想定される支援内容(仮説)")
add_footer(s, 15)

rows = [
    ("再発予防",
     "・抗凝固薬の確実服用(PT-INR定期確認)\n・血圧・脈拍の自己測定\n・脱水・便秘予防",
     "・かかりつけ医との情報共有\n・服薬カレンダー・残薬確認\n・本人 → 妻 → 訪看の三重チェック"),
    ("機能回復",
     "・訪問リハ(導入期) → 通所リハ(維持期)\n・自主トレメニュー\n・上肢機能の継続的評価",
     "・廃用予防の生活動作リスト\n・福祉用具(短下肢装具・杖)\n・住宅改修(手すり・トイレ)"),
    ("ADL再構築",
     "・更衣・整容・トイレ動作の自立支援\n・調理・買い物への参加\n・代償手段の獲得",
     "・OT/PTからの動作指導\n・妻への「待つ介助」の助言\n・「できるADL」と「しているADL」"),
    ("心理・社会",
     "・抑うつスクリーニング\n・役割の再構築\n・社会参加・運転再開評価",
     "・ピアサポート(脳卒中の会)\n・地域サロン・自治会復帰\n・運転外来との連携"),
]
col_widths = [Cm(5), Cm(13), Cm(13)]
headers = ["視点", "想定される支援内容", "実務での具体的アクション"]
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
    row_h = Cm(2.7)
    bg = STROKE_LIGHT if r_i % 2 == 0 else WHITE
    add_rect(s, start_x, y, sum(col_widths, Cm(0)), row_h, bg, line_color=LIGHT_GRAY)
    x = start_x
    for i, c in enumerate(row):
        align = PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT
        col = STROKE_COLOR if i == 0 else DARK_GRAY
        bold = (i == 0)
        size = 11 if i != 0 else 13
        add_text(s, x + Cm(0.2), y + Cm(0.2), col_widths[i] - Cm(0.4), row_h - Cm(0.4),
                 c, size=size, bold=bold, color=col, align=align,
                 anchor=MSO_ANCHOR.MIDDLE)
        x += col_widths[i]
    y += row_h


# =============================================================
#  SLIDE 16: Cさんのターニングポイント
# =============================================================
s = add_slide()
add_header(s, "3. Cさんの物語 ― 脳血管疾患", "C さんの転機 ― 8ヶ月目、孫の運動会")
add_footer(s, 16)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(3.5),
        "【場面】退院から 8 ヶ月、秋の運動会",
        "「じいちゃん、来てくれてありがとう」\n"
        "    車椅子で運動場にいたCさんに、小学生の孫が抱きついた。\n"
        "    Cさんは、初めて声を出して泣いた。",
        title_bg=STROKE_COLOR, body_bg=STROKE_LIGHT, body_size=14)

add_dialog(s, Cm(1.2), Cm(8.5), Cm(15.5),
           "Cさんから後日、CMへ",
           "あの日、生きてて良かった。\n もう少し、頑張ってみるよ。",
           speaker_color=STROKE_COLOR, body_bg=STROKE_LIGHT, height=Cm(3))

add_dialog(s, Cm(17.5), Cm(8.5), Cm(15),
           "後輩CMの振り返り",
           "ADLの数字じゃ、見えない目標があるんですね。\n これからは、本人の言葉から目標を作ります。",
           speaker_color=CM_COLOR, body_bg=CM_LIGHT, height=Cm(3))

add_box(s, Cm(1.2), Cm(12.5), Cm(31.4), Cm(4.5),
        "主任CMが伝えるべきこと",
        "■ ケアプランの目標は「○○できる」だけでなく、「○○のために」が大切\n"
        "■ 「孫の運動会に行く」が目標になれば、車椅子整備・移動手段・体力作りが手段になる\n"
        "■ 数値目標と意味目標、両方を持つこと ― これが「個別性のあるプラン」",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 17: 第4部 Dさん編 タイトル
# =============================================================
section_title(4, "Dさんの物語 ― 糖尿病の疾患別ケア",
              "Chapter Four: The Story of Mrs. D",
              accent_color=DM_COLOR)


# =============================================================
#  SLIDE 18: 糖尿病の特徴
# =============================================================
s = add_slide()
add_header(s, "4. Dさんの物語 ― 糖尿病", "糖尿病のケアで押さえる4つの視点")
add_footer(s, 18)

views = [
    ("V1", "血糖コントロール", "・服薬・インスリン管理\n・食事・運動の継続\n・自己血糖測定の習慣化"),
    ("V2", "合併症の予防",     "・網膜症 → 視力低下\n・腎症 → 透析\n・神経障害 → フットケア"),
    ("V3", "低血糖の管理",     "・シックデイ対応\n・夜間低血糖の予防\n・救急時の連絡体制"),
    ("V4", "生活全体の支援",   "・買い物・調理\n・楽しみとの両立\n・社会的孤立の予防"),
]
col_w = Cm(7.7)
col_h = Cm(11.5)
gap = Cm(0.3)
total_w = col_w * 4 + gap * 3
start_x = (SW - total_w) // 2
y = Cm(4.5)
for i, (label, title, body) in enumerate(views):
    x = start_x + (col_w + gap) * i
    add_rect(s, x, y, col_w, Cm(2.8), DM_COLOR)
    add_text(s, x, y + Cm(0.2), col_w, Cm(0.8),
             label, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, x + Cm(0.2), y + Cm(1.1), col_w - Cm(0.4), Cm(1.6),
             title, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(2.8), col_w, col_h - Cm(2.8), DM_LIGHT)
    add_text(s, x + Cm(0.3), y + Cm(3.0), col_w - Cm(0.6), col_h - Cm(3),
             body, size=12, color=DARK_GRAY)

add_text(s, Cm(1.2), Cm(16.5), SW - Cm(2.4), Cm(1),
         "▶ 糖尿病は「数十年単位の慢性疾患」。生活と一体化したケアが必須。",
         size=13, color=NAVY, bold=True)


# =============================================================
#  SLIDE 19: Dさんの「想定される支援内容」
# =============================================================
s = add_slide()
add_header(s, "4. Dさんの物語 ― 糖尿病", "D さんに 想定される支援内容(仮説)")
add_footer(s, 19)

rows = [
    ("血糖コントロール",
     "・インスリン手技と保管\n・自己血糖測定 (SMBG)\n・食事タイミング・量",
     "・訪問看護による週1確認\n・服薬カレンダー\n・主治医・糖尿病療養指導士と連携"),
    ("合併症の予防",
     "・眼科・腎機能の定期検査\n・フットケア(毎日観察)\n・血圧・脂質管理",
     "・複数科受診の情報集約 → CMが結節点\n・足の傷の早期発見ルート\n・受診同行・送迎の検討"),
    ("低血糖の管理",
     "・症状の自覚教育\n・補食の準備\n・シックデイルール",
     "・冷蔵庫に補食ストック\n・緊急連絡先カード携帯\n・近隣・息子との見守り体制"),
    ("生活全体の支援",
     "・楽しみとの両立\n・社会参加・役割継続\n・抑うつ・孤立予防",
     "・低糖質クッキング教室紹介\n・配食 + 自炊の組み合わせ\n・地域包括・サロンへ"),
]
col_widths = [Cm(5), Cm(13), Cm(13)]
headers = ["視点", "想定される支援内容", "実務での具体的アクション"]
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
    row_h = Cm(2.7)
    bg = DM_LIGHT if r_i % 2 == 0 else WHITE
    add_rect(s, start_x, y, sum(col_widths, Cm(0)), row_h, bg, line_color=LIGHT_GRAY)
    x = start_x
    for i, c in enumerate(row):
        align = PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT
        col = DM_COLOR if i == 0 else DARK_GRAY
        bold = (i == 0)
        size = 11 if i != 0 else 13
        add_text(s, x + Cm(0.2), y + Cm(0.2), col_widths[i] - Cm(0.4), row_h - Cm(0.4),
                 c, size=size, bold=bold, color=col, align=align,
                 anchor=MSO_ANCHOR.MIDDLE)
        x += col_widths[i]
    y += row_h


# =============================================================
#  SLIDE 20: Dさんと「楽しみ」
# =============================================================
s = add_slide()
add_header(s, "4. Dさんの物語 ― 糖尿病", "D さんの揺らぎ ― 「我慢の生活」を続けられるか")
add_footer(s, 20)

add_dialog(s, Cm(1.2), Cm(4.5), Cm(15.5),
           "Dさん",
           "毎日毎日、数字を測って、針を刺して。\n こんな生活、80歳まで続けられない。\n たまには、好きなもの食べたっていい。",
           speaker_color=DM_COLOR, body_bg=DM_LIGHT, height=Cm(3.5))

add_dialog(s, Cm(17.5), Cm(4.5), Cm(15),
           "息子",
           "母さん、また甘いものを?\n 入院したくないでしょう。\n 我慢して下さい。",
           speaker_color=NAVY, body_bg=ACCENT_BG, height=Cm(3.5))

add_box(s, Cm(1.2), Cm(8.5), Cm(31.4), Cm(8.5),
        "ここで主任CMが見ているもの",
        "■ 「コンプライアンス不良」と裁く前に、Dさんの30年の闘病史を尊重する\n"
        "    糖尿病は「ずっと我慢」と言われ続けてきた。心が疲れて当然。\n\n"
        "■ 「禁止」ではなく「工夫」を提案できるか?\n"
        "    ・週1回は好きな和菓子を、量と時間を決めて → 主治医・栄養士と相談\n"
        "    ・低糖質スイーツのレシピをDさん自身が探す → 趣味と療養の融合\n\n"
        "■ 家族の「正しさ」が本人を追い詰めることもある\n"
        "    息子に「数字より、母さんが続けられることを一緒に考えませんか」と伝える\n\n"
        "■ Dさんが諦めない、息子も付き合える、その間に立つのが ケアマネの仕事",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 21: 第5部 タイトル
# =============================================================
section_title(5, "アセスメントへの落とし込み",
              "Embedding into Assessment")


# =============================================================
#  SLIDE 22: 仮説→検証のサイクル
# =============================================================
s = add_slide()
add_header(s, "5. アセスメントへの落とし込み", "「想定される支援内容」を 仮説 として持ち込む")
add_footer(s, 22)

# サイクル図
boxes = [
    ("①\n手引きから\n仮説立て", TEAL),
    ("②\nアセスメントで\n本人の必要性検証", ORANGE),
    ("③\nケアプランへ\n落とし込み", TEAL),
    ("④\nモニタリングで\n再検証", ORANGE),
]
box_w = Cm(6.5)
box_h = Cm(3.5)
gap_x = Cm(1.4)
total_w = box_w * 4 + gap_x * 3
start_x = (SW - total_w) // 2
y = Cm(4.8)
for i, (txt, color) in enumerate(boxes):
    x = start_x + (box_w + gap_x) * i
    add_rect(s, x, y, box_w, box_h, color)
    add_text(s, x + Cm(0.2), y + Cm(0.3), box_w - Cm(0.4), box_h - Cm(0.6),
             txt, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if i < 3:
        arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                   x + box_w + Cm(0.1), y + Cm(1.3),
                                   gap_x - Cm(0.2), Cm(0.9))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = NAVY
        arrow.line.fill.background()

add_box(s, Cm(1.2), Cm(10), Cm(15.5), Cm(7),
        "Cさん(脳血管)の検証例",
        "■ 仮説 :「再発予防の服薬支援が必要」\n"
        "    → 検証 :妻が確実に管理できている → 訪看週1で十分(過剰サービス回避)\n\n"
        "■ 仮説 :「抑うつへの対応が必要」\n"
        "    → 検証 :GDS高値・閉じこもり傾向\n"
        "        → 訪問リハ + 心理面の声かけを内容に\n\n"
        "■ 仮説 :「ADL再構築」\n"
        "    → 検証 :「孫の運動会に行きたい」目標がある\n"
        "        → 車椅子・移動手段の支援に重点配分",
        title_bg=STROKE_COLOR, body_bg=STROKE_LIGHT, body_size=12)

add_box(s, Cm(17.5), Cm(10), Cm(15), Cm(7),
        "Dさん(糖尿病)の検証例",
        "■ 仮説 :「血糖管理 = 食事制限」\n"
        "    → 検証 :Dさんの楽しみが「料理」\n"
        "        → 制限ではなく低糖質料理教室の活用\n\n"
        "■ 仮説 :「インスリン手技確認が必要」\n"
        "    → 検証 :手技は問題なし。問題は孤立時の低血糖対応\n"
        "        → 訪看の内容を「低血糖教育」にシフト\n\n"
        "■ 仮説 :「フットケアが必要」\n"
        "    → 検証 :神経障害あり、自己観察困難\n"
        "        → 訪問看護に毎週の足観察を明記",
        title_bg=DM_COLOR, body_bg=DM_LIGHT, body_size=12)


# =============================================================
#  SLIDE 23: 課題分析標準項目との対応
# =============================================================
s = add_slide()
add_header(s, "5. アセスメントへの落とし込み", "課題分析標準項目で「どこを濃く書くか」")
add_footer(s, 23)

rows = [
    ("Cさん:脳血管疾患",
     "⑨健康状態(再発リスク・抗凝固)\n⑩ADL ⑪IADL(できる/しているの差)\n⑮排尿 ⑯排便(尿路感染リスク)\n⑲問題行動 → 抑うつとして再解釈\n⑳介護力(妻の介護観・腰痛)"),
    ("Dさん:糖尿病",
     "⑨健康状態(HbA1c・合併症)\n⑭社会との関わり(孤立・楽しみ)\n⑯排便 ⑰褥瘡(末梢神経障害・足)\n⑲問題行動 → 食行動として把握\n⑳介護力(息子の遠距離・関係性)"),
]
col_widths = [Cm(10), Cm(20)]
start_x = (SW - sum(col_widths, Cm(0))) // 2
start_y = Cm(4.5)

headers = ["事例", "重点的にアセスメントすべき項目"]
add_rect(s, start_x, start_y, sum(col_widths, Cm(0)), Cm(1.0), NAVY)
x = start_x
for i, h in enumerate(headers):
    add_text(s, x + Cm(0.2), start_y + Cm(0.1), col_widths[i] - Cm(0.4), Cm(0.8),
             h, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    x += col_widths[i]

y = start_y + Cm(1.0)
for r_i, row in enumerate(rows):
    row_h = Cm(4.5)
    bg = STROKE_LIGHT if r_i == 0 else DM_LIGHT
    color = STROKE_COLOR if r_i == 0 else DM_COLOR
    add_rect(s, start_x, y, sum(col_widths, Cm(0)), row_h, bg, line_color=LIGHT_GRAY)
    x = start_x
    for i, c in enumerate(row):
        align = PP_ALIGN.CENTER if i == 0 else PP_ALIGN.LEFT
        col = color if i == 0 else DARK_GRAY
        bold = (i == 0)
        size = 13 if i == 0 else 12
        add_text(s, x + Cm(0.3), y + Cm(0.3), col_widths[i] - Cm(0.6), row_h - Cm(0.6),
                 c, size=size, bold=bold, color=col, align=align,
                 anchor=MSO_ANCHOR.MIDDLE)
        x += col_widths[i]
    y += row_h

add_text(s, Cm(1.2), Cm(15.5), SW - Cm(2.4), Cm(2),
         "▶ 23項目は均等に書くのではなく、疾患特性に応じて 濃淡をつける。\n"
         "▶ 主任CMは後輩に「なぜこの項目をこの深さで書いたか」を説明させる訓練を。",
         size=13, color=NAVY, bold=True, line_spacing=1.3)


# =============================================================
#  SLIDE 24: 多職種連携
# =============================================================
s = add_slide()
add_header(s, "5. アセスメントへの落とし込み", "多職種から「情報を引き出す」のもアセスメント")
add_footer(s, 24)

# 中央にCM
center_x = SW // 2
center_y = Cm(10.5)
oval = s.shapes.add_shape(MSO_SHAPE.OVAL,
                          center_x - Cm(2.5), center_y - Cm(1.5),
                          Cm(5), Cm(3))
oval.fill.solid()
oval.fill.fore_color.rgb = NAVY
oval.line.fill.background()
tf = oval.text_frame
tf.margin_top = Cm(0)
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "ケアマネ"
r.font.name = FONT
r.font.size = Pt(18)
r.font.bold = True
r.font.color.rgb = WHITE
p2 = tf.add_paragraph()
p2.alignment = PP_ALIGN.CENTER
r2 = p2.add_run()
r2.text = "(情報の結節点)"
r2.font.name = FONT
r2.font.size = Pt(11)
r2.font.color.rgb = LIGHT_GRAY

# 周囲の職種
positions = [
    ("主治医",       "疾患の見通し・薬剤調整",     Cm(2),   Cm(5)),
    ("薬剤師",       "服薬・残薬・相互作用",       Cm(13),  Cm(5)),
    ("訪問看護",     "バイタル・症状の変化",       Cm(24),  Cm(5)),
    ("リハ職",       "ADL/IADLの実行力",          Cm(2),   Cm(15)),
    ("栄養士",       "食事・栄養・生活習慣",       Cm(13),  Cm(15)),
    ("地域包括",     "家族・地域資源",             Cm(24),  Cm(15)),
]
for name, role, x, y in positions:
    add_rect(s, x, y, Cm(7.9), Cm(0.9), TEAL)
    add_text(s, x + Cm(0.2), y, Cm(7.5), Cm(0.9),
             name, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(0.9), Cm(7.9), Cm(1.3), ACCENT_BG)
    add_text(s, x + Cm(0.2), y + Cm(1), Cm(7.5), Cm(1.1),
             role, size=11, color=DARK_GRAY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Cm(1.2), Cm(17.5), SW - Cm(2.4), Cm(1),
         "▶ アセスメントは1人で完成しない。CMは「問いを持って」多職種に声をかける役。",
         size=13, color=NAVY, bold=True)


# =============================================================
#  SLIDE 25: 第6部 タイトル
# =============================================================
section_title(6, "ケアプラン作成への落とし込み",
              "From Story to Care Plan")


# =============================================================
#  SLIDE 26: Cさんの第1表
# =============================================================
s = add_slide()
add_header(s, "6. ケアプランへの落とし込み", "Cさんのケアプラン:第1表(抜粋)")
add_footer(s, 26)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(3.8),
        "利用者及び家族の生活に対する意向",
        "本人:「孫の運動会にもう一度行きたい。\n"
        "        妻と近所のスーパーに歩いて買い物に行けたら、それで十分だ。」(R6.○月○日)\n\n"
        "妻 :「再発だけは絶対に避けたい。本人がやりたいことは、できる範囲で応援したい。」",
        title_bg=STROKE_COLOR, body_size=12)

add_box(s, Cm(1.2), Cm(8.5), Cm(15.5), Cm(4.5),
        "総合的な援助の方針",
        "■ 再発予防(服薬・血圧)を生活の基盤に据える\n"
        "■ 残存機能を活かしたADL再構築を、本人のペースで進める\n"
        "■ 「孫の運動会に行く」を中期目標に、生活全体を組み立てる\n"
        "■ 妻の介護負担を軽減し、夫婦の在宅生活を支える",
        title_bg=NAVY, body_size=12)

add_box(s, Cm(17.5), Cm(8.5), Cm(15), Cm(4.5),
        "認定情報・課題分析の理由",
        "・要介護2(R6.○月認定)\n"
        "・主治医意見書:脳梗塞後遺症(左片麻痺・BRSⅢ)\n"
        "・回復期病棟退院に伴う在宅生活設計\n"
        "・抑うつ傾向・社会的役割の喪失感あり",
        title_bg=ORANGE, body_bg=SOFT_ORANGE, body_size=12)

add_box(s, Cm(1.2), Cm(13.2), Cm(31.4), Cm(3.8),
        "第1表 作成のポイント(主任CMチェック)",
        "■ 「再発予防」と「本人の願い」を 並列 で書けているか?(片方に偏らない)\n"
        "■ 妻の意向と本人の意向を 分けて 記載しているか?\n"
        "■ 援助方針が、3ヶ月先・1年先を見据えた表現になっているか?",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=13)


# =============================================================
#  SLIDE 27: Cさんの第2表
# =============================================================
s = add_slide()
add_header(s, "6. ケアプランへの落とし込み", "Cさんのケアプラン:第2表(抜粋)")
add_footer(s, 27)

needs_data = [
    ("生活全般の解決すべき課題", "長期目標(12ヶ月)", "短期目標(3ヶ月)", "サービス内容"),
    ("妻と近所のスーパーまで\n歩いて買い物に行きたい",
     "片道500mを杖歩行で\n往復できる",
     "屋内・庭まで杖歩行が\n安定する",
     "訪問リハ(週2)\n自主トレ・住宅改修\n短下肢装具・杖"),
    ("再発せず、薬を確実に\n飲み続けたい",
     "再発・入院なく在宅\n生活を継続する",
     "PT-INRが治療域に\n安定する",
     "訪問看護(週1)\n服薬カレンダー\nかかりつけ医連携"),
    ("孫の運動会に\nもう一度行きたい",
     "10月の運動会に\n家族と一緒に参加できる",
     "車椅子で30分以上\n外出できる体力をつける",
     "通所リハ(週2)\n車椅子整備\n家族による外出支援"),
    ("妻にあまり負担をかけず、\n夫婦で穏やかに暮らしたい",
     "妻の介護負担感が\n軽減され、外出も可能に",
     "妻が月1回、自分の\n時間を持てるようになる",
     "短期入所(月1回)\n家族会・地域包括"),
]
col_widths = [Cm(8), Cm(7.5), Cm(7.5), Cm(8.4)]
start_x = (SW - sum(col_widths, Cm(0))) // 2
start_y = Cm(4.5)

add_rect(s, start_x, start_y, sum(col_widths, Cm(0)), Cm(1.2), STROKE_COLOR)
x = start_x
for i, h in enumerate(needs_data[0]):
    add_text(s, x + Cm(0.1), start_y + Cm(0.1), col_widths[i] - Cm(0.2), Cm(1),
             h, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    x += col_widths[i]

y = start_y + Cm(1.2)
for r_i, row in enumerate(needs_data[1:]):
    row_h = Cm(2.7)
    bg = STROKE_LIGHT if r_i % 2 == 0 else WHITE
    add_rect(s, start_x, y, sum(col_widths, Cm(0)), row_h, bg, line_color=LIGHT_GRAY)
    x = start_x
    for i, c in enumerate(row):
        add_text(s, x + Cm(0.15), y + Cm(0.2), col_widths[i] - Cm(0.3), row_h - Cm(0.4),
                 c, size=10.5, color=DARK_GRAY,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        x += col_widths[i]
    y += row_h


# =============================================================
#  SLIDE 28: Dさんのケアプラン
# =============================================================
s = add_slide()
add_header(s, "6. ケアプランへの落とし込み", "Dさんのケアプラン:第1表・第2表(抜粋)")
add_footer(s, 28)

# 第1表(コンパクト)
add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(3.5),
        "Dさん 第1表 意向・援助方針",
        "本人:「数字に縛られすぎず、好きな料理を続けたい。たまには甘いものも。」\n"
        "息子:「合併症で入院しないでほしい。母らしい生活を支えたい。」\n\n"
        "総合的な援助の方針:Dさんの「楽しみ」と「療養」を両立させる支援。\n"
        "孤立を防ぎ、合併症の早期発見ルートを確保する。",
        title_bg=DM_COLOR, body_bg=DM_LIGHT, body_size=12)

# 第2表(コンパクト)
needs_data = [
    ("ニーズ", "長期目標", "短期目標", "主なサービス内容"),
    ("料理を楽しみながら\n血糖をコントロールしたい",
     "HbA1c 7%台維持",
     "週1で和菓子を量を決めて摂取し続ける",
     "栄養指導/低糖質料理教室紹介"),
    ("低血糖で\n倒れないようにしたい",
     "低血糖救急搬送なく\n在宅生活継続",
     "補食をいつも携帯、\n症状を自覚できる",
     "訪問看護(週2)/緊急連絡カード"),
    ("足を守り、自分で\n歩き続けたい",
     "足の重大な合併症なく\n外出を継続",
     "毎日の足の自己観察\n+ 訪看の確認",
     "訪問看護(フットケア)/福祉用具"),
    ("孤立せず、料理仲間と\n楽しい時間を持ちたい",
     "週1回 外出・交流",
     "地域サロンに月2回参加",
     "通所サービス/地域包括との連携"),
]
col_widths = [Cm(8), Cm(6.5), Cm(8), Cm(8.9)]
start_x = (SW - sum(col_widths, Cm(0))) // 2
start_y = Cm(8.5)

add_rect(s, start_x, start_y, sum(col_widths, Cm(0)), Cm(1.0), DM_COLOR)
x = start_x
for i, h in enumerate(needs_data[0]):
    add_text(s, x + Cm(0.1), start_y + Cm(0.1), col_widths[i] - Cm(0.2), Cm(0.8),
             h, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    x += col_widths[i]

y = start_y + Cm(1.0)
for r_i, row in enumerate(needs_data[1:]):
    row_h = Cm(1.9)
    bg = DM_LIGHT if r_i % 2 == 0 else WHITE
    add_rect(s, start_x, y, sum(col_widths, Cm(0)), row_h, bg, line_color=LIGHT_GRAY)
    x = start_x
    for i, c in enumerate(row):
        add_text(s, x + Cm(0.15), y + Cm(0.15), col_widths[i] - Cm(0.3), row_h - Cm(0.3),
                 c, size=10, color=DARK_GRAY,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
        x += col_widths[i]
    y += row_h


# =============================================================
#  SLIDE 29: 二人を並べて比較
# =============================================================
s = add_slide()
add_header(s, "6. ケアプランへの落とし込み", "C さん × D さん ― 同じ「基本ケア」、違う「物語」")
add_footer(s, 29)

# 左:C
add_rect(s, Cm(1.2), Cm(4.3), Cm(15.5), Cm(0.9), STROKE_COLOR)
add_text(s, Cm(1.5), Cm(4.35), Cm(15), Cm(0.8),
         "C さん  ・  脳血管疾患",
         size=15, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, Cm(1.2), Cm(5.2), Cm(15.5), Cm(11.5), STROKE_LIGHT)
add_text(s, Cm(1.5), Cm(5.4), Cm(15), Cm(11.3),
         "【尊厳・自立】役割の再獲得が核\n   → 「孫の運動会」\n\n"
         "【全体像】夫婦単位で見る\n   → 妻の介護負担も同時に\n\n"
         "【意思決定】沈黙を選択肢として尊重\n   → 体験・時間を与える\n\n"
         "【リスク管理】再発予防が最優先\n   → 抗凝固・血圧・脱水\n\n"
         "■ 時間軸は「回復〜拡大」\n"
         "■ 主任CMの問い:\n"
         "   「数値と意味、両方の目標があるか?」",
         size=12, color=DARK_GRAY, line_spacing=1.35)

# 右:D
add_rect(s, Cm(17.5), Cm(4.3), Cm(15), Cm(0.9), DM_COLOR)
add_text(s, Cm(17.8), Cm(4.35), Cm(14.5), Cm(0.8),
         "D さん  ・  糖尿病",
         size=15, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_rect(s, Cm(17.5), Cm(5.2), Cm(15), Cm(11.5), DM_LIGHT)
add_text(s, Cm(17.8), Cm(5.4), Cm(14.5), Cm(11.3),
         "【尊厳・自立】楽しみとの両立が核\n   → 料理を続ける\n\n"
         "【全体像】独居 ・ 息子との関係\n   → 「正しさ」が壁にならない支援\n\n"
         "【意思決定】30年の闘病史を尊重\n   → 「我慢」ではなく「工夫」\n\n"
         "【リスク管理】合併症・低血糖が最優先\n   → 足・眼・救急対応\n\n"
         "■ 時間軸は「数十年の慢性」\n"
         "■ 主任CMの問い:\n"
         "   「禁止 ではなく 工夫 になっているか?」",
         size=12, color=DARK_GRAY, line_spacing=1.35)


# =============================================================
#  SLIDE 30: 第7部 タイトル
# =============================================================
section_title(7, "事例検討・グループ共有・まとめ",
              "Group Work and Closing")


# =============================================================
#  SLIDE 31: 複合疾患の事例
# =============================================================
s = add_slide()
add_header(s, "7. 事例検討", "事例:Eさん(78歳・男性) ― 脳卒中+糖尿病")
add_footer(s, 31)

add_box(s, Cm(1.2), Cm(4.5), Cm(31.4), Cm(5),
        "プロフィール",
        "・78歳 男性 要介護2\n"
        "・脳梗塞後遺症(右片麻痺・軽度失語)・2型糖尿病(HbA1c 7.8%)\n"
        "・妻(75歳・パートで週3日就労)と二人暮らし\n"
        "・元会社員、寡黙だが家族思い\n"
        "・「妻に迷惑をかけたくない。一人で何でもできるようになりたい」",
        title_bg=ORANGE, body_size=13)

add_box(s, Cm(1.2), Cm(9.7), Cm(31.4), Cm(7.3),
        "ワーク内容(3〜4人組・10分)",
        "Q1. Eさんに「基本ケア4つの柱」のうち、どこを最も濃く書きますか?(2分)\n\n"
        "Q2. 「脳血管疾患」と「糖尿病」の 想定される支援内容、どちらを優先しますか?\n"
        "       両立させるとしたら、どんな工夫が必要ですか?(4分)\n\n"
        "Q3. 妻の就労継続を支える支援を、ケアプラン第2表に1行で書いてみよう(2分)\n\n"
        "Q4. 共有 ・ 全体発表(2分)",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=13)


# =============================================================
#  SLIDE 32: ディスカッション論点
# =============================================================
s = add_slide()
add_header(s, "7. 事例検討", "複合疾患を扱うときの 4 つの論点")
add_footer(s, 32)

points = [
    ("論点 1", "優先順位の付け方",
     "・「命の危険」が最優先\n・両疾患のリスクを並べて整理\n・主治医と相談する場面を作る"),
    ("論点 2", "サービスの重複と統合",
     "・訪問看護で両疾患を診る\n・通所リハで運動+栄養指導\n・複数科受診の情報集約"),
    ("論点 3", "家族の理解と教育",
     "・妻に両疾患の特徴を伝える\n・低血糖と脳梗塞再発のサイン\n・緊急時の判断基準"),
    ("論点 4", "本人の意思の核を見つける",
     "・「一人で何でもできるように」\n  という言葉の背景を聴く\n・小さな成功体験の積み重ね"),
]
col_w = Cm(15.5)
row_h = Cm(5.5)
gap_x = Cm(0.4)
gap_y = Cm(0.3)
start_x = Cm(1.2)
start_y = Cm(4.5)
colors = [STROKE_COLOR, ORANGE, CM_COLOR, DM_COLOR]
for i, (label, title, body) in enumerate(points):
    col = i % 2
    row = i // 2
    x = start_x + (col_w + gap_x) * col
    y = start_y + (row_h + gap_y) * row
    color = colors[i]
    add_rect(s, x, y, col_w, Cm(1.2), color)
    add_text(s, x + Cm(0.2), y + Cm(0.1), Cm(3), Cm(1),
             label, size=12, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Cm(3.3), y + Cm(0.1), col_w - Cm(3.5), Cm(1),
             title, size=14, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, x, y + Cm(1.2), col_w, row_h - Cm(1.2), ACCENT_BG)
    add_text(s, x + Cm(0.3), y + Cm(1.4), col_w - Cm(0.6), row_h - Cm(1.4),
             body, size=12, color=DARK_GRAY)


# =============================================================
#  SLIDE 33: 主任CMへのエール
# =============================================================
s = add_slide()
add_header(s, "Closing", "後輩CMへ ― あなたが手法を伝える側になる")
add_footer(s, 33)

add_dialog(s, Cm(1.2), Cm(4.5), Cm(31.4),
           "ある主任CMの言葉",
           "適切なケアマネジメント手法は、答えを教えてくれません。\n"
           " でも、「次に何を問うか」を教えてくれます。\n"
           " 後輩には、答えではなく、問いの順序を伝えてあげてください。",
           speaker_color=NAVY, body_bg=ACCENT_BG, size=15, height=Cm(4))

add_box(s, Cm(1.2), Cm(9), Cm(31.4), Cm(8),
        "明日からの 3 ステップ",
        "STEP 1  自分の担当ケースを 1 件、「手引き」の項目に当てはめて棚卸ししてみる\n"
        "             → 抜けていた視点・濃淡を再確認する\n\n"
        "STEP 2  後輩のケース会議で、「想定される支援内容(仮説)」 → 「アセスメントで検証」\n"
        "             の順序で発表させる。司会としてこの構造を守る\n\n"
        "STEP 3  事業所内で「支援を入れなかった根拠」も記録に残す文化を作る\n"
        "             → これが「説明責任を果たすケアマネジメント」",
        title_bg=NAVY, body_bg=ACCENT_BG, body_size=14)


# =============================================================
#  SLIDE 34: まとめ
# =============================================================
s = add_slide()
add_header(s, "Summary", "本日の学びの整理")
add_footer(s, 34)

add_box(s, Cm(1.2), Cm(4.5), Cm(15.5), Cm(12.5),
        "C さんの物語が教えてくれたこと",
        "■ ADLの数値の先に、「本人の意味」がある\n   ― 「孫の運動会に行きたい」\n\n"
        "■ 「沈黙」も意思表示\n   ― 結論を急がず、体験と時間を提供\n\n"
        "■ 再発予防 と ADL再構築 は両輪\n   ― どちらかに偏らない\n\n"
        "■ 妻(家族介護者)は もう一人の利用者\n   ― 「夫婦単位」でアセスメント\n\n"
        "■ 「時間軸」を持つケアプランを\n   ― 退院 → 安定 → 拡大の見通し",
        title_bg=STROKE_COLOR, body_bg=STROKE_LIGHT, body_size=14)

add_box(s, Cm(17.5), Cm(4.5), Cm(15), Cm(12.5),
        "D さんの物語が教えてくれたこと",
        "■ 「我慢」ではなく「工夫」の提案を\n   ― 楽しみと療養の両立\n\n"
        "■ 30 年の闘病史を尊重する\n   ― 「コンプライアンス不良」と裁かない\n\n"
        "■ 家族の「正しさ」が壁になることも\n   ― 間に立つのが CM の仕事\n\n"
        "■ 合併症は CM が結節点として防ぐ\n   ― 多科受診の情報を集約する\n\n"
        "■ 慢性疾患は「数十年単位」の支援\n   ― 「卒業」より「伴走」が基本",
        title_bg=DM_COLOR, body_bg=DM_LIGHT, body_size=14)


# =============================================================
#  SLIDE 35: 共通の学び
# =============================================================
s = add_slide()
add_header(s, "Summary", "二人の物語に共通する ― 主任CMの核となる姿勢")
add_footer(s, 35)

points = [
    ("①", "仮説を持って、現場に立つ",
     "「想定される支援内容」を頭に入れて訪問する。\n"
     "ゼロから考えるのではなく、検証する姿勢で。"),
    ("②", "本人の言葉に翻訳する",
     "支援者の言葉(再発予防・コンプライアンス)を、\n"
     "本人の言葉(運動会に行く・料理を続ける)に置き換える。"),
    ("③", "プロセスを記録に残す",
     "「なぜ入れたか」だけでなく「なぜ入れなかったか」も。\n"
     "支援経過にこそ、説明責任の根拠が宿る。"),
    ("④", "後輩に「問いの順序」を伝える",
     "答えを与えるのではなく、構造化された問いを与える。\n"
     "これが事業所の力を底上げする。"),
]
y = Cm(4.5)
for label, title, body in points:
    # 番号
    add_rect(s, Cm(1.2), y, Cm(2), Cm(2.5), NAVY)
    add_text(s, Cm(1.2), y, Cm(2), Cm(2.5),
             label, size=32, bold=True, color=ORANGE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 内容
    add_rect(s, Cm(3.2), y, SW - Cm(4.4), Cm(2.5), ACCENT_BG)
    add_text(s, Cm(3.5), y + Cm(0.1), SW - Cm(5), Cm(1),
             title, size=16, bold=True, color=NAVY)
    add_text(s, Cm(3.5), y + Cm(1.05), SW - Cm(5), Cm(1.5),
             body, size=12, color=DARK_GRAY, line_spacing=1.25)
    y += Cm(2.8)


# =============================================================
#  SLIDE 36: 参考資料
# =============================================================
s = add_slide()
add_header(s, "Reference", "参考資料・関連ガイドライン")
add_footer(s, 36)

refs = [
    "■ 厚生労働省「適切なケアマネジメント手法の手引き」(脳血管疾患・糖尿病の章)",
    "■ 日本脳卒中学会「脳卒中治療ガイドライン」",
    "■ 日本糖尿病学会「糖尿病診療ガイドライン」「高齢者糖尿病診療ガイドライン」",
    "■ 厚生労働省「人生の最終段階における医療・ケアの決定プロセスに関するガイドライン」",
    "■ 日本ケアマネジメント学会「課題分析標準項目に関する解説」",
    "■ 国際生活機能分類(ICF) ― WHO",
    "■ 介護保険最新情報(課題分析標準項目 R5改訂)",
    "■ 「適切なケアマネジメント手法」事例集(地域別研修資料)",
]
y = Cm(4.5)
for r in refs:
    add_text(s, Cm(2), y, SW - Cm(4), Cm(1),
             r, size=15, color=DARK_GRAY)
    y += Cm(1.3)


# =============================================================
#  SLIDE 37: エピローグ
# =============================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, NAVY)
add_rect(s, 0, Cm(8), SW, Cm(0.15), ORANGE)

add_text(s, Cm(2), Cm(4), SW - Cm(4), Cm(1.5),
         "Epilogue",
         size=22, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)

add_text(s, Cm(2), Cm(6.5), SW - Cm(4), Cm(3),
         "C さんは今日、孫を助手席に乗せて\n"
         "近所のコンビニまで運転した。",
         size=22, color=WHITE, align=PP_ALIGN.CENTER, line_spacing=1.4)

add_text(s, Cm(2), Cm(11), SW - Cm(4), Cm(3),
         "D さんは今日、料理教室で\n"
         "低糖質の和菓子を仲間と作っていた。",
         size=22, color=WHITE, align=PP_ALIGN.CENTER, line_spacing=1.4)

add_text(s, Cm(2), Cm(15.5), SW - Cm(4), Cm(2),
         "二人の物語は、まだ続いていく。\nそしてあなたの担当ケースも、また明日から物語が動き出す。",
         size=15, color=LIGHT_GRAY, align=PP_ALIGN.CENTER, line_spacing=1.4)


# =============================================================
#  SLIDE 38: Thank you
# =============================================================
s = add_slide()
add_rect(s, 0, 0, SW, SH, NAVY)
add_rect(s, 0, Cm(8), SW, Cm(0.15), ORANGE)

add_text(s, Cm(2), Cm(6), SW - Cm(4), Cm(2),
         "Thank you",
         size=54, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

add_text(s, Cm(2), Cm(10.5), SW - Cm(4), Cm(3),
         "ご質問・ご意見、そしてあなた自身の「物語」を\n"
         "ぜひお聞かせください。",
         size=18, color=LIGHT_GRAY, align=PP_ALIGN.CENTER, line_spacing=1.4)


# ===== 保存 =====
output = "/home/user/kaigo-ga-app/slides/適切なケアマネジメント手法_脳血管糖尿病編_研修スライド.pptx"
prs.save(output)
print(f"Saved: {output}")
print(f"Total slides: {len(prs.slides)}")
