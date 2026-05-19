"""ハラスメントについてのスライド生成スクリプト"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# カラーパレット（ハラスメント＝赤系を基調に）
COLOR_PRIMARY = RGBColor(0x8B, 0x1E, 0x3F)      # ディープレッド
COLOR_ACCENT = RGBColor(0xE8, 0xA5, 0x3A)       # ゴールド
COLOR_LIGHT = RGBColor(0xF7, 0xF3, 0xEE)        # アイボリー
COLOR_DARK = RGBColor(0x22, 0x2B, 0x3A)         # 濃紺
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_GRAY = RGBColor(0x66, 0x70, 0x80)
COLOR_TEAL = RGBColor(0x2E, 0x7D, 0x7A)
COLOR_NAVY = RGBColor(0x1F, 0x3A, 0x5F)

FONT_JP = "Yu Gothic"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

blank_layout = prs.slide_layouts[6]


def add_bg(slide, color):
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    bg.shadow.inherit = False
    return bg


def add_text(slide, left, top, width, height, text, size=18, bold=False,
             color=COLOR_DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             font=FONT_JP):
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
        run = p.add_run()
        run.text = line
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
    return tb


def add_rect(slide, left, top, width, height, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    shp.shadow.inherit = False
    return shp


def add_round_rect(slide, left, top, width, height, fill, line=None):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    shp.shadow.inherit = False
    return shp


def page_header(slide, title, page_num, total):
    add_rect(slide, 0, 0, prs.slide_width, Inches(1.05), COLOR_PRIMARY)
    add_text(slide, Inches(0.6), Inches(0.22), Inches(11), Inches(0.7),
             title, size=28, bold=True, color=COLOR_WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(11.8), Inches(0.32), Inches(1.3), Inches(0.5),
             f"{page_num} / {total}", size=12, color=COLOR_WHITE,
             align=PP_ALIGN.RIGHT)
    add_rect(slide, 0, Inches(1.05), prs.slide_width, Inches(0.08), COLOR_ACCENT)
    add_text(slide, Inches(0.6), Inches(7.05), Inches(11), Inches(0.35),
             "ハラスメントを知る・防ぐ・なくす", size=10, color=COLOR_GRAY)


TOTAL = 12

# --------- スライド1: タイトル ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_PRIMARY)
add_rect(s, 0, Inches(5.6), prs.slide_width, Inches(0.15), COLOR_ACCENT)
add_rect(s, 0, Inches(5.85), prs.slide_width, Inches(0.05), COLOR_WHITE)

add_text(s, Inches(0.8), Inches(1.6), Inches(11.7), Inches(1.0),
         "ハラスメントについて", size=60, bold=True, color=COLOR_WHITE)
add_text(s, Inches(0.8), Inches(2.7), Inches(11.7), Inches(0.7),
         "Harassment Awareness & Prevention", size=24, color=COLOR_ACCENT)
add_text(s, Inches(0.8), Inches(3.7), Inches(11.7), Inches(1.4),
         "誰もが安心して働き、学び、暮らせる\n"
         "環境をつくるために",
         size=28, color=COLOR_WHITE)
add_text(s, Inches(0.8), Inches(6.4), Inches(11.7), Inches(0.4),
         "2026年5月", size=14, color=COLOR_LIGHT)

# --------- スライド2: 目次 ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_WHITE)
page_header(s, "目次", 2, TOTAL)

items = [
    ("01", "ハラスメントとは"),
    ("02", "主なハラスメントの種類"),
    ("03", "日本における現状（統計）"),
    ("04", "背景にある要因"),
    ("05", "被害者への影響"),
    ("06", "ハラスメントのサイン"),
    ("07", "法的枠組みと相談窓口"),
    ("08", "私たちにできること"),
    ("09", "まとめ"),
]
col_w = Inches(6.0)
for i, (num, label) in enumerate(items):
    col = i // 5
    row = i % 5
    left = Inches(0.7) + col * col_w
    top = Inches(1.6) + row * Inches(0.95)
    add_round_rect(s, left, top, Inches(0.7), Inches(0.7), COLOR_ACCENT)
    add_text(s, left, top, Inches(0.7), Inches(0.7), num,
             size=18, bold=True, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left + Inches(0.9), top + Inches(0.05),
             Inches(5.0), Inches(0.6), label,
             size=18, bold=True, color=COLOR_DARK,
             anchor=MSO_ANCHOR.MIDDLE)

# --------- スライド3: ハラスメントとは ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_WHITE)
page_header(s, "01. ハラスメントとは", 3, TOTAL)

add_round_rect(s, Inches(0.7), Inches(1.5), Inches(12), Inches(1.7),
               COLOR_LIGHT)
add_text(s, Inches(1.0), Inches(1.6), Inches(0.5), Inches(0.5),
         "“", size=48, bold=True, color=COLOR_ACCENT)
add_text(s, Inches(1.5), Inches(1.7), Inches(10.8), Inches(1.4),
         "相手の意に反する言動により、相手の尊厳を傷つけ、\n"
         "不快感や不利益、就業・学習環境の悪化を生じさせる行為",
         size=20, bold=True, color=COLOR_DARK)
add_text(s, Inches(1.5), Inches(2.8), Inches(10.8), Inches(0.4),
         "（厚生労働省 等の定義をもとに）",
         size=12, color=COLOR_GRAY)

# 3つのポイント
points = [
    ("受け手基準", "「冗談のつもり」でも\n相手が苦痛なら成立しうる"),
    ("関係性が問われる", "優位性・継続性・\n断りにくさが鍵となる"),
    ("組織の責任", "個人間トラブルではなく\n組織の問題として対処"),
]
for i, (title, desc) in enumerate(points):
    left = Inches(0.7) + i * Inches(4.13)
    top = Inches(3.6)
    add_round_rect(s, left, top, Inches(3.93), Inches(2.8), COLOR_WHITE,
                   line=COLOR_PRIMARY)
    add_rect(s, left, top, Inches(3.93), Inches(0.7), COLOR_PRIMARY)
    add_text(s, left, top, Inches(3.93), Inches(0.7), title,
             size=18, bold=True, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left + Inches(0.2), top + Inches(0.9),
             Inches(3.5), Inches(1.8), desc,
             size=15, color=COLOR_DARK, align=PP_ALIGN.CENTER)

# --------- スライド4: 主なハラスメントの種類 ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_WHITE)
page_header(s, "02. 主なハラスメントの種類", 4, TOTAL)

add_text(s, Inches(0.7), Inches(1.3), Inches(12), Inches(0.4),
         "職場・学校・家庭など、あらゆる場面で発生しうる多様なハラスメント。",
         size=14, color=COLOR_GRAY)

types = [
    ("パワハラ", "職務上の優位性を背景にした\n精神的・身体的攻撃"),
    ("セクハラ", "性的な言動による\n就業環境を害する行為"),
    ("マタハラ", "妊娠・出産・育児を\n理由とする不利益取扱い"),
    ("モラハラ", "言葉・態度による\n継続的な精神的暴力"),
    ("カスハラ", "顧客・取引先からの\n著しい迷惑行為"),
    ("アカハラ", "教育・研究上の\n地位を利用した嫌がらせ"),
    ("ジェンハラ", "性別役割を押し付ける\n差別的言動"),
    ("リモハラ", "テレワーク下での\n過度な監視・干渉"),
]
cols, rows = 4, 2
cell_w = Inches(3.0)
cell_h = Inches(2.1)
start_left = Inches(0.7)
start_top = Inches(1.85)
for i, (t, d) in enumerate(types):
    r = i // cols
    c = i % cols
    left = start_left + c * Inches(3.1)
    top = start_top + r * Inches(2.25)
    add_round_rect(s, left, top, cell_w, cell_h, COLOR_LIGHT)
    add_rect(s, left + Inches(0.15), top + Inches(0.25),
             Inches(0.12), Inches(0.45), COLOR_PRIMARY)
    add_text(s, left + Inches(0.35), top + Inches(0.2),
             cell_w - Inches(0.4), Inches(0.5), t,
             size=18, bold=True, color=COLOR_PRIMARY)
    add_text(s, left + Inches(0.35), top + Inches(0.85),
             cell_w - Inches(0.4), Inches(1.1), d,
             size=13, color=COLOR_DARK)

# --------- スライド5: 日本における現状 ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_WHITE)
page_header(s, "03. 日本における現状", 5, TOTAL)

add_text(s, Inches(0.7), Inches(1.3), Inches(12), Inches(0.5),
         "厚生労働省の実態調査では、依然として多くの人が被害を経験しています。",
         size=14, color=COLOR_GRAY)

stats = [
    ("過去3年でパワハラを受けた人", "約 3 人に 1 人", "全業種共通の課題"),
    ("セクハラを経験した人", "約 10 人に 1 人", "報告されにくい傾向"),
    ("カスハラを受けた人", "約 3 人に 1 人", "サービス業を中心に増加"),
]
for i, (label, num, desc) in enumerate(stats):
    left = Inches(0.7) + i * Inches(4.13)
    top = Inches(2.0)
    add_round_rect(s, left, top, Inches(3.93), Inches(3.0),
                   COLOR_PRIMARY)
    add_text(s, left, top + Inches(0.25), Inches(3.93), Inches(0.5),
             label, size=14, bold=True, color=COLOR_LIGHT,
             align=PP_ALIGN.CENTER)
    add_text(s, left, top + Inches(0.9), Inches(3.93), Inches(1.2),
             num, size=26, bold=True, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, left + Inches(1.2), top + Inches(2.15),
             Inches(1.53), Inches(0.04), COLOR_ACCENT)
    add_text(s, left, top + Inches(2.3), Inches(3.93), Inches(0.6),
             desc, size=12, color=COLOR_LIGHT,
             align=PP_ALIGN.CENTER)

add_round_rect(s, Inches(0.7), Inches(5.3), Inches(12), Inches(1.4),
               COLOR_LIGHT)
add_text(s, Inches(1.0), Inches(5.45), Inches(11.5), Inches(0.5),
         "■ 顕在化しにくい構造", size=15, bold=True, color=COLOR_PRIMARY)
add_text(s, Inches(1.0), Inches(5.9), Inches(11.5), Inches(0.8),
         "・被害を相談しても「何も変わらなかった」という声が多数\n"
         "・「自分にも非がある」と感じてしまい、沈黙してしまうケースも目立つ",
         size=13, color=COLOR_DARK)

# --------- スライド6: 背景にある要因 ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_WHITE)
page_header(s, "04. 背景にある要因", 6, TOTAL)

factors = [
    ("閉鎖的な組織風土", "声を上げづらく\n問題が表面化しにくい"),
    ("過度な上下関係", "「指導」と「攻撃」の\n線引きが曖昧になる"),
    ("長時間労働・高負荷", "余裕のなさが\n他者への攻撃性を生む"),
    ("価値観の多様化への遅れ", "性別・年齢・国籍など\n無意識の偏見が残る"),
    ("教育・研修の不足", "何がハラスメントか\n共通理解がない"),
    ("相談制度の不備", "通報しても\n適切に処理されない"),
]
for i, (t, d) in enumerate(factors):
    r = i // 3
    c = i % 3
    left = Inches(0.7) + c * Inches(4.13)
    top = Inches(1.6) + r * Inches(2.7)
    add_round_rect(s, left, top, Inches(3.93), Inches(2.5),
                   COLOR_WHITE, line=COLOR_PRIMARY)
    add_round_rect(s, left + Inches(0.3), top - Inches(0.2),
                   Inches(0.7), Inches(0.7), COLOR_PRIMARY)
    add_text(s, left + Inches(0.3), top - Inches(0.2),
             Inches(0.7), Inches(0.7), str(i + 1),
             size=18, bold=True, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left + Inches(1.2), top + Inches(0.05),
             Inches(2.6), Inches(0.6), t,
             size=16, bold=True, color=COLOR_PRIMARY,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left + Inches(0.3), top + Inches(1.0),
             Inches(3.3), Inches(1.3), d,
             size=13, color=COLOR_DARK)

# --------- スライド7: 被害者への影響 ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_WHITE)
page_header(s, "05. 被害者への影響", 7, TOTAL)

sections = [
    ("心の健康", COLOR_PRIMARY, [
        "・うつ・不安障害",
        "・自己肯定感の低下",
        "・PTSD・フラッシュバック",
        "・睡眠障害",
    ]),
    ("身体の健康", COLOR_ACCENT, [
        "・頭痛、腹痛、めまい",
        "・食欲不振、過食",
        "・慢性疲労",
        "・心身症",
    ]),
    ("仕事・学業", COLOR_TEAL, [
        "・集中力・生産性の低下",
        "・休職・退職・転職",
        "・キャリアの中断",
        "・経済的な損失",
    ]),
    ("人間関係", COLOR_NAVY, [
        "・人間不信、孤立",
        "・家族・友人との関係悪化",
        "・社会参加への不安",
        "・新しい職場での萎縮",
    ]),
]
positions = [
    (Inches(0.7), Inches(1.5)),
    (Inches(6.93), Inches(1.5)),
    (Inches(0.7), Inches(4.4)),
    (Inches(6.93), Inches(4.4)),
]
for (t, color, items), (left, top) in zip(sections, positions):
    add_round_rect(s, left, top, Inches(5.7), Inches(2.7),
                   COLOR_LIGHT)
    add_rect(s, left, top, Inches(0.18), Inches(2.7), color)
    add_text(s, left + Inches(0.4), top + Inches(0.15),
             Inches(5.2), Inches(0.5), t,
             size=20, bold=True, color=color)
    add_text(s, left + Inches(0.4), top + Inches(0.75),
             Inches(5.2), Inches(1.9), "\n".join(items),
             size=14, color=COLOR_DARK)

# --------- スライド8: ハラスメントのサイン ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_WHITE)
page_header(s, "06. ハラスメントのサイン", 8, TOTAL)

add_text(s, Inches(0.7), Inches(1.3), Inches(12), Inches(0.5),
         "同僚・部下・友人にこんな変化が見られたら、声をかけてみてください。",
         size=14, color=COLOR_GRAY)

signs = [
    "遅刻・欠勤・早退が増えた",
    "表情が乏しく、笑顔が減った",
    "特定の人物の前で急に萎縮する",
    "メール・会話で言い回しが防衛的になる",
    "業務上のミスや物忘れが目立つ",
    "体調不良の訴え、通院が増える",
    "飲み会や雑談を極端に避けるようになる",
    "「自分が悪い」と過度に自責する",
]
for i, sign in enumerate(signs):
    r = i // 2
    c = i % 2
    left = Inches(0.7) + c * Inches(6.13)
    top = Inches(2.0) + r * Inches(1.05)
    add_round_rect(s, left, top, Inches(5.93), Inches(0.85),
                   COLOR_LIGHT)
    add_round_rect(s, left + Inches(0.25), top + Inches(0.18),
                   Inches(0.5), Inches(0.5), COLOR_PRIMARY)
    add_text(s, left + Inches(0.25), top + Inches(0.18),
             Inches(0.5), Inches(0.5), "!",
             size=18, bold=True, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left + Inches(0.95), top, Inches(4.9), Inches(0.85),
             sign, size=15, color=COLOR_DARK,
             anchor=MSO_ANCHOR.MIDDLE)

add_round_rect(s, Inches(0.7), Inches(6.45), Inches(12), Inches(0.5),
               COLOR_PRIMARY)
add_text(s, Inches(0.7), Inches(6.45), Inches(12), Inches(0.5),
         "ポイント：「大丈夫？」より「最近どう？」とフラットに尋ねる。",
         size=14, bold=True, color=COLOR_WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# --------- スライド9: 法的枠組みと相談窓口 ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_WHITE)
page_header(s, "07. 法的枠組みと相談窓口", 9, TOTAL)

# 左：法的枠組み
add_round_rect(s, Inches(0.7), Inches(1.5), Inches(6), Inches(5.2),
               COLOR_LIGHT)
add_rect(s, Inches(0.7), Inches(1.5), Inches(6), Inches(0.6),
         COLOR_PRIMARY)
add_text(s, Inches(0.7), Inches(1.5), Inches(6), Inches(0.6),
         "主な法的枠組み",
         size=18, bold=True, color=COLOR_WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
laws = [
    ("労働施策総合推進法", "パワハラ防止を企業に義務化（2022年〜）"),
    ("男女雇用機会均等法", "セクハラ・マタハラ防止措置を義務化"),
    ("育児・介護休業法", "育児休業等に関するハラスメント防止"),
    ("民法（不法行為）", "損害賠償請求の根拠となる"),
    ("刑法", "暴行・脅迫・名誉毀損などに該当する場合"),
    ("カスハラ対策（東京都条例 等）", "顧客からの著しい迷惑行為への対応"),
]
for i, (name, desc) in enumerate(laws):
    top = Inches(2.3) + i * Inches(0.7)
    add_text(s, Inches(1.0), top, Inches(5.5), Inches(0.3), "● " + name,
             size=13, bold=True, color=COLOR_PRIMARY)
    add_text(s, Inches(1.2), top + Inches(0.3), Inches(5.3), Inches(0.3),
             desc, size=11, color=COLOR_GRAY)

# 右：相談窓口
add_round_rect(s, Inches(6.93), Inches(1.5), Inches(5.8), Inches(5.2),
               COLOR_LIGHT)
add_rect(s, Inches(6.93), Inches(1.5), Inches(5.8), Inches(0.6),
         COLOR_ACCENT)
add_text(s, Inches(6.93), Inches(1.5), Inches(5.8), Inches(0.6),
         "相談できる窓口（一例）",
         size=18, bold=True, color=COLOR_WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
contacts = [
    ("総合労働相談コーナー", "各都道府県労働局"),
    ("ハラスメント悩み相談室", "0120-714-864"),
    ("法テラス", "0570-078-374"),
    ("みんなの人権110番", "0570-003-110"),
    ("社内ハラスメント相談窓口", "各企業の窓口・産業医"),
]
for i, (name, tel) in enumerate(contacts):
    top = Inches(2.3) + i * Inches(0.85)
    add_text(s, Inches(7.2), top, Inches(5.3), Inches(0.4), name,
             size=14, bold=True, color=COLOR_DARK)
    add_text(s, Inches(7.2), top + Inches(0.38), Inches(5.3), Inches(0.4),
             "☎ " + tel, size=13, color=COLOR_PRIMARY, bold=True)

# --------- スライド10: 私たちにできること ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_WHITE)
page_header(s, "08. 私たちにできること", 10, TOTAL)

roles = [
    ("経営層・管理職", COLOR_PRIMARY, [
        "・トップの明確な宣言",
        "・研修を継続的に実施",
        "・相談・通報の窓口整備",
        "・処分基準を明確化する",
    ]),
    ("同僚・チーム", COLOR_ACCENT, [
        "・違和感を見過ごさない",
        "・被害者を孤立させない",
        "・「冗談だから」と流さない",
        "・記録に残し、共有する",
    ]),
    ("被害を受けた本人", COLOR_TEAL, [
        "・「自分が悪い」と思わない",
        "・日時・内容を記録する",
        "・第三者・専門窓口に相談",
        "・健康と安全を最優先に",
    ]),
    ("私たち一人ひとり", COLOR_NAVY, [
        "・正しい知識を持つ",
        "・無意識の偏見を見直す",
        "・受け手の感じ方を尊重",
        "・声を上げる人を守る",
    ]),
]
positions = [
    (Inches(0.7), Inches(1.5)),
    (Inches(6.93), Inches(1.5)),
    (Inches(0.7), Inches(4.4)),
    (Inches(6.93), Inches(4.4)),
]
for (t, color, items), (left, top) in zip(roles, positions):
    add_round_rect(s, left, top, Inches(5.7), Inches(2.7),
                   COLOR_WHITE, line=color)
    add_rect(s, left, top, Inches(5.7), Inches(0.55), color)
    add_text(s, left, top, Inches(5.7), Inches(0.55), t,
             size=18, bold=True, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, left + Inches(0.3), top + Inches(0.75),
             Inches(5.2), Inches(1.9), "\n".join(items),
             size=14, color=COLOR_DARK)

# --------- スライド11: まとめ ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_WHITE)
page_header(s, "09. まとめ", 11, TOTAL)

messages = [
    ("ハラスメントは「個人の問題」ではなく\n「組織と社会の問題」。",
     "属人的なトラブル処理ではなく、\n仕組みと文化で防ぐ視点が不可欠。"),
    ("「受け手がどう感じたか」を中心に考える。",
     "発信側の意図ではなく、\n相手の尊厳が守られているかが基準となる。"),
    ("沈黙させない、孤立させない。",
     "声を上げられる仕組みと、\n上げた人を守る文化を同時に育てていく。"),
]
for i, (title, desc) in enumerate(messages):
    top = Inches(1.55) + i * Inches(1.75)
    add_round_rect(s, Inches(0.7), top, Inches(12), Inches(1.55),
                   COLOR_LIGHT)
    add_rect(s, Inches(0.7), top, Inches(0.2), Inches(1.55), COLOR_PRIMARY)
    add_round_rect(s, Inches(1.1), top + Inches(0.3),
                   Inches(0.95), Inches(0.95), COLOR_PRIMARY)
    add_text(s, Inches(1.1), top + Inches(0.3),
             Inches(0.95), Inches(0.95), str(i + 1),
             size=28, bold=True, color=COLOR_WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(2.3), top + Inches(0.2),
             Inches(10.2), Inches(0.7), title,
             size=18, bold=True, color=COLOR_PRIMARY)
    add_text(s, Inches(2.3), top + Inches(0.85),
             Inches(10.2), Inches(0.7), desc,
             size=13, color=COLOR_DARK)

# --------- スライド12: 終わりに ---------
s = prs.slides.add_slide(blank_layout)
add_bg(s, COLOR_PRIMARY)
add_rect(s, 0, Inches(5.6), prs.slide_width, Inches(0.15), COLOR_ACCENT)

add_text(s, Inches(0.8), Inches(2.0), Inches(11.7), Inches(1.0),
         "Thank you", size=56, bold=True, color=COLOR_WHITE)
add_text(s, Inches(0.8), Inches(3.0), Inches(11.7), Inches(0.7),
         "ご清聴ありがとうございました。",
         size=22, color=COLOR_ACCENT)
add_text(s, Inches(0.8), Inches(4.2), Inches(11.7), Inches(1.5),
         "「気づくこと」と「動くこと」が、職場と社会を変えます。\n"
         "となりの誰かを、そしてあなた自身を守るために。",
         size=18, color=COLOR_WHITE)
add_text(s, Inches(0.8), Inches(6.5), Inches(11.7), Inches(0.4),
         "References: 厚生労働省「職場のハラスメントに関する実態調査」ほか",
         size=11, color=COLOR_LIGHT)

out = "/home/user/kaigo-ga-app/harassment.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
