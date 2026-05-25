"""法定研修制度改正への対応と主任ケアマネ更新研修の戦略 - PowerPoint生成スクリプト"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ====== カラーパレット ======
PRIMARY      = RGBColor(0x0F, 0x4C, 0x81)
PRIMARY_LIGHT= RGBColor(0x2D, 0x6F, 0xB3)
ACCENT       = RGBColor(0xF5, 0xA6, 0x23)
BG           = RGBColor(0xF5, 0xF7, 0xFA)
TEXT         = RGBColor(0x1F, 0x2D, 0x3D)
MUTED        = RGBColor(0x5A, 0x6B, 0x80)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
BORDER       = RGBColor(0xD9, 0xE1, 0xEC)
CARD         = RGBColor(0xFF, 0xFF, 0xFF)
CALLOUT_BG   = RGBColor(0xFF, 0xF8, 0xE6)
CALLOUT_TEXT = RGBColor(0x6B, 0x4E, 0x00)
TABLE_ALT    = RGBColor(0xF9, 0xFB, 0xFD)

FONT = "Yu Gothic"

# ====== プレゼン基本設定 (16:9) ======
prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height

blank = prs.slide_layouts[6]  # 完全な空白レイアウト


def add_bg(slide, color=BG):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.fill.solid(); bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    bg.shadow.inherit = False
    return bg


def add_text(slide, x, y, w, h, text, *, size=18, bold=False, color=TEXT,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=FONT):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
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


def add_rich_text(slide, x, y, w, h, runs, *, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                  line_spacing=None):
    """runs: list of lists of (text, {size, bold, color, font})."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0); tf.margin_right = Emu(0)
    tf.margin_top = Emu(0); tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    for i, line in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        for (txt, opts) in line:
            r = p.add_run()
            r.text = txt
            r.font.name = opts.get("font", FONT)
            r.font.size = Pt(opts.get("size", 18))
            r.font.bold = opts.get("bold", False)
            r.font.color.rgb = opts.get("color", TEXT)
    return tb


def add_rect(slide, x, y, w, h, *, fill=WHITE, line=None, line_width=0.75):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_width)
    shp.shadow.inherit = False
    return shp


def add_title(slide, text, page=None, section=None):
    """ページ見出し（左の縦帯 + テキスト）"""
    # 左の縦帯
    add_rect(slide, Inches(0.55), Inches(0.55), Inches(0.12), Inches(0.65), fill=ACCENT)
    add_text(slide, Inches(0.8), Inches(0.5), Inches(11.5), Inches(0.8),
             text, size=28, bold=True, color=PRIMARY)
    # 下線
    add_rect(slide, Inches(0.55), Inches(1.28), Inches(12.2), Inches(0.04), fill=PRIMARY)
    # フッタ
    if section or page:
        add_text(slide, Inches(0.55), Inches(7.05), Inches(8), Inches(0.3),
                 section or "", size=10, color=MUTED)
        if page:
            add_text(slide, Inches(11.5), Inches(7.05), Inches(1.3), Inches(0.3),
                     page, size=10, bold=True, color=PRIMARY, align=PP_ALIGN.RIGHT)
        add_rect(slide, Inches(0.55), Inches(7.0), Inches(12.2), Inches(0.015), fill=BORDER)


def add_card(slide, x, y, w, h, title, items, *, accent=False):
    border = ACCENT if accent else PRIMARY_LIGHT
    # カード本体
    card = add_rect(slide, x, y, w, h, fill=CARD)
    card.line.color.rgb = BORDER; card.line.width = Pt(0.75)
    # 左の縦帯
    add_rect(slide, x, y, Emu(45720), h, fill=border)  # 5pt
    # タイトル
    add_text(slide, x + Inches(0.25), y + Inches(0.12), w - Inches(0.4), Inches(0.4),
             title, size=15, bold=True, color=PRIMARY)
    # 本文
    body_y = y + Inches(0.55)
    body_h = h - Inches(0.65)
    runs = []
    for it in items:
        runs.append([("■  ", {"size": 11, "color": ACCENT, "bold": True}),
                     (it, {"size": 13, "color": TEXT})])
    add_rich_text(slide, x + Inches(0.25), body_y, w - Inches(0.4), body_h,
                  runs, line_spacing=1.3)


def add_callout(slide, x, y, w, h, header, body):
    box = add_rect(slide, x, y, w, h, fill=CALLOUT_BG)
    add_rect(slide, x, y, Emu(45720), h, fill=ACCENT)
    runs = [
        [(header, {"size": 14, "bold": True, "color": CALLOUT_TEXT}),
         (body,   {"size": 14, "color": CALLOUT_TEXT})]
    ]
    add_rich_text(slide, x + Inches(0.25), y + Inches(0.15), w - Inches(0.4), h - Inches(0.2),
                  runs, line_spacing=1.4, anchor=MSO_ANCHOR.MIDDLE)


def add_step(slide, x, y, w, h, num, title, desc):
    add_rect(slide, x, y, w, h, fill=CARD, line=BORDER)
    # トップバー
    add_rect(slide, x, y, w, Emu(50800), fill=PRIMARY_LIGHT)  # 4pt弱
    # 数字バッジ
    cir = slide.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.2), y + Inches(0.2),
                                 Inches(0.45), Inches(0.45))
    cir.fill.solid(); cir.fill.fore_color.rgb = PRIMARY
    cir.line.fill.background(); cir.shadow.inherit = False
    add_text(slide, x + Inches(0.2), y + Inches(0.21), Inches(0.45), Inches(0.45),
             str(num), size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
    # タイトル
    add_text(slide, x + Inches(0.2), y + Inches(0.75), w - Inches(0.3), Inches(0.4),
             title, size=13, bold=True, color=PRIMARY)
    # 説明
    add_text(slide, x + Inches(0.2), y + Inches(1.15), w - Inches(0.3), h - Inches(1.2),
             desc, size=11, color=MUTED)


def add_table(slide, x, y, w, h, headers, rows, col_widths=None):
    nrows = len(rows) + 1
    ncols = len(headers)
    table_shape = slide.shapes.add_table(nrows, ncols, x, y, w, h)
    tbl = table_shape.table

    # 列幅
    if col_widths:
        total = sum(col_widths)
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = int(w * cw / total)

    # ヘッダ
    for ci, ht in enumerate(headers):
        c = tbl.cell(0, ci)
        c.fill.solid(); c.fill.fore_color.rgb = PRIMARY
        c.text = ""
        tf = c.text_frame; tf.word_wrap = True
        tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1)
        tf.margin_top = Inches(0.05); tf.margin_bottom = Inches(0.05)
        p = tf.paragraphs[0]
        r = p.add_run(); r.text = ht
        r.font.name = FONT; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = WHITE
    # 本体
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            c = tbl.cell(ri, ci)
            c.fill.solid()
            c.fill.fore_color.rgb = TABLE_ALT if ri % 2 == 0 else CARD
            c.text = ""
            tf = c.text_frame; tf.word_wrap = True
            tf.margin_left = Inches(0.1); tf.margin_right = Inches(0.1)
            tf.margin_top = Inches(0.05); tf.margin_bottom = Inches(0.05)
            # マークダウン風の **bold** に簡易対応
            p = tf.paragraphs[0]
            segs = []
            buf = val; bold = False
            while "**" in buf:
                pre, _, rest = buf.partition("**")
                segs.append((pre, bold))
                bold = not bold
                buf = rest
            segs.append((buf, bold))
            for txt, b in segs:
                if not txt: continue
                r = p.add_run(); r.text = txt
                r.font.name = FONT; r.font.size = Pt(11); r.font.bold = b
                r.font.color.rgb = TEXT
    return tbl


def add_timeline(slide, x, y, w, h, items):
    """items: [(year, text), ...]"""
    n = len(items)
    # 縦ライン
    line_x = x + Inches(0.4)
    add_rect(slide, line_x, y + Inches(0.2), Emu(38100), h - Inches(0.4), fill=PRIMARY_LIGHT)
    row_h = (h - Inches(0.2)) / n
    for i, (yr, txt) in enumerate(items):
        cy = y + i * row_h + Inches(0.15)
        # ドット
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                     line_x - Inches(0.13), cy,
                                     Inches(0.32), Inches(0.32))
        dot.fill.solid(); dot.fill.fore_color.rgb = ACCENT
        dot.line.color.rgb = PRIMARY_LIGHT; dot.line.width = Pt(2)
        dot.shadow.inherit = False
        # 年
        add_text(slide, x + Inches(0.75), cy - Inches(0.02),
                 Inches(1.6), Inches(0.4),
                 yr, size=13, bold=True, color=PRIMARY)
        # 説明
        add_text(slide, x + Inches(2.4), cy - Inches(0.02),
                 w - Inches(2.6), row_h,
                 txt, size=12, color=TEXT)


# ============================================
# スライド生成
# ============================================

def slide_title():
    s = prs.slides.add_slide(blank)
    # 背景グラデ風（2色重ね）
    bg = add_rect(s, 0, 0, SW, SH, fill=PRIMARY)
    # 右下に薄い装飾
    accent_bar = add_rect(s, 0, SH - Inches(0.4), SW, Inches(0.4), fill=ACCENT)
    deco = slide.shapes.add_shape(MSO_SHAPE.OVAL, SW - Inches(3), -Inches(2),
                                   Inches(6), Inches(6)) if False else None

    # バッジ
    badge = add_rect(s, Inches(1.0), Inches(1.5), Inches(4.5), Inches(0.45),
                     fill=PRIMARY_LIGHT)
    badge.line.fill.background()
    add_text(s, Inches(1.0), Inches(1.5), Inches(4.5), Inches(0.45),
             "介護支援専門員向け 研修資料", size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # タイトル
    add_rich_text(s, Inches(1.0), Inches(2.3), Inches(11.3), Inches(2.2),
                  [[("法定研修制度改正への対応と", {"size": 36, "bold": True, "color": WHITE})],
                   [("主任ケアマネ更新研修の戦略", {"size": 36, "bold": True, "color": WHITE})]],
                  line_spacing=1.3)
    # サブタイトル
    add_rich_text(s, Inches(1.0), Inches(4.6), Inches(11.3), Inches(1.4),
                  [[("研修体系の見直しを踏まえた", {"size": 20, "color": WHITE})],
                   [("事業所内での計画的な人材育成", {"size": 20, "color": WHITE})]],
                  line_spacing=1.4)
    # 区切り
    add_rect(s, Inches(1.0), Inches(6.1), Inches(11.3), Emu(12700), fill=WHITE)
    # 対象
    add_text(s, Inches(1.0), Inches(6.25), Inches(11.3), Inches(0.5),
             "居宅介護支援事業所／地域包括支援センター　管理者・主任介護支援専門員 向け",
             size=13, color=WHITE)

def slide_agenda():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "本日のアジェンダ", page="2 / 21", section="法定研修制度改正への対応")
    items = [
        ("1.", "法定研修制度改正の背景と全体像", "なぜ今、見直しが必要なのか"),
        ("2.", "新研修体系のポイント", "カリキュラム・受講方法の変更点"),
        ("3.", "主任ケアマネ更新研修の現状と課題", "5年ごとの更新をどう乗り切るか"),
        ("4.", "事業所内での計画的人材育成戦略", "キャリアパス・OJT・受講体制"),
        ("5.", "実践事例とチェックリスト", "明日から取り組めるアクション"),
    ]
    y = Inches(1.6)
    for num, title, sub in items:
        add_text(s, Inches(0.8), y, Inches(0.6), Inches(0.5),
                 num, size=18, bold=True, color=ACCENT)
        add_rich_text(s, Inches(1.4), y, Inches(11), Inches(0.55),
                      [[(title, {"size": 18, "bold": True, "color": PRIMARY}),
                        ("　— " + sub, {"size": 14, "color": MUTED})]])
        y += Inches(0.55)
    add_callout(s, Inches(0.8), Inches(5.8), Inches(11.7), Inches(0.85),
                "本日のゴール：",
                "研修制度改正の要点を理解し、自事業所における 3〜5年スパン の人材育成計画を描けるようになる。")


def slide_chapter1_bg():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "第1章　法定研修制度改正の背景", page="3 / 21", section="第1章　改正の背景")
    cw = Inches(5.7); ch = Inches(2.5)
    add_card(s, Inches(0.8), Inches(1.55), cw, ch, "環境変化①　利用者ニーズの多様化",
             ["独居高齢者・認認介護世帯の増加",
              "医療ニーズの高い在宅利用者の増加",
              "ヤングケアラー・8050問題など複合課題"])
    add_card(s, Inches(6.85), Inches(1.55), cw, ch, "環境変化②　制度の深化",
             ["地域包括ケアシステムの深化・推進",
              "「適切なケアマネジメント手法」の普及",
              "科学的介護（LIFE）情報の活用"], accent=True)
    add_card(s, Inches(0.8), Inches(4.2), cw, ch, "現場の課題",
             ["研修時間が長く業務との両立が困難",
              "研修内容の重複・体系の分かりにくさ",
              "主任ケアマネ不足と更新の負担"])
    add_card(s, Inches(6.85), Inches(4.2), cw, ch, "制度的要請",
             ["厚労省「介護支援専門員資質向上方策」",
              "法定研修ガイドラインの改訂",
              "ICT活用（eラーニング）の推進"], accent=True)


def slide_schedule():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "改正スケジュールと法定研修体系", page="4 / 21", section="第1章　改正の背景")
    # 左カラム: タイムライン
    add_text(s, Inches(0.8), Inches(1.55), Inches(5.5), Inches(0.4),
             "施行と運用", size=16, bold=True, color=PRIMARY_LIGHT)
    add_timeline(s, Inches(0.8), Inches(2.05), Inches(5.7), Inches(3.6),
                 [("令和5年4月", "厚労省「介護支援専門員資質向上事業ガイドライン」公表"),
                  ("令和6年4月", "新ガイドライン施行。オンライン研修は手引き等を踏まえ各実施機関が活用"),
                  ("運用",       "実際の研修実施時期・運用は都道府県ごとの募集要項を確認")])
    add_text(s, Inches(0.8), Inches(5.7), Inches(5.7), Inches(0.8),
             "※施行時期・募集回数・運用方法は都道府県・実施機関により異なるため、所管自治体の最新通知を必ず確認のこと。",
             size=11, color=MUTED)
    # 右カラム: 表
    add_text(s, Inches(6.85), Inches(1.55), Inches(5.7), Inches(0.4),
             "法定研修の全体像", size=16, bold=True, color=PRIMARY_LIGHT)
    add_table(s, Inches(6.85), Inches(2.05), Inches(5.7), Inches(4.2),
              ["研修名", "対象"],
              [["実務研修", "新規登録予定者"],
               ["専門研修課程I", "実務経験概ね2年〜"],
               ["専門研修課程II", "更新時（概ね5年）"],
               ["主任介護支援専門員研修", "主任ケアマネ希望者"],
               ["主任介護支援専門員更新研修", "主任ケアマネ5年ごと"],
               ["再研修", "登録失効者の再登録時"]],
              col_widths=[3, 4])


def slide_main_points():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "新研修体系の主要な改正ポイント", page="5 / 21", section="第2章　新研修体系のポイント")
    sw_each = Inches(2.32); sh_each = Inches(2.6)
    gap = Inches(0.1)
    x0 = Inches(0.8)
    steps = [
        ("1", "カリキュラム整理", "科目間の重複を整理し、実践演習を中心とした構成へ。"),
        ("2", "オンライン研修活用", "知識習得部分はオンライン化が進展。集合研修は演習・対話中心に。"),
        ("3", "適切なケアマネジメント手法", "適切なケアマネジメント手法を踏まえた実践的な視点が重視。"),
        ("4", "多職種連携の強化", "医療・介護・地域資源との連携演習を充実。"),
        ("5", "評価・振り返り", "研修記録シート・修了評価等で到達目標の確認・振り返りを重視。"),
    ]
    for i, (n, t, d) in enumerate(steps):
        x = x0 + i * (sw_each + gap)
        add_step(s, x, Inches(1.7), sw_each, sh_each, n, t, d)
    add_callout(s, Inches(0.8), Inches(4.6), Inches(11.7), Inches(1.6),
                "事業所インパクト：",
                "受講時間の総量に大きな変化はないものの、「事前学習・オンライン受講・職場での実践」がセットになり、勤務シフトと学習時間の管理がより重要になります。")


def slide_elearning():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "オンライン研修活用の留意点", page="6 / 21", section="第2章　新研修体系のポイント")
    add_card(s, Inches(0.8), Inches(1.6), Inches(5.85), Inches(2.4),
             "メリット",
             ["受講機会の地域差の縮小",
              "業務との両立がしやすい",
              "繰り返し視聴で理解定着",
              "移動コスト・時間の削減"])
    add_card(s, Inches(6.7), Inches(1.6), Inches(5.85), Inches(2.4),
             "注意すべき点",
             ["受講時間の確保（勤務時間扱いの整理）",
              "視聴環境（PC・ネット・静音）の準備",
              "個人学習に陥りがちで質問機会が減少",
              "修了要件（テスト・課題）の確認"], accent=True)
    add_text(s, Inches(0.8), Inches(4.2), Inches(11.7), Inches(0.4),
             "事業所として整えるべきこと", size=16, bold=True, color=PRIMARY_LIGHT)
    items = [
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("就業規則上の「研修時間の取扱い」を明確化（労働時間／自己研鑽の区分）",
          {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("受講可能な個室・PC・ヘッドセットの確保", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("視聴後の事業所内シェア会で学びを定着", {"size": 14, "color": TEXT})],
    ]
    add_rich_text(s, Inches(0.9), Inches(4.7), Inches(11.5), Inches(2),
                  items, line_spacing=1.5)


def slide_renewal_current():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "第3章　主任ケアマネ更新研修の現状", page="7 / 21",
              section="第3章　主任ケアマネ更新研修")
    add_text(s, Inches(0.8), Inches(1.55), Inches(11.7), Inches(0.4),
             "制度の概要", size=16, bold=True, color=PRIMARY_LIGHT)
    add_table(s, Inches(0.8), Inches(2.05), Inches(11.7), Inches(2.6),
              ["項目", "内容"],
              [["更新サイクル", "主任介護支援専門員の更新は **5年ごと**"],
               ["研修時間", "国の実施要綱上、合計 **46時間以上**（都道府県により異なる）"],
               ["主な受講要件", "都道府県により異なる。国の要綱では、研修の企画・講師・ファシリテーター経験、法定外研修への年4回以上の参加、学会発表、認定ケアマネジャー、都道府県が適当と認める者等が例示"],
               ["未受講のリスク", "主任介護支援専門員の配置要件に影響し、**特定事業所加算や地域包括支援センターの人員体制**に支障が生じる可能性"]],
              col_widths=[3, 9])
    add_callout(s, Inches(0.8), Inches(4.95), Inches(11.7), Inches(1.7),
                "経営インパクト：",
                "主任ケアマネが失効した場合、事業所の主任配置数・加算区分等によっては、特定事業所加算や地域包括支援センターの人員体制に支障が生じる可能性がある。事業所運営に関わるリスクとして管理が必要。")


def slide_renewal_issues():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "更新研修における4つの課題", page="8 / 21",
              section="第3章　主任ケアマネ更新研修")
    cw = Inches(5.85); ch = Inches(2.3)
    add_card(s, Inches(0.8), Inches(1.6), cw, ch, "① 時間的負担",
             ["46時間以上の受講時間に加え、レポート・事例提出の準備が必要",
              "日常業務と並行しての学習負担が大きい"])
    add_card(s, Inches(6.7), Inches(1.6), cw, ch, "② 受講機会の限定",
             ["開催回数・定員・選考方法は都道府県・実施機関により異なる",
              "受講できない年が続くと更新期限に間に合わないリスク"], accent=True)
    add_card(s, Inches(0.8), Inches(4.1), cw, ch, "③ 業務継続との両立",
             ["研修日のケース対応・モニタリング・担当者会議の調整",
              "事業所全体での体制整備が不可欠"])
    add_card(s, Inches(6.7), Inches(4.1), cw, ch, "④ 受講要件の管理",
             ["「主任ケアマネとしての実務経験」「指導実績」など",
              "5年間にわたって意識的に積み上げる必要"], accent=True)


def slide_renewal_strategy():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "更新研修への戦略的アプローチ", page="9 / 21",
              section="第3章　主任ケアマネ更新研修")
    add_text(s, Inches(0.8), Inches(1.55), Inches(11.7), Inches(0.4),
             "「5年前から逆算」する受講計画", size=16, bold=True, color=PRIMARY_LIGHT)
    add_timeline(s, Inches(0.8), Inches(2.1), Inches(11.7), Inches(3.6),
                 [("更新5年前",   "更新期限・要件をリスト化／業務経験ログの開始"),
                  ("更新3〜2年前", "指導実績の蓄積（実習指導者、事例検討会講師など）／受講要件チェック"),
                  ("更新1年前",   "受講申込／代替要員・シフト計画の確定／事業所内引き継ぎ準備"),
                  ("受講中",     "業務カバー体制の実施／学習時間の確保"),
                  ("修了後",     "事業所内へのフィードバック／次の5年への引き継ぎ")])
    add_callout(s, Inches(0.8), Inches(5.85), Inches(11.7), Inches(0.85),
                "ポイント：",
                "更新管理を「個人任せ」にせず、事業所の人材台帳で見える化することが第一歩。")


def slide_register():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "人材台帳（更新管理シート）の例", page="10 / 21",
              section="第3章　主任ケアマネ更新研修")
    add_table(s, Inches(0.8), Inches(1.55), Inches(11.7), Inches(2.6),
              ["氏名", "資格区分", "有効期限", "次回研修", "受講要件状況", "担当"],
              [["A 主任CM", "主任更新", "R8.3.31", "R7年度 主任更新研修", "○ 指導実績充足", "管理者"],
               ["B CM", "専門II", "R9.3.31", "R8年度 専門研修II", "―", "主任A"],
               ["C CM", "専門I", "R10.3.31", "R7年度 専門研修I", "―", "主任A"],
               ["D 新人CM", "実務研修修了", "R11.3.31", "R8年度 専門研修I", "OJT継続中", "主任A"]],
              col_widths=[2, 2, 2, 3, 3, 1.5])
    add_text(s, Inches(0.8), Inches(4.35), Inches(11.7), Inches(0.4),
             "運用のポイント", size=16, bold=True, color=PRIMARY_LIGHT)
    runs = [
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("年度初め（4月）に全員分を更新し、管理者・主任で共有", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("受講申込時期の3ヶ月前にアラート（カレンダー登録）", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("「受講要件」欄は主任更新で必要な指導実績等を具体的に記載", {"size": 14, "color": TEXT})],
    ]
    add_rich_text(s, Inches(0.9), Inches(4.85), Inches(11.5), Inches(2), runs, line_spacing=1.5)


def slide_chapter4():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "第4章　事業所内での計画的人材育成", page="11 / 21",
              section="第4章　事業所内人材育成")
    add_callout(s, Inches(0.8), Inches(1.55), Inches(11.7), Inches(1.2),
                "基本認識：",
                "法定研修は「最低限の共通基盤」に過ぎない。事業所内のOJT・キャリアパス・組織文化がそろって、はじめて研修効果は実務に定着する。")
    add_text(s, Inches(0.8), Inches(3.0), Inches(11.7), Inches(0.4),
             "人材育成戦略の3本柱", size=16, bold=True, color=PRIMARY_LIGHT)
    sw_each = Inches(3.85); sh_each = Inches(2.8)
    add_step(s, Inches(0.8), Inches(3.55), sw_each, sh_each, "1",
             "キャリアパス設計", "新人→中堅→主任→指導者まで、段階ごとの役割と目標を明確化。")
    add_step(s, Inches(4.75), Inches(3.55), sw_each, sh_each, "2",
             "計画的な研修受講", "法定研修＋事業所内研修＋外部研修を年間計画に組み込む。")
    add_step(s, Inches(8.7), Inches(3.55), sw_each, sh_each, "3",
             "OJTとフォロー", "同行訪問・ケアプラン点検・面談で日常業務に学びを統合。")


def slide_career():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "キャリアパスの設計例（事業所モデルの一例）", page="12 / 21",
              section="第4章　事業所内人材育成")
    add_text(s, Inches(0.8), Inches(1.45), Inches(11.7), Inches(0.35),
             "※下記の経験年数・段階区分は事業所内モデルの例であり、法令上の基準ではありません。",
             size=11, color=MUTED)
    add_table(s, Inches(0.8), Inches(1.85), Inches(11.7), Inches(2.8),
              ["段階", "経験年数（例）", "役割・期待", "受講すべき研修"],
              [["導入期", "0〜1年", "担当ケース20件以下／OJTで基礎習得", "実務研修・新任者研修"],
               ["基礎期", "2〜4年", "標準ケースを自立して担当", "専門研修課程I／法定外スキル研修"],
               ["応用期", "5〜7年", "困難事例・若手OJT支援", "専門研修課程II／事例検討会講師"],
               ["主任期", "8年〜", "地域連携・事業所運営・指導", "主任介護支援専門員研修"],
               ["指導者期", "主任後", "法定研修講師／地域ケア会議", "主任更新研修／指導者養成"]],
              col_widths=[1.5, 1.5, 4, 4])
    add_text(s, Inches(0.8), Inches(4.75), Inches(11.7), Inches(0.4),
             "各段階で必須の仕組み", size=16, bold=True, color=PRIMARY_LIGHT)
    runs = [
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("目標設定面談（年2回）と処遇との連動", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("段階移行の判定基準を文書化（属人化させない）", {"size": 14, "color": TEXT})],
    ]
    add_rich_text(s, Inches(0.9), Inches(5.25), Inches(11.5), Inches(1.5), runs, line_spacing=1.5)


def slide_annual_plan():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "年間研修計画の立て方", page="13 / 21",
              section="第4章　事業所内人材育成")
    # 左
    add_text(s, Inches(0.8), Inches(1.55), Inches(5.85), Inches(0.4),
             "3層構造で計画する", size=16, bold=True, color=PRIMARY_LIGHT)
    add_card(s, Inches(0.8), Inches(2.05), Inches(5.85), Inches(1.4),
             "① 法定研修（必須）",
             ["都道府県告知をもとに4月時点で受講者・時期を確定"])
    add_card(s, Inches(0.8), Inches(3.5), Inches(5.85), Inches(1.4),
             "② 事業所内研修（月例）",
             ["事例検討会・制度学習会・適切なケアマネジメント手法の勉強会"], accent=True)
    add_card(s, Inches(0.8), Inches(4.95), Inches(5.85), Inches(1.4),
             "③ 外部研修・自己研鑽",
             ["協会研修・学会・専門領域（医療・認知症・看取り等）"])
    # 右
    add_text(s, Inches(6.85), Inches(1.55), Inches(5.85), Inches(0.4),
             "個人別研修計画書（フォーマット）", size=16, bold=True, color=PRIMARY_LIGHT)
    runs = [
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("本人の到達目標（1年・3年）", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("受講予定研修（法定／法定外）", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("OJT項目とメンター", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("振り返り（半期ごと）", {"size": 14, "color": TEXT})],
    ]
    add_rich_text(s, Inches(6.95), Inches(2.1), Inches(5.7), Inches(2.5), runs, line_spacing=1.5)
    add_callout(s, Inches(6.85), Inches(4.7), Inches(5.85), Inches(1.5),
                "運用のコツ：",
                "4月作成 → 9月中間レビュー → 翌年3月評価、のPDCAを回す。")


def slide_attendance():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "受講機会を確保する職場体制", page="14 / 21",
              section="第4章　事業所内人材育成")
    cw = Inches(5.85); ch = Inches(2.4)
    add_card(s, Inches(0.8), Inches(1.6), cw, ch, "シフト・業務面",
             ["研修日の担当ケース引継ぎルール整備",
              "主任ケアマネ間の相互バックアップ",
              "モニタリング訪問の前倒し調整"])
    add_card(s, Inches(6.7), Inches(1.6), cw, ch, "制度・処遇面",
             ["研修受講を業務扱いとする規程",
              "受講料・交通費の事業所負担",
              "主任研修修了者への資格手当・役割手当等（事業所内処遇）"], accent=True)
    add_card(s, Inches(0.8), Inches(4.15), cw, ch, "環境・ICT面",
             ["eラーニング受講用の静音スペース",
              "ケアプランソフトでの業務効率化",
              "記録のテンプレート化で記入時間短縮"])
    add_card(s, Inches(6.7), Inches(4.15), cw, ch, "コミュニケーション面",
             ["受講予定の事前共有・カレンダー化",
              "利用者・家族への事前説明",
              "受講後のシェア会で全員に還元"], accent=True)


def slide_ojt():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "OJTの体系化　— 4つの基本ツール", page="15 / 21",
              section="第4章　事業所内人材育成")
    sw_each = Inches(2.85); sh_each = Inches(2.5)
    x0 = Inches(0.8); gap = Inches(0.1)
    items = [
        ("1", "同行訪問", "初回・困難事例で主任が同行。終了後に必ず振り返り。"),
        ("2", "ケアプラン点検", "毎月、主任が新人プランをレビュー。点検視点を標準化。"),
        ("3", "事例検討会", "月1回、全員参加。発表者持ち回りで指導力も育成。"),
        ("4", "1on1面談", "月1回30分。業務状況・キャリア・メンタルまで対話。"),
    ]
    for i, (n, t, d) in enumerate(items):
        add_step(s, x0 + i * (sw_each + gap), Inches(1.65), sw_each, sh_each, n, t, d)
    add_callout(s, Inches(0.8), Inches(4.4), Inches(11.7), Inches(1.2),
                "OJTを「指導者の善意」に任せない：",
                "記録様式・チェックリスト・点検視点を事業所共通のフォーマットとして整備し、誰が指導しても一定の質を担保する。")
    add_text(s, Inches(0.8), Inches(5.75), Inches(11.7), Inches(0.4),
             "主任ケアマネにとってのメリット", size=15, bold=True, color=PRIMARY_LIGHT)
    runs = [
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("OJT・事例検討・研修企画等の実績は、都道府県の受講要件や実績確認に関連する場合がある", {"size": 13, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("指導を通じた自身の振り返りで更新研修の事例提出が容易に", {"size": 13, "color": TEXT})],
    ]
    add_rich_text(s, Inches(0.9), Inches(6.2), Inches(11.5), Inches(0.8), runs, line_spacing=1.4)


def slide_retention():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "研修効果を実務に定着させる工夫", page="16 / 21",
              section="第4章　事業所内人材育成")
    add_text(s, Inches(0.8), Inches(1.55), Inches(11.7), Inches(0.4),
             "「受けっぱなし」を防ぐ3つの仕掛け", size=16, bold=True, color=PRIMARY_LIGHT)
    cw = Inches(5.85); ch = Inches(2.0)
    add_card(s, Inches(0.8), Inches(2.05), cw, ch, "① 事前学習",
             ["研修テキストに目を通し、持ち帰りたい問いを1つ決めてから参加"])
    add_card(s, Inches(6.7), Inches(2.05), cw, ch, "② 事業所内シェア会",
             ["研修翌週に30分。「明日から変えること」を1つ宣言してもらう"], accent=True)
    add_card(s, Inches(0.8), Inches(4.2), cw, ch, "③ 3ヶ月後フォロー",
             ["「変えると宣言したこと」が継続しているか、1on1で確認"])
    add_card(s, Inches(6.7), Inches(4.2), cw, ch, "＋α ポートフォリオ",
             ["受講記録・事例・指導実績を蓄積。主任更新時の提出資料が一気に楽に"], accent=True)


def slide_external():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "事業所の外に学びの場をつくる", page="17 / 21",
              section="第4章　事業所内人材育成")
    # 左
    add_text(s, Inches(0.8), Inches(1.55), Inches(5.85), Inches(0.4),
             "地域・多職種との連携", size=16, bold=True, color=PRIMARY_LIGHT)
    runs1 = [
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("地域ケア会議への積極参加（事例提供・コメント）", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("地域包括支援センター主催の研修・連絡会", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("医療機関の退院前カンファレンスを学びの機会に", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("近隣事業所との合同事例検討会", {"size": 14, "color": TEXT})],
    ]
    add_rich_text(s, Inches(0.9), Inches(2.1), Inches(5.7), Inches(3), runs1, line_spacing=1.6)
    # 右
    add_text(s, Inches(6.85), Inches(1.55), Inches(5.85), Inches(0.4),
             "外部資源の活用", size=16, bold=True, color=PRIMARY_LIGHT)
    runs2 = [
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("都道府県・市町村介護支援専門員協会の研修", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("日本ケアマネジメント学会など学術団体", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("「適切なケアマネジメント手法」関連の研修", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("主任更新研修の講師依頼を受ける側になる", {"size": 14, "color": TEXT})],
    ]
    add_rich_text(s, Inches(6.95), Inches(2.1), Inches(5.7), Inches(3), runs2, line_spacing=1.6)
    add_callout(s, Inches(0.8), Inches(5.3), Inches(11.7), Inches(1.4),
                "外部活動は更新の追い風：",
                "地域ケア会議の事例提供や講師経験、法定外研修への参加等は、都道府県によっては主任更新研修の受講要件・実績確認に関連する場合がある。所管自治体の要綱を確認のこと。")


def slide_case():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "第5章　モデルケース　A居宅介護支援事業所（架空事例）", page="18 / 21",
              section="第5章　モデルケース")
    add_text(s, Inches(0.8), Inches(1.45), Inches(11.7), Inches(0.35),
             "※本事例は、本資料の解説を目的としたモデルケースであり、特定の事業所の実績ではありません。",
             size=11, color=MUTED)
    add_card(s, Inches(0.8), Inches(1.85), Inches(5.85), Inches(1.7),
             "事業所プロフィール（モデル）",
             ["常勤ケアマネ6名（うち主任3名）",
              "担当件数 約180件",
              "特定事業所加算（I）取得"])
    add_card(s, Inches(6.7), Inches(1.85), Inches(5.85), Inches(1.7),
             "取り組み前の課題（モデル）",
             ["主任の更新時期が重なり業務麻痺",
              "研修参加が個人任せで非効率",
              "新人OJTが属人化"], accent=True)
    add_text(s, Inches(0.8), Inches(3.75), Inches(11.7), Inches(0.4),
             "実施した3つの施策", size=16, bold=True, color=PRIMARY_LIGHT)
    runs = [
        [("1.  ", {"size": 14, "color": ACCENT, "bold": True}),
         ("人材台帳の整備と5年スパンの更新ローテーション", {"size": 14, "color": TEXT})],
        [("2.  ", {"size": 14, "color": ACCENT, "bold": True}),
         ("月例事例検討会の定例化と発表者ローテ", {"size": 14, "color": TEXT})],
        [("3.  ", {"size": 14, "color": ACCENT, "bold": True}),
         ("個人別研修計画書と半期面談の制度化", {"size": 14, "color": TEXT})],
    ]
    add_rich_text(s, Inches(0.9), Inches(4.25), Inches(11.5), Inches(1.5), runs, line_spacing=1.6)
    add_callout(s, Inches(0.8), Inches(5.85), Inches(11.7), Inches(0.95),
                "想定される効果（モデル）：",
                "受講機会の確実な確保／主任更新の取りこぼし防止／加算要件の安定的維持／OJT標準化による定着率改善")


def slide_checklist():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "明日から始めるチェックリスト", page="19 / 21",
              section="第5章　モデルケース")
    cw = Inches(5.85); ch = Inches(2.4)
    add_card(s, Inches(0.8), Inches(1.6), cw, ch, "□ 制度理解",
             ["新ガイドラインの確認",
              "都道府県の施行スケジュール把握",
              "eラーニング受講方法の確認"])
    add_card(s, Inches(6.7), Inches(1.6), cw, ch, "□ 更新管理",
             ["人材台帳の作成・更新",
              "主任の更新期限のアラート設定",
              "受講要件の充足状況の点検"], accent=True)
    add_card(s, Inches(0.8), Inches(4.15), cw, ch, "□ 育成体制",
             ["キャリアパス図の作成",
              "個人別研修計画書の運用",
              "OJTフォーマットの整備"])
    add_card(s, Inches(6.7), Inches(4.15), cw, ch, "□ 学習文化",
             ["月例事例検討会の定例化",
              "受講後シェア会の実施",
              "ポートフォリオの作成支援"], accent=True)


def slide_references():
    s = prs.slides.add_slide(blank)
    add_bg(s)
    add_title(s, "本資料の根拠と利用上の注意", page="20 / 21", section="出典・注記")
    add_text(s, Inches(0.8), Inches(1.55), Inches(11.7), Inches(0.4),
             "主な根拠資料", size=16, bold=True, color=PRIMARY_LIGHT)
    runs1 = [
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("厚生労働省「介護支援専門員資質向上事業ガイドライン」（令和5年4月版、令和6年4月施行）",
          {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("厚生労働省「主任介護支援専門員更新研修実施要綱」", {"size": 14, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("厚生労働省「介護支援専門員に係るオンライン研修の手引き」等",
          {"size": 14, "color": TEXT})],
    ]
    add_rich_text(s, Inches(0.9), Inches(2.05), Inches(11.6), Inches(1.8),
                  runs1, line_spacing=1.6)
    add_callout(s, Inches(0.8), Inches(3.95), Inches(11.7), Inches(1.5),
                "利用上の注意：",
                "本資料は、上記ガイドラインおよび更新研修実施要綱を踏まえた実務整理である。実際の受講要件・募集時期・実施方法・オンライン研修の取扱いは、各都道府県および研修実施機関の最新通知を必ず確認すること。")
    add_text(s, Inches(0.8), Inches(5.6), Inches(11.7), Inches(0.4),
             "本資料の位置づけ", size=16, bold=True, color=PRIMARY_LIGHT)
    runs2 = [
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("事業所内研修・管理者向け勉強会の素材として作成", {"size": 13, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("第5章のモデルケースは解説目的の架空事例", {"size": 13, "color": TEXT})],
        [("■  ", {"size": 12, "color": ACCENT, "bold": True}),
         ("キャリアパス・経験年数区分は事業所モデルの例であり法令上の基準ではない",
          {"size": 13, "color": TEXT})],
    ]
    add_rich_text(s, Inches(0.9), Inches(6.05), Inches(11.6), Inches(0.9),
                  runs2, line_spacing=1.4)


def slide_closing():
    s = prs.slides.add_slide(blank)
    add_rect(s, 0, 0, SW, SH, fill=PRIMARY)
    add_rect(s, 0, SH - Inches(0.4), SW, Inches(0.4), fill=ACCENT)
    add_text(s, Inches(1), Inches(1.4), Inches(11.3), Inches(1.2),
             "本日のまとめ", size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    runs = [
        [("法定研修改正は ", {"size": 22, "color": WHITE}),
         ("「研修体系の再設計」 ", {"size": 22, "color": ACCENT, "bold": True}),
         ("の好機。", {"size": 22, "color": WHITE})],
        [("主任ケアマネ更新は ", {"size": 22, "color": WHITE}),
         ("「5年前から」 ", {"size": 22, "color": ACCENT, "bold": True}),
         ("の戦略で乗り切る。", {"size": 22, "color": WHITE})],
        [("人材育成は ", {"size": 22, "color": WHITE}),
         ("「個人任せ」から「組織の仕組み」 ", {"size": 22, "color": ACCENT, "bold": True}),
         ("へ。", {"size": 22, "color": WHITE})],
    ]
    add_rich_text(s, Inches(0.8), Inches(2.9), Inches(11.7), Inches(2.8),
                  runs, align=PP_ALIGN.CENTER, line_spacing=1.7)
    add_text(s, Inches(0.8), Inches(6.0), Inches(11.7), Inches(0.6),
             "ご清聴ありがとうございました", size=18, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Inches(11.5), Inches(6.85), Inches(1.5), Inches(0.4),
             "21 / 21", size=11, color=WHITE, align=PP_ALIGN.RIGHT)


# ====== 実行 ======
slide_title()
slide_agenda()
slide_chapter1_bg()
slide_schedule()
slide_main_points()
slide_elearning()
slide_renewal_current()
slide_renewal_issues()
slide_renewal_strategy()
slide_register()
slide_chapter4()
slide_career()
slide_annual_plan()
slide_attendance()
slide_ojt()
slide_retention()
slide_external()
slide_case()
slide_checklist()
slide_references()
slide_closing()

out = "/home/user/kaigo-ga-app/slides/法定研修制度改正への対応と主任ケアマネ更新研修の戦略.pptx"
prs.save(out)
print(f"OK: {out}")
