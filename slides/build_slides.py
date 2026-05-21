# -*- coding: utf-8 -*-
"""令和8年度 第3回 介護支援専門員研修会 予習スライド生成スクリプト。

成果物: 令和8年度第3回ケアマネ研修_予習スライド_成年後見制度.pptx (21枚)
内容は2026年5月時点の一次情報調査に基づく。要確認箇所は本文に明示。
"""
import os
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from design import (
    new_deck, _blank, page_bg, header, footer_note, add_text, add_rect,
    chip, memo_panel,
    NAVY, NAVY_DARK, TEAL, TEAL_LIGHT, GOLD, INK, GRAY, LIGHT, LINE, WHITE,
    MEMO_BG, MEMO_LINE, SLIDE_W, SLIDE_H,
)

IN = Inches


def card(slide, left, top, w, h, title, accent, body_lines,
         title_size=14, body_size=12):
    """見出しバー付きカード。body_lines: [(text,bold)] のリスト。"""
    add_rect(slide, left, top, w, h, WHITE, line=LINE, line_w=1.0, radius=0.04)
    add_rect(slide, left, top, w, IN(0.46), accent, radius=0.04)
    add_rect(slide, left, top + IN(0.30), w, IN(0.16), accent)
    add_text(slide, left + IN(0.14), top, w - IN(0.24), IN(0.46),
             [(title, title_size, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
    paras = []
    for seg in body_lines:
        txt = seg[0]
        bold = seg[1] if len(seg) > 1 else False
        col = seg[2] if len(seg) > 2 else INK
        paras.append([("• " if not bold else "", body_size, accent, True),
                      (txt, body_size, col, bold)])
    add_text(slide, left + IN(0.16), top + IN(0.58), w - IN(0.30),
             h - IN(0.70), paras, line_spacing=1.2, space_after=4)


def labelrow(slide, left, top, w, label, value, lw=IN(2.3),
             vsize=13, accent=TEAL):
    add_rect(slide, left, top, lw, IN(0.40), accent)
    add_text(slide, left + IN(0.10), top, lw - IN(0.18), IN(0.40),
             [(label, 12, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, left + lw, top, w - lw, IN(0.40), LIGHT, line=LINE,
             line_w=0.75)
    add_text(slide, left + lw + IN(0.12), top, w - lw - IN(0.22), IN(0.40),
             [(value, vsize, INK, False)], anchor=MSO_ANCHOR.MIDDLE)


# ===================================================================
# Slide 01 — 表紙
# ===================================================================
def s01(prs):
    s = _blank(prs)
    page_bg(s, NAVY)
    add_rect(s, 0, IN(4.92), SLIDE_W, IN(2.58), WHITE)
    add_rect(s, 0, IN(4.92), SLIDE_W, IN(0.06), GOLD)
    # 上部
    add_rect(s, IN(0.9), IN(0.78), IN(0.12), IN(0.62), GOLD)
    add_text(s, IN(1.15), IN(0.74), IN(10), IN(0.4),
             [("令和8年度 第3回 介護支援専門員研修会", 15, TEAL_LIGHT, True)])
    add_text(s, IN(1.15), IN(1.06), IN(7.5), IN(0.4),
             [("主催：日本橋おとしより相談センター", 12, RGBC(0xB8, 0xC6, 0xD8))])
    # タイトル
    add_text(s, IN(0.9), IN(1.95), IN(11.6), IN(1.7),
             [[("高齢者を取り巻く", 40, WHITE, True)],
              [("様々な問題への支援", 40, WHITE, True)]],
             line_spacing=1.1)
    add_text(s, IN(0.95), IN(3.5), IN(11.4), IN(0.5),
             [("〜 介護支援専門員と法律専門職との協力関係の確立を目指して 〜",
               16, GOLD, True)])
    # 予習バッジ
    add_rect(s, IN(0.95), IN(4.18), IN(3.5), IN(0.5), GOLD, radius=0.5)
    add_text(s, IN(0.95), IN(4.18), IN(3.5), IN(0.5),
             [("予 習 資 料（受講者個人用）", 13, NAVY_DARK, True)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 下部情報帯
    info = [
        ("日　時", "令和8年6月25日(木) 14:00〜16:00"),
        ("会　場", "人形町区民館"),
        ("講　師", "司法書士 大谷雅彦 氏"),
    ]
    x = IN(0.95)
    for lab, val in info:
        add_text(s, x, IN(5.25), IN(3.8), IN(0.3),
                 [(lab, 11, TEAL, True)])
        add_text(s, x, IN(5.52), IN(3.9), IN(0.5),
                 [(val, 14, INK, True)])
        x += IN(3.95)
    add_text(s, IN(0.95), IN(6.18), IN(11.4), IN(0.35),
             [("成年後見センター・リーガルサポート東京支部 中央地区リーダー",
               11, GRAY)])
    add_rect(s, IN(0.95), IN(6.62), IN(11.43), Pt(1), LINE)
    add_text(s, IN(0.95), IN(6.74), IN(11.4), IN(0.4),
             [("対象：中央区で活動する介護支援専門員（定員30名）　／　"
               "作成：2026年5月　主任介護支援専門員（予習用に独自作成）",
               10, GRAY)])


def RGBC(r, g, b):
    from pptx.dml.color import RGBColor
    return RGBColor(r, g, b)


# ===================================================================
# Slide 02 — 本資料の位置づけ
# ===================================================================
def s02(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "はじめに", "本資料の位置づけ 〜 これは「予習メモ」です", 2)
    add_rect(s, IN(0.8), top, IN(7.5), IN(1.0), TEAL_LIGHT, radius=0.05)
    add_text(s, IN(1.0), top + IN(0.10), IN(7.1), IN(0.8),
             [[("本資料は研修本番で使うものではありません。", 13, NAVY, True)],
              [("受講者（主任ケアマネ）が当日に備えて作成した個人予習資料です。",
                12, INK, False)]], line_spacing=1.2)

    add_text(s, IN(0.8), top + IN(1.2), IN(7.6), IN(0.35),
             [("◆ 3つの目的", 14, TEAL, True)])
    purposes = [
        ("① 先回りして整理", "講師（司法書士）が話すであろう論点を、"
         "自分の頭で事前に組み立てておく。"),
        ("② 当日に突き合わせ", "講師持参スライドと本資料を比較し、"
         "理解の差・抜け漏れをその場で埋める。"),
        ("③ 疑問を質問化", "予習で残った疑問を質問リスト化し（P19-20）、"
         "当日に確実に解消する。"),
    ]
    y = top + IN(1.58)
    for h_, b_ in purposes:
        add_rect(s, IN(0.8), y, IN(7.5), IN(0.86), WHITE, line=LINE, line_w=1)
        add_rect(s, IN(0.8), y, IN(0.10), IN(0.86), TEAL)
        add_text(s, IN(1.05), y + IN(0.07), IN(7.1), IN(0.35),
                 [(h_, 13, NAVY, True)])
        add_text(s, IN(1.05), y + IN(0.38), IN(7.1), IN(0.45),
                 [(b_, 11.5, INK, False)], line_spacing=1.15)
        y += IN(0.92)

    add_rect(s, IN(0.8), y + IN(0.05), IN(7.5), IN(0.72), RGBC(0xFC,0xF3,0xDD),
             line=GOLD, line_w=1)
    add_text(s, IN(1.0), y + IN(0.12), IN(7.1), IN(0.6),
             [[("⚠ 注意：", 11.5, GOLD, True),
               ("内容は2026年5月時点の調査ベース。「要確認」表示は一次情報の"
                "最終確認が必要な箇所。", 11, INK, False)],
              [("　　　　 制度の正確な説明は研修当日の講師説明を正とすること。",
                11, INK, False)]], line_spacing=1.2)

    memo_panel(s, IN(8.55), top, IN(3.98), IN(4.95), "気づき・予習中の疑問")


# ===================================================================
# Slide 03 — 研修会概要
# ===================================================================
def s03(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "研修会概要", "令和8年度 第3回 介護支援専門員研修会", 3)
    rows = [
        ("研 修 名", "令和8年度 第3回 介護支援専門員研修会"),
        ("主　　催", "日本橋おとしより相談センター"),
        ("日　　時", "令和8年6月25日(木) 14:00〜16:00"),
        ("会　　場", "人形町区民館"),
        ("対　　象", "中央区で活動する介護支援専門員（定員30名／主任ケアマネ含む実務者）"),
        ("テ ー マ", "高齢者を取り巻く様々な問題への支援"),
        ("講　　師", "司法書士 大谷雅彦 氏（リーガルサポート東京支部 中央地区リーダー）"),
    ]
    y = top
    for lab, val in rows:
        labelrow(s, IN(0.8), y, IN(8.0), lab, val, lw=IN(1.7), vsize=12)
        y += IN(0.50)

    add_rect(s, IN(0.8), y + IN(0.08), IN(8.0), IN(1.25),
             TEAL_LIGHT, radius=0.05)
    add_text(s, IN(1.0), y + IN(0.16), IN(7.6), IN(0.35),
             [("◆ 想定される講師の話題（予習の的を絞る）", 13, NAVY, True)])
    add_text(s, IN(1.0), y + IN(0.50), IN(7.6), IN(0.7),
             [[("・ 成年後見制度（法定後見・任意後見、2026年改正の動向）",
                12, INK, False)],
              [("・ 身寄りのない認知症高齢者への支援（身元保証・医療同意・"
                "死後事務・連携）", 12, INK, False)]], line_spacing=1.25)

    memo_panel(s, IN(9.05), top, IN(3.48), IN(5.0), "当日メモ")
    footer_note(s, "※ サブタイトル：介護支援専門員と法律専門職との協力関係の確立を目指して")


# ===================================================================
# Slide 04 — 典型事例3つ（導入）
# ===================================================================
def s04(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "導入", "ケアマネが直面する典型事例 ── あなたなら、どうつなぐ？", 4)
    add_text(s, IN(0.8), top, IN(11.7), IN(0.4),
             [("下記はいずれも「法律専門職との連携」が鍵になる場面。"
               "本資料は、この3事例に答えるための予習です。", 12, GRAY, False)])
    cw, gap = IN(3.78), IN(0.2)
    x = IN(0.8)
    cases = [
        ("事例A 金銭管理の限界", TEAL,
         [("独居・認知症が進行。預貯金を管理できず公共料金が滞納。",),
          ("親族は遠方の甥のみ。本人は通帳の場所も曖昧。",),
          ("→ 日常生活自立支援事業？ 補助・保佐・後見？",
           False, NAVY)]),
        ("事例B 入院と身元保証", GOLD,
         [("身寄りなし高齢者が緊急入院。病院は「身元保証人を」と要求。",),
          ("手術の同意・入院費の支払いは誰が担うのか。",),
          ("→ 身元保証・医療同意の「主体」を整理する。",
           False, NAVY)]),
        ("事例C 単身死亡後の事務", NAVY,
         [("利用者が単身で死亡。火葬・残置物・年金停止が宙に浮く。",),
          ("生前に死後事務の備えがなかった。",),
          ("→ 死後事務は誰が、どこまでできるのか。",
           False, NAVY)]),
    ]
    for title, ac, body in cases:
        card(s, x, top + IN(0.5), cw, IN(3.15), title, ac, body,
             title_size=13, body_size=11.5)
        x += cw + gap

    add_rect(s, IN(0.8), top + IN(3.85), IN(11.73), IN(0.62),
             RGBC(0xFC,0xF3,0xDD), line=GOLD, line_w=1)
    add_text(s, IN(1.0), top + IN(3.93), IN(11.4), IN(0.5),
             [("共通する問い：ケアマネは「保証人代わり」になるのではなく、"
               "“機能ごと”に適切な制度・専門職へどうつなぐか。",
               12.5, INK, True)], anchor=MSO_ANCHOR.MIDDLE)
    footer_note(s, "当日メモ：自分の担当ケースで似た事例は？　→")


# ===================================================================
# Slide 05 — 成年後見制度の全体像
# ===================================================================
def s05(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "制度の基礎", "成年後見制度の全体像", 5)
    add_text(s, IN(0.8), top, IN(8.0), IN(0.7),
             [("判断能力が不十分な人の "
               "財産管理 と 身上保護 を、法的な権限をもって支える制度。"
               "本人の意思決定支援と権利擁護が理念。", 13, INK, False)],
             line_spacing=1.25)
    # 2本柱
    y = top + IN(0.85)
    add_text(s, IN(0.8), y, IN(8), IN(0.35), [("◆ 制度は2本柱", 13, TEAL, True)])
    card(s, IN(0.8), y + IN(0.4), IN(3.9), IN(1.85), "法定後見", NAVY,
         [("判断能力が低下した後に家庭裁判所が選任",),
          ("後見・保佐・補助 の3類型",),
          ("取消権あり",)], body_size=11.5)
    card(s, IN(4.9), y + IN(0.4), IN(3.9), IN(1.85), "任意後見", TEAL,
         [("判断能力があるうちに本人が契約（公正証書）",),
          ("後見人を自分で選べる",),
          ("効力発生は監督人選任後／取消権なし",)], body_size=11.5)
    # 数字パネル
    yy = y + IN(2.45)
    add_rect(s, IN(0.8), yy, IN(8.0), IN(1.35), TEAL_LIGHT, radius=0.05)
    add_text(s, IN(1.0), yy + IN(0.08), IN(7.6), IN(0.3),
             [("◆ 利用者数（令和6年12月末・全国）", 12.5, NAVY, True)])
    stats = [("約25.4万人", "総数"), ("17.9万", "後見"),
             ("5.5万", "保佐"), ("1.7万", "補助"), ("0.3万", "任意後見")]
    sx = IN(1.0)
    for num, lab in stats:
        add_text(s, sx, yy + IN(0.40), IN(1.55), IN(0.45),
                 [(num, 17, TEAL, True)])
        add_text(s, sx, yy + IN(0.85), IN(1.55), IN(0.3),
                 [(lab, 10.5, GRAY, False)])
        sx += IN(1.5)

    memo_panel(s, IN(9.05), top, IN(3.48), IN(5.0), "当日メモ")
    footer_note(s, "出典：最高裁「成年後見関係事件の概況―令和6年1〜12月―」"
                   "（数値は出典PDFで要確認）")


# ===================================================================
# Slide 06 — 利用促進基本計画（第二期）
# ===================================================================
def s06(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "制度の最新像 ①", "成年後見制度利用促進基本計画（第二期）の進捗", 6)
    add_rect(s, IN(0.8), top, IN(8.0), IN(0.78), NAVY, radius=0.05)
    add_text(s, IN(1.0), top + IN(0.06), IN(7.6), IN(0.7),
             [[("第二期計画：令和4年3月閣議決定 ／ 期間 令和4〜8年度",
                12.5, WHITE, True)],
              [("＝ 令和8年度（今年度）が計画の最終年度。次期計画の議論も視野。",
                11, TEAL_LIGHT, False)]], line_spacing=1.2)
    items = [
        ("理念", "「尊厳のある本人らしい生活の継続」「地域社会への参加」を図る"
         "権利擁護支援の推進。"),
        ("地域連携ネットワーク", "中核機関を整備し、本人を支えるチーム・"
         "協議会づくりを進める。"),
        ("中間検証", "計画中間年に検証を実施。中間検証報告書を2025年3月に"
         "とりまとめ。"),
        ("市町村計画", "策定率は令和4年4月時点62.8%。KPIは全1,741市町村"
         "（最新の策定率は要確認）。"),
        ("身寄りなし支援", "令和6年度〜、身寄りのない高齢者等への相談・調整"
         "窓口の総合支援パッケージを試行。"),
    ]
    y = top + IN(0.95)
    for h_, b_ in items:
        add_rect(s, IN(0.8), y, IN(2.1), IN(0.66), TEAL)
        add_text(s, IN(0.9), y, IN(1.95), IN(0.66),
                 [(h_, 11.5, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
        add_rect(s, IN(2.9), y, IN(5.9), IN(0.66), LIGHT, line=LINE, line_w=0.75)
        add_text(s, IN(3.04), y + IN(0.04), IN(5.65), IN(0.6),
                 [(b_, 11, INK, False)], anchor=MSO_ANCHOR.MIDDLE,
                 line_spacing=1.12)
        y += IN(0.74)

    memo_panel(s, IN(9.05), top, IN(3.48), IN(5.0), "当日メモ")
    footer_note(s, "出典：厚労省 成年後見制度利用促進／第二期基本計画。"
                   "策定率の最新値は要確認。")


# ===================================================================
# Slide 07 — 法定後見 3類型の比較表
# ===================================================================
def s07(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "制度の最新像 ②", "法定後見 3類型の比較 ── 後見・保佐・補助", 7)
    x0 = IN(0.8)
    cols = [IN(2.05), IN(3.22), IN(3.22), IN(3.24)]
    heads = ["", "後見", "保佐", "補助"]
    hcolors = [NAVY, NAVY, TEAL, RGBC(0x4A,0x9A,0x6A)]
    # header row
    x = x0
    for i, htxt in enumerate(heads):
        add_rect(s, x, top, cols[i], IN(0.46), hcolors[i])
        if htxt:
            add_text(s, x, top, cols[i], IN(0.46), [(htxt, 14, WHITE, True)],
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        x += cols[i]
    rows = [
        ("対象者", "判断能力を欠くのが\n通常の状態",
         "判断能力が\n著しく不十分", "判断能力が\n不十分"),
        ("代理権", "財産に関する\n包括的な代理権",
         "家裁が定める特定の行為\n（本人の同意が必要）",
         "家裁が定める特定の行為\n（本人の同意が必要）"),
        ("同意権・取消権", "日常生活の行為以外を\n広く取消し可",
         "民法13条1項所定の\n重要な行為", "申立て範囲で家裁が\n定めた特定行為のみ"),
        ("医師の鑑定", "原則 必要", "原則 必要", "原則 不要（診断書等で可）"),
    ]
    y = top + IN(0.46)
    rh = IN(0.92)
    for ri, (lab, c1, c2, c3) in enumerate(rows):
        bg = WHITE if ri % 2 == 0 else LIGHT
        x = x0
        add_rect(s, x, y, cols[0], rh, NAVY_DARK, line=WHITE, line_w=1)
        add_text(s, x + IN(0.06), y, cols[0] - IN(0.1), rh,
                 [(lab, 11.5, WHITE, True)], align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        x += cols[0]
        for ci, ctxt in enumerate([c1, c2, c3]):
            add_rect(s, x, y, cols[ci + 1], rh, bg, line=LINE, line_w=0.75)
            paras = [[(ln, 11, INK, False)] for ln in ctxt.split("\n")]
            add_text(s, x + IN(0.08), y, cols[ci + 1] - IN(0.16), rh, paras,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                     line_spacing=1.1)
            x += cols[ci + 1]
        y += rh

    add_rect(s, x0, y + IN(0.10), IN(11.73), IN(0.56), TEAL_LIGHT, radius=0.06)
    add_text(s, x0 + IN(0.15), y + IN(0.15), IN(11.4), IN(0.5),
             [("申立てができる人：本人・配偶者・四親等内の親族・"
               "市町村長・検察官 等（類型により任意後見受任者等も）",
               11.5, NAVY, True)], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x0, y + IN(0.78), IN(11.7), IN(0.3),
             [("予習ポイント：類型は「判断能力の程度」で決まる。"
               "2026年改正で将来は『補助』へ一元化の方向（P11参照）。",
               11, GOLD, True)])


# ===================================================================
# Slide 08 — 申立ての流れ・費用・期間
# ===================================================================
def s08(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "制度の最新像 ③", "申立ての流れ・費用・期間", 8)
    # 流れ
    add_text(s, IN(0.8), top, IN(11.7), IN(0.3),
             [("◆ 申立ての流れ（申立先＝本人住所地の家庭裁判所）", 13, TEAL, True)])
    steps = ["診断書取得\n申立て準備", "家裁へ\n申立て", "調査・鑑定\n（必要時）",
             "審　判", "確定・登記\n後見等開始"]
    sw = IN(2.16)
    x = IN(0.8)
    for i, st in enumerate(steps):
        ac = NAVY if i < 4 else TEAL
        add_rect(s, x, top + IN(0.4), sw, IN(0.78), ac, radius=0.08)
        paras = [[(ln, 11, WHITE, True)] for ln in st.split("\n")]
        add_text(s, x, top + IN(0.4), sw, IN(0.78), paras,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 line_spacing=1.05)
        if i < 4:
            add_text(s, x + sw - IN(0.04), top + IN(0.4), IN(0.3), IN(0.78),
                     [("▶", 14, GOLD, True)], align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)
        x += sw + IN(0.20)

    y = top + IN(1.45)
    card(s, IN(0.8), y, IN(5.75), IN(2.5), "費用の目安", TEAL,
         [("申立手数料：収入印紙 800円",),
          ("（保佐・補助の権限付与は申立てごとに各800円追加）",),
          ("登記手数料：収入印紙 2,600円",),
          ("連絡用の郵便切手・診断書作成費 ほか",),
          ("鑑定費用：実施は全体の約4%。実施時で多くは10万円以下",)],
         body_size=11.5)
    card(s, IN(6.78), y, IN(5.75), IN(2.5), "期間・件数", NAVY,
         [("審理期間：4か月以内に約94%が終局（2か月以内 約72%）",),
          ("（鑑定実施時はさらに期間を要する）",),
          ("申立件数：令和6年 約4.2万件（前年比 約+2%）",),
          ("内訳：後見開始 約2.9万 / 保佐 約0.9万 / 補助 約0.3万 ほか",),
          ("費用負担が困難な低所得者には利用支援事業の助成あり",)],
         body_size=11.5)

    add_rect(s, IN(0.8), y + IN(2.65), IN(11.73), IN(0.5), RGBC(0xFC,0xF3,0xDD),
             line=GOLD, line_w=1)
    add_text(s, IN(1.0), y + IN(2.65), IN(11.4), IN(0.5),
             [("当日確認したい：中央区の家裁で申立て→開始までの“実際の”所要期間。",
               11.5, INK, True)], anchor=MSO_ANCHOR.MIDDLE)
    footer_note(s, "出典：最高裁・法務省・厚労省。件数・費用は出典資料で要確認。")


# ===================================================================
# Slide 09 — 選任後の実務
# ===================================================================
def s09(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "制度の最新像 ④", "選任後の実務 ── 後見人が「できること／できないこと」", 9)
    card(s, IN(0.8), top, IN(3.85), IN(3.35), "できること", TEAL,
         [("財産の管理・保全",),
          ("契約の代理（施設入所契約 等）",),
          ("預貯金など収支の管理",),
          ("家庭裁判所への定期報告",),
          ("本人の生活・療養看護に関する事務（身上保護）",)],
         body_size=11.5)
    card(s, IN(4.85), top, IN(3.85), IN(3.35), "できない・要注意", GOLD,
         [("医療行為への同意権はない（P13）",),
          ("身元保証人にはなれない",),
          ("居住用不動産の処分は家裁の許可が必要",),
          ("日用品の購入など日常の行為は本人が行う",),
          ("本人の意思・自己決定の尊重が大前提",)],
         body_size=11.5)
    card(s, IN(8.9), top, IN(3.63), IN(3.35), "死後事務（限定的）", NAVY,
         [("民法873条の2（平成28年改正）",),
          ("家裁の許可等を要件に、",),
          ("相続財産の保存／弁済期到来の債務弁済／",),
          ("火葬・埋葬の契約 等が可能",),
          ("対象は成年後見人のみ。保佐・補助は対象外",)],
         body_size=11)

    add_rect(s, IN(0.8), top + IN(3.5), IN(11.73), IN(0.95),
             TEAL_LIGHT, radius=0.05)
    add_text(s, IN(1.0), top + IN(3.58), IN(11.4), IN(0.3),
             [("◆ 後見人等と本人の関係（令和6年）", 12.5, NAVY, True)])
    add_text(s, IN(1.0), top + IN(3.9), IN(11.4), IN(0.5),
             [[("親族 約17% ／ 親族以外（専門職等）約83%。",
                12, INK, True),
               ("　専門職の内訳は司法書士＞弁護士＞社会福祉士の順。",
                11.5, INK, False)]], line_spacing=1.2)
    footer_note(s, "出典：最高裁「成年後見関係事件の概況―令和6年―」。"
                   "割合は出典PDFで要確認。")


# ===================================================================
# Slide 10 — 任意後見と法定後見の使い分け
# ===================================================================
def s10(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "制度の最新像 ⑤", "任意後見と法定後見の使い分け", 10)
    x0 = IN(0.8)
    cols = [IN(2.6), IN(4.55), IN(4.58)]
    heads = ["", "任意後見", "法定後見"]
    x = x0
    for i, h_ in enumerate(heads):
        ac = [WHITE, TEAL, NAVY][i]
        add_rect(s, x, top, cols[i], IN(0.5), ac)
        if h_:
            add_text(s, x, top, cols[i], IN(0.5), [(h_, 14, WHITE, True)],
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        x += cols[i]
    rows = [
        ("利用の起点", "判断能力があるうちに\n本人が自ら契約",
         "判断能力が低下した後に\n家庭裁判所が選任"),
        ("後見人を選ぶ", "本人が選べる", "家裁が選任（本人は選べない）"),
        ("契約の方式", "公正証書で契約", "（契約ではなく審判）"),
        ("効力の発生", "家裁が任意後見監督人を\n選任した時から",
         "審判の確定時から"),
        ("取消権", "なし", "あり"),
    ]
    y = top + IN(0.5)
    rh = IN(0.74)
    for ri, (lab, a, b) in enumerate(rows):
        bg = WHITE if ri % 2 == 0 else LIGHT
        add_rect(s, x0, y, cols[0], rh, NAVY_DARK, line=WHITE, line_w=1)
        add_text(s, x0 + IN(0.05), y, cols[0] - IN(0.1), rh,
                 [(lab, 11.5, WHITE, True)], align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        x = x0 + cols[0]
        for ci, ctxt in enumerate([a, b]):
            add_rect(s, x, y, cols[ci + 1], rh, bg, line=LINE, line_w=0.75)
            paras = [[(ln, 11, INK, False)] for ln in ctxt.split("\n")]
            add_text(s, x + IN(0.1), y, cols[ci + 1] - IN(0.2), rh, paras,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                     line_spacing=1.1)
            x += cols[ci + 1]
        y += rh

    add_rect(s, x0, y + IN(0.12), IN(11.73), IN(0.78), TEAL_LIGHT, radius=0.06)
    add_text(s, x0 + IN(0.15), y + IN(0.18), IN(11.4), IN(0.7),
             [[("使い分けの軸：", 12.5, NAVY, True),
               ("「将来に備える」段階＝任意後見、"
                "「すでに判断能力が低下」＝法定後見。",
                12, INK, False)],
              [("　ケアマネは、本人が元気なうちに任意後見・"
                "死後事務委任の選択肢を話題化できる立場にある。",
                11.5, GRAY, False)]], line_spacing=1.25)
    footer_note(s, "当日メモ：任意後見を勧めるべき利用者像は？")


# ===================================================================
# Slide 11 — 2026年 制度大改正の動向（最重要）
# ===================================================================
def s11(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "制度の最新像 ⑥【最重要】",
                 "成年後見制度の見直し ── 約25年ぶりの大改正が進行中", 11)
    add_rect(s, IN(0.8), top, IN(11.73), IN(0.62), GOLD, radius=0.05)
    add_text(s, IN(1.0), top, IN(11.4), IN(0.62),
             [("研修当日（2026年6月）に“いま動いている”最大の論点。"
               "現行制度と改正動向を二段で押さえる。", 12.5, NAVY_DARK, True)],
             anchor=MSO_ANCHOR.MIDDLE)
    # 経緯タイムライン
    add_text(s, IN(0.8), top + IN(0.72), IN(8), IN(0.3),
             [("◆ これまでの経緯", 13, TEAL, True)])
    tl = [
        ("2025.6", "法制審議会の部会が中間試案を取りまとめ（パブコメ実施）"),
        ("2026.2.12", "法制審が「改正に関する要綱」を法務大臣に答申"),
        ("2026.4", "政府が改正法案を閣議決定・国会提出（報道ベース／要確認）"),
        ("2026.5 現在", "国会で審議中。施行は2028年度中の見込み（要確認）"),
    ]
    y = top + IN(1.04)
    for d, t in tl:
        add_rect(s, IN(0.8), y, IN(1.7), IN(0.40), NAVY)
        add_text(s, IN(0.8), y, IN(1.7), IN(0.40), [(d, 10.5, WHITE, True)],
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, IN(2.65), y, IN(6.1), IN(0.40), [(t, 10.5, INK, False)],
                 anchor=MSO_ANCHOR.MIDDLE)
        y += IN(0.46)

    add_text(s, IN(0.8), y + IN(0.04), IN(8), IN(0.3),
             [("◆ 見直しの柱", 13, TEAL, True)])
    pillars = [
        "終身制の廃止 ── 判断能力が回復しなくても、必要がなくなれば終了できる",
        "終了・見直しの柔軟化 ── あらかじめ期間を区切る制度ではなく、"
        "必要性がなくなった時点で家裁が終了できる仕組み（要確認）",
        "3類型を「補助」へ一元化 ── 必要な範囲のみ権限を付与するオーダーメイド型",
        "後見人の交代を容易に ／ デジタル遺言の創設（報道ベース）",
    ]
    yy = y + IN(0.38)
    for p in pillars:
        add_rect(s, IN(0.8), yy + IN(0.06), IN(0.13), IN(0.13), GOLD)
        add_text(s, IN(1.05), yy, IN(7.55), IN(0.46), [(p, 10.5, INK, False)],
                 line_spacing=1.12)
        yy += IN(0.46)

    memo_panel(s, IN(9.05), top + IN(0.72), IN(3.48), IN(4.42),
               "当日メモ：施行時期・現場影響")
    footer_note(s, "出典：法務省 法制審議会民法（成年後見等関係）部会／"
                   "日弁連・リーガルサポート声明。閣議決定・施行日は要確認。")


# ===================================================================
# Slide 12 — 身寄りなし高齢者支援の論点整理
# ===================================================================
def s12(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "現場の本丸 ①", "身寄りのない高齢者支援 ── 論点の整理", 12)
    add_text(s, IN(0.8), top, IN(8.0), IN(0.55),
             [("「身寄りがない」課題は一括りにせず、"
               "“機能”に分解して対応先を考えるのが厚労省ガイドラインの発想。",
               12, INK, False)], line_spacing=1.2)
    add_text(s, IN(0.8), top + IN(0.6), IN(8), IN(0.3),
             [("◆ 4つの困りどころ", 13, TEAL, True)])
    quad = [
        ("入院・入所の身元保証", "緊急連絡先、入院計画書の同意、"
         "物品準備、費用支払い 等"),
        ("医療同意・延命の判断", "手術同意・人生の最終段階の医療を"
         "誰が決めるか"),
        ("金銭管理", "日々の支払い、預貯金・年金の管理"),
        ("死後事務", "火葬・埋葬、残置物、年金停止 等"),
    ]
    qx = [IN(0.8), IN(4.6)]
    qy = [top + IN(0.95), top + IN(2.25)]
    cols = [TEAL, GOLD, NAVY, RGBC(0x4A,0x9A,0x6A)]
    for i, (h_, b_) in enumerate(quad):
        x = qx[i % 2]
        y = qy[i // 2]
        add_rect(s, x, y, IN(3.6), IN(1.15), WHITE, line=LINE, line_w=1,
                 radius=0.06)
        add_rect(s, x, y, IN(0.12), IN(1.15), cols[i])
        add_text(s, x + IN(0.2), y + IN(0.08), IN(3.3), IN(0.32),
                 [(h_, 12, NAVY, True)])
        add_text(s, x + IN(0.2), y + IN(0.42), IN(3.3), IN(0.65),
                 [(b_, 10.5, INK, False)], line_spacing=1.15)

    add_rect(s, IN(0.8), top + IN(3.6), IN(8.0), IN(0.95), TEAL_LIGHT,
             radius=0.05)
    add_text(s, IN(1.0), top + IN(3.68), IN(7.6), IN(0.85),
             [[("ケアマネの立ち位置：", 12, NAVY, True),
               ("自ら保証人になるのではなく、", 11.5, INK, False)],
              [("　機能ごとに制度・専門職・行政へ“つなぐ”ことが役割。"
                "厚労省2019年ガイドライン参照。", 11.5, INK, False)]],
             line_spacing=1.25)

    memo_panel(s, IN(9.05), top, IN(3.48), IN(5.0), "当日メモ")
    footer_note(s, "出典：厚労省「身寄りがない人の入院及び医療に係る意思決定が"
                   "困難な人への支援に関するガイドライン」(2019)")


# ===================================================================
# Slide 13 — 医療同意・身元保証の現状と限界
# ===================================================================
def s13(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "現場の本丸 ②", "医療同意・身元保証 ── 現状と限界", 13)
    card(s, IN(0.8), top, IN(5.75), IN(2.2), "身元保証", TEAL,
         [("身元保証人がいないことのみを理由とした",),
          ("入院拒否は不可（応召義務に抵触）",),
          ("── 厚労省医政局通知 平成30年4月27日",),
          ("民間の高齢者等終身サポート事業者には",),
          ("適正運営のガイドライン（2024年）あり",)], body_size=11)
    card(s, IN(6.78), top, IN(5.75), IN(2.2), "医療同意", GOLD,
         [("成年後見人に医療行為への同意権はない",),
          ("（一身専属的な判断とされる ＝ 現行法の整理）",),
          ("家族がいても「代理決定」ではなく",),
          ("本人意思の推定が出発点",),
          ("→ 誰が決めるのか、が現場の難所",)], body_size=11)

    y = top + IN(2.4)
    add_rect(s, IN(0.8), y, IN(11.73), IN(1.55), TEAL_LIGHT, radius=0.05)
    add_text(s, IN(1.0), y + IN(0.08), IN(11.3), IN(0.3),
             [("◆ 意思決定のよりどころ ── ACP（人生会議）",
               12.5, NAVY, True)])
    add_text(s, IN(1.0), y + IN(0.4), IN(11.3), IN(1.05),
             [[("厚労省「人生の最終段階における医療・ケアの決定プロセスに"
                "関するガイドライン」(2018改訂)。", 11.5, INK, False)],
              [("本人・家族・医療ケアチームが繰り返し話し合い、その都度"
                "文書化。身寄りなしの場合は本人意思を推定し、", 11.5, INK, False)],
              [("困難ならチームで本人にとって最善の方針を決定する。",
                11.5, INK, False)]], line_spacing=1.25)

    add_rect(s, IN(0.8), y + IN(2.1), IN(11.73), IN(0.5), RGBC(0xFC,0xF3,0xDD),
             line=GOLD, line_w=1)
    add_text(s, IN(1.0), y + IN(2.1), IN(11.4), IN(0.5),
             [("予習ポイント：「後見人がいれば医療同意もしてもらえる」は誤解。"
               "ここは要確認の最重要点。", 11.5, INK, True)],
             anchor=MSO_ANCHOR.MIDDLE)
    footer_note(s, "出典：厚労省 医政局通知／人生の最終段階の医療・ケア"
                   "決定プロセスガイドライン／終身サポート事業者ガイドライン")


# ===================================================================
# Slide 14 — 死後事務委任契約の位置づけ
# ===================================================================
def s14(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "現場の本丸 ③", "死後事務 ── 後見人の権限と「死後事務委任契約」", 14)
    card(s, IN(0.8), top, IN(5.75), IN(2.35), "成年後見人の死後事務", NAVY,
         [("民法873条の2（平成28年改正）",),
          ("家裁の許可等を要件に限定的に可：",),
          ("① 相続財産の保存に必要な行為",),
          ("② 弁済期が来た債務の弁済（医療費等）",),
          ("③ 火葬・埋葬に関する契約 ※葬儀主催は不可",)], body_size=11)
    card(s, IN(6.78), top, IN(5.75), IN(2.35), "死後事務委任契約", TEAL,
         [("後見終了後の広い事務に備える生前契約",),
          ("対応範囲を本人と取り決められる",),
          ("（葬儀・納骨・行政手続・残置物処理 等）",),
          ("任意後見契約とあわせて結ぶことが多い",),
          ("元気なうちの備えが鍵",)], body_size=11)

    y = top + IN(2.55)
    add_text(s, IN(0.8), y, IN(11.7), IN(0.3),
             [("◆ 単身死亡後の実務の流れ", 12.5, TEAL, True)])
    flow = ["死亡届の提出", "火葬許可\n（24時間後）", "引取り手なし→\n市区町村が埋火葬",
            "年金停止\n受給権者死亡届", "残置物の処理"]
    x = IN(0.8)
    fw = IN(2.16)
    for i, f in enumerate(flow):
        add_rect(s, x, y + IN(0.35), fw, IN(0.7), LIGHT, line=LINE, line_w=1,
                 radius=0.08)
        paras = [[(ln, 10, INK, False)] for ln in f.split("\n")]
        add_text(s, x, y + IN(0.35), fw, IN(0.7), paras,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 line_spacing=1.05)
        if i < len(flow) - 1:
            add_text(s, x + fw - IN(0.04), y + IN(0.35), IN(0.3), IN(0.7),
                     [("▶", 12, TEAL, True)], align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)
        x += fw + IN(0.20)

    add_rect(s, IN(0.8), y + IN(1.2), IN(11.73), IN(0.5), RGBC(0xFC,0xF3,0xDD),
             line=GOLD, line_w=1)
    add_text(s, IN(1.0), y + IN(1.2), IN(11.4), IN(0.5),
             [("引取り手がない場合：墓地埋葬法第9条／行旅病人及行旅死亡人"
               "取扱法により市区町村が埋火葬を実施。", 11.5, INK, True)],
             anchor=MSO_ANCHOR.MIDDLE)
    footer_note(s, "出典：法務省（民法873条の2）／墓地埋葬法／日本年金機構")


# ===================================================================
# Slide 15 — 金銭管理支援の階層
# ===================================================================
def s15(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "現場の本丸 ④", "金銭管理支援の階層 ── つなぎ先の選び方", 15)
    add_text(s, IN(0.8), top, IN(8.2), IN(0.4),
             [("判断能力の程度と「必要な法律行為の範囲」で支援先を選ぶ。"
               "上ほど軽度、下ほど重度。", 12, INK, False)])
    tiers = [
        ("見守り＋自己管理", "親族・地域・ケアチームによる見守り。"
         "本人が自分で管理できる段階。", RGBC(0x6E,0xA8,0xB8)),
        ("日常生活自立支援事業（社協）", "契約を理解できる力は必要。"
         "福祉サービス利用援助・日常的金銭管理・書類等預かりに限定。利用は有料。",
         TEAL),
        ("補助", "判断能力が不十分。家裁が定める特定の行為のみ支援。",
         RGBC(0x4A,0x9A,0x6A)),
        ("保佐", "判断能力が著しく不十分。重要な法律行為に同意・取消し。",
         RGBC(0x3C,0x6E,0x8F)),
        ("後見", "判断能力を欠くのが通常。財産に関する包括的な支援。", NAVY),
    ]
    # 左の矢印
    add_rect(s, IN(0.8), top + IN(0.5), IN(0.42), IN(3.55),
             RGBC(0xE0,0xE6,0xEC))
    add_text(s, IN(0.78), top + IN(0.55), IN(0.5), IN(0.4),
             [("軽", 11, GRAY, True)], align=PP_ALIGN.CENTER)
    add_text(s, IN(0.78), top + IN(3.62), IN(0.5), IN(0.4),
             [("重", 11, GRAY, True)], align=PP_ALIGN.CENTER)
    add_text(s, IN(0.72), top + IN(1.6), IN(0.6), IN(1.0),
             [[("判", 9.5, GRAY, True)], [("断", 9.5, GRAY, True)],
              [("力", 9.5, GRAY, True)]], align=PP_ALIGN.CENTER,
             line_spacing=1.0)
    y = top + IN(0.5)
    indent = IN(0.0)
    for i, (h_, b_, c) in enumerate(tiers):
        x = IN(1.45) + Emu(int(indent))
        w = IN(7.6) - Emu(int(indent))
        add_rect(s, x, y, w, IN(0.66), c, radius=0.05)
        add_text(s, x + IN(0.15), y + IN(0.03), IN(2.7), IN(0.6),
                 [(h_, 11.5, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x + IN(2.95), y + IN(0.02), w - IN(3.1), IN(0.62),
                 [(b_, 9.8, WHITE, False)], anchor=MSO_ANCHOR.MIDDLE,
                 line_spacing=1.1)
        y += IN(0.71)
        indent += Inches(0.0)

    memo_panel(s, IN(9.3), top, IN(3.23), IN(5.0), "選定の判断メモ")
    footer_note(s, "予習ポイント：日常生活自立支援事業と成年後見は対立でなく"
                   "段階。移行のタイミングが論点（→質問P19）。")


# ===================================================================
# Slide 16 — 市町村長申立ての要件と運用
# ===================================================================
def s16(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "現場の本丸 ⑤", "市町村長申立て ── 要件と中央区での運用", 16)
    card(s, IN(0.8), top, IN(5.75), IN(2.45), "法的根拠と要件", NAVY,
         [("根拠：老人福祉法 第32条",),
          ("（知的障害者・精神障害者にも同様規定）",),
          ("本人・親族が申立てできず、本人の福祉のため",),
          ("特に必要な場合に市町村長が申立て",),
          ("運用：まず2親等内の親族の有無を確認",)], body_size=11)
    card(s, IN(6.78), top, IN(5.75), IN(2.45), "件数と支援", TEAL,
         [("市区町村長申立て：令和6年 約1.0万件",),
          ("申立人別では最多（全体の約24%）",),
          ("背景に身寄りなし高齢者の増加",),
          ("成年後見制度利用支援事業：",),
          ("申立費用・後見人報酬を助成（要件は自治体ごと）",)], body_size=11)

    y = top + IN(2.65)
    add_rect(s, IN(0.8), y, IN(11.73), IN(1.55), TEAL_LIGHT, radius=0.05)
    add_text(s, IN(1.0), y + IN(0.08), IN(11.3), IN(0.3),
             [("◆ 中央区での運用（窓口の整理／詳細は当日・公式サイトで要確認）",
               12.5, NAVY, True)])
    add_text(s, IN(1.0), y + IN(0.42), IN(11.3), IN(1.05),
             [[("・ 区長申立ての担当：中央区 福祉保健部 高齢者福祉課",
                11.5, INK, False)],
              [("・ 入口の相談：おとしより相談センター（地域包括）／"
                "成年後見支援センター「すてっぷ中央」(中央区社協)",
                11.5, INK, False)],
              [("・ 利用支援事業による費用・報酬助成あり（対象要件は区へ確認）",
                11.5, INK, False)]], line_spacing=1.3)
    footer_note(s, "出典：老人福祉法／最高裁概況（件数は要確認）／"
                   "中央区公式・中央区社協（窓口は公式サイトで要確認）")


# ===================================================================
# Slide 17 — 法律専門職との連携フロー
# ===================================================================
def s17(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "連携実務 ①", "法律専門職との連携フロー ── 誰に、何を相談する？", 17)
    # 上段：ケアマネ → 入口
    add_rect(s, IN(0.8), top, IN(2.8), IN(0.7), NAVY, radius=0.08)
    add_text(s, IN(0.8), top, IN(2.8), IN(0.7),
             [("ケアマネジャー", 12.5, WHITE, True)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, IN(3.6), top, IN(0.5), IN(0.7), [("▶", 15, GOLD, True)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, IN(4.15), top, IN(4.3), IN(0.7), TEAL, radius=0.08)
    add_text(s, IN(4.15), top, IN(4.3), IN(0.7),
             [[("地域の入口窓口", 11.5, WHITE, True)],
              [("おとしより相談センター／すてっぷ中央", 9.5, WHITE, False)]],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)
    add_text(s, IN(8.45), top, IN(0.5), IN(0.7), [("▶", 15, GOLD, True)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, IN(9.0), top, IN(3.53), IN(0.7), NAVY, radius=0.08)
    add_text(s, IN(9.0), top, IN(3.53), IN(0.7),
             [("法律専門職・専門職団体へ", 11, WHITE, True)],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 3専門職カード
    y = top + IN(0.95)
    pros = [
        ("司法書士", "リーガルサポート", TEAL,
         ["財産管理・登記・契約手続が中心",
          "第三者後見人として選任が最多",
          "後見人候補者の推薦窓口あり"]),
        ("弁護士", "日弁連／法テラス", NAVY,
         ["紛争性の高い案件に強い",
          "相続争い・虐待・財産被害・訴訟",
          "低所得者は法テラスの援助"]),
        ("社会福祉士", "ぱあとなあ", RGBC(0x4A,0x9A,0x6A),
         ["身上保護・生活支援に強い",
          "福祉的視点のきめ細かな支援",
          "後見人候補者を登録・紹介"]),
    ]
    x = IN(0.8)
    cw = IN(3.78)
    for name, org, ac, lines in pros:
        add_rect(s, x, y, cw, IN(2.55), WHITE, line=LINE, line_w=1, radius=0.05)
        add_rect(s, x, y, cw, IN(0.62), ac, radius=0.05)
        add_rect(s, x, y + IN(0.4), cw, IN(0.22), ac)
        add_text(s, x, y + IN(0.02), cw, IN(0.4), [(name, 14, WHITE, True)],
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, x, y + IN(0.36), cw, IN(0.26), [(org, 9.5, WHITE, False)],
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        paras = [[("• ", 10.5, ac, True), (ln, 10.5, INK, False)]
                 for ln in lines]
        add_text(s, x + IN(0.16), y + IN(0.76), cw - IN(0.3), IN(1.7), paras,
                 line_spacing=1.25, space_after=4)
        x += cw + IN(0.20)

    add_rect(s, IN(0.8), y + IN(2.7), IN(11.73), IN(0.62), TEAL_LIGHT,
             radius=0.06)
    add_text(s, IN(1.0), y + IN(2.7), IN(11.4), IN(0.62),
             [("三者とも後見受任が可能。紛争性・財産管理・身上保護のどれが"
               "中心かで相談先を見立て、迷えば入口窓口に相談。",
               11.5, NAVY, True)], anchor=MSO_ANCHOR.MIDDLE)
    footer_note(s, "出典：リーガルサポート／日弁連／法テラス／日本社会福祉士会")


# ===================================================================
# Slide 18 — 中央区の社会資源マップ
# ===================================================================
def s18(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "連携実務 ②", "中央区の社会資源マップ", 18)
    add_text(s, IN(0.8), top, IN(11.7), IN(0.32),
             [("※ 住所・電話番号は調査時点のもの。研修当日・各公式サイトで"
               "最新情報を要確認。", 10.5, GOLD, True)])
    res = [
        ("成年後見支援センター\n「すてっぷ中央」", TEAL,
         ["運営：中央区社会福祉協議会",
          "後見の相談／日常生活自立支援事業",
          "中央区八丁堀4-1-5",
          "TEL 03-3206-0567（要確認）"]),
        ("おとしより相談センター", NAVY,
         ["中央区の地域包括支援センター",
          "区内6か所に配置",
          "日本橋圏域：日本橋／人形町",
          "高齢者支援の総合相談の入口"]),
        ("中央区 高齢者福祉課", RGBC(0x4A,0x9A,0x6A),
         ["福祉保健部の担当課",
          "区長申立ての担当窓口",
          "成年後見制度利用支援事業",
          "成年後見制度利用促進審議会"]),
        ("リーガルサポート東京支部", RGBC(0x3C,0x6E,0x8F),
         ["司法書士による後見支援",
          "新宿区四谷本塩町",
          "TEL 03-3353-8191（要確認）",
          "成年後見ホットライン 03-5379-1888"]),
    ]
    positions = [(IN(0.8), top + IN(0.45)), (IN(6.78), top + IN(0.45)),
                 (IN(0.8), top + IN(2.7)), (IN(6.78), top + IN(2.7))]
    for (name, ac, lines), (x, y) in zip(res, positions):
        w, h = IN(5.75), IN(2.05)
        add_rect(s, x, y, w, h, WHITE, line=LINE, line_w=1, radius=0.04)
        add_rect(s, x, y, IN(0.13), h, ac)
        nameparas = [[(ln, 12.5, NAVY, True)] for ln in name.split("\n")]
        add_text(s, x + IN(0.28), y + IN(0.1), w - IN(0.45), IN(0.7),
                 nameparas, line_spacing=1.05)
        ny = y + (IN(0.78) if "\n" in name else IN(0.5))
        paras = [[("• ", 10.5, ac, True), (ln, 10.5, INK, False)]
                 for ln in lines]
        add_text(s, x + IN(0.28), ny, w - IN(0.45), h - IN(0.85), paras,
                 line_spacing=1.22, space_after=3)
    footer_note(s, "出典：中央区公式サイト／中央区社協／リーガルサポート"
                   "東京支部（窓口情報は要確認）")


# ===================================================================
# Slide 19 / 20 — 講師への質問リスト（メモ欄付き）
# ===================================================================
def question_slide(prs, num, kicker, title, qs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, kicker, title, num)
    add_text(s, IN(0.8), top, IN(7.1), IN(0.32),
             [("予習で残った疑問。当日、講師の説明・回答を右の余白にメモする。",
               11, GRAY, False)])
    y = top + IN(0.42)
    for i, q in enumerate(qs):
        h = IN(0.84)
        add_rect(s, IN(0.8), y, IN(7.15), h, WHITE, line=LINE, line_w=1,
                 radius=0.04)
        add_rect(s, IN(0.8), y, IN(0.66), h, NAVY)
        add_text(s, IN(0.8), y, IN(0.66), h, [(q[0], 15, WHITE, True)],
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, IN(1.62), y + IN(0.07), IN(6.2), h - IN(0.14),
                 [(q[1], 11, INK, False)], line_spacing=1.18,
                 anchor=MSO_ANCHOR.MIDDLE)
        # 右のメモ枠
        add_rect(s, IN(8.1), y, IN(4.43), h, MEMO_BG, line=MEMO_LINE,
                 line_w=1)
        ly = y + IN(0.24)
        while ly < y + h - IN(0.1):
            add_rect(s, IN(8.25), ly, IN(4.13), Pt(0.75), MEMO_LINE)
            ly += IN(0.32)
        y += h + IN(0.10)
    add_text(s, IN(8.1), top + IN(0.1), IN(4.4), IN(0.3),
             [("← 当日メモ（講師の回答）", 9.5, RGBC(0x9A,0x86,0x3A), True)])


def s19(prs):
    question_slide(prs, 19, "予習の総仕上げ ①",
                   "講師への質問リスト ① ── 制度・申立て・改正",
                   [("Q1", "2026年改正で3類型が「補助」へ一元化された後、"
                            "現場の運用は具体的にどう変わるか。施行時期の見込みは。"),
                    ("Q2", "終身制の廃止（終われる後見）で、いったん始めた"
                            "後見を終了できる具体的な要件・手続きは。"),
                    ("Q3", "鑑定が必要になるのはどんなケースか。"
                            "中央区での申立て〜開始までの実際の所要期間は。"),
                    ("Q4", "任意後見契約を勧めるべき利用者像と、"
                            "契約から効力発生までで実務がつまずきやすい点は。"),
                    ("Q5", "後見人（特に専門職）とケアマネの役割分担で、"
                            "現場で摩擦が生じやすいのはどんな場面か。")])


def s20(prs):
    question_slide(prs, 20, "予習の総仕上げ ②",
                   "講師への質問リスト ② ── 身寄りなし支援・連携・中央区",
                   [("Q6", "身寄りなし利用者の入院で身元保証会社を使う場合、"
                            "良い事業者をどう見極めるか（終身サポートGLの活用）。"),
                    ("Q7", "医療同意が要る場面で後見人も家族もいない時、"
                            "現場として誰とどう意思決定を進めるべきか。"),
                    ("Q8", "死後事務委任契約のないまま利用者が単身死亡した場合、"
                            "ケアマネ・社協・区の現実的な動き方は。"),
                    ("Q9", "日常生活自立支援事業から後見へ移行を判断する"
                            "タイミング・見極めの目安は。"),
                    ("Q10", "中央区で区長申立てを検討する際の相談先・"
                             "標準的な所要期間、すてっぷ中央との連携の実際は。")])


# ===================================================================
# Slide 21 — 参考文献・出典URL
# ===================================================================
def s21(prs):
    s = _blank(prs)
    page_bg(s)
    top = header(s, "巻末", "参考文献・出典URL（2026年5月 調査時点）", 21)
    groups = [
        ("制度全般・統計", [
            "最高裁 成年後見関係事件の概況　https://www.courts.go.jp/toukei_siryou/siryo/kouken/index.html",
            "裁判所 成年後見制度の概要　https://www.courts.go.jp/saiban/koukenp00/koukenp1/index.html",
            "法務省 成年後見制度Q&A　https://www.moj.go.jp/MINJI/a01.html",
            "厚労省 成年後見はやわかり　https://guardianship.mhlw.go.jp/",
        ]),
        ("利用促進・制度見直し", [
            "厚労省 成年後見制度利用促進　https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000202622_00017.html",
            "法務省 法制審議会 民法（成年後見等関係）部会　https://www.moj.go.jp/shingi1/housei02_003007_00008",
            "日弁連 改正要綱に対する会長声明(2026.2.12)　https://www.nichibenren.or.jp/document/statement/year/2026/260212.html",
        ]),
        ("身寄りなし支援・医療・死後事務", [
            "厚労省 身寄りがない人の入院・意思決定支援ガイドライン(2019)　https://www.mhlw.go.jp/content/000516181.pdf",
            "厚労省 人生の最終段階における医療・ケアの決定プロセスGL(2018)",
            "法務省 成年後見人の死後事務(民法873条の2)　https://www.moj.go.jp/MINJI/minji07_00196.html",
        ]),
        ("連携・中央区の窓口", [
            "リーガルサポート東京支部　https://legal-support-tokyo.jp/",
            "日本社会福祉士会 ぱあとなあ　https://www.jacsw.or.jp/citizens/seinenkoken/",
            "中央区 成年後見制度の利用促進　https://www.city.chuo.lg.jp/kenkouiryou/fukushiippan/chiikifukushi/koukenseido/index.html",
            "中央区社協 すてっぷ中央　https://www.shakyo-chuo-city.jp/jigyo/stepchuo",
        ]),
    ]
    y = top
    for gtitle, urls in groups:
        add_rect(s, IN(0.8), y, IN(11.73), IN(0.34), NAVY)
        add_text(s, IN(0.95), y, IN(11.5), IN(0.34),
                 [(gtitle, 11.5, WHITE, True)], anchor=MSO_ANCHOR.MIDDLE)
        yy = y + IN(0.40)
        for u in urls:
            add_rect(s, IN(0.95), yy + IN(0.06), IN(0.10), IN(0.10), TEAL)
            add_text(s, IN(1.15), yy, IN(11.2), IN(0.26),
                     [(u, 9.3, INK, False)])
            yy += IN(0.265)
        y = yy + IN(0.12)
    add_text(s, IN(0.8), IN(7.04), IN(11.7), IN(0.3),
             [("※ 官公庁PDFは自動取得が制限されたため一部数値は二次情報経由。"
               "研修資料への引用前に各原典で要確認。", 9, GRAY, False)])


# ===================================================================
def main():
    prs = new_deck()
    for fn in [s01, s02, s03, s04, s05, s06, s07, s08, s09, s10,
               s11, s12, s13, s14, s15, s16, s17, s18, s19, s20, s21]:
        fn(prs)
    out_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out = os.path.join(out_dir, "令和8年度第3回ケアマネ研修_予習スライド_成年後見制度.pptx")
    prs.save(out)
    print("saved:", out, "/ slides:", len(prs.slides._sldIdLst))


if __name__ == "__main__":
    main()
