# -*- coding: utf-8 -*-
"""株式会社蒼蓮あて 請求書（提出用PDF）を生成する。日本語フォントを埋め込む。"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor, black, white

BASE = os.path.dirname(os.path.abspath(__file__))
JP_TTF = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
pdfmetrics.registerFont(TTFont("JP", JP_TTF))
FONT = "JP"
NAVY = HexColor("#1F3864")
BAND = HexColor("#EAEFF7")
GRAY = HexColor("#BFBFBF")
DARK = HexColor("#404040")

W, H = A4
L = 45
R = W - 45


def yen(n):
    return f"¥{n:,}"


def build(path, inv_no, issue_date, due_date, items, remarks=None):
    c = canvas.Canvas(path, pagesize=A4)

    def text(x, y, s, size=10, color=black):
        c.setFillColor(color); c.setFont(FONT, size); c.drawString(x, y, s)

    def rtext(x, y, s, size=10, color=black):
        c.setFillColor(color); c.setFont(FONT, size); c.drawRightString(x, y, s)

    def ctext(x, y, s, size=10, color=black):
        c.setFillColor(color); c.setFont(FONT, size); c.drawCentredString(x, y, s)

    # ---- タイトル ----
    y = H - 60
    ctext(W / 2, y, "請求書", size=24, color=NAVY)
    y -= 12
    c.setStrokeColor(NAVY); c.setLineWidth(1.4); c.line(L, y, R, y)

    # ---- 宛先（左）----
    y -= 34
    text(L, y, "株式会社蒼蓮　御中", size=14)
    c.setStrokeColor(HexColor("#808080")); c.setLineWidth(0.6)
    c.line(L, y - 4, L + 250, y - 4)
    text(L, y - 18, "ご担当：小笹彩加 様", size=9.5)
    text(L, y - 33, "〒542-0012 大阪市中央区谷町6丁目3番25号 LOK09号", size=8.5, color=DARK)

    # ---- 請求メタ（右）----
    my = H - 106
    for label, val in [("請求日", issue_date), ("請求書番号", inv_no), ("支払期限", due_date)]:
        rtext(R - 95, my, label, size=9.5, color=NAVY)
        rtext(R, my, val, size=9.5)
        my -= 15

    # ---- 発行者（右）----
    iy = my - 14
    rtext(R, iy, "勝矢 夏帆", size=12)
    rtext(R, iy - 15, "〒550-0003 大阪府大阪市西区京町堀2-9-18", size=8.5, color=DARK)
    rtext(R, iy - 27, "ジュネーゼ京町堀パークサイド601", size=8.5, color=DARK)

    # ---- 金額計算 ----
    total = sum(q * u for _, q, u in items)
    adj = round(total * 0.08)
    grand = total + adj

    # ---- ご請求金額 ----
    by = iy - 56
    box_h = 34
    c.setFillColor(NAVY); c.rect(L, by - box_h, 150, box_h, fill=1, stroke=0)
    c.setFillColor(BAND); c.rect(L + 150, by - box_h, 210, box_h, fill=1, stroke=0)
    ctext(L + 75, by - box_h + 12, "ご請求金額（税込）", size=11, color=white)
    ctext(L + 255, by - box_h + 10, yen(grand), size=16, color=NAVY)
    text(L, by - box_h - 20, "下記のとおりご請求申し上げます。", size=9.5)

    # ---- 明細テーブル ----
    ty = by - box_h - 40
    c_no = (L, L + 30)
    c_item = (L + 30, L + 310)
    c_qty = (L + 310, L + 350)
    c_unit = (L + 350, L + 425)
    c_amt = (L + 425, R)
    xs = [L, c_item[0], c_qty[0], c_unit[0], c_amt[0], R]

    header_h = 22
    c.setFillColor(NAVY); c.rect(L, ty - header_h, R - L, header_h, fill=1, stroke=0)
    ctext((c_no[0] + c_no[1]) / 2, ty - header_h + 7, "No.", size=10.5, color=white)
    text(c_item[0] + 6, ty - header_h + 7, "品目", size=10.5, color=white)
    ctext((c_qty[0] + c_qty[1]) / 2, ty - header_h + 7, "数量", size=10.5, color=white)
    ctext((c_unit[0] + c_unit[1]) / 2, ty - header_h + 7, "単価", size=10.5, color=white)
    ctext((c_amt[0] + c_amt[1]) / 2, ty - header_h + 7, "金額", size=10.5, color=white)

    row_h = 24
    n = max(len(items), 8)
    cur = ty - header_h
    c.setStrokeColor(GRAY); c.setLineWidth(0.5)
    for i in range(n):
        rb = cur - row_h
        c.line(L, rb, R, rb)
        if i < len(items):
            name, qty, unit = items[i]
            ctext((c_no[0] + c_no[1]) / 2, rb + 8, str(i + 1), size=9.5)
            text(c_item[0] + 6, rb + 8, name, size=9.5)
            ctext((c_qty[0] + c_qty[1]) / 2, rb + 8, f"{qty:,}", size=9.5)
            rtext(c_unit[1] - 6, rb + 8, yen(unit), size=9.5)
            rtext(c_amt[1] - 6, rb + 8, yen(qty * unit), size=9.5)
        cur = rb
    for x in xs:
        c.line(x, ty, x, cur)
    c.line(L, ty, R, ty)

    # ---- 集計行の定義 ----
    rows = [("小計", total, False),
            ("調整額（インボイス経過措置 8%）", adj, False),
            ("合計（税込）", grand, True)]

    def fit_size(s, max_w, max_size, min_size=7.0):
        sz = max_size
        while sz > min_size and c.stringWidth(s, FONT, sz) > max_w:
            sz -= 0.5
        return sz

    if remarks:
        # 備考（左）＋ 集計（右）を同じ高さで配置
        rm_right = L + 250          # 備考ボックス右端
        pad = 6
        line_h = 11.5
        head_h = 18
        body_h = len(remarks) * line_h + 10
        block_h = head_h + body_h
        rh = block_h / 3.0
        block_top = cur - 8

        # 備考ヘッダ
        c.setFillColor(NAVY)
        c.rect(L, block_top - head_h, rm_right - L, head_h, fill=1, stroke=0)
        text(L + 6, block_top - head_h + 5, "備考", size=10, color=white)
        # 備考ボックス枠
        c.setStrokeColor(GRAY); c.setLineWidth(0.5)
        c.rect(L, block_top - block_h, rm_right - L, body_h, fill=0, stroke=1)
        # 備考本文
        ly = block_top - head_h - pad - 5
        for line in remarks:
            if line:
                sz = fit_size(line, rm_right - L - 2 * pad, 9.0, 7.0)
                text(L + pad, ly, line, size=sz, color=black)
            ly -= line_h

        # 集計（右・金額列に揃える。3行で備考と同じ高さ）
        sy = block_top
        for label, val, strong in rows:
            bot = sy - rh
            midy = bot + rh / 2 - 4
            if strong:
                c.setFillColor(BAND); c.rect(c_amt[0], bot, c_amt[1] - c_amt[0], rh, fill=1, stroke=0)
            c.setStrokeColor(GRAY); c.setLineWidth(0.5)
            c.rect(c_amt[0], bot, c_amt[1] - c_amt[0], rh, fill=0, stroke=1)
            col = NAVY if strong else black
            lsz = fit_size(label, c_amt[0] - 8 - rm_right, 10.0, 7.0)
            rtext(c_amt[0] - 8, midy, label, size=11 if strong else lsz, color=col)
            rtext(c_amt[1] - 6, midy, yen(val), size=11 if strong else 10, color=col)
            sy = bot
        sy = block_top - block_h
    else:
        # 集計のみ（金額列に揃えて縦積み）
        sy = cur - 6
        rh = 22
        for label, val, strong in rows:
            bot = sy - rh
            if strong:
                c.setFillColor(BAND); c.rect(c_amt[0], bot, c_amt[1] - c_amt[0], rh, fill=1, stroke=0)
            c.setStrokeColor(GRAY); c.setLineWidth(0.5)
            c.rect(c_amt[0], bot, c_amt[1] - c_amt[0], rh, fill=0, stroke=1)
            col = NAVY if strong else black
            rtext(c_amt[0] - 8, bot + 7, label, size=11 if strong else 10, color=col)
            rtext(c_amt[1] - 6, bot + 7, yen(val), size=11 if strong else 10, color=col)
            sy = bot

    # ---- お振込先 ----
    by2 = sy - 30
    text(L, by2, "お振込先", size=11, color=NAVY)
    c.setStrokeColor(NAVY); c.setLineWidth(0.8); c.line(L, by2 - 4, L + 250, by2 - 4)
    for i, line in enumerate(["三井住友銀行　姫路支店", "普通　9554406",
                              "口座名義：勝矢 夏帆（カツヤ カホ）"]):
        text(L, by2 - 18 - i * 14, line, size=9.5)

    c.showPage(); c.save()
    print(os.path.basename(path), "合計", grand)


july_items = [
    ("7/22 Westin横浜 対面ミーティング", 1, 5000),
    ("7/22 交通費（新幹線 新大阪→新横浜）", 1, 14190),
    ("7/22 交通費（現地移動 PASMO）", 1, 708),
    ("7/23 Six Senses Kyoto 交通費（新幹線 新横浜→京都）", 1, 13300),
    ("7/23 交通費（現地移動 PASMO）", 1, 1348),
]
august_items = [
    ("8/4 Aloft大阪 報酬", 1, 60000),
    ("8/4 交通費（バス 土佐堀2丁目→堂島）", 1, 210),
]

july_remarks = [
    "7/22 現地移動交通費（内訳）",
    "新幹線 新大阪→新横浜（スマートEX）¥14,190",
    "JR 大阪→新大阪 ¥180　新横浜→菊名 ¥155",
    "東急菊名→みなとみらい ¥373",
    "",
    "7/23 現地移動交通費（内訳）",
    "新幹線 新横浜→京都（スマートEX）¥13,300",
    "みなとみらい→東急菊名 ¥373",
    "菊名→新横浜 ¥155　JR 京都→大阪 ¥580",
    "地下鉄 梅田→阿波座 ¥240",
]
build(os.path.join(BASE, "請求書_株式会社蒼蓮_2026年7月分.pdf"),
      "2026-07", "2026/07/31", "", july_items, remarks=july_remarks)
build(os.path.join(BASE, "請求書_株式会社蒼蓮_2026年8月分.pdf"),
      "2026-08", "2026/08/31", "", august_items)
