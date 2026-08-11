"""株式会社蒼蓮あて 2026年8月分 請求書（Aloft大阪）を生成する。"""
import os
import openpyxl
from openpyxl.styles import Font, Alignment

BASE = os.path.dirname(os.path.abspath(__file__))
src = os.path.join(BASE, "invoice_template.xlsx")
wb = openpyxl.load_workbook(src)
ws = wb.active
FONT = "Arial"

def setv(coord, value, size=None, bold=None, color=None, align=None, wrap=None):
    c = ws[coord]
    c.value = value
    f = c.font
    c.font = Font(name=FONT,
                  size=size if size is not None else f.size,
                  bold=bold if bold is not None else f.bold,
                  color=color if color is not None else f.color)
    a = c.alignment
    c.alignment = Alignment(horizontal=align or a.horizontal,
                            vertical=a.vertical or "center",
                            wrap_text=wrap if wrap is not None else a.wrap_text)

# ---- 請求先 ----
setv("A5", "株式会社蒼蓮　御中")
setv("A6", "ご担当：小笹彩加 様")
setv("A7", "〒542-0012 大阪市中央区谷町6丁目3番25号 LOK09号")

# ---- メタ ----
setv("E4", "2026-08")
setv("E5", "2026/08/11")   # 請求日
setv("E6", "")             # 支払期限：記入

# ---- 発行者 ----
setv("A12", "勝矢 夏帆")
setv("A13", "〒550-0003 大阪府大阪市西区京町堀2-9-18 ジュネーゼ京町堀パークサイド601")
setv("A14", "TEL / Email（任意・必要に応じてご記入）")
setv("A15", "※適格請求書発行事業者 未登録（インボイス経過措置により調整額8%で計上）")

# ---- お振込先 ----
setv("A33", "三井住友銀行　姫路支店")
setv("A34", "普通　9554406")
setv("A35", "口座名義：勝矢 夏帆（カツヤ カホ）")

# ---- 明細（8月分）----
items = [
    (18, "8/4 Aloft（報酬）", 1, 60000),
    (19, "8/4 交通費 バス（土佐堀2丁目→堂島）※8月以降は事務所発の交通費を含む", 1, 210),
]
for r in range(18, 26):
    ws[f"B{r}"].value = None
    ws[f"C{r}"].value = None
    ws[f"D{r}"].value = None
for r, name, qty, unit in items:
    setv(f"B{r}", name, size=9, align="left", wrap=True)
    ws[f"C{r}"].value = qty
    ws[f"D{r}"].value = unit
    ws.row_dimensions[r].height = 28

# ---- 集計ラベルを「調整額（経過措置8%）」に変更 ----
setv("C27", "調整額率（経過措置80%）", size=11, align="right")
ws["E27"].value = 0.08
setv("C28", "調整額（消費税相当）", size=11, align="right")
ws["E28"].value = "=ROUND(E26*E27,0)"

# ---- 備考 ----
note = ("※ インボイス制度の経過措置（80%控除）に基づき、消費税10%ではなく調整額8%（小計×8%）で計上しています。\n"
        "※ 8月以降は事務所までの交通費も請求対象に含めています。\n"
        "※ 源泉徴収は「なし」で計上しています。必要な場合は源泉徴収税額欄に金額を入力すると合計から差し引かれます。\n"
        "※ ご提出：毎月5日までにメール（To: ayakaozasa@soren.co.jp / CC: ayakaanzo@soren.co.jp）")
setv("A38", note, size=9, wrap=True)

out = os.path.join(BASE, "invoice_soren_2026-08.xlsx")
wb.save(out)
print("saved", out)
sub = sum(v for *_, v in items)
adj = round(sub * 0.08)
print("小計:", sub, "調整額8%:", adj, "合計:", sub + adj)
