from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OUT = r"C:\Users\ncucc\Desktop\20261030\個人行程與任務管理系統_PRD.docx"
doc=Document()
sec=doc.sections[0]
sec.top_margin=Inches(.65); sec.bottom_margin=Inches(.65); sec.left_margin=Inches(.78); sec.right_margin=Inches(.78)
PURPLE="6738E8"; INK="29263A"; MUTED="6F6A78"; PALE="F5F0FF"; ALT="FBFAFD"

def font(run,size=9.3,bold=False,color=INK):
    run.font.name="Microsoft JhengHei"
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"),"Microsoft JhengHei")
    run.font.size=Pt(size); run.bold=bold; run.font.color.rgb=RGBColor.from_string(color)
    return run

styles=doc.styles
for name in ["Normal","Title","Subtitle","Heading 1","Heading 2"]:
    s=styles[name]; s.font.name="Microsoft JhengHei"; s._element.rPr.rFonts.set(qn("w:eastAsia"),"Microsoft JhengHei")
styles["Normal"].font.size=Pt(9.3); styles["Normal"].font.color.rgb=RGBColor.from_string(INK)
styles["Normal"].paragraph_format.space_after=Pt(4); styles["Normal"].paragraph_format.line_spacing=1.1
styles["Title"].font.size=Pt(25); styles["Title"].font.bold=True; styles["Title"].font.color.rgb=RGBColor.from_string(INK)
styles["Title"].paragraph_format.space_after=Pt(3)
styles["Subtitle"].font.size=Pt(12); styles["Subtitle"].font.color.rgb=RGBColor.from_string(PURPLE)
styles["Heading 1"].font.size=Pt(15); styles["Heading 1"].font.bold=True; styles["Heading 1"].font.color.rgb=RGBColor.from_string(PURPLE)
styles["Heading 1"].paragraph_format.space_before=Pt(12); styles["Heading 1"].paragraph_format.space_after=Pt(4)
styles["Heading 2"].font.size=Pt(11); styles["Heading 2"].font.bold=True; styles["Heading 2"].font.color.rgb=RGBColor.from_string(INK)
styles["Heading 2"].paragraph_format.space_before=Pt(7); styles["Heading 2"].paragraph_format.space_after=Pt(3)
for style_name in ["Title","Heading 1","Heading 2"]:
    ppr=styles[style_name]._element.get_or_add_pPr()
    for tag in ["w:pBdr","w:numPr"]:
        node=ppr.find(qn(tag))
        if node is not None: ppr.remove(node)

def shade(cell,color):
    shd=OxmlElement("w:shd"); shd.set(qn("w:fill"),color); cell._tc.get_or_add_tcPr().append(shd)
def margins(cell):
    tcMar=OxmlElement("w:tcMar")
    for k,v in [("top","80"),("start","100"),("bottom","80"),("end","100")]:
        el=OxmlElement("w:"+k); el.set(qn("w:w"),v); el.set(qn("w:type"),"dxa"); tcMar.append(el)
    cell._tc.get_or_add_tcPr().append(tcMar)
def add_table(headers,rows,widths):
    t=doc.add_table(rows=1,cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.LEFT; t.autofit=False
    tbl_ind=OxmlElement("w:tblInd"); tbl_ind.set(qn("w:w"),"580"); tbl_ind.set(qn("w:type"),"dxa"); t._tbl.tblPr.append(tbl_ind)
    repeat_header=OxmlElement("w:tblHeader"); repeat_header.set(qn("w:val"),"true"); t.rows[0]._tr.get_or_add_trPr().append(repeat_header)
    for col,w in zip(t.columns,widths): col.width=Inches(w)
    for j,h in enumerate(headers):
        c=t.rows[0].cells[j]; shade(c,PURPLE); margins(c); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.space_after=Pt(0); font(p.add_run(h),8.2,True,"FFFFFF")
    for i,row in enumerate(rows):
        cells=t.add_row().cells
        for c,w in zip(cells,widths): c.width=Inches(w)
        for j,value in enumerate(row):
            c=cells[j]; margins(c); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if i%2==1: shade(c,ALT)
            p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.03
            font(p.add_run(str(value)),8.1)
    doc.add_paragraph().paragraph_format.space_after=Pt(0)
    return t
def para(text="",style=None,size=9.3,color=INK,bold=False):
    if isinstance(style,(int,float)): size=style; style=None
    p=doc.add_paragraph(style=style)
    if text: font(p.add_run(text),size,bold,color)
    return p
def bullet(text):
    p=doc.add_paragraph(style="List Bullet"); p.paragraph_format.space_after=Pt(2); font(p.add_run(text),9)
def number(text):
    p=doc.add_paragraph(style="List Number"); p.paragraph_format.space_after=Pt(2); font(p.add_run(text),9)

h=sec.header.paragraphs[0]; h.alignment=WD_ALIGN_PARAGRAPH.RIGHT; font(h.add_run("產品需求文件  |  日序 DAYFLOW"),8,True,PURPLE)
f=sec.footer.paragraphs[0]; f.alignment=WD_ALIGN_PARAGRAPH.CENTER; font(f.add_run("個人行程與任務管理系統  •  版本 1.0  •  2026-10-30"),8,False,MUTED)

p=doc.add_paragraph(style="Title"); font(p.add_run("個人行程與任務管理系統"),25,True,INK)
p=doc.add_paragraph(style="Subtitle"); font(p.add_run("產品需求文件  PRD"),12,True,PURPLE)
para("版本 1.0    |    需求整理日期 2026 年 10 月 30 日    |    產品代稱 日序 DAYFLOW",size=8.5,color=MUTED)
# Summary callout
t=doc.add_table(rows=1,cols=1); t.alignment=WD_TABLE_ALIGNMENT.CENTER
c=t.cell(0,0); shade(c,PALE); margins(c)
p=c.paragraphs[0]; p.paragraph_format.space_after=Pt(3); font(p.add_run("產品摘要"),10.5,True,PURPLE)
p=c.add_paragraph(); p.paragraph_format.space_after=Pt(0)
font(p.add_run("日序是一套供個人使用的行程與任務管理介面，首頁聚焦當日安排。使用者可把有時間的活動放進時間軸，也能保留未排時間的待辦，並用重複規則、重要星號、完成狀態、所在地天氣與每日笑話支援每日規劃。"),9.1)
para("本文件定義第一版產品範圍、互動規則與驗收條件。現有交付物為可操作的瀏覽器前端原型；系統通知、雲端同步等尚待後續技術決策。",9)

para("1 產品概述","Heading 1")
para("目標使用者","Heading 2")
para("單一個人使用者，透過桌面或手機瀏覽器檢視每日安排、記錄任務並切換日期。第一版不包含多人協作或團隊權限。")
para("產品目標","Heading 2")
bullet("在一個首頁同時掌握今日行程、未排程任務與完成進度。")
bullet("降低新增、調整、完成與刪除安排的操作成本。")
bullet("以明亮純色和背景小插圖建立活潑、易辨識的個人化體驗。")
para("設計原則","Heading 2")
bullet("今天優先：首頁呈現所選日期的時間軸、未排程任務及摘要。")
bullet("彈性安排：任務可只設定日期；行程支援時段或全天。")
bullet("視覺鮮明：紫、桃紅、珊瑚橘、青藍、明黃純色；背景使用低透明度插圖，不使用漸層。")

para("2 第一版範圍","Heading 1")
add_table(["範圍","功能"],[
("納入 MVP","今日首頁、日期切換、月曆選日、行程時間軸、未排程任務、摘要與完成狀態。"),
("納入 MVP","新增／編輯／刪除行程和任務；日期、時段、全天、顏色、重複、提醒設定、重要星號。"),
("納入 MVP","所在地天氣、每日笑話、響應式版面、瀏覽器本機資料保存。"),
("暫不納入","登入、雲端同步、多裝置資料合併、分享協作、真正的系統通知。"),
("待決策","重複項目單次修改規則、提醒提前時間、資料備份、天氣地點保存方式。")],[1.25,5.6])

para("3 主要使用流程","Heading 1")
para("檢視每日安排","Heading 2")
for x in ["開啟首頁，查看日期、笑話、天氣和當日摘要。","在時間軸查看有時段的行程，在未排時間清單查看待辦。","透過核取按鈕標記完成，使用日期箭頭、月曆或「回到今天」切換日期。"]: number(x)
para("建立與維護項目","Heading 2")
for x in ["選擇行程或任務，輸入名稱與日期。","行程可設定開始／結束時間或全天；任務可不排時段。","可設定重複、開始時提醒、標籤顏色與重要星號。","從項目操作選單編輯或刪除；重要項目以黃色星號標示。"]: number(x)
para("取得天氣","Heading 2")
para("瀏覽器請求位置權限，取得座標後呼叫 Open-Meteo Forecast API，顯示目前天氣、目前氣溫、當日高低溫與降雨機率。切換日期時，若服務提供該日預報則更新顯示。")

para("4 功能需求","Heading 1")
requirements=[
("FR-01 首頁","顯示所選日期、每日笑話、天氣卡、行程數、未完成任務數與完成率；預設為裝置當地日期。"),
("FR-02 日期導覽","可前後切換、回到今天或月曆選日；切換後更新日期內容、摘要與笑話。"),
("FR-03 時間軸","行程依開始時間排序，顯示名稱、時段、顏色、提醒／重複狀態及完成控制。"),
("FR-04 未排程任務","顯示所選日期無時段任務，每項可完成、編輯或刪除。"),
("FR-05 建立項目","名稱與日期必填；行程支援時段、全天和顏色；任務可不設定時段。"),
("FR-06 編輯／刪除","可編輯既有內容；刪除前確認；編輯保留原完成狀態。"),
("FR-07 重複","提供不重複、每天、平日、每週。重複項目的單次例外修改方式待決策。"),
("FR-08 提醒","可設定開始時提醒。原型只保存設定並呈現狀態，沒有排程系統通知。"),
("FR-09 重要事項","行程與任務均可設定重要；重要時名稱旁顯示黃色星號，可在編輯時取消。"),
("FR-10 完成狀態","行程與任務可切換完成；完成樣式與每日完成率同步更新。"),
("FR-11 天氣","經使用者授權定位後查詢天氣；定位或網路失敗時顯示錯誤及重試操作。"),
("FR-12 每日笑話","依所選日期顯示固定笑話；同日重新載入內容不變，不同日期可輪替。"),
("FR-13 本機保存","使用 localStorage；資料保留於相同瀏覽器來源，清除瀏覽器資料可能造成遺失。")]
for ident,desc in requirements:
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(3)
    font(p.add_run(ident+"  "),9.1,True,PURPLE); font(p.add_run(desc),9.1)

para("5 資料欄位","Heading 1")
add_table(["欄位","適用","規則"],[
("id","全部","唯一識別碼，由系統產生。"),("title","全部","必填，最多 70 字。"),("type","全部","event 或 task。"),("date","全部","必填，格式 YYYY-MM-DD。"),("start / end","行程","留空代表全天；有時間時結束須晚於開始。"),("color","行程","亮紫、珊瑚橘、青藍、桃紅。"),("repeat","全部","none、daily、weekdays、weekly。"),("remind","全部","布林設定；原型不觸發系統通知。"),("important","全部","布林設定；true 時顯示星號。"),("done","全部","布林完成狀態。")],[1.15,1.15,4.55])

para("6 介面與視覺需求","Heading 1")
bullet("桌面版使用左側導覽與月曆、中央時間軸、右側未排程任務及週概覽。")
bullet("手機版改為單欄，保留新增、完成、日期切換和天氣操作。")
bullet("背景插圖包含太陽、星星、雲朵和花朵，低透明度、不攔截點擊；卡片輕微透出背景。")
bullet("純色高對比主題，不使用漸層；重要狀態除顏色外須有星號符號。")
bullet("表單提供必填與時間錯誤提示；互動控件需有清楚標籤。")

para("7 驗收條件","Heading 1")
for x in [
"新增行程後，項目出現在正確日期和時間位置，且可重新開啟編輯。",
"新增無時段任務後，項目顯示於未排程清單，並可完成、編輯及刪除。",
"開始時間不早於結束時間時，阻止儲存並提供修正提示。",
"切換重要狀態儲存後，星號正確顯示或移除；重新載入後狀態保留。",
"完成或取消完成項目後，樣式、待辦數與完成率更新。",
"拒絕定位或天氣服務失敗時，其他功能仍可用，且天氣區提供重試。",
"切換日期後，日期內容、笑話及可用預報日期一致。",
"桌面與手機尺寸下，主要操作無水平溢出，文字與按鈕清楚可辨。"]: number(x)

para("8 非功能需求與限制","Heading 1")
add_table(["面向","需求與限制"],[
("易用性","新增與編輯共用表單；完成可直接切換；刪除須二次確認。"),
("響應式","支援常見桌面與手機瀏覽器尺寸。"),
("隱私","定位需經瀏覽器授權；座標會傳給天氣服務查詢天氣。原型不建立帳號或伺服器個人資料。"),
("可靠性","天氣依賴位置權限與網路；行程依賴瀏覽器本機儲存。"),
("提醒","設定目前只記錄偏好，正式版須設計通知權限、排程及提前時間。"),
("同步","目前沒有帳號、雲端備份或跨裝置同步。")],[1.25,5.6])

para("9 待決策事項","Heading 1")
add_table(["議題","建議方向"],[
("重複項目","定義僅此一次或整個系列的編輯／刪除，並支援單次例外。"),
("提醒行為","決定開始前提醒選項、提前分鐘數，以及權限遭拒時的替代方式。"),
("同步與備份","確認是否需要登入、雲端同步及匯出／匯入。"),
("天氣地點","確認每次定位、保存地點或手動搜尋城市的方式。"),
("有時段任務","若任務支援排時段，定義其進入時間軸及衝突規則。"),
("歷史日期天氣","超出預報日期範圍時，決定是否提供歷史資料或顯示無預報。")],[1.55,5.3])

para("附錄 原型行為基準","Heading 1")
para("原型包含日期導覽、重複規則、重要星號、行程／任務新增編輯刪除、完成狀態、瀏覽器定位天氣、每日笑話及純色插圖背景。正式產品仍需補齊通知、重複系列例外、資料同步及備份策略。")

doc.core_properties.title="個人行程與任務管理系統 產品需求文件"
doc.core_properties.subject="日序 DAYFLOW PRD 版本 1.0"
doc.core_properties.author="Product Team"
doc.save(OUT)
print(OUT)

