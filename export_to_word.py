# -*- coding: utf-8 -*-
"""
將肝癌用藥存活分析結果輸出為專業 Word 報告檔 (.docx)
包含：標題、描述性統計表、Group 1 & 2 KM 壽命表、嵌入存活曲線圖、Log-Rank 逐步計算表、Cox 風險模型與生統建議。
"""

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import os

def set_cell_background(cell, fill_hex):
    """設定儲存格背景顏色"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """設定儲存格邊距 (單位: dxa, 20 dxa = 1 pt)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def create_styled_table(doc, headers, data, col_widths=None):
    """建立具有專業美觀樣式的表格"""
    table = doc.add_table(rows=len(data) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # 標題列
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "1F497D") # 沉穩深藍
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=140, right=140)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255) # 白字
            run.font.name = 'Microsoft JhengHei'
            run.font.size = Pt(9.5)
            
    # 資料列
    for r_idx, row in enumerate(data):
        row_cells = table.rows[r_idx + 1].cells
        bg_color = "F2F5F8" if r_idx % 2 == 1 else "FFFFFF" # 交替底色
        for c_idx, val in enumerate(row):
            row_cells[c_idx].text = str(val)
            set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=90, bottom=90, left=130, right=130)
            p = row_cells[c_idx].paragraphs[0]
            # 數值置中，文字置中
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.name = 'Microsoft JhengHei'
                run.font.size = Pt(9.0)
                run.font.color.rgb = RGBColor(50, 50, 50)
                
    # 設定欄寬
    if col_widths:
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = Inches(width)
                
    doc.add_paragraph() # 間隔空行
    return table

def generate_word_report():
    doc = docx.Document()
    
    # 版面設定 (標準 A4, 邊界 2 cm)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    # 設定預設字型
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Microsoft JhengHei'
    font.size = Pt(11)
    font.color.rgb = RGBColor(40, 40, 40)
    
    # 1. 主標題
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("肝癌用藥存活分析統計報告\n(Kaplan-Meier Survival Analysis Report)")
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(31, 73, 125)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("課程：生物統計學實習  |  分析模型：Kaplan-Meier 乘積極限法 & Log-Rank 檢定")
    r_sub.font.size = Pt(10)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(120, 120, 120)
    
    doc.add_paragraph() # 空行
    
    # 2. 研究概況與資料摘要
    h1 = doc.add_heading("一、 研究背景與資料概況 (Study Overview)", level=1)
    h1.style.font.color.rgb = RGBColor(31, 73, 125)
    
    p1 = doc.add_paragraph(
        "本研究評估肝癌用藥試驗（Survival test example）之病患存活情形，"
        "資料共收錄 12 位病患，隨機分配至 Group 1 與 Group 2 兩組（每組各 6 位受試者）。"
        "追蹤時間以「月份 (Follow-up Months)」記錄，臨床事件定義為死亡（Status: 1=死亡事件, 0=設限 Censored）。"
    )
    p1.paragraph_format.line_spacing = 1.2
    
    summary_headers = ["分析指標 (Metric)", "Group 1", "Group 2", "全體受試者 (Overall)"]
    summary_data = [
        ["受試者人數 (Total N)", "6 人", "6 人", "12 人"],
        ["死亡事件人數 (Events, d)", "3 人 (50.0%)", "6 人 (100.0%)", "9 人 (75.0%)"],
        ["設限人數 (Censored, c)", "3 人 (50.0%)", "0 人 (0.0%)", "3 人 (25.0%)"],
        ["平均追蹤時間 (Mean ± SD)", "6.03 ± 1.13 個月", "5.53 ± 0.99 個月", "5.78 ± 1.05 個月"],
        ["追蹤月份範圍 (Min ~ Max)", "4.5 ~ 7.2 個月", "4.3 ~ 7.0 個月", "4.3 ~ 7.2 個月"],
        ["中位存活時間 (Median Survival)", "6.5 個月", "5.5 個月", "5.5 個月"],
        ["Log-Rank 檢定 (Mantel-Cox)", "—", "—", "Chi-sq = 2.304, df = 1, p = 0.129"],
        ["Cox 風險比 HR (Group 2 vs 1)", "1.00 (基準對照)", "2.81 (95% CI: 0.69 ~ 11.46)", "Cox p = 0.150"]
    ]
    create_styled_table(doc, summary_headers, summary_data, [2.4, 1.4, 1.6, 2.0])
    
    # 3. Kaplan-Meier 存活率估計表
    h2 = doc.add_heading("二、 Kaplan-Meier 存活率估計表 (Life Tables)", level=1)
    h2.style.font.color.rgb = RGBColor(31, 73, 125)
    
    p2 = doc.add_paragraph(
        "Kaplan-Meier (KM) 乘積極限法透過條件存活機率之連乘計算累積存活率 S(t)。"
        "各時間點之標準誤 (SE) 採用 Greenwood 公式估算，並提供 95% 信賴區間："
    )
    p2.paragraph_format.line_spacing = 1.2
    
    # Group 1 Table
    doc.add_heading("1. Group 1 壽命表 (Group 1 Life Table)", level=2)
    g1_headers = ["時間 t (月)", "期初在險數 (n)", "事件數 (d)", "設限數 (c)", "區間存活率 (1-d/n)", "累積存活率 S(t)", "Greenwood SE", "95% 信賴區間"]
    g1_data = [
        ["0.0", "6", "0", "0", "1.0000", "1.0000 (100%)", "0.0000", "1.0000 ~ 1.0000"],
        ["4.5", "6", "1", "0", "0.8333", "0.8333 (83.3%)", "0.1521", "0.2731 ~ 0.9747"],
        ["5.0", "5", "1", "0", "0.8000", "0.6667 (66.7%)", "0.1925", "0.1946 ~ 0.9044"],
        ["6.0", "4", "0", "1", "1.0000", "0.6667 (66.7%)", "0.1925", "0.1946 ~ 0.9044"],
        ["6.5", "3", "1", "0", "0.6667", "0.4444 (44.4%)", "0.2222", "0.0662 ~ 0.7849"],
        ["7.0", "2", "0", "1", "1.0000", "0.4444 (44.4%)", "0.2222", "0.0662 ~ 0.7849"],
        ["7.2", "1", "0", "1", "1.0000", "0.4444 (44.4%)", "0.2222", "0.0662 ~ 0.7849"]
    ]
    create_styled_table(doc, g1_headers, g1_data, [0.8, 1.0, 0.8, 0.8, 1.1, 1.2, 1.0, 1.3])
    
    p_g1_note = doc.add_paragraph("★ Group 1 中位存活時間 (Median Survival Time) 為 6.5 個月（於 6.5 個月時累積存活率首度降至 50% 以下）。")
    p_g1_note.runs[0].font.bold = True
    p_g1_note.runs[0].font.color.rgb = RGBColor(31, 73, 125)
    
    # Group 2 Table
    doc.add_heading("2. Group 2 壽命表 (Group 2 Life Table)", level=2)
    g2_headers = ["時間 t (月)", "期初在險數 (n)", "事件數 (d)", "設限數 (c)", "區間存活率 (1-d/n)", "累積存活率 S(t)", "Greenwood SE", "95% 信賴區間"]
    g2_data = [
        ["0.0", "6", "0", "0", "1.0000", "1.0000 (100%)", "0.0000", "1.0000 ~ 1.0000"],
        ["4.3", "6", "1", "0", "0.8333", "0.8333 (83.3%)", "0.1521", "0.2731 ~ 0.9747"],
        ["5.0", "5", "2", "0", "0.6000", "0.5000 (50.0%)", "0.2041", "0.1109 ~ 0.8037"],
        ["5.5", "3", "1", "0", "0.6667", "0.3333 (33.3%)", "0.1925", "0.0461 ~ 0.6756"],
        ["6.4", "2", "1", "0", "0.5000", "0.1667 (16.7%)", "0.1521", "0.0077 ~ 0.5168"],
        ["7.0", "1", "1", "0", "0.0000", "0.0000 (0.0%)", "0.0000", "0.0000 ~ 0.0000"]
    ]
    create_styled_table(doc, g2_headers, g2_data, [0.8, 1.0, 0.8, 0.8, 1.1, 1.2, 1.0, 1.3])
    
    p_g2_note = doc.add_paragraph("★ Group 2 中位存活時間 (Median Survival Time) 為 5.5 個月（於 5.0 個月時剛好為 50%，5.5 個月時降至 33.33%）。")
    p_g2_note.runs[0].font.bold = True
    p_g2_note.runs[0].font.color.rgb = RGBColor(192, 0, 0)
    
    # 4. 存活曲線圖
    h3 = doc.add_heading("三、 Kaplan-Meier 存活曲線圖 (Survival Curves)", level=1)
    h3.style.font.color.rgb = RGBColor(31, 73, 125)
    
    p3 = doc.add_paragraph(
        "下圖呈現兩組之存活階梯曲線。藍線代表 Group 1，紅線代表 Group 2。"
        "曲線附帶 95% 信賴區間帶、設限符號（+ 標記於 6.0、7.0、7.2 個月），"
        "圖表下方附有受試在險人數表（Number at Risk Table），符合臨床醫學期刊發布規範："
    )
    p3.paragraph_format.line_spacing = 1.2
    
    # 插入圖片
    img_path = 'km_survival_curve.png'
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=Inches(6.2))
        
        p_caption = doc.add_paragraph()
        p_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_cap = p_caption.add_run("圖 1：肝癌用藥存活率曲線圖（附 95% 信賴區間與在險人數表）")
        r_cap.font.size = Pt(9.5)
        r_cap.font.bold = True
        r_cap.font.color.rgb = RGBColor(80, 80, 80)
        
    doc.add_paragraph()
    
    # 5. Log-Rank 檢定
    h4 = doc.add_heading("四、 假設檢定：Log-Rank 檢定 (Mantel-Cox Test)", level=1)
    h4.style.font.color.rgb = RGBColor(31, 73, 125)
    
    p4 = doc.add_paragraph(
        "Log-Rank 檢定用於檢定兩組母體存活曲線是否相同。\n"
        "• 虛無假設 (H0)：兩組之母體存活曲線相同，即 S1(t) = S2(t)。\n"
        "• 對立假設 (H1)：兩組之母體存活曲線不相同，即 S1(t) ≠ S2(t)。\n\n"
        "在每個死亡事件發生時間點 t，依據超幾何分佈計算期望事件數 E1i 與變異數 Vari："
    )
    p4.paragraph_format.line_spacing = 1.2
    
    lr_headers = ["事件月份 t", "Group 1 在險 (n1)", "Group 1 觀察 (O1)", "Group 1 期望 (E1)", "Group 2 在險 (n2)", "Group 2 觀察 (O2)", "Group 2 期望 (E2)", "合計在險 (N)", "合計事件 (D)", "超幾何變異數 (Var)"]
    lr_data = [
        ["4.3", "6", "0", "0.5000", "6", "1", "0.5000", "12", "1", "0.2500"],
        ["4.5", "6", "1", "0.5455", "5", "0", "0.4545", "11", "1", "0.2479"],
        ["5.0", "5", "1", "1.5000", "5", "2", "1.5000", "10", "3", "0.5833"],
        ["5.5", "4", "0", "0.5714", "3", "1", "0.4286", "7", "1", "0.2449"],
        ["6.4", "3", "0", "0.6000", "2", "1", "0.4000", "5", "1", "0.2400"],
        ["6.5", "3", "1", "0.7500", "1", "0", "0.2500", "4", "1", "0.1875"],
        ["7.0", "2", "0", "0.6667", "1", "1", "0.3333", "3", "1", "0.2222"],
        ["總計 (Sum)", "—", "O1 = 3", "E1 = 5.1336", "—", "O2 = 6", "E2 = 3.8664", "—", "9", "Var = 1.9758"]
    ]
    create_styled_table(doc, lr_headers, lr_data, [0.7, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7, 0.8])
    
    p_formula = doc.add_paragraph()
    p_formula.paragraph_format.line_spacing = 1.3
    r_form = p_formula.add_run(
        "【卡方統計量計算式】\n"
        "Chi-square (χ²) = (O1 - E1)² / Var\n"
        "               = (3 - 5.1336)² / 1.9758\n"
        "               = (-2.1336)² / 1.9758\n"
        "               = 4.5520 / 1.9758 = 2.3040\n\n"
        "• 自由度 (df) = 1\n"
        "• 檢定 p 值 (p-value) = 0.1291"
    )
    r_form.font.name = 'Consolas'
    r_form.font.size = Pt(10)
    r_form.font.color.rgb = RGBColor(20, 20, 20)
    
    p_concl = doc.add_paragraph(
        "統計結論：在顯著水準 α = 0.05 下，Log-Rank 檢定統計量 χ² = 2.304，p = 0.1291 > 0.05，"
        "未達到統計上之顯著水準，因此「不拒絕虛無假設 (Fail to reject H0)」。"
        "現有樣本證據尚不足以斷定兩組之存活率存在顯著差異。"
    )
    p_concl.runs[0].font.bold = True
    
    # 6. Cox 比例風險模型與結論
    h5 = doc.add_heading("五、 Cox 比例風險模型與生物統計學結論", level=1)
    h5.style.font.color.rgb = RGBColor(31, 73, 125)
    
    p5 = doc.add_paragraph(
        "進一步建立單變量 Cox 比例風險迴歸模型（Cox Proportional Hazards Model），以 Group 1 為對照基準：\n"
        "• 風險比 Hazard Ratio (HR)：2.81\n"
        "• 95% 信賴區間 (95% CI)：0.69 ~ 11.46\n"
        "• Wald 檢定 p-value：0.1499\n\n"
        "【綜合生統分析與實習報告建議】\n"
        "1. 臨床趨勢 (Clinical Trend)：Group 1 的中位存活時間為 6.5 個月，且於追蹤截止時仍有 50%（3位）患者存活；"
        "相較之下，Group 2 的中位存活時間為 5.5 個月，死亡率達 100%。點估計顯示 Group 2 死亡風險約為 Group 1 的 2.81 倍。\n"
        "2. 統計檢定力 (Statistical Power)：雖然臨床曲線上呈現明顯差距，但由於總樣本數僅 12 位（每組 6 位），"
        "微小的樣本量導致統計檢定力嚴重不足（易發生第二型錯誤 Type II Error）。\n"
        "3. 實務改善建議：建議在後續臨床試驗或報告結論中註明：若欲達到 80% 檢定力（Power = 0.80, α = 0.05），"
        "應擴大試驗收案規模以驗證藥物之顯著療效。"
    )
    p5.paragraph_format.line_spacing = 1.3
    
    # 儲存檔案
    word_filename = "肝癌用藥存活分析報告_Survival_Analysis_Report.docx"
    doc.save(word_filename)
    print(f"[成功] Word 報告已儲存為：{word_filename}")
    
    # 同步儲存英檔名備用
    doc.save("Survival_Analysis_Report.docx")
    print("[成功] 同步備份為：Survival_Analysis_Report.docx")

if __name__ == '__main__':
    generate_word_report()
