import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def generate_modul_docx():
    doc = docx.Document()
    
    # Page setup
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
        # Differentiate first page (cover page has no header/footer)
        section.different_first_page_header_footer = True
        
        # Subsequent pages header
        header = section.header
        p_hd = header.paragraphs[0]
        p_hd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_hd.text = "Modul Ebook Lengkap: Meta Ads & AI Agent Automation"
        p_hd.runs[0].font.name = 'Calibri'
        p_hd.runs[0].font.size = Pt(8.5)
        p_hd.runs[0].font.color.rgb = RGBColor(148, 163, 184)
        
        # Subsequent pages footer with 2-column table: Title left, Page numbers right
        footer = section.footer
        p_ft_old = footer.paragraphs[0]
        p_ft_old.text = "" # clear default text
        
        ft_tbl = footer.add_table(rows=1, cols=2, width=Inches(6.9))
        ft_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Subtle top border line above footer
        ft_tblPr = ft_tbl._tbl.tblPr
        ft_borders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>
                <w:left w:val="none"/>
                <w:bottom w:val="none"/>
                <w:right w:val="none"/>
                <w:insideH w:val="none"/>
                <w:insideV w:val="none"/>
            </w:tblBorders>
        ''')
        ft_tblPr.append(ft_borders)
        
        # Left cell: Title & Author
        c_left = ft_tbl.cell(0, 0)
        c_left.width = Inches(4.5)
        p_l = c_left.paragraphs[0]
        p_l.paragraph_format.space_before = Pt(4)
        p_l.paragraph_format.space_after = Pt(0)
        r_l = p_l.add_run("Modul Meta Ads & AI Agent Automation | Muhammad Syahmi (syahmiaof)")
        r_l.font.name = 'Calibri'
        r_l.font.size = Pt(8.5)
        r_l.font.color.rgb = RGBColor(148, 163, 184)
        
        # Right cell: Dynamic Page Numbering (Muka Surat X daripada Y)
        c_right = ft_tbl.cell(0, 1)
        c_right.width = Inches(2.4)
        p_r = c_right.paragraphs[0]
        p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_r.paragraph_format.space_before = Pt(4)
        p_r.paragraph_format.space_after = Pt(0)
        
        r_pg1 = p_r.add_run("Muka Surat ")
        r_pg1.font.name = 'Calibri'
        r_pg1.font.size = Pt(8.5)
        r_pg1.font.color.rgb = RGBColor(148, 163, 184)
        
        # Add PAGE field
        run_fld_pg = p_r.add_run()
        fld_pg_1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
        fld_pg_instr = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
        fld_pg_sep = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
        fld_pg_t = parse_xml(r'<w:t %s>1</w:t>' % nsdecls('w'))
        fld_pg_end = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
        run_fld_pg._r.append(fld_pg_1)
        run_fld_pg._r.append(fld_pg_instr)
        run_fld_pg._r.append(fld_pg_sep)
        run_fld_pg._r.append(fld_pg_t)
        run_fld_pg._r.append(fld_pg_end)
        run_fld_pg.font.name = 'Calibri'
        run_fld_pg.font.size = Pt(8.5)
        run_fld_pg.font.color.rgb = RGBColor(148, 163, 184)
        
        r_pg2 = p_r.add_run(" daripada ")
        r_pg2.font.name = 'Calibri'
        r_pg2.font.size = Pt(8.5)
        r_pg2.font.color.rgb = RGBColor(148, 163, 184)
        
        # Add NUMPAGES field
        run_fld_np = p_r.add_run()
        fld_np_1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
        fld_np_instr = parse_xml(r'<w:instrText %s xml:space="preserve"> NUMPAGES </w:instrText>' % nsdecls('w'))
        fld_np_sep = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
        fld_np_t = parse_xml(r'<w:t %s>1</w:t>' % nsdecls('w'))
        fld_np_end = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
        run_fld_np._r.append(fld_np_1)
        run_fld_np._r.append(fld_np_instr)
        run_fld_np._r.append(fld_np_sep)
        run_fld_np._r.append(fld_np_t)
        run_fld_np._r.append(fld_np_end)
        run_fld_np.font.name = 'Calibri'
        run_fld_np.font.size = Pt(8.5)
        run_fld_np.font.color.rgb = RGBColor(148, 163, 184)

    # Style Helpers
    def set_cell_bg(cell, fill_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        tcPr.append(shd)

    def set_cell_padding(cell, top=140, bottom=140, left=180, right=180):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def set_callout_border(cell, color_hex="0668E1", size="32"):
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="none"/>
                <w:left w:val="single" w:sz="{size}" w:space="0" w:color="{color_hex}"/>
                <w:bottom w:val="none"/>
                <w:right w:val="none"/>
            </w:tcBorders>
        ''')
        tcPr.append(borders)

    def set_card_border(cell, color_hex="E2E8F0", size="8"):
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="{size}" w:space="0" w:color="{color_hex}"/>
                <w:left w:val="single" w:sz="{size}" w:space="0" w:color="{color_hex}"/>
                <w:bottom w:val="single" w:sz="{size}" w:space="0" w:color="{color_hex}"/>
                <w:right w:val="single" w:sz="{size}" w:space="0" w:color="{color_hex}"/>
            </w:tcBorders>
        ''')
        tcPr.append(borders)

    def set_clean_table_borders(table, border_color="E2E8F0"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>
                <w:bottom w:val="single" w:sz="12" w:space="0" w:color="{border_color}"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>
                <w:insideV w:val="none"/>
                <w:left w:val="none"/>
                <w:right w:val="none"/>
            </w:tblBorders>
        ''')
        tblPr.append(borders)

    def add_chapter_title(num, title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        
        r1 = p.add_run(f"BAB {num}: ")
        r1.font.name = 'Calibri'
        r1.font.size = Pt(16)
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(6, 104, 225) # Meta Blue
        
        r2 = p.add_run(title)
        r2.font.name = 'Calibri'
        r2.font.size = Pt(16)
        r2.font.bold = True
        r2.font.color.rgb = RGBColor(15, 23, 42) # Slate 900
        return p

    def add_subheading(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        
        r = p.add_run(title)
        r.font.name = 'Calibri'
        r.font.size = Pt(12.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def add_callout(paragraphs_list, title=None, border_color="0668E1", bg_color="F0F7FF", title_color=RGBColor(6, 104, 225)):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.9)
        set_cell_bg(cell, bg_color)
        set_callout_border(cell, border_color, size="32")
        set_cell_padding(cell, top=140, bottom=140, left=180, right=180)
        
        first = True
        if title:
            p = cell.paragraphs[0]
            first = False
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(title)
            r.bold = True
            r.font.name = 'Calibri'
            r.font.size = Pt(10.5)
            r.font.color.rgb = title_color
            
        for text in paragraphs_list:
            if first:
                p = cell.paragraphs[0]
                first = False
            else:
                p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(text)
            r.font.name = 'Calibri'
            r.font.size = Pt(10)
            r.font.color.rgb = RGBColor(51, 65, 85)
            
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(4)

    def add_code_block(code_text, title=None):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.9)
        set_cell_bg(cell, "0F172A") # Dark slate
        set_card_border(cell, "1E293B", size="8")
        set_cell_padding(cell, top=140, bottom=140, left=180, right=180)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        
        if title:
            r_title = p.add_run(f"# {title}\n")
            r_title.font.name = 'Consolas'
            r_title.font.size = Pt(9.5)
            r_title.font.bold = True
            r_title.font.color.rgb = RGBColor(52, 211, 153)
            
        r_code = p.add_run(code_text)
        r_code.font.name = 'Consolas'
        r_code.font.size = Pt(8.8)
        r_code.font.color.rgb = RGBColor(226, 232, 240)
        
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(4)

    # ==========================================
    # 1. COVER / HERO SECTION (PAGE 1)
    # ==========================================
    p_badge = doc.add_paragraph()
    p_badge.paragraph_format.space_before = Pt(10)
    p_badge.paragraph_format.space_after = Pt(4)
    r_badge = p_badge.add_run("EDISI TERKINI 2026 • PANDUAN PRAKTIKAL INDUSTRI")
    r_badge.font.name = 'Calibri'
    r_badge.font.size = Pt(9.5)
    r_badge.font.bold = True
    r_badge.font.color.rgb = RGBColor(124, 58, 237) # Violet

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(4)
    r_title1 = p_title.add_run("MODUL LENGKAP: META ADS & API AUTOMATION")
    r_title1.font.name = 'Calibri'
    r_title1.font.size = Pt(23)
    r_title1.font.bold = True
    r_title1.font.color.rgb = RGBColor(6, 104, 225) # Meta Blue

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(10)
    r_sub = p_sub.add_run("Dari Butang Klik Hingga Token Kekal & Integrasi AI Agent Melalui MCP: Panduan Praktikal Untuk Orang Bukan Teknikal")
    r_sub.font.name = 'Calibri'
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    # Insert Image if available
    img_path = os.path.join(r"C:\Users\USER\Desktop\modulmetaads", "hero-bg.png")
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.paragraph_format.space_before = Pt(2)
        p_img.paragraph_format.space_after = Pt(10)
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_picture(img_path, width=Inches(6.6))

    # Metadata Card Table
    meta_tbl = doc.add_table(rows=2, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_tbl.autofit = False
    
    meta_data = [
        [("Penulis / Builder", "Muhammad Syahmi (syahmiaof)"), ("Organisasi", "Pasukan Orkestrasi Ejen AI DFK INC")],
        [("Laman Web Modul", "https://modul-meta-ads.syahmiaof.my"), ("Protokol & Versi", "Meta Graph API v21.0 & Antigravity MCP")]
    ]
    for r_idx, row in enumerate(meta_data):
        for c_idx, (k, v) in enumerate(row):
            cell = meta_tbl.cell(r_idx, c_idx)
            cell.width = Inches(3.4)
            set_cell_bg(cell, "F8FAFC")
            set_card_border(cell, "E2E8F0", "6")
            set_cell_padding(cell, top=80, bottom=80, left=120, right=120)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r_k = p.add_run(f"{k}: ")
            r_k.font.bold = True
            r_k.font.size = Pt(9)
            r_k.font.color.rgb = RGBColor(100, 116, 139)
            r_v = p.add_run(v)
            r_v.font.size = Pt(9)
            r_v.font.color.rgb = RGBColor(30, 41, 59)

    # ==========================================
    # 2. TABLE OF CONTENTS (PAGE 2)
    # ==========================================
    doc.add_page_break()

    p_toc_head = doc.add_paragraph()
    p_toc_head.paragraph_format.space_before = Pt(14)
    p_toc_head.paragraph_format.space_after = Pt(4)
    r_toc_head = p_toc_head.add_run("Jadual Kandungan Modul (Table of Contents)")
    r_toc_head.font.name = 'Calibri'
    r_toc_head.font.size = Pt(18)
    r_toc_head.font.bold = True
    r_toc_head.font.color.rgb = RGBColor(15, 23, 42)

    p_toc_desc = doc.add_paragraph()
    p_toc_desc.paragraph_format.space_before = Pt(0)
    p_toc_desc.paragraph_format.space_after = Pt(10)
    r_toc_desc = p_toc_desc.add_run("Setiap bab disusun pada halaman tersendiri dengan rujukan nombor muka surat yang tepat:")
    r_toc_desc.font.size = Pt(10)
    r_toc_desc.font.color.rgb = RGBColor(100, 116, 139)

    # Chapters & explicit page targets (Every chapter starts on fresh page)
    toc_items = [
        ("Prakata", "Prakata & Panduan Cara Guna Ebook Ini", 3),
        ("Bab 1", "Kamus Istilah 'Orang Awam' & Analogi Pejabat Korporat", 4),
        ("Bab 2", "Membina Facebook Page & Mencari 'Page ID'", 5),
        ("Bab 3", "Menyediakan Meta Business Suite & Akaun Iklan (Ad Account ID)", 6),
        ("Bab 4", "Daftar Meta Developer App (Laluan Rasmi API) & App Secret", 7),
        ("Bab 5", "System User & Token Akses Kekal (Never-Expiring Access Token)", 8),
        ("Bab 6", "Pengujian Endpoint API Menggunakan Graph API Explorer", 9),
        ("Bab 7", "Dapatkan Instant Form ID Untuk Kempen Prospek (Lead Generation)", 10),
        ("Bab 8", "Susun Semua ID & Token Ke Dalam Fail Tetapan .env", 11),
        ("Bab 9", "Integrasi Antigravity & Protokol MCP (Orkestrasi Ejen AI Autonomi)", 12),
        ("Bab 10", "Penyelesaian Masalah / Troubleshooting & Jadual Kod Ralat Meta API", 14),
        ("Semakan", "Senarai Semak Akhir (Checklist Sebelum Present / Terbit)", 15)
    ]

    toc_tbl = doc.add_table(rows=len(toc_items)+1, cols=3)
    toc_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_clean_table_borders(toc_tbl, "CBD5E1")

    # Header Row
    th_labels = ["Bahagian / Bab", "Tajuk Kandungan Modul", "Muka Surat"]
    th_widths = [Inches(1.3), Inches(4.3), Inches(1.3)]
    for c_i, (h_txt, w) in enumerate(zip(th_labels, th_widths)):
        cell = toc_tbl.cell(0, c_i)
        cell.width = w
        set_cell_bg(cell, "0F172A")
        set_cell_padding(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        if c_i == 2:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(h_txt)
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for idx, (b_label, b_title, b_pg) in enumerate(toc_items):
        cell_lbl = toc_tbl.cell(idx+1, 0)
        cell_lbl.width = Inches(1.3)
        set_cell_padding(cell_lbl, top=65, bottom=65, left=100, right=60)
        p0 = cell_lbl.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(b_label)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = RGBColor(6, 104, 225)

        cell_title = toc_tbl.cell(idx+1, 1)
        cell_title.width = Inches(4.3)
        set_cell_padding(cell_title, top=65, bottom=65, left=60, right=80)
        p1 = cell_title.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(b_title)
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = RGBColor(30, 41, 59)

        cell_pg = toc_tbl.cell(idx+1, 2)
        cell_pg.width = Inches(1.3)
        set_cell_padding(cell_pg, top=65, bottom=65, left=60, right=100)
        p2 = cell_pg.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(f"Halaman {b_pg}")
        r2.font.bold = True
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = RGBColor(71, 85, 105)

        if idx % 2 == 1:
            set_cell_bg(cell_lbl, "F8FAFC")
            set_cell_bg(cell_title, "F8FAFC")
            set_cell_bg(cell_pg, "F8FAFC")

    # ==========================================
    # 3. PRAKATA (PAGE 3)
    # ==========================================
    doc.add_page_break()

    p_pk = doc.add_paragraph()
    p_pk.paragraph_format.space_before = Pt(12)
    p_pk.paragraph_format.space_after = Pt(8)
    r_pk = p_pk.add_run("Prakata & Cara Guna Ebook Ini")
    r_pk.font.name = 'Calibri'
    r_pk.font.size = Pt(16)
    r_pk.font.bold = True
    r_pk.font.color.rgb = RGBColor(15, 23, 42)

    add_callout([
        "Modul ini ditulis khas untuk usahawan, pelajar, dan sesiapa sahaja yang tiada latar belakang pengaturcaraan (non-technical).",
        "Banyak tutorial di internet terlalu teknikal dan mengelirukan — menggunakan istilah seperti endpoint, OAuth, payload, webhook tanpa menerangkan di mana letaknya butang sebenar di skrin.",
        "Dalam modul ini, anda akan dipimpin langkah demi langkah (click-by-click):",
        "  • Di mana nak klik tetikus anda di portal Meta.",
        "  • Warna dan nama butang yang perlu ditekan.",
        "  • Apa maksud setiap kod/ID yang anda jumpa.",
        "  • Macam mana nak dapatkan 'Kunci Digital' (Token) yang tidak luput.",
        "  • Bagaimana AI Agent di Antigravity menggunakan MCP untuk membaca ads dan membina kempen secara autonomi.",
        "  • Macam mana nak hubungkan semuanya tanpa risiko akaun disekat atau bajet terbakar."
    ], title="MENGAPA MODUL INI DIBINA?", border_color="0668E1", bg_color="EFF6FF", title_color=RGBColor(6, 104, 225))

    # ==========================================
    # 4. BAB 1: KAMUS ISTILAH (PAGE 4)
    # ==========================================
    doc.add_page_break()
    add_chapter_title("1", "Kamus Istilah 'Orang Awam' & Analogi Pejabat")
    
    p_b1_desc = doc.add_paragraph()
    p_b1_desc.paragraph_format.space_after = Pt(8)
    r = p_b1_desc.add_run("Sebelum kita menekan sebarang butang tetapan di portal Meta, fahami terlebih dahulu istilah penting ini dengan Analogi Pejabat Korporat:")
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(51, 65, 85)

    b1_headers = ["Istilah Teknikal", "Maksud Sebenar", "Analogi Mudah", "Di Mana Anda Jumpa?"]
    b1_data = [
        ("Facebook Page", "Halaman perniagaan rasmi anda di Facebook.", "Papan Tanda Kedai anda di tepi jalan.", "facebook.com/{nama-page}"),
        ("Meta Business Portfolio", "Pusat kawalan bisnes yang memegang Page, akaun iklan, dan staf.", "Ibu Pejabat yang memegang semua cawangan.", "business.facebook.com"),
        ("Meta Developer App", "Aplikasi digital untuk benarkan sistem komputer bercakap dengan Facebook.", "Laluan Staf Belakang untuk lori kargo keluar masuk.", "developers.facebook.com"),
        ("App ID & App Secret", "Nombor pengenalan unik dan kata laluan rahsia bagi aplikasi developer anda.", "No. IC & PIN Keselamatan aplikasi anda.", "Dalam App Settings Meta Developer."),
        ("System User", "Akaun pengguna digital (bukan manusia) khas untuk robot/automasi.", "Pekerja Robotik yang bertugas khas 24 jam.", "Meta Business Settings -> System Users."),
        ("Access Token", "Kod teks panjang (mula dengan EAA...). Bertindak sebagai kunci digital.", "Kunci Akses VIP pejabat.", "Dijana daripada System User / Graph Explorer."),
        ("Page ID", "Nombor siri unik 15-16 digit mewakili Facebook Page.", "Nombor Geran Kedai anda.", "Bahagian About Page anda."),
        ("Ad Account ID", "Nombor akaun iklan (format: act_123...).", "Nombor Akaun Bank iklan.", "Meta Ads Manager / Business Settings."),
        ("MCP", "Model Context Protocol. Protokol penyambung universal.", "Soket Palam USB-C untuk otak AI.", "Konfigurasi Antigravity (mcp_config.json)."),
        ("Antigravity IDE", "Pusat arahan dan persekitaran kerja ejen AI.", "Bilik Kawalan Utama Operasi.", "Antigravity Desktop / CLI.")
    ]

    t1 = doc.add_table(rows=len(b1_data)+1, cols=4)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_clean_table_borders(t1, "CBD5E1")

    # Header Row
    for c_i, h in enumerate(b1_headers):
        cell = t1.cell(0, c_i)
        set_cell_bg(cell, "0F172A")
        set_cell_padding(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    # Data Rows
    for r_i, row in enumerate(b1_data):
        for c_i, val in enumerate(row):
            cell = t1.cell(r_i+1, c_i)
            if (r_i % 2) == 1:
                set_cell_bg(cell, "F8FAFC")
            set_cell_padding(cell, top=60, bottom=60, left=100, right=100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if c_i == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(6, 104, 225)
            elif c_i == 2:
                r.font.color.rgb = RGBColor(16, 185, 129)
            elif c_i == 3:
                r.font.name = 'Consolas'
                r.font.size = Pt(8)
                r.font.color.rgb = RGBColor(71, 85, 105)
            else:
                r.font.color.rgb = RGBColor(51, 65, 85)

    # ==========================================
    # 5. BAB 2: FACEBOOK PAGE & PAGE ID (PAGE 5)
    # ==========================================
    doc.add_page_break()
    add_chapter_title("2", "Membina Facebook Page & Mencari 'Page ID'")
    
    add_subheading("Langkah 2.1: Bina Facebook Page (Jika Belum Ada)")
    p2_1 = doc.add_paragraph()
    p2_1.paragraph_format.space_after = Pt(6)
    p2_1.add_run("1. Buka pelayar web dan layari pautan rasmi penciptaan halaman: https://www.facebook.com/pages/create\n"
                 "2. Pada panel sebelah kiri, lengkapkan maklumat asas:\n"
                 "   • Page Name: Masukkan nama syarikat (Contoh: Demo SME Sdn Bhd).\n"
                 "   • Category: Pilih kategori perniagaan yang sesuai.\n"
                 "   • Bio: Tulis 1-2 ayat ringkas mengenai bisnes anda.\n"
                 "3. Tekan butang biru 'Create Page' di bahagian bawah kiri.\n"
                 "4. Muat naik gambar profil (Logo rasmi) dan gambar cover (Banner promosi).\n"
                 "5. Tekan butang 'Save' / 'Done'.")

    add_subheading("Langkah 2.2: Kat Mana Nak Cari 'Page ID'? (Paling Penting!)")
    add_callout([
        "Ramai orang tersilap memasukkan nama Page, sedangkan sistem automasi dan API Meta hanya mengenali nombor siri ID unik.",
        "Langkah mencarinya:",
        "1. Masuk ke Facebook Page anda sebagai profil Page tersebut.",
        "2. Di menu sebelah kiri, klik tab 'About' (Tentang).",
        "3. Skrol ke bawah dan klik tab 'Page Transparency' (Ketelusan Halaman).",
        "4. Anda akan melihat kotak maklumat bertulis contoh: Page ID: 1389752850879310"
    ], title="PANDUAN MENCARI PAGE ID", border_color="F59E0B", bg_color="FFFBEB", title_color=RGBColor(217, 119, 6))

    add_callout([
        "Salin nombor Page ID ini dan simpan dalam Notepad anda. Labelkan pembolehubah ini sebagai META_PAGE_ID."
    ], title="TINDAKAN ANDA", border_color="8B5CF6", bg_color="F5F3FF", title_color=RGBColor(124, 58, 237))

    # ==========================================
    # 6. BAB 3: META BUSINESS SUITE & AKAUN IKLAN (PAGE 6)
    # ==========================================
    doc.add_page_break()
    add_chapter_title("3", "Menyediakan Meta Business Suite & Akaun Iklan")

    add_subheading("Langkah 3.1: Buka Meta Business Suite")
    p3_1 = doc.add_paragraph()
    p3_1.paragraph_format.space_after = Pt(6)
    p3_1.add_run("1. Buka https://business.facebook.com/ di pelayar web anda.\n"
                 "2. Di penjuru atas kiri, pastikan nama Business Portfolio / Perniagaan anda dipilih.\n"
                 "3. Di menu bar sisi sebelah kiri paling bawah, cari dan klik ikon gear bertulis 'Settings' atau terus layari https://business.facebook.com/settings.")

    add_subheading("Langkah 3.2: Dapatkan / Cipta 'Ad Account ID'")
    p3_2 = doc.add_paragraph()
    p3_2.paragraph_format.space_after = Pt(6)
    p3_2.add_run("Sistem AI memerlukan Ad Account ID untuk merangka dan memantau kempen iklan.\n"
                 "1. Di Business Settings, lihat menu kiri di bawah bahagian 'Accounts'.\n"
                 "2. Klik 'Ad Accounts' (Akaun Iklan).\n"
                 "3. Jika akaun sudah sedia ada: Salin nombor ID di sebelah nama akaun. Format: act_123456789...\n"
                 "4. Jika belum ada akaun: Klik butang Add -> Create a new ad account.\n"
                 "   • Timezone: Wajib pilih GMT+08:00 Asia/Kuala_Lumpur.\n"
                 "   • Currency: Pilih MYR (Ringgit Malaysia).")

    add_callout([
        "Simpan nombor akaun ini sebagai META_AD_ACCOUNT_ID. Pastikan anda menyertakan awalan 'act_' sekiranya menggunakan API Meta Ads."
    ], title="NOTA PENTING", border_color="0668E1", bg_color="F0F7FF", title_color=RGBColor(6, 104, 225))

    # ==========================================
    # 7. BAB 4: DAFTAR DEVELOPER APP (PAGE 7)
    # ==========================================
    doc.add_page_break()
    add_chapter_title("4", "Daftar Meta Developer App (Laluan Rasmi API)")
    
    p4_desc = doc.add_paragraph()
    p4_desc.paragraph_format.space_after = Pt(6)
    p4_desc.add_run("Untuk membolehkan kod komputer dan ejen AI bercakap dengan platform Meta, kita perlu mendaftarkan satu 'Aplikasi Penghubung' rasmi.")

    add_subheading("Langkah 4.1: Portal Meta for Developers")
    p4_1 = doc.add_paragraph()
    p4_1.paragraph_format.space_after = Pt(6)
    p4_1.add_run("1. Layari portal pemaju: https://developers.facebook.com/\n"
                 "2. Log masuk menggunakan akaun Facebook peribadi anda.\n"
                 "3. Klik 'Get Started' dan selesaikan pengesahan jika pertama kali mendaftar.\n"
                 "4. Pergi ke menu penjuru kanan atas bertulis 'My Apps'.")

    add_subheading("Langkah 4.2: Cipta Aplikasi Baharu")
    p4_2 = doc.add_paragraph()
    p4_2.paragraph_format.space_after = Pt(6)
    p4_2.add_run("1. Klik butang hijau/biru 'Create App'.\n"
                 "2. Pada bahagian Use Case: Pilih 'Other' -> Klik Next.\n"
                 "3. Pada bahagian App Type: Pilih 'Business' -> Klik Next.\n"
                 "4. Masukkan App Name: Contoh: Demo_SME_Automation (Jangan gunakan perkataan 'Meta' atau 'Facebook' kerana ia akan ditolak).\n"
                 "5. Pilih Business Account syarikat anda -> Klik butang 'Create App'.")

    add_subheading("Langkah 4.3: Di Mana Nak Ambil App ID & Secret?")
    p4_3 = doc.add_paragraph()
    p4_3.paragraph_format.space_after = Pt(6)
    p4_3.add_run("1. Di menu sebelah kiri dashboard aplikasi, klik App settings -> Basic.\n"
                 "2. Salin App ID -> Simpan sebagai pembolehubah META_APP_ID.\n"
                 "3. Klik butang 'Show' pada kotak tersembunyi untuk memaparkan App Secret.\n"
                 "4. Salin kod rahsia tersebut -> Simpan sebagai META_APP_SECRET.")

    add_callout([
        "AMARAN KESELAMATAN SIBER:",
        "App Secret ibarat nombor PIN kad ATM peribadi anda. Jangan sesekali berkongsi kod ini di media sosial, WhatsApp awam, atau di dalam fail GitHub awam tanpa kawalan privasi!"
    ], title="AMARAN KESELAMATAN", border_color="EF4444", bg_color="FEF2F2", title_color=RGBColor(220, 38, 38))

    # ==========================================
    # 8. BAB 5: SYSTEM USER & TOKEN AKSES KEKAL (PAGE 8)
    # ==========================================
    doc.add_page_break()
    add_chapter_title("5", "System User & Token Akses Kekal (The Holy Grail)")

    add_callout([
        "THE HOLY GRAIL: Mengapa kita wajib menggunakan System User?",
        "Token akses biasa yang dijana melalui profil manusia akan luput selepas 1 jam atau maksimum 60 hari. Apabila token luput, automasi sistem AI anda akan mati di tengah malam!",
        "Sebaliknya, token yang dijana daripada System User berstatus Admin adalah KEKAL (Never-Expiring) selagi anda tidak membatalkannya secara manual."
    ], title="KENAPA TOKEN KEKAL PENTING?", border_color="8B5CF6", bg_color="F5F3FF", title_color=RGBColor(124, 58, 237))

    add_subheading("Langkah 5.1: Cipta System User")
    p5_1 = doc.add_paragraph()
    p5_1.paragraph_format.space_after = Pt(6)
    p5_1.add_run("1. Buka Meta Business Settings: https://business.facebook.com/settings\n"
                 "2. Pada menu kiri, navigasi ke Users -> System Users.\n"
                 "3. Klik butang Add.\n"
                 "4. Masukkan nama pekerja digital: Contoh: Jebat_Bot.\n"
                 "5. Tetapkan System User Role: Wajib pilih 'Admin'.\n"
                 "6. Klik Create System User.")

    add_subheading("Langkah 5.2: Hubungkan Aset (Assign Assets) - Wajib!")
    p5_2 = doc.add_paragraph()
    p5_2.paragraph_format.space_after = Pt(6)
    p5_2.add_run("1. Klik nama bot yang baharu dicipta, kemudian tekan butang 'Assign Assets'.\n"
                 "2. Pages: Tandakan Facebook Page anda -> Aktifkan suis 'Full Control'.\n"
                 "3. Ad Accounts: Tandakan Akaun Iklan anda -> Aktifkan suis 'Manage Campaigns'.\n"
                 "4. Klik butang 'Save Changes'.")

    add_subheading("Langkah 5.3: Jana Token Akses Tanpa Had Luput (Generate Token)")
    p5_3 = doc.add_paragraph()
    p5_3.paragraph_format.space_after = Pt(6)
    p5_3.add_run("1. Klik butang 'Generate New Token'.\n"
                 "2. Select App: Pilih nama aplikasi yang didaftarkan dalam Bab 4.\n"
                 "3. Token Expiration: Pilih pilihan 'Never'.\n"
                 "4. Tandakan kebenaran (Permissions) berikut (Sangat Kritikal):\n"
                 "   • pages_show_list, pages_read_engagement, pages_manage_posts, pages_manage_engagement\n"
                 "   • ads_management, ads_read, leads_retrieval, pages_manage_ads\n"
                 "5. Klik butang Generate Token.\n"
                 "6. Salin kod panjang yang bermula dengan huruf 'EAA...' dan simpan serta-merta sebagai META_SYSTEM_USER_TOKEN.")

    add_callout([
        "PERINGATAN: Meta tidak akan memaparkan semula token ini selepas anda menutup tetingkap dialog! Sekiranya terlepas, anda perlu menjana semula token baharu."
    ], title="SIMPAN SERTA-MERTA", border_color="F59E0B", bg_color="FFFBEB", title_color=RGBColor(217, 119, 6))

    # ==========================================
    # 9. BAB 6: UJI TOKEN (API EXPLORER) (PAGE 9)
    # ==========================================
    doc.add_page_break()
    add_chapter_title("6", "Pengujian Endpoint API Menggunakan Graph API Explorer")
    
    p6_desc = doc.add_paragraph()
    p6_desc.paragraph_format.space_after = Pt(6)
    p6_desc.add_run("Sebelum memasang token ke dalam kod program sebenar, kita perlu mengesahkan sama ada token tersebut sah dan mempunyai akses penuh ke pelayan Meta.")

    add_subheading("Langkah Ujian:")
    p6_steps = doc.add_paragraph()
    p6_steps.paragraph_format.space_after = Pt(6)
    p6_steps.add_run("1. Buka alat Graph API Explorer rasmi: https://developers.facebook.com/tools/explorer/\n"
                     "2. Padamkan Access Token sementara yang ada, dan tampal (paste) Token System User (EAA...) anda.\n"
                     "3. Di bar pertanyaan URL (Kaedah GET), taip perintah berikut:")

    add_code_block("GET me?fields=name,id", title="PANGGILAN GRAPH API")

    p6_res = doc.add_paragraph()
    p6_res.paragraph_format.space_after = Pt(4)
    p6_res.add_run("4. Klik butang biru 'Submit'. Sekiranya token berfungsi dengan sempurna, anda akan menerima respon HTTP 200 OK:")

    add_code_block('{\n  "name": "Demo_SME_Automation",\n  "id": "1122334455667788"\n}', title="RESPON JSON (STATUS: 200 OK)")

    # ==========================================
    # 10. BAB 7: INSTANT FORM ID (PAGE 10)
    # ==========================================
    doc.add_page_break()
    add_chapter_title("7", "Dapatkan Instant Form ID Untuk Kempen Prospek")

    p7_desc = doc.add_paragraph()
    p7_desc.paragraph_format.space_after = Pt(6)
    p7_desc.add_run("Sekiranya perniagaan anda menggunakan iklan jenis Lead Generation (kumpul nama, nombor telefon, dan e-mel prospek), anda memerlukan nombor siri Instant Form.")

    add_subheading("Langkah Membina & Mengambil ID Borang:")
    p7_steps = doc.add_paragraph()
    p7_steps.paragraph_format.space_after = Pt(6)
    p7_steps.add_run("1. Di Meta Business Suite, buka menu All Tools -> Pilih 'Instant Forms'.\n"
                     "2. Klik butang 'Create Form' -> Pilih 'New Form'.\n"
                     "3. Namakan borang anda: Contoh: Borang_Pertanyaan_Produk.\n"
                     "4. Pada bahagian soalan, tambah ruang kenalan wajib: Contact Fields -> Phone Number.\n"
                     "5. Masukkan URL Dasar Privasi (Privacy Policy) laman web anda.\n"
                     "6. Klik butang 'Create Form' untuk menerbitkan borang.\n"
                     "7. Di jadual senarai borang, salin nombor ID borang tersebut.")

    add_callout([
        "Simpan nombor siri borang ini dan labelkan sebagai META_LEAD_FORM_ID dalam konfigurasi automasi anda."
    ], title="TINDAKAN ANDA", border_color="0668E1", bg_color="F0F7FF", title_color=RGBColor(6, 104, 225))

    # ==========================================
    # 11. BAB 8: SUSUN FAIL .ENV (PAGE 11)
    # ==========================================
    doc.add_page_break()
    add_chapter_title("8", "Susun Semua ID & Token Ke Dalam Fail .env")

    p8_desc = doc.add_paragraph()
    p8_desc.paragraph_format.space_after = Pt(6)
    p8_desc.add_run("Semua maklumat pengenalan dan kunci keselamatan yang dikumpulkan daripada Bab 2 hingga Bab 7 kini sedia dimasukkan ke dalam fail pembolehubah persekitaran (Environment Variables / .env). Fail ini bertindak seperti kabinet keselamatan khas yang menyimpan semua rahsia sistem.")

    env_content = (
        "# ==============================================================================\n"
        "# TETAPAN KREDENSIAL AUTOMATION SISTEM OPERASI AI (AI-MOS)\n"
        "# PROJEK: DEMO SME SDN BHD\n"
        "# ==============================================================================\n\n"
        "# 1. DARI BAB 4 (META DEVELOPER APP)\n"
        "META_APP_ID=1122334455667788\n"
        "META_APP_SECRET=a1b2c3d4e5f67890abcdef1234567890\n"
        "META_GRAPH_VERSION=v21.0\n\n"
        "# 2. DARI BAB 5 (SYSTEM USER NEVER-EXPIRING TOKEN)\n"
        "META_SYSTEM_USER_TOKEN=EAABcdefgh1234567890long_token_here_never_expire\n\n"
        "# 3. DARI BAB 2 (FACEBOOK PAGE ID)\n"
        "META_PAGE_ID=1389752850879310\n\n"
        "# 4. DARI BAB 3 (AD ACCOUNT ID)\n"
        "META_AD_ACCOUNT_ID=act_123456789012345\n\n"
        "# 5. DARI BAB 7 (INSTANT LEAD FORM ID)\n"
        "META_LEAD_FORM_ID=1456204426381022\n\n"
        "# 6. KAWALAN KESELAMATAN SILING BAJET (HARD-LOCK SAFETY)\n"
        "MAX_DAILY_AD_BUDGET_MYR=15.00\n"
        "DEFAULT_AD_STATUS=PAUSED"
    )
    add_code_block(env_content, title="TEMPLAT FAIL .env")

    # ==========================================
    # 12. BAB 9: INTEGRASI AI AGENT & MCP (PAGE 12 - 13)
    # ==========================================
    doc.add_page_break()
    add_chapter_title("9", "Integrasi Antigravity & MCP (AI Agent Autonomi)")

    p9_desc = doc.add_paragraph()
    p9_desc.paragraph_format.space_after = Pt(6)
    p9_desc.add_run("Ini adalah bahagian paling moden dan berkuasa dalam modul ini. Kita tidak lagi bercakap tentang skrip kod statik yang kaku, sebaliknya bagaimana Ejen AI moden berfikir, merancang, dan mengambil tindakan secara autonomi menggunakan protokol MCP (Model Context Protocol).")

    add_subheading("9.1 Kenapa Menggunakan MCP & Ejen AI?")
    
    comp_tbl = doc.add_table(rows=4, cols=2)
    comp_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_clean_table_borders(comp_tbl, "CBD5E1")

    # Headers
    cell_l_h = comp_tbl.cell(0, 0)
    set_cell_bg(cell_l_h, "FEE2E2")
    set_cell_padding(cell_l_h, top=80, bottom=80, left=120, right=120)
    p = cell_l_h.paragraphs[0]
    r = p.add_run("Cara Lama (Manual / Skrip Kod Biasa)")
    r.font.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(185, 28, 28)

    cell_r_h = comp_tbl.cell(0, 1)
    set_cell_bg(cell_r_h, "DCFCE7")
    set_cell_padding(cell_r_h, top=80, bottom=80, left=120, right=120)
    p = cell_r_h.paragraphs[0]
    r = p.add_run("Cara Moden (Orkestrasi Ejen AI + MCP)")
    r.font.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(21, 128, 61)

    comp_rows = [
        ("Perlu mengupah pengaturcara menulis beratus-ratus baris skrip kaku.", "LLM (Otak Pintar): Membaca konteks metrik & merangka copywriting secara dinamik."),
        ("Bila dokumentasi API berubah sedikit, skrip ralat dan perlu dikod semula.", "MCP (Tangan & Mata): Menghubungkan fungsi AI secara terus kepada endpoint Meta Ads."),
        ("Manusia masih terpaksa membuka Ads Manager setiap pagi untuk semak baki.", "Arahan bahasa biasa: 'Tolong audit iklan hari ni, kalau CTR teruk cadangkan ayat baru.'")
    ]
    for r_i, (c_old, c_new) in enumerate(comp_rows):
        cell_l = comp_tbl.cell(r_i+1, 0)
        cell_r = comp_tbl.cell(r_i+1, 1)
        set_cell_padding(cell_l, top=70, bottom=70, left=120, right=120)
        set_cell_padding(cell_r, top=70, bottom=70, left=120, right=120)
        if r_i % 2 == 1:
            set_cell_bg(cell_l, "FFF5F5")
            set_cell_bg(cell_r, "F0FDF4")
        p_l = cell_l.paragraphs[0]
        p_l.add_run(f"• {c_old}").font.size = Pt(9)
        p_r = cell_r.paragraphs[0]
        p_r.add_run(f"• {c_new}").font.size = Pt(9)

    add_subheading("9.2 Empat Objektif Utama & Jangkaan Hasil")
    p9_objs = doc.add_paragraph()
    p9_objs.paragraph_format.space_after = Pt(6)
    p9_objs.add_run(
        "1. Audit Masa Nyata (Real-Time): Ejen AI memanggil tool 'meta_get_insights'. Jika nilai CTR kurang daripada 1%, sistem menghantar amaran kerana audiens mula jemu (Ad Fatigue).\n"
        "2. Draf Kempen Pintar: Anda memberi arahan ringkas: 'Buat draf kempen restoran'. AI memanggil 'meta_create_campaign_draft' dengan status PAUSED bagi menjamin sifar risiko kewangan.\n"
        "3. Kawalan Bajet (Hard-Lock Safety): Ejen AI menyemak baki menggunakan 'meta_get_ad_account_details' dan menghalang sebarang bajet melebihi siling (cth: RM15/hari) tanpa kelulusan manusia.\n"
        "4. Diagnosis Ralat (Self-Healing): Sekiranya iklan ditolak (Ad Rejected), AI mentafsir kod ralat rasmi Meta dan membetulkan salinan iklan mengikut dasar pengiklanan secara automatik."
    )

    # Page break to let Architecture & Dialog start cleanly on Page 13
    doc.add_page_break()

    add_subheading("9.3 Kitaran Autonomi Latar Belakang (System Architecture)")
    arch_diagram = (
        "┌────────────────────────────────────────────────────────────────────────┐\n"
        "│ 1. PENGGUNA MEMBERI ARAHAN DALAM BAHASA MELAYU                         │\n"
        "│    \"Juru Audit, semak akaun iklan aku hari ni. Ada jalan ke tak?\"      │\n"
        "└───────────────────────────────────┬────────────────────────────────────┘\n"
        "                                    ▼\n"
        "┌────────────────────────────────────────────────────────────────────────┐\n"
        "│ 2. ANTIGRAVITY ORCHESTRATION (OTAK AI MENILAI ARAHAN)                  │\n"
        "│    - Ejen memilih Tool MCP: meta_get_ad_account_details & insights.    │\n"
        "└───────────────────────────────────┬────────────────────────────────────┘\n"
        "                                    ▼\n"
        "┌────────────────────────────────────────────────────────────────────────┐\n"
        "│ 3. META ADS MCP SERVER (JAMBATAN PROTOKOL)                            │\n"
        "│    - Mengambil Token System User daripada fail .env.                  │\n"
        "│    - Menghantar panggilan rasmi ke Meta Graph API (v21.0).             │\n"
        "└───────────────────────────────────┬────────────────────────────────────┘\n"
        "                                    ▼\n"
        "┌────────────────────────────────────────────────────────────────────────┐\n"
        "│ 4. PELAYAN META ADS (FACEBOOK SERVER)                                  │\n"
        "│    - Memulangkan data: Spend: RM0.00, Status: PAUSED, Leads: 0.        │\n"
        "└───────────────────────────────────┬────────────────────────────────────┘\n"
        "                                    ▼\n"
        "┌────────────────────────────────────────────────────────────────────────┐\n"
        "│ 5. EJEN AI MEMBUAT KESIMPULAN & LAPORAN BAHASA MANUSIA                 │\n"
        "│    \"Boss, akaun iklan selamat. Kempen PAUSED. Belanja RM0.00.\"         │\n"
        "└────────────────────────────────────────────────────────────────────────┘"
    )
    add_code_block(arch_diagram, title="CARTA ALIRAN ORKESTRASI SISTEM (ASCII ARCHITECTURE)")

    add_subheading("9.4 Senario Sebenar: Dialog Anda Bersama Ejen AI")
    add_callout([
        "Pengguna: 'Juru Audit, tolong buat audit akaun iklan aku. Kenapa rasa macam lead masuk sikit sangat minggu ni?'",
        "Pemikiran AI: Memanggil tool MCP: meta_list_campaigns & meta_get_insights...",
        "Respon AI: 'Boss, saya dah selesai audit akaun. Ini dapatan utama:',",
        "  1. Kempen SME Leads sedang aktif dengan bajet harian RM15.00/hari.",
        "  2. Frekuensi paparan telah mencecah 3.4 (audiens sasaran telah melihat iklan lebih 3 kali).",
        "  3. Kadar CTR telah jatuh ke 0.65% — tanda jelas masalah Ad Fatigue (audiens jemu).",
        "  Cadangan: Ejen Pereka telah menyediakan 2 visual kreatif baharu. Adakah bos bersetuju untuk saya kemas kini visual dalam draf iklan sekarang?"
    ], title="SIMULASI PERBUALAN EJEN PINTAR", border_color="10B981", bg_color="ECFDF5", title_color=RGBColor(5, 150, 105))

    # ==========================================
    # 13. BAB 10: TROUBLESHOOTING & ERROR CODES (PAGE 14)
    # ==========================================
    doc.add_page_break()
    add_chapter_title("10", "Penyelesaian Masalah (Error Troubleshooting)")

    p10_desc = doc.add_paragraph()
    p10_desc.paragraph_format.space_after = Pt(6)
    p10_desc.add_run("Jadual diagnostik pantas bagi menyelesaikan kod ralat lazim Meta Graph API dan sambungan pelayan MCP:")

    err_headers = ["Kod Ralat", "Mesej Lazim", "Sebab Utama", "Cara Selesai Segera"]
    err_data = [
        ("Error 190", "Invalid OAuth access token", "Token salah salin, atau menggunakan token sementara yang telah luput.", "Jana token baharu melalui System User. Pastikan memilih tempoh 'Never'."),
        ("Error 200 / 100", "Permissions error / Missing scope", "System User belum diberikan kebenaran bagi fungsi yang diminta.", "Semak Bab 5.3. Tandakan kebenaran ads_management, ads_read, dan leads_retrieval."),
        ("Error 273", "Cannot create ad for this page", "System User belum disambungkan ke aset Facebook Page atau Ad Account.", "Masuk ke Business Settings -> System Users -> Klik butang 'Assign Assets'."),
        ("Error 1487843", "Dev mode app restriction", "Meta App masih dalam status 'In Development' (belum Live).", "Gunakan tetapan muat naik 'published: false' sebelum menyambungkan pos ke iklan."),
        ("Timeout", "Tool call failed / Server unavailable", "Pelayan MCP lokal belum dihidupkan atau port komunikasi salah.", "Pastikan port fail .env tepat dan jalankan saringan 'health-check' pelayan MCP.")
    ]

    t_err = doc.add_table(rows=len(err_data)+1, cols=4)
    t_err.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_clean_table_borders(t_err, "CBD5E1")

    for c_i, h in enumerate(err_headers):
        cell = t_err.cell(0, c_i)
        set_cell_bg(cell, "7F1D1D") # Dark Red
        set_cell_padding(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for r_i, row in enumerate(err_data):
        for c_i, val in enumerate(row):
            cell = t_err.cell(r_i+1, c_i)
            if (r_i % 2) == 1:
                set_cell_bg(cell, "FEF2F2")
            set_cell_padding(cell, top=65, bottom=65, left=100, right=100)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(8.5)
            if c_i == 0:
                r.font.bold = True
                r.font.name = 'Consolas'
                r.font.color.rgb = RGBColor(220, 38, 38)
            elif c_i == 3:
                r.font.bold = True
                r.font.color.rgb = RGBColor(21, 128, 61)
            else:
                r.font.color.rgb = RGBColor(51, 65, 85)

    # ==========================================
    # 14. SENARAI SEMAK AKHIR (CHECKLIST) (PAGE 15)
    # ==========================================
    doc.add_page_break()

    p_ck_h = doc.add_paragraph()
    p_ck_h.paragraph_format.space_before = Pt(14)
    p_ck_h.paragraph_format.space_after = Pt(4)
    r_ck_h = p_ck_h.add_run("Senarai Semak Akhir (Final Checklist)")
    r_ck_h.font.name = 'Calibri'
    r_ck_h.font.size = Pt(16)
    r_ck_h.font.bold = True
    r_ck_h.font.color.rgb = RGBColor(6, 104, 225)

    p_ck_sub = doc.add_paragraph()
    p_ck_sub.paragraph_format.space_after = Pt(10)
    p_ck_sub.add_run("Gunakan senarai semak ini sebelum pembentangan projek kepada pensyarah, penerbitan modul, atau pelancaran sistem sebenar:")

    ck_items = [
        "Saya faham perbezaan teknikal antara Page ID, Ad Account ID, App ID, App Secret, dan Access Token.",
        "Saya mempunyai System User berstatus Admin dengan aset Page dan Ad Account yang telah disambungkan (Assign Assets).",
        "Saya mempunyai Token Akses Kekal (Never-Expiring) yang bermula dengan huruf 'EAA...' dan disimpan dengan selamat.",
        "Saya faham bagaimana kombinasi Antigravity + MCP membolehkan Ejen AI membaca metrik ads dan membuat draf kempen secara automatik.",
        "Semua draf kempen iklan dikunci kepada status PAUSED bagi memastikan keselamatan kewangan mutlak (RM 0.00 perbelanjaan tanpa kelulusan).",
        "Saya tahu cara menjawab soalan pensyarah sekiranya ditanya mengenai etika keselamatan kecerdasan buatan (AI Guardrails) dan kawalan bajet."
    ]

    t_ck = doc.add_table(rows=len(ck_items), cols=2)
    t_ck.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_clean_table_borders(t_ck, "E2E8F0")

    for idx, item_txt in enumerate(ck_items):
        c0 = t_ck.cell(idx, 0)
        c0.width = Inches(0.6)
        set_cell_padding(c0, top=70, bottom=70, left=80, right=60)
        p0 = c0.paragraphs[0]
        r0 = p0.add_run("[   ]")
        r0.font.bold = True
        r0.font.name = 'Consolas'
        r0.font.size = Pt(11)
        r0.font.color.rgb = RGBColor(6, 104, 225)

        c1 = t_ck.cell(idx, 1)
        c1.width = Inches(6.3)
        set_cell_padding(c1, top=70, bottom=70, left=60, right=100)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(f"{idx+1}. {item_txt}")
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = RGBColor(30, 41, 59)
        if idx % 2 == 1:
            set_cell_bg(c0, "F8FAFC")
            set_cell_bg(c1, "F8FAFC")

    # Final Closing Footer Note
    p_close = doc.add_paragraph()
    p_close.paragraph_format.space_before = Pt(24)
    p_close.paragraph_format.space_after = Pt(4)
    p_close.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_close = p_close.add_run("Modul Ebook ini dibangunkan secara rasmi oleh Pasukan Orkestrasi Ejen AI DFK INC / Muhammad Syahmi (syahmiaof).")
    r_close.font.size = Pt(9)
    r_close.font.italic = True
    r_close.font.color.rgb = RGBColor(100, 116, 139)

    out_file = os.path.join(r"C:\Users\USER\Desktop\modulmetaads", "Modul_Meta_Ads_dan_AI_Agent_Automation.docx")
    doc.save(out_file)
    print(f"Document successfully created at: {out_file}")

if __name__ == '__main__':
    generate_modul_docx()
