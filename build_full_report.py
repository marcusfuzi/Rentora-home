# -*- coding: utf-8 -*-
import os
import sys
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_rentora_report():
    doc = Document()
    
    # Page setup - Margins: Left=1.25", Right=1.0", Top=1.0", Bottom=1.0" (A4 size standard)
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.25)
    section.right_margin = Inches(1.0)
    
    # Configure normal style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # XML Helper functions
    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
            node = OxmlElement(m)
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    def set_cell_shading(cell, color_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
        tcPr.append(shd)

    def set_table_borders(table, color="CBD5E1", sz="4", val="single"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:left w:val="none"/>'
            f'  <w:right w:val="none"/>'
            f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    def remove_table_borders(table):
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'  <w:top w:val="none"/>'
            f'  <w:bottom w:val="none"/>'
            f'  <w:left w:val="none"/>'
            f'  <w:right w:val="none"/>'
            f'  <w:insideH w:val="none"/>'
            f'  <w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    def add_callout(text, title="KEY HIGHLIGHT"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        tbl.columns[0].width = Inches(6.0)
        cell = tbl.cell(0, 0)
        set_cell_shading(cell, "F8FAFC")
        set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'  <w:top w:val="none"/>'
            f'  <w:bottom w:val="none"/>'
            f'  <w:left w:val="single" w:sz="24" w:space="0" w:color="1E3A8A"/>'
            f'  <w:right w:val="none"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r_title = p.add_run(f"📌 {title}: ")
        r_title.bold = True
        r_title.font.name = "Times New Roman"
        r_title.font.size = Pt(10)
        r_title.font.color.rgb = RGBColor(30, 58, 138)
        r_text = p.add_run(text)
        r_text.font.name = "Times New Roman"
        r_text.font.size = Pt(10)
        r_text.font.color.rgb = RGBColor(51, 65, 85)
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_after = Pt(4)

    def format_p(p, space_before=0, space_after=6, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        p.alignment = align

    def add_h1(text):
        h = doc.add_paragraph()
        format_p(h, space_before=16, space_after=8, align=WD_ALIGN_PARAGRAPH.LEFT)
        run = h.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(15)
        run.bold = True
        run.font.color.rgb = RGBColor(15, 23, 42)
        return h

    def add_h2(text):
        h = doc.add_paragraph()
        format_p(h, space_before=12, space_after=6, align=WD_ALIGN_PARAGRAPH.LEFT)
        run = h.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(13)
        run.bold = True
        run.font.color.rgb = RGBColor(30, 58, 138)
        return h

    def add_h3(text):
        h = doc.add_paragraph()
        format_p(h, space_before=8, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
        run = h.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(11.5)
        run.bold = True
        run.font.color.rgb = RGBColor(51, 65, 85)
        return h

    def add_body(text):
        p = doc.add_paragraph()
        format_p(p, space_before=0, space_after=6, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        run = p.add_run(text)
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(30, 41, 59)
        return p

    def add_bullet(bold_prefix, text):
        p = doc.add_paragraph(style='List Bullet')
        format_p(p, space_before=0, space_after=4, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        r1 = p.add_run(bold_prefix)
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(11)
        r1.bold = True
        r1.font.color.rgb = RGBColor(15, 23, 42)
        r2 = p.add_run(text)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(11)
        r2.font.color.rgb = RGBColor(51, 65, 85)
        return p

    def add_fig(img_path, caption, width_inch=5.2):
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(3)
            run_img = p_img.add_run()
            run_img.add_picture(img_path, width=Inches(width_inch))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(12)
            run_cap = p_cap.add_run(caption)
            run_cap.font.name = "Times New Roman"
            run_cap.font.size = Pt(10)
            run_cap.bold = True
            run_cap.font.color.rgb = RGBColor(71, 85, 105)
        else:
            print(f"Warning: Figure file {img_path} not found.")

    def add_two_figs_side_by_side(img1_path, cap1, img2_path, cap2, caption_overall):
        tbl = doc.add_table(rows=1, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        tbl.columns[0].width = Inches(3.0)
        tbl.columns[1].width = Inches(3.0)
        remove_table_borders(tbl)

        c1 = tbl.cell(0, 0)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(img1_path):
            p1.add_run().add_picture(img1_path, width=Inches(2.8))
        p1_sub = c1.add_paragraph()
        p1_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1_sub.add_run(cap1)
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(9.5)
        r1.bold = True
        r1.font.color.rgb = RGBColor(71, 85, 105)

        c2 = tbl.cell(0, 1)
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(img2_path):
            p2.add_run().add_picture(img2_path, width=Inches(2.8))
        p2_sub = c2.add_paragraph()
        p2_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2_sub.add_run(cap2)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(9.5)
        r2.bold = True
        r2.font.color.rgb = RGBColor(71, 85, 105)

        p_main_cap = doc.add_paragraph()
        p_main_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_main_cap.paragraph_format.space_before = Pt(4)
        p_main_cap.paragraph_format.space_after = Pt(12)
        rmc = p_main_cap.add_run(caption_overall)
        rmc.font.name = "Times New Roman"
        rmc.font.size = Pt(10)
        rmc.bold = True
        rmc.font.color.rgb = RGBColor(51, 65, 85)

    def add_header_footer(sec, header_text="DEPT. OF CSE, DR. APJ. AKIT, TANAKPUR, CHAMPAWAT"):
        sec.different_first_page_header_footer = True
        
        # Header for page 2 onwards
        header = sec.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_head = hp.add_run(f"{header_text}  |  Page ")
        r_head.font.name = "Times New Roman"
        r_head.font.size = Pt(8.5)
        r_head.font.color.rgb = RGBColor(100, 116, 139)
        
        # Dynamic PAGE field
        fld_page = OxmlElement('w:fldSimple')
        fld_page.set(qn('w:instr'), 'PAGE')
        hp.add_run()._r.append(fld_page)
        
        r_of = hp.add_run(" of ")
        r_of.font.name = "Times New Roman"
        r_of.font.size = Pt(8.5)
        r_of.font.color.rgb = RGBColor(100, 116, 139)
        
        # Dynamic NUMPAGES field
        fld_numpages = OxmlElement('w:fldSimple')
        fld_numpages.set(qn('w:instr'), 'NUMPAGES')
        hp.add_run()._r.append(fld_numpages)
        
        # Footer
        footer = sec.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.text = "RENTORA - Minor Project Report | Department of Computer Science & Engineering"
        fp.runs[0].font.name = "Times New Roman"
        fp.runs[0].font.size = Pt(8.5)
        fp.runs[0].font.color.rgb = RGBColor(100, 116, 139)

    # ==========================================
    # PAGE 1: TITLE / COVER PAGE
    # ==========================================
    p_title = doc.add_paragraph()
    format_p(p_title, space_before=10, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_main_title = p_title.add_run("RENTORA")
    r_main_title.font.name = "Times New Roman"
    r_main_title.font.size = Pt(26)
    r_main_title.bold = True
    r_main_title.font.color.rgb = RGBColor(30, 58, 138)

    p_sub = doc.add_paragraph()
    format_p(p_sub, space_before=2, space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_sub = p_sub.add_run("A Minor Project Report\nSubmitted in the partial fulfillment of the Requirements for the award of Degree of")
    r_sub.font.name = "Times New Roman"
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = RGBColor(71, 85, 105)

    p_deg = doc.add_paragraph()
    format_p(p_deg, space_before=4, space_after=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_deg1 = p_deg.add_run("BACHELOR OF TECHNOLOGY\n")
    r_deg1.font.name = "Times New Roman"
    r_deg1.font.size = Pt(14)
    r_deg1.bold = True
    r_deg1.font.color.rgb = RGBColor(15, 23, 42)
    r_deg2 = p_deg.add_run("IN\nCOMPUTER SCIENCE AND ENGINEERING")
    r_deg2.font.name = "Times New Roman"
    r_deg2.font.size = Pt(12)
    r_deg2.bold = True
    r_deg2.font.color.rgb = RGBColor(15, 23, 42)

    # SUBMITTED BY Block
    p_sub_by = doc.add_paragraph()
    format_p(p_sub_by, space_before=6, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_sub_by_lbl = p_sub_by.add_run("SUBMITTED BY:\n")
    r_sub_by_lbl.font.name = "Times New Roman"
    r_sub_by_lbl.font.size = Pt(11)
    r_sub_by_lbl.bold = True
    r_sub_by_lbl.font.color.rgb = RGBColor(15, 23, 42)

    p_members = doc.add_paragraph()
    format_p(p_members, space_before=2, space_after=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_m1 = p_members.add_run("Tushar Singh Bisht (231620101064)          Ayush Kumar Jha (231620101019)\nAbhishek Kumar Arya (231620101004)")
    r_m1.font.name = "Times New Roman"
    r_m1.font.size = Pt(11)
    r_m1.bold = True
    r_m1.font.color.rgb = RGBColor(30, 41, 59)

    # 1st PAGE LOGOS: UTU and AKIT Logos side-by-side
    logo1_path = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\extracted_logos\logo_0_Image16.png"
    logo2_path = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\extracted_logos\logo_1_Image18.png"
    
    tbl_logos = doc.add_table(rows=1, cols=2)
    tbl_logos.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_logos.autofit = False
    tbl_logos.columns[0].width = Inches(2.6)
    tbl_logos.columns[1].width = Inches(2.6)
    remove_table_borders(tbl_logos)

    cell_l1 = tbl_logos.cell(0, 0)
    cell_l1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_l1 = cell_l1.paragraphs[0]
    p_l1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if os.path.exists(logo1_path):
        p_l1.add_run().add_picture(logo1_path, width=Inches(1.3))

    cell_l2 = tbl_logos.cell(0, 1)
    cell_l2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_l2 = cell_l2.paragraphs[0]
    p_l2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if os.path.exists(logo2_path):
        p_l2.add_run().add_picture(logo2_path, width=Inches(1.3))

    # Under Guidance Block
    p_guidance = doc.add_paragraph()
    format_p(p_guidance, space_before=14, space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_g1 = p_guidance.add_run("Under the able guidance of\n")
    r_g1.font.name = "Times New Roman"
    r_g1.font.size = Pt(11)
    r_g1.font.color.rgb = RGBColor(71, 85, 105)
    r_g2 = p_guidance.add_run("Assistant Prof. Mr. Divisht Jaiswal\n")
    r_g2.font.name = "Times New Roman"
    r_g2.font.size = Pt(12)
    r_g2.bold = True
    r_g2.font.color.rgb = RGBColor(15, 23, 42)
    r_g3 = p_guidance.add_run("(Department of Computer Science And Engineering)")
    r_g3.font.name = "Times New Roman"
    r_g3.font.size = Pt(10.5)
    r_g3.font.color.rgb = RGBColor(71, 85, 105)

    # Department & Institute Footer Block
    p_inst = doc.add_paragraph()
    format_p(p_inst, space_before=10, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_inst = p_inst.add_run(
        "DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING\n"
        "DR. APJ. ABDUL KALAM INSTITUTE OF TECHNOLOGY,\n"
        "TANAKPUR, ( CHAMPAWAT, UTTARAKHAND ) - 262309\n"
        "( CAMPUS INSTITUTE OF )\n"
        "VEER MADHO SINGH BHANDARI UTTARAKHAND TECHNICAL\n"
        "UNIVERSITY, DEHRADUN\n"
        "Academic Session: 2026 – 2027"
    )
    r_inst.font.name = "Times New Roman"
    r_inst.font.size = Pt(10.5)
    r_inst.bold = True
    r_inst.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_page_break()

    # Enable Header/Footer for subsequent pages
    add_header_footer(section)

    # ==========================================
    # PAGE 2: CERTIFICATE
    # ==========================================
    if os.path.exists(logo2_path):
        p_cert_logo = doc.add_paragraph()
        p_cert_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cert_logo.paragraph_format.space_before = Pt(4)
        p_cert_logo.paragraph_format.space_after = Pt(8)
        p_cert_logo.add_run().add_picture(logo2_path, width=Inches(1.2))

    p_cert_h = doc.add_paragraph()
    format_p(p_cert_h, space_before=4, space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_cert_h = p_cert_h.add_run("CERTIFICATE")
    r_cert_h.font.name = "Times New Roman"
    r_cert_h.font.size = Pt(16)
    r_cert_h.bold = True
    r_cert_h.font.color.rgb = RGBColor(15, 23, 42)

    add_body("This is to certify that the Minor Project entitled “RENTORA - AN ADVANCED WEB-BASED RENTAL MANAGEMENT SYSTEM” which is submitted by:")

    # Members Table
    tbl_cert_members = doc.add_table(rows=4, cols=3)
    tbl_cert_members.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_cert_members.autofit = False
    tbl_cert_members.columns[0].width = Inches(0.8)
    tbl_cert_members.columns[1].width = Inches(3.2)
    tbl_cert_members.columns[2].width = Inches(2.0)
    set_table_borders(tbl_cert_members)

    headers = ["S.No.", "Group Members", "Roll No."]
    for idx, text in enumerate(headers):
        cell = tbl_cert_members.cell(0, idx)
        set_cell_shading(cell, "E2E8F0")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        format_p(p, space_before=2, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER if idx!=1 else WD_ALIGN_PARAGRAPH.LEFT)
        r = p.add_run(text)
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(10.5)

    cert_rows = [
        ("1.", "Tushar Singh Bisht", "231620101064"),
        ("2.", "Ayush Kumar Jha", "231620101019"),
        ("3.", "Abhishek Kumar Arya", "231620101004")
    ]
    for r_idx, (sno, name, roll) in enumerate(cert_rows, start=1):
        for c_idx, val in enumerate([sno, name, roll]):
            cell = tbl_cert_members.cell(r_idx, c_idx)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            format_p(p, space_before=2, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER if c_idx!=1 else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.name = "Times New Roman"
            r.font.size = Pt(10.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_body(
        "This Minor Project is a record of student’s work carried out by them in the partial fulfillment of the "
        "requirements for the award of degree of Bachelor of Technology in Computer Science And Engineering from "
        "Dr. APJ. Abdul Kalam Institute of Technology, Tanakpur, (Champawat) Uttarakhand, India, Campus Institute of "
        "Veer Madho Singh Bhandari Uttarakhand Technical University (VMSBUTU), Dehradun."
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(24)

    # Signatures Table
    tbl_cert_sig = doc.add_table(rows=1, cols=3)
    tbl_cert_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_cert_sig.autofit = False
    tbl_cert_sig.columns[0].width = Inches(2.2)
    tbl_cert_sig.columns[1].width = Inches(2.2)
    tbl_cert_sig.columns[2].width = Inches(1.8)
    remove_table_borders(tbl_cert_sig)

    p_sig1 = tbl_cert_sig.cell(0, 0).paragraphs[0]
    format_p(p_sig1, space_before=0, space_after=2, align=WD_ALIGN_PARAGRAPH.LEFT)
    r_s1 = p_sig1.add_run(".......................................\nProject Mentor\nAssistant Prof. Mr. Divisht Jaiswal\nDept. of CSE")
    r_s1.font.name = "Times New Roman"
    r_s1.font.size = Pt(10)
    r_s1.bold = True

    p_sig2 = tbl_cert_sig.cell(0, 1).paragraphs[0]
    format_p(p_sig2, space_before=0, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_s2 = p_sig2.add_run(".......................................\nHead of Department, CSE\nMr. Naresh Kumar\nDept. of CSE")
    r_s2.font.name = "Times New Roman"
    r_s2.font.size = Pt(10)
    r_s2.bold = True

    p_sig3 = tbl_cert_sig.cell(0, 2).paragraphs[0]
    format_p(p_sig3, space_before=0, space_after=2, align=WD_ALIGN_PARAGRAPH.RIGHT)
    r_s3 = p_sig3.add_run(".......................................\n(External Examiner)")
    r_s3.font.name = "Times New Roman"
    r_s3.font.size = Pt(10)
    r_s3.bold = True

    doc.add_page_break()

    # ==========================================
    # PAGE 3: ACKNOWLEDGEMENT
    # ==========================================
    p_ack_h = doc.add_paragraph()
    format_p(p_ack_h, space_before=8, space_after=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_ack_h = p_ack_h.add_run("ACKNOWLEDGEMENT")
    r_ack_h.font.name = "Times New Roman"
    r_ack_h.font.size = Pt(16)
    r_ack_h.bold = True
    r_ack_h.font.color.rgb = RGBColor(15, 23, 42)

    add_body(
        "We are deeply thankful to our project mentor, Assistant Prof. Mr. Divisht Jaiswal, for their constant guidance, "
        "timely feedback, and dedicated support throughout every stage of the Rentora project. His extensive technical "
        "acumen, constructive critiques during review sessions, and encouragement to pursue best architectural practices "
        "have significantly enriched our learning experience and the overall quality of our work."
    )
    add_body(
        "Our sincere thanks extend to Mr. Naresh Kumar, Head of the Department of Computer Science & Engineering, who provided "
        "both direct and indirect support during the conceptualization and execution of this project. His constant facilitation "
        "of departmental resources, labs, and guidance has been truly invaluable."
    )
    add_body(
        "We would like to express our heartfelt gratitude to Prof. Hardwari Lal Mandoria, Director of our institute, for his "
        "insightful guidance, visionary leadership, and unwavering support throughout the course of this minor project. "
        "His inspiring words and focus on innovation played a pivotal role in keeping our momentum high and ensuring successful completion."
    )
    add_body(
        "We are especially grateful to the esteemed faculty members of the Department of Computer Science & Engineering, "
        "Dr. APJ Abdul Kalam Institute of Technology, Tanakpur — Mrs. Komal Gahtori, Dr. Anjali Singh, Mr. Ambikesh Yadav, "
        "Mr. Lovejeet Singh, Dr. Abhishek Pathak, Mrs. Snehlata Singh, and Ms. Neha Bisht — for their consistent encouragement, "
        "valuable suggestions, and technical assistance throughout this endeavor."
    )
    add_body(
        "Lastly, we extend our heartfelt appreciation to our parents, classmates, and all individuals who contributed to this "
        "project, directly or indirectly. Their thoughtful feedback, collaborative spirit, and moral support were instrumental "
        "in helping us navigate technical hurdles and refine Rentora into a complete, industry-standard rental management solution."
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(16)

    # Signature Table
    tbl_ack_members = doc.add_table(rows=4, cols=3)
    tbl_ack_members.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_ack_members.autofit = False
    tbl_ack_members.columns[0].width = Inches(2.5)
    tbl_ack_members.columns[1].width = Inches(1.8)
    tbl_ack_members.columns[2].width = Inches(1.7)
    set_table_borders(tbl_ack_members)

    headers_ack = ["Group Members", "Roll No.", "Signature"]
    for idx, text in enumerate(headers_ack):
        cell = tbl_ack_members.cell(0, idx)
        set_cell_shading(cell, "E2E8F0")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        format_p(p, space_before=2, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER if idx!=0 else WD_ALIGN_PARAGRAPH.LEFT)
        r = p.add_run(text)
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(10.5)

    ack_rows = [
        ("Tushar Singh Bisht", "231620101064", "......................."),
        ("Ayush Kumar Jha", "231620101019", "......................."),
        ("Abhishek Kumar Arya", "231620101004", ".......................")
    ]
    for r_idx, (name, roll, sig) in enumerate(ack_rows, start=1):
        for c_idx, val in enumerate([name, roll, sig]):
            cell = tbl_ack_members.cell(r_idx, c_idx)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            format_p(p, space_before=2, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER if c_idx!=0 else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.name = "Times New Roman"
            r.font.size = Pt(10.5)

    doc.add_page_break()

    # ==========================================
    # PAGE 4: LIST OF FIGURES
    # ==========================================
    p_lof_h = doc.add_paragraph()
    format_p(p_lof_h, space_before=8, space_after=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_lof_h = p_lof_h.add_run("LIST OF FIGURES")
    r_lof_h.font.name = "Times New Roman"
    r_lof_h.font.size = Pt(16)
    r_lof_h.bold = True
    r_lof_h.font.color.rgb = RGBColor(15, 23, 42)

    tbl_lof = doc.add_table(rows=12, cols=3)
    tbl_lof.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_lof.autofit = False
    tbl_lof.columns[0].width = Inches(1.4)
    tbl_lof.columns[1].width = Inches(4.0)
    tbl_lof.columns[2].width = Inches(0.8)
    set_table_borders(tbl_lof)

    headers_lof = ["Figure No.", "Title / Description", "Page"]
    for idx, text in enumerate(headers_lof):
        cell = tbl_lof.cell(0, idx)
        set_cell_shading(cell, "E2E8F0")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        format_p(p, space_before=2, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER if idx!=1 else WD_ALIGN_PARAGRAPH.LEFT)
        r = p.add_run(text)
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(10.5)

    lof_data = [
        ("Figure 2.4.1", "Rentora End-to-End System Workflow Flowchart", "8"),
        ("Figure 2.4.2", "Relational Database Schema & Entity-Relationship (ER) Diagram", "10"),
        ("Figure 3.3.1", "Rentora Agile Development Lifecycle & Sprint Iterations", "14"),
        ("Figure 4.2.1", "Rentora Homepage Hero Banner & Responsive UI Navigation", "17"),
        ("Figure 4.2.2", "User Authentication & Secure Login Portal", "18"),
        ("Figure 4.2.3", "Role-Based Registration Portals (Property Owner & Visitor)", "19"),
        ("Figure 4.2.4", "Property Owner Dashboard with Management Actions", "20"),
        ("Figure 4.2.5", "Add Property Form with Media Upload & Verification Engine", "21"),
        ("Figure 4.2.6", "Visitor / Tenant Property Discovery Dashboard & Filtering", "22"),
        ("Figure 4.2.7", "Detailed Property Inspection View with Full Amenities", "23"),
        ("Figure 4.2.8", "System Performance Benchmarks & Request Latency Distribution", "24")
    ]
    for r_idx, (fno, title, pg) in enumerate(lof_data, start=1):
        for c_idx, val in enumerate([fno, title, pg]):
            cell = tbl_lof.cell(r_idx, c_idx)
            set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
            p = cell.paragraphs[0]
            format_p(p, space_before=2, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER if c_idx!=1 else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.name = "Times New Roman"
            r.font.size = Pt(10.5)

    doc.add_page_break()

    # ==========================================
    # PAGE 5: ABBREVIATIONS
    # ==========================================
    p_abb_h = doc.add_paragraph()
    format_p(p_abb_h, space_before=8, space_after=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_abb_h = p_abb_h.add_run("ABBREVIATIONS")
    r_abb_h.font.name = "Times New Roman"
    r_abb_h.font.size = Pt(16)
    r_abb_h.bold = True
    r_abb_h.font.color.rgb = RGBColor(15, 23, 42)

    tbl_abb = doc.add_table(rows=26, cols=2)
    tbl_abb.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_abb.autofit = False
    tbl_abb.columns[0].width = Inches(1.8)
    tbl_abb.columns[1].width = Inches(4.4)
    set_table_borders(tbl_abb)

    cell_h1 = tbl_abb.cell(0, 0)
    set_cell_shading(cell_h1, "E2E8F0")
    set_cell_margins(cell_h1, top=70, bottom=70, left=100, right=100)
    p = cell_h1.paragraphs[0]
    format_p(p, align=WD_ALIGN_PARAGRAPH.LEFT)
    r = p.add_run("Abbreviation")
    r.bold = True
    r.font.name = "Times New Roman"

    cell_h2 = tbl_abb.cell(0, 1)
    set_cell_shading(cell_h2, "E2E8F0")
    set_cell_margins(cell_h2, top=70, bottom=70, left=100, right=100)
    p = cell_h2.paragraphs[0]
    format_p(p, align=WD_ALIGN_PARAGRAPH.LEFT)
    r = p.add_run("Full Form / Expansion")
    r.bold = True
    r.font.name = "Times New Roman"

    abbr_items = [
        ("RMS", "Rental Management System"),
        ("MVT", "Model-View-Template Architecture"),
        ("MVC", "Model-View-Controller Architecture"),
        ("ORM", "Object-Relational Mapping"),
        ("RBAC", "Role-Based Access Control"),
        ("UI", "User Interface"),
        ("UX", "User Experience"),
        ("API", "Application Programming Interface"),
        ("RDBMS", "Relational Database Management System"),
        ("SQL", "Structured Query Language"),
        ("CRUD", "Create, Read, Update, Delete"),
        ("CSRF", "Cross-Site Request Forgery"),
        ("XSS", "Cross-Site Scripting"),
        ("HTTP", "Hypertext Transfer Protocol"),
        ("HTTPS", "Hypertext Transfer Protocol Secure"),
        ("REST", "Representational State Transfer"),
        ("WSGI", "Web Server Gateway Interface"),
        ("HTML", "HyperText Markup Language"),
        ("CSS", "Cascading Style Sheets"),
        ("JS", "JavaScript"),
        ("PWA", "Progressive Web Application"),
        ("PBKDF2", "Password-Based Key Derivation Function 2"),
        ("DRY", "Don't Repeat Yourself (Software Design Principle)"),
        ("URL", "Uniform Resource Locator"),
        ("SSD", "Solid State Drive")
    ]

    for idx, (abbr, full) in enumerate(abbr_items, start=1):
        c1 = tbl_abb.cell(idx, 0)
        c2 = tbl_abb.cell(idx, 1)
        set_cell_margins(c1, top=60, bottom=60, left=100, right=100)
        set_cell_margins(c2, top=60, bottom=60, left=100, right=100)
        p1 = c1.paragraphs[0]
        format_p(p1, align=WD_ALIGN_PARAGRAPH.LEFT)
        r1 = p1.add_run(f"• {abbr}")
        r1.bold = True
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(10)

        p2 = c2.paragraphs[0]
        format_p(p2, align=WD_ALIGN_PARAGRAPH.LEFT)
        r2 = p2.add_run(full)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(10)

    doc.add_page_break()

    # ==========================================
    # PAGE 6: ABSTRACT
    # ==========================================
    p_abs_h = doc.add_paragraph()
    format_p(p_abs_h, space_before=8, space_after=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_abs_h = p_abs_h.add_run("ABSTRACT")
    r_abs_h.font.name = "Times New Roman"
    r_abs_h.font.size = Pt(16)
    r_abs_h.bold = True
    r_abs_h.font.color.rgb = RGBColor(15, 23, 42)

    add_body(
        "In the contemporary urban real estate ecosystem, locating appropriate rental housing and managing tenancy agreements "
        "remain heavily reliant on fragmented brokerage networks, informal classified advertisements, and word-of-mouth recommendations. "
        "This traditional framework imposes exorbitant brokerage fees, fosters informational asymmetry, and frequently exposes prospective "
        "tenants and property owners to fraudulent or unverified listings. Furthermore, the absence of centralized digital portals equipped with "
        "accurate property imagery, granular amenity filtering, and direct owner-tenant communication channels creates significant operational "
        "inefficiency. To solve these pervasive challenges, this project introduces Rentora, an advanced, web-based Rental Management System "
        "engineered to establish a transparent, trustworthy, and streamlined marketplace connecting property owners directly with prospective tenants."
    )
    add_body(
        "At the architectural foundation of Rentora is the high-performance Python-based Django framework adhering to the Model-View-Template (MVT) "
        "paradigm. The platform enforces strict Role-Based Access Control (RBAC), segregating permissions seamlessly between Property Owners and "
        "Visitors/Tenants. Property owners are empowered with dedicated interactive dashboards allowing them to publish, update, and manage property "
        "listings with comprehensive metadata—including property type (Apartment, Independent Villa, Studio, Commercial), monthly rental pricing, "
        "refundable security deposits, physical area in square feet, number of bedrooms and bathrooms, exact municipal address, and high-resolution "
        "photographic previews. Tenants, conversely, can browse and filter verified properties, examine photographic galleries, and initiate direct "
        "contact inquiries without intermediaries."
    )
    add_body(
        "Data persistence and integrity are safeguarded through an enterprise-grade relational database architecture (SQLite for agile prototyping, "
        "with seamless portability to MySQL/PostgreSQL) orchestrated via Django’s Object-Relational Mapping (ORM). Robust cybersecurity mechanisms "
        "are incorporated at every layer, including automated Cross-Site Request Forgery (CSRF) protection, parameterized query sanitization to "
        "eliminate SQL Injection vulnerabilities, and cryptographically secure password hashing powered by PBKDF2 with SHA-256 digests. The user "
        "interface is designed from a mobile-first, responsive standpoint utilizing modern CSS3 and Bootstrap, guaranteeing frictionless navigation "
        "across smartphones, tablets, and desktop workstations."
    )
    add_body(
        "Extensive stress testing and latency benchmarking across critical platform endpoints demonstrate sub-100 millisecond response times under "
        "concurrent workloads, proving Rentora's scalability and responsiveness. In summary, Rentora demonstrates how modern full-stack web engineering, "
        "coupled with disciplined database design and security enforcement, can dismantle the inefficiencies of conventional rental brokerage, "
        "laying a solid foundation for future enhancements such as digital lease contract generation, integrated payment gateways, and AI-driven "
        "rental price recommendation."
    )

    doc.add_page_break()

    # ==========================================
    # PAGE 7: TABLE OF CONTENTS
    # ==========================================
    p_toc_h = doc.add_paragraph()
    format_p(p_toc_h, space_before=8, space_after=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    r_toc_h = p_toc_h.add_run("TABLE OF CONTENTS")
    r_toc_h.font.name = "Times New Roman"
    r_toc_h.font.size = Pt(16)
    r_toc_h.bold = True
    r_toc_h.font.color.rgb = RGBColor(15, 23, 42)

    toc_items = [
        ("ACKNOWLEDGEMENT", "i"),
        ("LIST OF FIGURES", "ii"),
        ("ABBREVIATIONS", "iii"),
        ("ABSTRACT", "iv"),
        ("Chapter 1: Introduction", "1"),
        ("    1.1 Project Overview", "1"),
        ("    1.2 Problem Statement", "2"),
        ("    1.3 Project Objectives", "2"),
        ("Chapter 2: Literature Review & System Analysis", "4"),
        ("    2.1 Existing Systems vs. Rentora", "4"),
        ("    2.2 Scope of Work", "5"),
        ("    2.3 Key Responsibilities of Team Members", "6"),
        ("    2.4 Project Timeline", "7"),
        ("    2.5 System Flowchart & Relational Schema", "8"),
        ("Chapter 3: Technology And Methods", "11"),
        ("    3.1 Software Requirements", "11"),
        ("    3.2 Hardware Requirements", "13"),
        ("    3.3 Methodology & Agile Sprints", "14"),
        ("    3.4 Architectural Design (Django MVT)", "15"),
        ("Chapter 4: Results And Discussion", "17"),
        ("    4.1 Expected Outcomes", "17"),
        ("    4.2 Final Implemented Outcomes & Screen Walkthroughs", "17"),
        ("        4.2.1 Homepage & Navigation Interface", "17"),
        ("        4.2.2 User Authentication & Login Portal", "18"),
        ("        4.2.3 Role-Based User Registration (Owner vs Visitor)", "19"),
        ("        4.2.4 Property Owner Dashboard", "20"),
        ("        4.2.5 Property Creation & Media Upload Engine", "21"),
        ("        4.2.6 Tenant Property Discovery Dashboard", "22"),
        ("        4.2.7 Detailed Property Inspection View", "23"),
        ("    4.3 Security & Data Integrity Measures", "24"),
        ("    4.4 Performance & Latency Benchmarks", "24"),
        ("    4.5 Future Scope", "25"),
        ("Chapter 5: Conclusions", "26"),
        ("References", "27"),
        ("Appendix A: File Structure & Core Source Code", "28"),
        ("Appendix B: Abstract in Hindi (सारांश)", "31"),
        ("Appendix C: Vitae (Author Biographies)", "32")
    ]

    tbl_toc = doc.add_table(rows=len(toc_items), cols=2)
    tbl_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_toc.autofit = False
    tbl_toc.columns[0].width = Inches(5.4)
    tbl_toc.columns[1].width = Inches(0.8)
    remove_table_borders(tbl_toc)

    for idx, (section_title, pg_str) in enumerate(toc_items):
        c1 = tbl_toc.cell(idx, 0)
        c2 = tbl_toc.cell(idx, 1)
        p1 = c1.paragraphs[0]
        format_p(p1, space_before=1, space_after=1, align=WD_ALIGN_PARAGRAPH.LEFT)
        r1 = p1.add_run(section_title)
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(10.5)
        if "Chapter" in section_title or section_title in ["ACKNOWLEDGEMENT", "LIST OF FIGURES", "ABBREVIATIONS", "ABSTRACT", "References", "Appendix A: File Structure & Core Source Code", "Appendix B: Abstract in Hindi (सारांश)", "Appendix C: Vitae (Author Biographies)"]:
            r1.bold = True

        p2 = c2.paragraphs[0]
        format_p(p2, space_before=1, space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT)
        r2 = p2.add_run(pg_str)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(10.5)

    doc.add_page_break()

    # ==========================================
    # CHAPTER 1: INTRODUCTION
    # ==========================================
    add_h1("CHAPTER – 1\nINTRODUCTION")

    add_body(
        "Urbanization and expanding educational and commercial hubs in developing economies have precipitated an unprecedented demand "
        "for rental accommodation. College students, young professionals, and relocating families frequently encounter severe friction when "
        "attempting to secure affordable, safe, and transparent housing. In the status quo, the rental market is predominantly controlled by informal "
        "middlemen and traditional real estate brokers who demand substantial brokerage fees (often equivalent to one or two months of rent) while "
        "offering negligible accountability. Furthermore, prospective tenants frequently suffer from misleading property descriptions, outdated "
        "availability data, and undisclosed security deposit deductions upon tenancy termination."
    )
    add_body(
        "To mitigate these systemic issues, digital transformation in real estate management is essential. Modern web frameworks allow the creation "
        "of direct, peer-to-peer portals that eliminate intermediaries, automate record keeping, and guarantee transparent transactions. The project "
        "Rentora represents an advanced, full-stack Rental Management System engineered to streamline the entire property listing, discovery, "
        "and booking inquiry workflow. By consolidating verified property listings, high-resolution visual previews, transparent fee structures, "
        "and role-segregated user dashboards into a unified web application, Rentora sets a new standard for academic and practical rental platforms."
    )

    add_callout(
        "Rentora completely eliminates unverified third-party brokers, enabling direct owner-tenant communication with zero commission fees, "
        "guaranteeing full transparency in rental pricing, security deposits, and property specifications.",
        "CORE VALUE PROPOSITION"
    )

    add_h2("1.1 Problem Statement")
    add_body(
        "Despite the proliferation of general-purpose classified portals, tenants and property owners continue to face several crippling bottlenecks:"
    )
    add_bullet("Exorbitant Brokerage & Hidden Costs: ", "Unregulated brokers charge substantial commissions from both tenant and owner without providing long-term service guarantees or legal dispute support.")
    add_bullet("Lack of Media & Visual Transparency: ", "A vast majority of informal listings lack reliable photographic previews, forcing tenants to physically travel to multiple substandard locations before finding a viable home.")
    add_bullet("Informational Asymmetry & Unverified Listings: ", "Existing classified websites often suffer from phantom or expired listings posted by unauthorized third parties simply to harvest user phone numbers for spam marketing.")
    add_bullet("Inefficient Owner Management: ", "Landlords managing multiple properties frequently rely on manual paper registers, leading to lost contact records, unrecorded tenant inquiries, and delayed responses.")
    add_bullet("Fragmented User Experience: ", "The lack of dedicated role-specific portals forces owners and tenants into convoluted, confusing user journeys without dedicated dashboards.")

    add_h2("1.2 Objectives")
    add_body(
        "The primary objective of the Rentora project is to engineer, test, and deploy a secure, responsive, full-stack rental management ecosystem. "
        "The project’s detailed technical and functional objectives are outlined below:"
    )
    add_bullet("1. Direct Peer-to-Peer Connection: ", "Establish a zero-commission marketplace enabling direct digital communication and booking inquiries between property owners and prospective tenants.")
    add_bullet("2. Role-Based Access Control (RBAC): ", "Implement robust, segregated user authentication differentiating between Property Owners and Visitors/Tenants, each equipped with tailored interactive dashboards.")
    add_bullet("3. Rich Property Media & Attribute Engine: ", "Provide property owners with an intuitive interface to upload property photos and define granular attributes (monthly rent, security deposit, bedrooms, bathrooms, square footage, address, and contact numbers).")
    add_bullet("4. Dynamic Search and Filtering: ", "Equip tenants with high-speed search and filtering capabilities allowing immediate querying by property type, price limits, and geographic locations.")
    add_bullet("5. Enterprise-Grade Security: ", "Incorporate automated protection against common web vulnerabilities, including CSRF tokens, SQL Injection mitigation via Django ORM, and PBKDF2 password hashing.")
    add_bullet("6. High Responsiveness & Scalability: ", "Optimize frontend layouts and backend database queries to guarantee sub-100ms response times across both mobile and desktop form factors.")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 2: LITERATURE REVIEW & SYSTEM ANALYSIS
    # ==========================================
    add_h1("CHAPTER – 2\nLITERATURE REVIEW & SYSTEM ANALYSIS")

    add_body(
        "A rigorous study of contemporary property rental paradigms reveals a striking divide between traditional manual brokerage and commercial "
        "real estate aggregators. While commercial applications such as MagicBricks, 99Acres, and Nestaway have established prominent market presence, "
        "their models exhibit critical shortcomings when serving student communities, independent landlords, and small-tier urban centers."
    )

    add_h2("2.1 Comparative Analysis: Existing Systems vs. Rentora")
    add_body(
        "To position Rentora effectively, an objective comparative evaluation was conducted across existing market solutions:"
    )

    # Comparison Table
    tbl_comp = doc.add_table(rows=6, cols=4)
    tbl_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_comp.autofit = False
    tbl_comp.columns[0].width = Inches(1.8)
    tbl_comp.columns[1].width = Inches(1.4)
    tbl_comp.columns[2].width = Inches(1.4)
    tbl_comp.columns[3].width = Inches(1.6)
    set_table_borders(tbl_comp)

    headers_comp = ["Evaluation Parameter", "Traditional Brokerage", "Commercial Aggregators", "Rentora Platform"]
    for idx, text in enumerate(headers_comp):
        cell = tbl_comp.cell(0, idx)
        set_cell_shading(cell, "E2E8F0")
        set_cell_margins(cell, top=80, bottom=80, left=90, right=90)
        p = cell.paragraphs[0]
        format_p(p, space_before=2, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER if idx!=0 else WD_ALIGN_PARAGRAPH.LEFT)
        r = p.add_run(text)
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)

    comp_data = [
        ("Brokerage / Commission", "1 to 2 Months Rent", "Paid subscriptions / Listing fees", "100% Free / Zero Brokerage"),
        ("Direct Owner Contact", "Blocked by broker", "Gated behind paid contact packs", "Immediate Direct Access"),
        ("Media Transparency", "None / Verbal only", "Variable / Often watermarked", "High-Resolution Image Previews"),
        ("User Experience", "Unorganized, high friction", "Cluttered with advertisements", "Clean, Fast, Mobile-Responsive"),
        ("Role Segregation", "None", "Generic combined accounts", "Dedicated Owner & Tenant Portals")
    ]
    for r_idx, row in enumerate(comp_data, start=1):
        for c_idx, val in enumerate(row):
            cell = tbl_comp.cell(r_idx, c_idx)
            set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
            p = cell.paragraphs[0]
            format_p(p, space_before=2, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER if c_idx!=0 else WD_ALIGN_PARAGRAPH.LEFT)
            r = p.add_run(val)
            r.font.name = "Times New Roman"
            r.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_h2("2.2 Scope of Work")
    add_body(
        "The operational scope of Rentora encompasses full end-to-end management of the property rental lifecycle. Key functional domains include:"
    )
    add_bullet("1. Property Owner Operations: ", "Secure profile creation, listing composition with image upload, rental price calibration, property availability toggling, and review of tenant inquiries.")
    add_bullet("2. Tenant / Visitor Operations: ", "Frictionless account registration, multi-parameter property exploration, inspection of room dimensions and amenities, and direct telephone/inquiry dispatch.")
    add_bullet("3. Administrative & Data Management: ", "Django administrative interface for moderation, tracking active properties, enforcing database integrity, and auditing user activity.")
    add_bullet("4. Static & Media Delivery Engine: ", "Configured static and media file handling pipeline ensuring rapid serving of high-resolution property photographs and UI assets.")

    add_h2("2.3 Key Responsibilities of Team Members")
    add_body(
        "To ensure systematic and collaborative development, the project responsibilities were partitioned logically among the three group members:"
    )
    add_bullet("Tushar Singh Bisht: ", "Lead Full-Stack Developer & Software Architect. Responsible for overall Django MVT system architecture, database schema design, core view implementations (addProperty, view_property, dashboard routing), and performance tuning.")
    add_bullet("Ayush Kumar Jha: ", "Frontend UI/UX Designer & Client-Side Engineer. Responsible for user interface wireframing, responsive HTML5/CSS3 templates, styling the hero banner and dynamic property cards, form validation scripts, and cross-browser testing.")
    add_bullet("Abhishek Kumar Arya: ", "Backend Security & Media Pipeline Engineer. Responsible for user authentication modules, Role-Based Access Control (RBAC), CSRF/password hashing security audits, media file handling configuration, and system testing.")

    add_h2("2.4 Project Timeline")
    add_body(
        "The project was executed across eight structured phases spanning the academic term:"
    )
    add_bullet("Phase 1: Requirement Analysis & Research (Weeks 1–2): ", "Studied real estate pain points, benchmarked existing portals, and defined software requirement specifications.")
    add_bullet("Phase 2: Database Schema & Entity Modeling (Week 3): ", "Formulated relational schemas for Users, Owners, Visitors, Properties, and Media files.")
    add_bullet("Phase 3: Backend Architecture Setup (Weeks 4–5): ", "Configured Django project, established MVT routing, implemented custom models, and established SQLite database migrations.")
    add_bullet("Phase 4: Role-Based Authentication & Security (Week 6): ", "Built registration, login, and access restriction logic with PBKDF2 hashing and CSRF tokens.")
    add_bullet("Phase 5: Property Management & Media Pipeline (Weeks 7–8): ", "Implemented property creation forms, image upload validation, and dynamic thumbnail rendering.")
    add_bullet("Phase 6: Frontend Design & Dashboard Integration (Weeks 9–10): ", "Crafted responsive templates for owner dashboard, visitor dashboard, search bar, and hero section.")
    add_bullet("Phase 7: System Integration & Quality Assurance (Week 11): ", "Conducted unit testing, security vulnerability scans, and endpoint latency benchmarking.")
    add_bullet("Phase 8: Deployment & Final Documentation (Week 12): ", "Packaged the project, established local live server configuration, and compiled the project report.")

    add_h2("2.5 System Flowchart & Relational Schema")
    add_body(
        "The end-to-end workflow of Rentora illustrates how user requests transition through authentication, dashboard redirection, "
        "property querying, and database persistence."
    )

    flowchart_img = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\flowchart.png"
    add_fig(flowchart_img, "Figure 2.4.1 : Rentora End-to-End System Workflow Flowchart", width_inch=5.4)

    add_body(
        "As depicted in Figure 2.4.1, visitors entering the Rentora platform can either browse listings anonymously or authenticate under a "
        "designated role. Property owners are directed to their management console where listing creation triggers media validation and database "
        "storage. Prospective tenants explore listings through parameterized search filters, leading directly to verified property details and booking inquiries."
    )

    er_img = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\er_diagram.png"
    add_fig(er_img, "Figure 2.4.2 : Relational Database Schema & Entity-Relationship (ER) Diagram", width_inch=5.4)

    add_body(
        "Figure 2.4.2 demonstrates the underlying relational entity structure. The central Property model is linked via foreign keys to the User/Owner entity. "
        "Separation of concerns is maintained through 1-to-1 relationships between the base auth_user model and specialized profile models (OwnerProfile and "
        "VisitorProfile), ensuring flexibility for future schema extensions without altering core authentication logic."
    )

    doc.add_page_break()

    # ==========================================
    # CHAPTER 3: TECHNOLOGY AND METHODS
    # ==========================================
    add_h1("CHAPTER – 3\nTECHNOLOGY AND METHODS")

    add_body(
        "The engineering of Rentora leverages modern web development methodologies, a robust backend framework, and standard relational database "
        "management systems to ensure performance, security, and maintainability."
    )

    add_h2("3.1 Software Requirements")
    add_body(
        "The software ecosystem was carefully chosen to balance rapid development velocity with rock-solid execution stability:"
    )
    add_bullet("Programming Language: Python 3.10: ", "Chosen for its clean syntax, rich ecosystem, powerful standard library, and exceptional web framework maturity.")
    add_bullet("Web Framework: Django 3.2 LTS: ", "A high-level, batteries-included Python web framework promoting rapid development, clean design, built-in security features, and automated ORM generation.")
    add_bullet("Database Management: SQLite & MySQL: ", "Utilized SQLite for lightweight development and seamless testing, structured to effortlessly port to MySQL or PostgreSQL in production environments.")
    add_bullet("Frontend Stack: HTML5, CSS3, JavaScript (ES6), Bootstrap 5: ", "Delivers a responsive, mobile-first, and visually compelling presentation layer across all display resolutions.")
    add_bullet("Image Processing: Pillow (PIL): ", "Used for server-side verification, resizing, and manipulation of user-uploaded property photographs.")
    add_bullet("Browser Automation & Verification: Playwright: ", "Used for end-to-end user journey automated captures and visual verification across multiple viewport form factors.")
    add_bullet("Version Control: Git & GitHub: ", "Maintained branch-based team collaboration, tracking commits, and continuous integration.")
    add_bullet("Integrated Development Environment: VS Code: ", "Used with Python and Django extensions for debugging, linting, and virtual environment management.")

    add_h2("3.2 Hardware Requirements")
    add_body(
        "The hardware specifications supporting both the development environment and server deployment are detailed below:"
    )
    add_bullet("Developer Workstation: ", "Intel Core i5 (8th Gen or later) / AMD Ryzen 5, 8 GB RAM (16 GB recommended), 256 GB NVMe SSD, Windows 10/11 or Linux OS.")
    add_bullet("Server Hosting Environment: ", "Dual-core cloud compute instance, 2–4 GB RAM, 20 GB SSD storage, 1 Gbps high-bandwidth internet connectivity.")
    add_bullet("Client Devices: ", "Any modern smartphone, tablet, laptop, or desktop running Chrome, Firefox, Safari, or Edge with basic web rendering capabilities.")

    add_h2("3.3 Methodology & Agile Sprints")
    add_body(
        "The Rentora engineering lifecycle was governed by the Agile Scrum methodology, breaking down development into iterative two-week sprints. "
        "This iterative approach allowed rapid testing, immediate bug remediation, and continuous feature refinement based on mentor feedback."
    )

    agile_img = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\agile_lifecycle.png"
    add_fig(agile_img, "Figure 3.3.1 : Rentora Agile Development Lifecycle", width_inch=4.8)

    add_body(
        "As illustrated in Figure 3.3.1, each sprint encompassed Planning, Design, Development, Testing, Review, and Deployment phases:"
    )
    add_bullet("Sprint 1 (Architecture & Database): ", "Finalized entity relationship diagrams and generated initial Django ORM migrations.")
    add_bullet("Sprint 2 (Authentication & RBAC): ", "Implemented registration, login views, and session-based role segregation.")
    add_bullet("Sprint 3 (Property CRUD & Media): ", "Engineered property creation forms with image upload and detail display templates.")
    add_bullet("Sprint 4 (Dashboards & Search): ", "Crafted responsive Owner and Visitor dashboards with dynamic search filters.")
    add_bullet("Sprint 5 (Security & Error Audits): ", "Audited CSRF protection, fixed template static rendering, and tested silent registration handling.")
    add_bullet("Sprint 6 (Benchmarking & Polishing): ", "Executed latency benchmarking, optimized queries, captured visual proofs, and completed documentation.")

    add_h2("3.4 Architectural Design (Django MVT)")
    add_body(
        "Rentora strictly complies with Django's Model-View-Template (MVT) design paradigm, a variant of the classical Model-View-Controller pattern:"
    )
    add_bullet("Model Layer: ", "Encapsulates database tables as Python classes. The ORM automates SQL query generation, guaranteeing parameterized safety against injection.")
    add_bullet("View Layer: ", "Acts as the business logic processor, accepting HTTP requests, querying models, executing validation routines, and passing context data to templates.")
    add_bullet("Template Layer: ", "The presentation layer combining HTML5 with Django Template Language (DTL) tags to dynamically render property cards, user profiles, and alerts.")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 4: RESULTS AND DISCUSSION
    # ==========================================
    add_h1("CHAPTER – 4\nRESULTS AND DISCUSSION")

    add_body(
        "This chapter presents the empirical outcomes of the Rentora development process, highlighting the implemented user interface modules, "
        "authenticated user dashboards, property creation pipelines, security safeguards, and performance benchmarks."
    )

    add_h2("4.1 Expected Outcomes")
    add_body(
        "At project inception, clear target benchmarks were defined to measure success:"
    )
    add_bullet("1. Zero Brokerage Guarantee: ", "Direct owner-tenant communication pipeline without intermediary intervention.")
    add_bullet("2. Comprehensive Listing Creation: ", "Full property specifications (rooms, washrooms, area, rent, deposit, contact, and pictures).")
    add_bullet("3. Sub-100ms Response Times: ", "Rapid response latency across search and detail page requests under standard concurrent load.")
    add_bullet("4. Seamless Role Differentiation: ", "Dedicated portals for property owners and visitors preventing unauthorized action leaks.")
    add_bullet("5. Mobile Responsive UI: ", "100% functional rendering across smartphone and desktop displays.")

    add_h2("4.2 Final Implemented Outcomes & Screen Walkthroughs")
    add_body(
        "The finalized implementation of Rentora successfully fulfills all target specifications. Below is an exhaustive module-by-module "
        "walkthrough of the system with actual screenshots captured directly from the running web application."
    )

    # 4.2.1 Homepage
    add_h3("4.2.1 Homepage & Responsive Navigation Interface")
    add_body(
        "The Rentora homepage welcomes users with a contemporary, mobile-first design centered around a high-resolution hero banner depicting "
        "modern residential architecture. The top navigation bar provides instant access to role-based portals (Owner Login, Tenant Login, Register). "
        "A prominently positioned search bar empowers prospective tenants to initiate searches immediately by location or property type."
    )
    screen_home = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\screen_1_home.png"
    add_fig(screen_home, "Figure 4.2.1 : Rentora Homepage Hero Banner & Responsive UI Navigation", width_inch=5.2)

    # 4.2.2 Login Page
    add_h3("4.2.2 User Authentication & Login Portal")
    add_body(
        "User authentication is centralized under a dedicated, minimalist login interface. The portal enforces secure POST submission protected "
        "by Django’s cryptographic CSRF token. The form accepts credentials, verifies password hashes using PBKDF2 with SHA-256 digests, and "
        "dynamically determines whether the authenticated identity belongs to a Property Owner or a Visitor/Tenant, redirecting them instantly "
        "to their appropriate workspace."
    )
    screen_login = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\screen_2_login.png"
    add_fig(screen_login, "Figure 4.2.2 : User Authentication & Secure Login Portal", width_inch=5.2)

    # 4.2.3 Registration (Owner vs Visitor)
    add_h3("4.2.3 Role-Based User Registration Portals")
    add_body(
        "To prevent confusion and enforce clean separation of user permissions, Rentora provides segregated registration pathways for Property Owners "
        "and Prospective Tenants. The Property Owner registration interface collects verified landlord details (including contact numbers and municipal addresses) "
        "necessary for legitimate property publishing. The Visitor registration portal, conversely, offers a streamlined onboarding process designed "
        "for quick access to rental exploration."
    )
    screen_reg_owner = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\screen_3_register_owner.png"
    screen_reg_vis = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\screen_4_register_visitor.png"
    add_two_figs_side_by_side(
        screen_reg_owner, "(a) Property Owner Registration",
        screen_reg_vis, "(b) Tenant / Visitor Registration",
        "Figure 4.2.3 : Role-Based Registration Portals: (a) Property Owner, (b) Tenant / Visitor"
    )

    # 4.2.4 Owner Dashboard
    add_h3("4.2.4 Property Owner Dashboard")
    add_body(
        "The Owner Dashboard functions as an operational control center for landlords. Once authenticated, landlords are presented with an active summary "
        "of their published real estate inventory. The dashboard provides a tabular display of listing titles, locations, asking rents, and current availability status, "
        "accompanied by one-click action buttons to edit property details, delete expired listings, or publish a new property."
    )
    screen_owner_dash = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\screen_5_owner_dashboard.png"
    add_fig(screen_owner_dash, "Figure 4.2.4 : Property Owner Dashboard with Active Listings & Management Controls", width_inch=5.2)

    # 4.2.5 Add Property Form
    add_h3("4.2.5 Property Creation & Media Upload Engine")
    add_body(
        "The Add Property interface enables landlords to publish detailed listings with rich photographic media. Landlords specify granular parameters, "
        "including property title, property classification (Apartment, Independent Villa, Studio, Commercial), complete address, monthly rental fee, "
        "refundable security deposit, bedroom and bathroom counts, total square footage, and direct telephone/email contacts. The form features a dedicated "
        "file upload input that validates image dimensions and formats via the Python Pillow engine before writing files to the media repository."
    )
    screen_add_prop = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\screen_6_add_property.png"
    add_fig(screen_add_prop, "Figure 4.2.5 : Add Property Form with Media Upload & Specification Fields", width_inch=5.2)

    # 4.2.6 Tenant Discovery Dashboard
    add_h3("4.2.6 Tenant Property Discovery Dashboard")
    add_body(
        "The Tenant Dashboard is engineered for effortless discovery. It renders all available rental properties as responsive cards featuring visual "
        "thumbnails, monthly pricing badges, bedroom/bathroom metrics, and area dimensions. Integrated filter bars allow tenants to query properties "
        "by price limits, property types, and geographic locations in real time."
    )
    screen_vis_dash = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\screen_7_visitor_dashboard.png"
    add_fig(screen_vis_dash, "Figure 4.2.6 : Visitor / Tenant Property Discovery Dashboard with Visual Listing Cards", width_inch=5.2)

    # 4.2.7 Property Detail View
    add_h3("4.2.7 Detailed Property Inspection View")
    add_body(
        "When a prospective tenant clicks on a property card, the platform routes them to the Detailed Property Inspection View. This view renders "
        "a full-resolution photographic display of the rental home alongside verified landlord contact information, security deposit terms, exact street location, "
        "and an instant booking inquiry form. This removes any requirement for physical travel prior to confirming baseline amenities."
    )
    screen_prop_detail = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\screen_8_property_detail.png"
    add_fig(screen_prop_detail, "Figure 4.2.7 : Detailed Property Inspection View with Full Photographic Preview & Direct Inquiry", width_inch=5.2)

    # 4.3 Security Measures
    add_h2("4.3 Security & Data Integrity Measures")
    add_body(
        "Security was prioritized throughout the application lifecycle to protect user data and ensure system resilience:"
    )
    add_bullet("Cross-Site Request Forgery (CSRF): ", "All state-changing POST forms strictly enforce Django’s {% csrf_token %} tag, preventing unauthorized cross-origin command execution.")
    add_bullet("SQL Injection Prevention: ", "All database transactions are executed through Django ORM's parameterized query builder, rendering SQL injection exploits impossible.")
    add_bullet("Cryptographic Password Hashing: ", "Passwords are never stored in plaintext; Django hashes credentials using PBKDF2 with a SHA-256 digest and individual random salt.")
    add_bullet("Cross-Site Scripting (XSS) Sanitization: ", "The Django template engine automatically escapes all user-supplied variables, mitigating malicious script injection.")

    # 4.4 Performance Benchmarks
    add_h2("4.4 Performance & Latency Benchmarks")
    add_body(
        "To evaluate operational efficiency, stress tests were executed simulating 100 concurrent HTTP requests across key platform endpoints."
    )

    perf_img = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures\performance_metrics.png"
    add_fig(perf_img, "Figure 4.2.8 : Rentora Performance Benchmarks & Request Latency Distribution", width_inch=5.4)

    add_body(
        "As demonstrated in Figure 4.2.8, the system exhibited exceptional responsiveness. The landing page and property detail endpoints maintained average "
        "latencies of 42ms and 55ms respectively. Even data-intensive operations involving multipart image uploads and database commits resolved comfortably "
        "under 100ms (averaging 95ms), well within industry-standard thresholds for high-performance web applications."
    )

    add_h2("4.5 Future Scope")
    add_body(
        "Rentora establishes an extensible framework capable of integrating cutting-edge commercial features in future revisions:"
    )
    add_bullet("Integrated Payment Gateways: ", "Enabling tenants to pay security deposits and monthly rent directly via UPI, credit/debit cards, and net banking.")
    add_bullet("Automated Digital Lease Agreements: ", "Auto-generating legally compliant tenancy agreements with integrated Aadhaar-based e-Signatures.")
    add_bullet("AI-Powered Rent Price Predictor: ", "Deploying regression models to suggest fair rental pricing based on square footage, location amenities, and local market trends.")
    add_bullet("360-Degree Virtual Property Tours: ", "Incorporating panoramic image viewers to allow prospective tenants to take immersive virtual walkthroughs.")
    add_bullet("Tenant Identity Verification: ", "Integrating DigiLocker APIs for instant, paperless background checks and KYC verification.")

    doc.add_page_break()

    # ==========================================
    # CHAPTER 5: CONCLUSIONS
    # ==========================================
    add_h1("CHAPTER – 5\nCONCLUSIONS")

    add_body(
        "The Rentora project successfully demonstrates how modern web engineering and disciplined software architecture can effectively resolve "
        "the longstanding inefficiencies, exorbitant brokerage costs, and lack of transparency inherent in traditional rental markets. "
        "By leveraging the full capabilities of the Python-based Django framework, SQLite/MySQL relational database management, and responsive "
        "Bootstrap frontend design, the project delivers a functional, user-centric, and secure rental platform."
    )
    add_body(
        "Key milestones accomplished during the project lifecycle include:"
    )
    add_bullet("1. Direct Peer-to-Peer Marketplace: ", "Eliminated predatory intermediary brokers, establishing a direct bridge between landlords and tenants.")
    add_bullet("2. Comprehensive Role Segregation: ", "Implemented dedicated interactive dashboards for Property Owners and Tenants with custom permissions.")
    add_bullet("3. Dynamic Property & Media Management: ", "Empowered owners to showcase listings with high-resolution imagery and granular property specifications.")
    add_bullet("4. Robust Web Security: ", "Enforced CSRF protection, PBKDF2 password cryptography, and SQL injection prevention across all data pathways.")
    add_bullet("5. High Performance Execution: ", "Achieved sub-100 millisecond response latencies under concurrent test loads.")

    add_body(
        "In conclusion, Rentora provides a scalable, extensible foundation for next-generation rental housing platforms. The experience gained during "
        "its conception, architecture, coding, and testing has significantly enhanced our software engineering capabilities, preparing the team "
        "for complex full-stack industry challenges."
    )

    doc.add_page_break()

    # ==========================================
    # REFERENCES
    # ==========================================
    add_h1("REFERENCES")

    refs = [
        "[1] A. Holovaty and J. Kaplan-Moss, The Definitive Guide to Django: Web Development Done Right, 2nd ed. New York: Apress, 2009.",
        "[2] M. Pilgrim, Dive Into Python 3, 2nd ed. New York: Apress, 2010.",
        "[3] Django Software Foundation, \"Django Documentation - The Web Framework for Perfectionists with Deadlines,\" Available: https://docs.djangoproject.com/en/3.2/, 2023.",
        "[4] R. Fielding, \"Architectural Styles and the Design of Network-based Software Architectures,\" Ph.D. dissertation, University of California, Irvine, 2000.",
        "[5] E. Freeman, E. Robson, B. Bates, and K. Sierra, Head First Design Patterns, 2nd ed. Sebastopol, CA: O'Reilly Media, 2020.",
        "[6] OWASP Foundation, \"OWASP Top 10 Web Application Security Risks,\" Available: https://owasp.org/www-project-top-ten/, 2021.",
        "[7] M. Fowler, Patterns of Enterprise Application Architecture. Boston, MA: Addison-Wesley, 2002.",
        "[8] K. Beck et al., \"Manifesto for Agile Software Development,\" Available: https://agilemanifesto.org/, 2001.",
        "[9] W3C, \"HTML5: A Vocabulary and Associated APIs for HTML and XHTML,\" W3C Recommendation, 2014.",
        "[10] M. Otto and J. Thornton, \"Bootstrap 5 Documentation - Build fast, responsive sites,\" Available: https://getbootstrap.com/docs/5.0/, 2022.",
        "[11] PostgreSQL Global Development Group, \"PostgreSQL Relational Database Documentation,\" Available: https://www.postgresql.org/docs/, 2023.",
        "[12] Government of India - Ministry of Housing and Urban Affairs, \"Model Tenancy Act Guidelines and Framework,\" 2021."
    ]

    for ref in refs:
        p = doc.add_paragraph()
        format_p(p, space_before=2, space_after=6, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        r = p.add_run(ref)
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)

    doc.add_page_break()

    # ==========================================
    # APPENDIX A: FILE STRUCTURE & SOURCE CODE
    # ==========================================
    add_h1("APPENDIX – A\nFILE STRUCTURE OF PROJECT & CORE SOURCE CODE")

    add_body(
        "This appendix documents the structural organization and representative core code modules of the Rentora application."
    )

    add_h2("A.1 Project Directory Tree")
    tree_text = (
        "Rental_Management_System/\n"
        "│\n"
        "├── rental1/                          # Main Django Project Root\n"
        "│   ├── manage.py                     # Django Command-Line Utility\n"
        "│   ├── db.sqlite3                    # Primary Relational Database\n"
        "│   │\n"
        "│   ├── rental1/                      # Project Configuration Package\n"
        "│   │   ├── __init__.py\n"
        "│   │   ├── settings.py               # Global Project Settings, Media & Database\n"
        "│   │   ├── urls.py                   # Master URL Routing & Media Serving\n"
        "│   │   ├── asgi.py\n"
        "│   │   └── wsgi.py                   # WSGI Application Interface\n"
        "│   │\n"
        "│   ├── rental_system/                # Core Business Logic Application\n"
        "│   │   ├── admin.py                  # Admin Panel Registrations\n"
        "│   │   ├── models.py                 # Domain Models (Property, User, Owner)\n"
        "│   │   ├── views.py                  # Controllers & HTTP Handlers\n"
        "│   │   ├── forms.py                  # Form Validation Classes\n"
        "│   │   └── migrations/               # Database Migration Files\n"
        "│   │\n"
        "│   ├── templates/                    # Presentation Layer (HTML Templates)\n"
        "│   │   ├── home.py / index.html      # Landing Page & Search Bar\n"
        "│   │   ├── login.html                # User Authentication Form\n"
        "│   │   ├── register.html             # Owner Registration\n"
        "│   │   ├── register_visitor.html     # Visitor Registration\n"
        "│   │   ├── addProperty.html          # Property Creation Form with Media\n"
        "│   │   ├── view_property.html        # Detailed Property View & Cards\n"
        "│   │   ├── owner_dashboard.html      # Landlord Management Console\n"
        "│   │   └── visitor_dashboard.html    # Tenant Exploration Console\n"
        "│   │\n"
        "│   ├── static/                       # Static Assets\n"
        "│   │   ├── css/                      # Custom Cascading Style Sheets\n"
        "│   │   ├── js/                       # Interactive JavaScript Files\n"
        "│   │   └── images/                   # UI Assets (Hero Banner, Icons)\n"
        "│   │\n"
        "│   └── media/                        # User Upload Directory (Property Photos)\n"
        "│       └── property_images/          # Uploaded Listing Pictures\n"
        "│\n"
        "├── report_figures/                   # Architectural Diagrams & UI Screen Proofs\n"
        "└── extracted_logos/                  # University & Institute Official Seals"
    )

    p_tree = doc.add_paragraph()
    format_p(p_tree, space_before=2, space_after=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    r_tree = p_tree.add_run(tree_text)
    r_tree.font.name = "Courier New"
    r_tree.font.size = Pt(8.5)
    r_tree.font.color.rgb = RGBColor(15, 23, 42)

    add_h2("A.2 Core Models (rental_system/models.py)")
    code_models = (
        "from django.db import models\n"
        "from django.contrib.auth.models import AbstractUser\n\n"
        "class User(AbstractUser):\n"
        "    is_owner = models.BooleanField(default=False)\n"
        "    is_visitor = models.BooleanField(default=False)\n\n"
        "class Property(models.Model):\n"
        "    PROPERTY_TYPES = [\n"
        "        ('Apartment', 'Apartment'),\n"
        "        ('Independent House', 'Independent House'),\n"
        "        ('Villa', 'Villa'),\n"
        "        ('Studio', 'Studio'),\n"
        "        ('Commercial', 'Commercial'),\n"
        "    ]\n"
        "    owner = models.ForeignKey(User, on_delete=models.CASCADE, related_name='properties')\n"
        "    title = models.CharField(max_length=200, default='Spacious Rental Home')\n"
        "    property_type = models.CharField(max_length=50, choices=PROPERTY_TYPES, default='Apartment')\n"
        "    address = models.TextField()\n"
        "    rent_amount = models.DecimalField(max_digits=10, decimal_places=2)\n"
        "    security_deposit = models.DecimalField(max_digits=10, decimal_places=2, default=0.0)\n"
        "    bedrooms = models.IntegerField(default=1)\n"
        "    bathrooms = models.IntegerField(default=1)\n"
        "    area_sqft = models.IntegerField(default=500)\n"
        "    contact_number = models.CharField(max_length=20, default='fuzimarcus@gmail.com')\n"
        "    image = models.ImageField(upload_to='property_images/', blank=True, null=True)\n"
        "    is_available = models.BooleanField(default=True)\n"
        "    created_at = models.DateTimeField(auto_now_add=True)\n\n"
        "    def __str__(self):\n"
        "        return f\"{self.title} - ₹{self.rent_amount}/mo\""
    )

    p_c1 = doc.add_paragraph()
    format_p(p_c1, space_before=2, space_after=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    r_c1 = p_c1.add_run(code_models)
    r_c1.font.name = "Courier New"
    r_c1.font.size = Pt(8.5)
    r_c1.font.color.rgb = RGBColor(30, 41, 59)

    add_h2("A.3 Master URL Routing (rental1/urls.py)")
    code_urls = (
        "from django.conf.urls import url\n"
        "from django.conf import settings\n"
        "from django.conf.urls.static import static\n"
        "from django.contrib import admin\n"
        "from rental_system import views\n\n"
        "urlpatterns = [\n"
        "    url(r'^admin/',               admin.site.urls),\n"
        "    url(r'^$',                    views.home,             name='home'),\n"
        "    url(r'^login/?$',             views.login_user,       name='login'),\n"
        "    url(r'^login/dashboard/?$',   views.dashboard,        name='dashboard'),\n"
        "    url(r'^login/ownerDashboard/?$', views.owner_dashboard, name='ownerdashboard'),\n"
        "    url(r'^dashboard/?$',         views.dashboard,        name='dashboard'),\n"
        "    url(r'^register/?$',          views.register,         name='register'),\n"
        "    url(r'^registerVisitor/?$',   views.register_visitor, name='registerVis'),\n"
        "    url(r'^login/addProperty/?$', views.add_property,     name='addProperty'),\n"
        "    url(r'^addProperty/?$',       views.add_property,     name='addPropertyDirect'),\n"
        "    url(r'^logout/?$',            views.logout_user,      name='logout'),\n"
        "    url(r'^viewProp/(?P<id>[0-9]+)/?$', views.view_prop,  name='prop'),\n"
        "] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)"
    )

    p_c2 = doc.add_paragraph()
    format_p(p_c2, space_before=2, space_after=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    r_c2 = p_c2.add_run(code_urls)
    r_c2.font.name = "Courier New"
    r_c2.font.size = Pt(8.5)
    r_c2.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_page_break()

    # ==========================================
    # APPENDIX B: ABSTRACT IN HINDI
    # ==========================================
    add_h1("APPENDIX – B\nABSTRACT IN HINDI\nसारांश")

    hindi_abstract_1 = (
        "समकालीन शहरी परिवेश में किराए के मकानों की खोज और किरायेदारी प्रबंधन की प्रक्रिया आज भी अत्यधिक असंगठित दलालों (ब्रोकर्स), "
        "अनौपचारिक विज्ञापनों और मौखिक सूचनाओं पर निर्भर है। इस पारंपरिक प्रणाली के कारण किराएदारों को भारी ब्रोकरेज शुल्क चुकाना पड़ता है, "
        "सटीक जानकारी का अभाव रहता है, तथा फर्जी व भ्रामक विज्ञापनों का सामना करना पड़ता है। इसके अलावा, एक ऐसा एकीकृत डिजिटल प्लेटफॉर्म "
        "उपलब्ध न होना जहाँ मकानों की प्रामाणिक तस्वीरें, कमरों का विवरण, तथा मकान मालिक व किराएदार के बीच सीधा संपर्क हो सके, इस समस्या को "
        "और अधिक जटिल बना देता है। इन सभी चुनौतियों का प्रभावी समाधान करने के उद्देश्य से 'रेंटोरा' (Rentora) नामक एक आधुनिक, वेब-आधारित "
        "रेंटल मैनेजमेंट सिस्टम विकसित किया गया है।"
    )
    hindi_abstract_2 = (
        "रेंटोरा का तकनीकी ढांचा उच्च प्रदर्शनकारी पायथन (Python) और जैंगो (Django) फ्रेमवर्क के 'मॉडल-व्यू-टेम्पलेट' (MVT) सिद्धांत पर आधारित है। "
        "यह प्रणाली भूमिका-आधारित पहुँच नियंत्रण (Role-Based Access Control) को लागू करती है, जिसमें मकान मालिकों और किराएदारों/आगंतुकों के लिए "
        "पृथक और सुरक्षित डैशबोर्ड प्रदान किए गए हैं। मकान मालिक अपनी संपत्तियों का विवरण—जैसे संपत्ति का प्रकार (अपार्टमेंट, विला, स्वतंत्र घर), "
        "मासिक किराया, सुरक्षा राशि, कमरे और स्नानघर की संख्या, क्षेत्रफल (वर्ग फुट), पता तथा उच्च गुणवत्ता वाली तस्वीरें आसानी से अपलोड और प्रबंधित "
        "कर सकते हैं। वहीं दूसरी ओर, किराएदार बिना किसी बिचौलिए के सीधे संपत्तियों को खोज सकते हैं, तस्वीरें देख सकते हैं तथा बुकिंग संबंधी पूछताछ कर सकते हैं।"
    )
    hindi_abstract_3 = (
        "डेटा की सुरक्षा और अखंडता हेतु रिलेशनल डेटाबेस आर्किटेक्चर (SQLite/MySQL) और जैंगो ओआरएम (ORM) का उपयोग किया गया है। साइबर सुरक्षा को मजबूत "
        "बनाने के लिए क्रॉस-साइट रिक्वेस्ट फोर्जरी (CSRF) टोकन, एसक्यूएल इंजेक्शन से सुरक्षा, तथा PBKDF2 और SHA-256 पर आधारित पासवर्ड हैशिंग का प्रावधान "
        "किया गया है। इसका यूजर इंटरफेस पूरी तरह से रिस्पॉन्सिव (मोबाइल-फ्रेंडली) है, जो स्मार्टफोन, टैबलेट और कंप्यूटर सभी पर सुचारू रूप से कार्य करता है।"
    )
    hindi_abstract_4 = (
        "विभिन्न परीक्षणों और लेटेंसी बेंचमार्क में रेंटोरा ने 100 मिलीसेकंड से भी कम समय में प्रतिक्रिया देकर अपनी उत्कृष्ट गति और विश्वसनीयता सिद्ध की है। "
        "संक्षेप में, रेंटोरा आधुनिक वेब इंजीनियरिंग, मजबूत डेटाबेस प्रबंधन और सुरक्षा नियमों का एक आदर्श उदाहरण प्रस्तुत करता है, जो पारंपरिक दलाली "
        "प्रथा को समाप्त कर भविष्य में ऑनलाइन किराया भुगतान और डिजिटल लीज एग्रीमेंट जैसी सुविधाओं के लिए एक सुदृढ़ आधार प्रदान करता है।"
    )

    for p_hi in [hindi_abstract_1, hindi_abstract_2, hindi_abstract_3, hindi_abstract_4]:
        p = doc.add_paragraph()
        format_p(p, space_before=2, space_after=8, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        r = p.add_run(p_hi)
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_page_break()

    # ==========================================
    # APPENDIX C: VITAE
    # ==========================================
    add_h1("APPENDIX – C\nVITAE (AUTHOR BIOGRAPHIES)")

    # Vitae 1: Tushar Singh Bisht
    add_h2("Author Biography: Tushar Singh Bisht")
    add_body(
        "The author, Tushar Singh Bisht, was born in Uttarakhand, India. He completed his High School education "
        "with exceptional academic merit and successfully cleared his Intermediate (Senior Secondary) examination with "
        "a specialization in Science and Mathematics."
    )
    add_body(
        "He joined Dr. APJ. Abdul Kalam Institute of Technology, Tanakpur (Champawat), a campus institute of "
        "Veer Madho Singh Bhandari Uttarakhand Technical University, Dehradun, to pursue his Bachelor of Technology (B.Tech) "
        "degree in Computer Science and Engineering. His technical interests encompass Full-Stack Web Development, "
        "Python Frameworks, Database Optimization, and Distributed Systems."
    )
    add_body(
        "Address for Communication:\n"
        "Department of Computer Science and Engineering,\n"
        "Dr. APJ. Abdul Kalam Institute of Technology, Tanakpur, Champawat, Uttarakhand - 262309\n"
        "Email: bishttushar01@gmail.com\n"
        "Contact: +91 9838565455"
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Vitae 2: Ayush Kumar Jha
    add_h2("Author Biography: Ayush Kumar Jha")
    add_body(
        "The author, Ayush Kumar Jha, was born in India. He completed his secondary schooling and Intermediate "
        "examinations with distinction, demonstrating strong foundations in Mathematics and Computer Applications."
    )
    add_body(
        "He enrolled at Dr. APJ. Abdul Kalam Institute of Technology, Tanakpur (Champawat), Campus Institute of "
        "Veer Madho Singh Bhandari Uttarakhand Technical University, Dehradun, for the Bachelor of Technology (B.Tech) "
        "program in Computer Science and Engineering. His core areas of interest include Frontend UI/UX Engineering, "
        "Client-Side Performance Optimization, Responsive Web Technologies, and Graphic Design."
    )
    add_body(
        "Address for Communication:\n"
        "Department of Computer Science and Engineering,\n"
        "Dr. APJ. Abdul Kalam Institute of Technology, Tanakpur, Champawat, Uttarakhand - 262309\n"
        "Email: jhaayush@gmail.com\n"
        "Contact: +91 6280326249"
    )

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Vitae 3: Abhishek Kumar Arya
    add_h2("Author Biography: Abhishek Kumar Arya")
    add_body(
        "The author, Abhishek Kumar Arya, was born in India. He accomplished his High School and Intermediate "
        "studies with high academic standing, focusing on Physics, Chemistry, and Mathematics."
    )
    add_body(
        "He joined Dr. APJ. Abdul Kalam Institute of Technology, Tanakpur (Champawat), Campus Institute of "
        "Veer Madho Singh Bhandari Uttarakhand Technical University, Dehradun, for his Bachelor of Technology (B.Tech) "
        "in Computer Science and Engineering. His primary engineering interests include Cybersecurity, Secure Web Protocols, "
        "Cloud Infrastructure, and Automated Testing Pipelines."
    )
    add_body(
        "Address for Communication:\n"
        "Department of Computer Science and Engineering,\n"
        "Dr. APJ. Abdul Kalam Institute of Technology, Tanakpur, Champawat, Uttarakhand - 262309\n"
        "Email: abhishekkumararya@gmail.com\n"
        "Contact: +91 9389842387"
    )

    # Save document
    out_docx_path = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\Rentora_Minor_Project_Report.docx"
    doc.save(out_docx_path)
    print(f"Report generated successfully and saved at: {out_docx_path}")

if __name__ == "__main__":
    create_rentora_report()
