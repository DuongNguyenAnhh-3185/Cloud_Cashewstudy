# -*- coding: utf-8 -*-
"""
Script tạo tài liệu Word hoàn chỉnh:
BÁO CÁO PHÂN TÍCH HỆ THỐNG VÀ LẬP PHƯƠNG ÁN CHUYỂN ĐỔI, TÍCH HỢP ĐIỆN TOÁN ĐÁM MÂY (CLOUD)
CHO ỨNG DỤNG QUẢN LÝ TÀI LIỆU (DOCUMENT MANAGEMENT SYSTEM)
Tên tệp xuất: phan_tich_va_lap_phuong_an.docx
"""

import os
import sys

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import matplotlib.pyplot as plt
import matplotlib.patches as patches
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

WORKSPACE_DIR = r"d:\Nam4\Mobile\TH1"
IMG_DIR = os.path.join(WORKSPACE_DIR, "assets", "diagrams")
os.makedirs(IMG_DIR, exist_ok=True)
DOCX_OUT = os.path.join(WORKSPACE_DIR, "phan_tich_va_lap_phuong_an.docx")

# ==============================================================================
# PHẦN 1: TẠO CÁC SƠ ĐỒ ĐỒ HỌA CHUYÊN NGHIỆP (300 DPI)
# ==============================================================================

def generate_traditional_vs_cloud_arch_diagram():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 7), dpi=300)
    fig.patch.set_facecolor('#F8FAFC')

    # BÊN TRÁI: KIẾN TRÚC TRUYỀN THỐNG (ON-PREMISE / CLIENT-CENTRIC)
    ax1.set_xlim(0, 100)
    ax1.set_ylim(0, 100)
    ax1.axis('off')
    ax1.set_facecolor('#F8FAFC')
    ax1.text(50, 96, "MÔ HÌNH TRUYỀN THỐNG (ON-PREMISE)", ha='center', fontsize=11, fontweight='bold', color='#991B1B')
    ax1.text(50, 92, "(Cục bộ / Giới hạn mở rộng / Nghẽn băng thông)", ha='center', fontsize=8.5, color='#7F1D1D', style='italic')

    # Frontend
    b_fe1 = patches.FancyBboxPatch((10, 74), 80, 14, boxstyle="round,pad=0.8,rounding_size=2",
                                  facecolor="#DBEAFE", edgecolor="#1D4ED8", linewidth=1.8)
    ax1.add_patch(b_fe1)
    ax1.text(50, 83, "CLIENT / FRONTEND", ha='center', va='center', fontsize=10, fontweight='bold', color="#1E3A8A")
    ax1.text(50, 77.5, "Mobile App (Flutter) / Web Browser", ha='center', va='center', fontsize=8.5, color="#2563EB")

    # Backend
    b_be1 = patches.FancyBboxPatch((10, 48), 80, 16, boxstyle="round,pad=0.8,rounding_size=2",
                                  facecolor="#FEE2E2", edgecolor="#DC2626", linewidth=2)
    ax1.add_patch(b_be1)
    ax1.text(50, 59, "ON-PREMISE BACKEND SERVER", ha='center', va='center', fontsize=10, fontweight='bold', color="#991B1B")
    ax1.text(50, 53, "API Server (Node/Spring) + Service Logic", ha='center', va='center', fontsize=8.5, color="#B91C1C")
    ax1.text(50, 49.5, "⚠️ ĐIỂM NGHẼN: Chịu tải cả Metadata + File nhị phân", ha='center', va='center', fontsize=7.5, fontweight='bold', color="#7F1D1D")

    # Database & Storage
    b_db1 = patches.FancyBboxPatch((10, 22), 37, 16, boxstyle="round,pad=0.8,rounding_size=2",
                                  facecolor="#FEF3C7", edgecolor="#D97706", linewidth=1.8)
    ax1.add_patch(b_db1)
    ax1.text(28.5, 32, "LOCAL DATABASE", ha='center', va='center', fontsize=9, fontweight='bold', color="#92400E")
    ax1.text(28.5, 26, "SQLite / MySQL Cục bộ\n(Giới hạn dung lượng)", ha='center', va='center', fontsize=7.5, color="#B45309")

    b_st1 = patches.FancyBboxPatch((53, 22), 37, 16, boxstyle="round,pad=0.8,rounding_size=2",
                                  facecolor="#FEE2E2", edgecolor="#B91C1C", linewidth=1.8)
    ax1.add_patch(b_st1)
    ax1.text(71.5, 32, "LOCAL FILE STORAGE", ha='center', va='center', fontsize=9, fontweight='bold', color="#991B1B")
    ax1.text(71.5, 26, "Ổ cứng đĩa vật lý / NAS\n(Nguy cơ hỏng hóc/Disk Full)", ha='center', va='center', fontsize=7.5, color="#7F1D1D")

    # Risk box
    b_risk = patches.FancyBboxPatch((10, 3), 80, 13, boxstyle="round,pad=0.5,rounding_size=1",
                                   facecolor="#FEF2F2", edgecolor="#F87171", linewidth=1.2, linestyle='--')
    ax1.add_patch(b_risk)
    ax1.text(50, 11, "HẠN CHẾ CỐ HỮU:", ha='center', fontsize=8, fontweight='bold', color="#991B1B")
    ax1.text(50, 6.5, "• Dễ mất dữ liệu (SPOF) • Chi phí CapEx cao\n• Không mở rộng linh hoạt • Nghẽn băng thông upload", 
             ha='center', fontsize=7.5, color="#7F1D1D")

    arrow_p = dict(arrowstyle="<|-|>", color="#DC2626", lw=2, mutation_scale=12)
    ax1.annotate('', xy=(50, 64), xytext=(50, 74), arrowprops=arrow_p)
    ax1.text(51, 69, "HTTP Request + File Bytes (Nặng)", fontsize=7.5, color="#991B1B", va='center')

    ax1.annotate('', xy=(28.5, 38), xytext=(35, 48), arrowprops=arrow_p)
    ax1.annotate('', xy=(71.5, 38), xytext=(65, 48), arrowprops=arrow_p)

    # BÊN PHẢI: KIẾN TRÚC TÍCH HỢP CLOUD (MODERN CLOUD DMS)
    ax2.set_xlim(0, 100)
    ax2.set_ylim(0, 100)
    ax2.axis('off')
    ax2.set_facecolor('#F8FAFC')
    ax2.text(50, 96, "MÔ HÌNH TÍCH HỢP CLOUD (CLOUD-INTEGRATED)", ha='center', fontsize=11, fontweight='bold', color='#15803D')
    ax2.text(50, 92, "(Đàn hồi vô hạn / Tách biệt dữ liệu / Độ bền 99.999999999%)", ha='center', fontsize=8.5, color='#166534', style='italic')

    # Frontend
    b_fe2 = patches.FancyBboxPatch((10, 74), 80, 14, boxstyle="round,pad=0.8,rounding_size=2",
                                  facecolor="#DBEAFE", edgecolor="#1D4ED8", linewidth=1.8)
    ax2.add_patch(b_fe2)
    ax2.text(50, 83, "CLIENT (Flutter Mobile / Web App)", ha='center', va='center', fontsize=10, fontweight='bold', color="#1E3A8A")
    ax2.text(50, 77.5, "Tải trực tiếp lên Cloud qua Pre-signed URL", ha='center', va='center', fontsize=8.5, color="#2563EB")

    # Cloud Backend
    b_be2 = patches.FancyBboxPatch((10, 48), 38, 16, boxstyle="round,pad=0.8,rounding_size=2",
                                  facecolor="#DCFCE7", edgecolor="#16A34A", linewidth=1.8)
    ax2.add_patch(b_be2)
    ax2.text(29, 58, "CLOUD BACKEND", ha='center', va='center', fontsize=9.5, fontweight='bold', color="#166534")
    ax2.text(29, 52, "AWS ECS / Cloud Run\n(Chỉ xử lý Auth & Meta)", ha='center', va='center', fontsize=7.5, color="#15803D")

    # Cloud Storage S3
    b_s3 = patches.FancyBboxPatch((52, 48), 38, 16, boxstyle="round,pad=0.8,rounding_size=2",
                                 facecolor="#FEF3C7", edgecolor="#D97706", linewidth=2)
    ax2.add_patch(b_s3)
    ax2.text(71, 58, "OBJECT STORAGE (S3)", ha='center', va='center', fontsize=9.5, fontweight='bold', color="#92400E")
    ax2.text(71, 52, "Amazon S3 / Azure Blob\n(Lưu tệp nhị phân - Vô hạn)", ha='center', va='center', fontsize=7.5, color="#B45309")

    # Database Cloud
    b_db2 = patches.FancyBboxPatch((10, 22), 38, 16, boxstyle="round,pad=0.8,rounding_size=2",
                                  facecolor="#E0E7FF", edgecolor="#4338CA", linewidth=1.8)
    ax2.add_patch(b_db2)
    ax2.text(29, 32, "CLOUD DATABASE (RDS)", ha='center', va='center', fontsize=9, fontweight='bold', color="#3730A3")
    ax2.text(29, 26, "Amazon RDS PostgreSQL\n(Multi-AZ, Tự sao lưu)", ha='center', va='center', fontsize=7.5, color="#4338CA")

    # CDN CloudFront
    b_cdn = patches.FancyBboxPatch((52, 22), 38, 16, boxstyle="round,pad=0.8,rounding_size=2",
                                  facecolor="#F3E8FF", edgecolor="#7E22CE", linewidth=1.8)
    ax2.add_patch(b_cdn)
    ax2.text(71, 32, "EDGE CDN (CloudFront)", ha='center', va='center', fontsize=9, fontweight='bold', color="#6B21A8")
    ax2.text(71, 26, "Phân phối tài liệu toàn cầu\n(Độ trễ < 50ms, Cache)", ha='center', va='center', fontsize=7.5, color="#7E22CE")

    # Benefits box
    b_ben = patches.FancyBboxPatch((10, 3), 80, 13, boxstyle="round,pad=0.5,rounding_size=1",
                                  facecolor="#F0FDF4", edgecolor="#4ADE80", linewidth=1.2)
    ax2.add_patch(b_ben)
    ax2.text(50, 11, "LỢI ÍCH VƯỢT TRỘI:", ha='center', fontsize=8, fontweight='bold', color="#166534")
    ax2.text(50, 6.5, "• Độ bền 99.999999999% • Tách luồng upload trực tiếp\n• Mã hóa AES-256 • Chi phí tối ưu Pay-as-you-go", 
             ha='center', fontsize=7.5, color="#15803D")

    # Arrows for Cloud
    ax2.annotate('', xy=(29, 64), xytext=(29, 74), arrowprops=dict(arrowstyle="-|>", color="#16A34A", lw=2, mutation_scale=12))
    ax2.text(20, 69, "1. Xin Pre-signed URL", fontsize=7, color="#166534")

    ax2.annotate('', xy=(71, 64), xytext=(71, 74), arrowprops=dict(arrowstyle="-|>", color="#D97706", lw=2.5, mutation_scale=14))
    ax2.text(72, 69, "2. Upload File trực tiếp (S3)", fontsize=7, color="#D97706", fontweight='bold')

    ax2.annotate('', xy=(29, 38), xytext=(29, 48), arrowprops=dict(arrowstyle="<|-|>", color="#4338CA", lw=1.8, mutation_scale=10))
    ax2.annotate('', xy=(71, 38), xytext=(71, 48), arrowprops=dict(arrowstyle="<|-|>", color="#7E22CE", lw=1.8, mutation_scale=10))

    img_path = os.path.join(IMG_DIR, "traditional_vs_cloud_arch.png")
    plt.tight_layout()
    plt.savefig(img_path, dpi=300, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close()
    return img_path

def generate_cloud_data_flow_diagram():
    fig, ax = plt.subplots(figsize=(11, 6.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')

    ax.text(50, 96, "LUỒNG DỮ LIỆU TẢI LÊN & TRUY XUẤT TÀI LIỆU TRÊN NỀN TẢNG CLOUD (PRE-SIGNED URL)",
            ha='center', fontsize=11, fontweight='bold', color='#1E3A8A')
    ax.text(50, 92, "Cơ chế Direct-to-S3 Upload giải phóng hoàn toàn tắc nghẽn băng thông của Backend API",
            ha='center', fontsize=8.5, color='#475569', style='italic')

    # Entities
    b_client = patches.FancyBboxPatch((4, 35), 22, 45, boxstyle="round,pad=1,rounding_size=2",
                                     facecolor="#DBEAFE", edgecolor="#2563EB", linewidth=2)
    ax.add_patch(b_client)
    ax.text(15, 65, "CLIENT\n(Mobile / Web)", ha='center', va='center', fontsize=10.5, fontweight='bold', color="#1E3A8A")
    ax.text(15, 47, "• Chọn tệp đính kèm\n• Gửi metadata\n• Upload nhị phân\n• Nhận thông báo", 
            ha='center', va='center', fontsize=8, color="#1E40AF")

    b_api = patches.FancyBboxPatch((38, 55), 24, 25, boxstyle="round,pad=1,rounding_size=2",
                                  facecolor="#DCFCE7", edgecolor="#16A34A", linewidth=2)
    ax.add_patch(b_api)
    ax.text(50, 72, "BACKEND API\n(AWS ECS / Fargate)", ha='center', va='center', fontsize=10, fontweight='bold', color="#166534")
    ax.text(50, 62, "• Xác thực JWT\n• Sinh Pre-signed URL\n• Lưu metadata vào RDS", 
            ha='center', va='center', fontsize=8, color="#15803D")

    b_db = patches.FancyBboxPatch((38, 15), 24, 25, boxstyle="round,pad=1,rounding_size=2",
                                 facecolor="#E0E7FF", edgecolor="#4338CA", linewidth=2)
    ax.add_patch(b_db)
    ax.text(50, 31, "CLOUD DATABASE\n(Amazon RDS PostgreSQL)", ha='center', va='center', fontsize=9.5, fontweight='bold', color="#3730A3")
    ax.text(50, 22, "• Lưu UUID, Title, Path S3\n• Audit logs, Status", 
            ha='center', va='center', fontsize=8, color="#4338CA")

    b_s3 = patches.FancyBboxPatch((74, 55), 22, 25, boxstyle="round,pad=1,rounding_size=2",
                                 facecolor="#FEF3C7", edgecolor="#D97706", linewidth=2)
    ax.add_patch(b_s3)
    ax.text(85, 72, "AMAZON S3\n(Object Storage)", ha='center', va='center', fontsize=10, fontweight='bold', color="#92400E")
    ax.text(85, 62, "• Lưu trữ tệp vật lý\n• Mã hóa AES-256\n• Bucket Private 100%", 
            ha='center', va='center', fontsize=8, color="#B45309")

    b_cdn = patches.FancyBboxPatch((74, 15), 22, 25, boxstyle="round,pad=1,rounding_size=2",
                                  facecolor="#F3E8FF", edgecolor="#7E22CE", linewidth=2)
    ax.add_patch(b_cdn)
    ax.text(85, 31, "CLOUDFRONT CDN\n(Edge Network)", ha='center', va='center', fontsize=9.5, fontweight='bold', color="#6B21A8")
    ax.text(85, 22, "• Cache tài liệu tĩnh\n• Tải siêu tốc (<50ms)\n• Giảm tải cho S3", 
            ha='center', va='center', fontsize=8, color="#7E22CE")

    # Arrows with steps
    arr_f = dict(arrowstyle="-|>", color="#1D4ED8", lw=2, mutation_scale=12)
    arr_s = dict(arrowstyle="-|>", color="#D97706", lw=2.5, mutation_scale=14)
    arr_g = dict(arrowstyle="-|>", color="#16A34A", lw=2, mutation_scale=12)
    arr_p = dict(arrowstyle="-|>", color="#7E22CE", lw=2, mutation_scale=12)

    # Step 1: Client -> API (Request Upload)
    ax.annotate('', xy=(37, 72), xytext=(27, 72), arrowprops=arr_f)
    ax.text(32, 74, "1. Request Upload (Tên, Size)", fontsize=7.5, color="#1E40AF", ha='center')

    # Step 2: API -> Client (Return Pre-signed URL)
    ax.annotate('', xy=(27, 63), xytext=(37, 63), arrowprops=arr_g)
    ax.text(32, 60, "2. Trả Pre-signed PUT URL", fontsize=7.5, color="#166534", ha='center')

    # Step 3: Client -> S3 (Direct Upload)
    ax.annotate('', xy=(73, 76), xytext=(27, 76), arrowprops=arr_s)
    ax.text(50, 79, "3. Upload File trực tiếp lên S3 (HTTP PUT - Băng thông tối đa)", fontsize=8, fontweight='bold', color="#D97706", ha='center')

    # Step 4: Client -> API (Confirm success)
    ax.annotate('', xy=(37, 57), xytext=(27, 57), arrowprops=arr_f)
    ax.text(32, 53, "4. Báo hoàn tất upload", fontsize=7.5, color="#1E40AF", ha='center')

    # Step 5: API -> RDS (Save Metadata)
    ax.annotate('', xy=(50, 41), xytext=(50, 54), arrowprops=arr_g)
    ax.text(51, 47, "5. Lưu Metadata (S3 Key, UUID)", fontsize=7.5, color="#166534")

    # Step 6: S3 -> CDN (Cache)
    ax.annotate('', xy=(85, 41), xytext=(85, 54), arrowprops=arr_p)
    ax.text(86, 47, "6. Đồng bộ CDN Edge", fontsize=7.5, color="#7E22CE")

    # Step 7: Client <- CDN (Fast Download)
    ax.annotate('', xy=(27, 38), xytext=(73, 38), arrowprops=arr_p)
    ax.text(50, 35, "7. Tải / Xem tài liệu tốc độ cao qua CDN CloudFront Signed URL", fontsize=8, fontweight='bold', color="#7E22CE", ha='center')

    img_path = os.path.join(IMG_DIR, "cloud_data_flow_presigned.png")
    plt.tight_layout()
    plt.savefig(img_path, dpi=300, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close()
    return img_path

# ==============================================================================
# PHẦN 2: HELPER FUNCTIONS FORMATTING DOCX
# ==============================================================================

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_table_borders(table, color="CBD5E1"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def add_callout(doc, text, title=None, border_color="0284C7", bg_color="F0F9FF"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>
            <w:top w:val="none"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if title:
        run_title = p.add_run(f"📌 {title}\n")
        run_title.font.name = "Calibri"
        run_title.font.size = Pt(10.5)
        run_title.font.bold = True
        run_title.font.color.rgb = RGBColor(15, 76, 129)
    run_text = p.add_run(text)
    run_text.font.name = "Calibri"
    run_text.font.size = Pt(10)
    run_text.font.color.rgb = RGBColor(51, 65, 85)
    
    empty_p = doc.add_paragraph()
    empty_p.paragraph_format.space_before = Pt(0)
    empty_p.paragraph_format.space_after = Pt(4)

def format_table_header(row, col_widths, titles):
    for i, title in enumerate(titles):
        cell = row.cells[i]
        set_cell_background(cell, "1E3A8A")
        set_cell_margins(cell, top=110, bottom=110, left=140, right=140)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(title)
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        if col_widths and i < len(col_widths):
            cell.width = col_widths[i]

def add_table_row(row, col_widths, values, is_even=False):
    fill = "F8FAFC" if is_even else "FFFFFF"
    for i, val in enumerate(values):
        cell = row.cells[i]
        set_cell_background(cell, fill)
        set_cell_margins(cell, top=70, bottom=70, left=120, right=120)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(val)
        run.font.name = "Calibri"
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(30, 41, 59)
        if col_widths and i < len(col_widths):
            cell.width = col_widths[i]

def add_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(10)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(9.5)
    run.font.italic = True
    run.font.color.rgb = RGBColor(100, 116, 139)

def add_sec_h1(doc, title):
    h = doc.add_heading(level=1)
    h.paragraph_format.space_before = Pt(16)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(title)
    run.font.name = "Calibri"
    run.font.size = Pt(15)
    run.font.bold = True
    run.font.color.rgb = RGBColor(30, 58, 138)
    return h

def add_sec_h2(doc, title):
    h = doc.add_heading(level=2)
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(4)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(title)
    run.font.name = "Calibri"
    run.font.size = Pt(12.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(2, 132, 199)
    return h

def add_sec_h3(doc, title):
    h = doc.add_heading(level=3)
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(3)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(title)
    run.font.name = "Calibri"
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RGBColor(15, 23, 42)
    return h

def add_body_p(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.2
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor(51, 65, 85)
    return p

def add_bullet_p(doc, bold_prefix, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    run_b = p.add_run(bold_prefix)
    run_b.font.name = "Calibri"
    run_b.font.size = Pt(10)
    run_b.font.bold = True
    run_b.font.color.rgb = RGBColor(15, 23, 42)
    
    run_t = p.add_run(text)
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(10)
    run_t.font.color.rgb = RGBColor(51, 65, 85)
    return p

def add_code_block(doc, code_text):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "1E293B")
    set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:left w:val="single" w:sz="16" w:space="0" w:color="38BDF8"/>
            <w:top w:val="none"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run(code_text)
    run.font.name = "Consolas"
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(241, 245, 249)
    
    empty_p = doc.add_paragraph()
    empty_p.paragraph_format.space_before = Pt(0)
    empty_p.paragraph_format.space_after = Pt(4)

# ==============================================================================
# PHẦN 3: XÂY DỰNG TÀI LIỆU DOCX "phan_tich_va_lap_phuong_an.docx"
# ==============================================================================

def build_cloud_doc():
    print("-> Đang vẽ các sơ đồ kiến trúc Cloud 300 DPI...")
    arch_img = generate_traditional_vs_cloud_arch_diagram()
    flow_img = generate_cloud_data_flow_diagram()

    print("-> Đang khởi tạo tài liệu Word: phan_tich_va_lap_phuong_an.docx...")
    doc = Document()
    
    # Thiết lập lề chuẩn trang A4
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)

    # --------------------------------------------------------------------------
    # TRANG TIÊU ĐỀ / BÌA BÁO CÁO
    # --------------------------------------------------------------------------
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(10)
    p_inst.paragraph_format.space_after = Pt(2)
    r_inst = p_inst.add_run("HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG (PTIT)\n")
    r_inst.font.name = "Calibri"
    r_inst.font.size = Pt(12)
    r_inst.font.bold = True
    r_inst.font.color.rgb = RGBColor(30, 58, 138)
    
    r_dept = p_inst.add_run("KHOA CÔNG NGHỆ THÔNG TIN - BỘ MÔN MẠNG & HỆ THỐNG PHÂN TÁN\n")
    r_dept.font.name = "Calibri"
    r_dept.font.size = Pt(11)
    r_dept.font.bold = True
    r_dept.font.color.rgb = RGBColor(71, 85, 105)

    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_line.paragraph_format.space_after = Pt(28)
    r_line = p_line.add_run("─────────── ★ ───────────")
    r_line.font.color.rgb = RGBColor(148, 163, 184)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(12)
    r_title = p_title.add_run("BÁO CÁO KỸ THUẬT\n")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(30, 58, 138)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(26)
    r_sub = p_sub.add_run("PHÂN TÍCH HỆ THỐNG VÀ LẬP PHƯƠNG ÁN CHUYỂN ĐỔI,\nTÍCH HỢP ĐIỆN TOÁN ĐÁM MÂY (CLOUD) CHO ỨNG DỤNG QUẢN LÝ TÀI LIỆU\n")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(13.5)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(2, 132, 199)

    add_callout(doc, 
        "Báo cáo phân tích chuyên sâu các thành phần kiến trúc của hệ thống Quản lý tài liệu (DMS); nhận diện các điểm nghẽn nghiêm trọng khi vận hành trên hạ tầng truyền thống (On-Premise); đề xuất phương án chuyển đổi tích hợp Điện toán đám mây (Public Cloud - AWS/GCP); thiết kế sơ đồ kiến trúc luồng dữ liệu tối ưu Direct-to-S3 Upload qua Pre-signed URL; và đánh giá đa chiều về bảo mật, chi phí (Pay-as-you-go) cùng hiệu suất vận hành toàn cầu.",
        title="TÓM TẮT ĐỀ TÀI (EXECUTIVE SUMMARY)",
        border_color="0284C7", bg_color="F0F9FF"
    )

    info_tbl = doc.add_table(rows=5, cols=2)
    info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(info_tbl, "E2E8F0")
    col_w = [Inches(2.3), Inches(4.2)]
    
    meta_data = [
        ("Đề tài / Nhiệm vụ:", "Phân tích & Lập phương án tích hợp Cloud cho hệ thống Quản lý Tài liệu"),
        ("Đối tượng nghiên cứu:", "Ứng dụng Quản lý Tài liệu Học tập (Document Management System - DMS)"),
        ("Sinh viên thực hiện:", "Thu Hoài (Username GitHub: ThuHoai10)"),
        ("Mô hình đám mây đề xuất:", "Public Cloud (Amazon Web Services - S3, CloudFront, RDS, ECS)"),
        ("Tệp tài liệu lưu trữ:", "phan_tich_va_lap_phuong_an.docx"),
    ]
    for idx, (lbl, val) in enumerate(meta_data):
        row = info_tbl.rows[idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        set_cell_background(cell_lbl, "F8FAFC")
        set_cell_background(cell_val, "FFFFFF")
        set_cell_margins(cell_lbl, top=60, bottom=60, left=100, right=100)
        set_cell_margins(cell_val, top=60, bottom=60, left=100, right=100)
        
        p0 = cell_lbl.paragraphs[0]
        r0 = p0.add_run(lbl)
        r0.font.name = "Calibri"
        r0.font.size = Pt(10)
        r0.font.bold = True
        r0.font.color.rgb = RGBColor(30, 41, 59)
        
        p1 = cell_val.paragraphs[0]
        r1 = p1.add_run(val)
        r1.font.name = "Calibri"
        r1.font.size = Pt(10)
        r1.font.color.rgb = RGBColor(71, 85, 105)
        
        cell_lbl.width = col_w[0]
        cell_val.width = col_w[1]

    doc.add_page_break()

    # --------------------------------------------------------------------------
    # MỤC LỤC
    # --------------------------------------------------------------------------
    add_sec_h1(doc, "MỤC LỤC BÁO CÁO")
    toc_items = [
        "1. LIỆT KÊ VÀ PHÂN TÍCH CÁC THÀNH PHẦN CỐT LÕI CỦA ỨNG DỤNG QUẢN LÝ TÀI LIỆU",
        "    1.1. Tầng Giao diện Người dùng (Frontend / Client Layer)",
        "    1.2. Tầng Xử lý Nghiệp vụ (Backend / Application Logic Layer)",
        "    1.3. Tầng Cơ sở Dữ liệu (Database Layer - Structured Metadata)",
        "    1.4. Tầng Lưu trữ Tệp nhị phân (File Storage Layer - Unstructured Data)",
        "    1.5. Luồng dữ liệu và Ma trận tương tác giữa 4 thành phần",
        "2. XÁC ĐỊNH CÁC ĐIỂM NGHẼN VÀ HẠN CHẾ KHI VẬN HÀNH TRÊN HẠ TẦNG TRUYỀN THỐNG",
        "    2.1. Điểm nghẽn về khả năng mở rộng dung lượng (Storage Scalability Bottleneck)",
        "    2.2. Điểm lỗi đơn lẻ (SPOF) và nguy cơ mất mát dữ liệu do thiên tai/sự cố",
        "    2.3. Tắc nghẽn băng thông và quá tải tài nguyên máy chủ ứng dụng (Bandwidth Chokepoint)",
        "    2.4. Rào cản truy cập từ xa, độ trễ địa lý và thiếu cơ chế CDN",
        "    2.5. Hạn chế về bảo mật, kiểm toán và mã hóa tĩnh (Encryption at Rest)",
        "    2.6. Gánh nặng chi phí đầu tư ban đầu (CapEx) và chi phí bảo trì vận hành",
        "3. LỰA CHỌN MÔ HÌNH TRIỂN KHAI CLOUD VÀ CÁC DỊCH VỤ CỤ THỂ",
        "    3.1. Đánh giá và lựa chọn Mô hình Triển khai Cloud (Public vs. Private vs. Hybrid)",
        "    3.2. Lựa chọn Dịch vụ Lưu trữ Đối tượng (Object Storage): Amazon S3",
        "    3.3. Lựa chọn Dịch vụ Mạng Phân phối Nội dung (CDN): Amazon CloudFront",
        "    3.4. Lựa chọn Dịch vụ Cơ sở Dữ liệu Quản lý Đám mây: Amazon RDS PostgreSQL",
        "    3.5. Lựa chọn Dịch vụ Tính toán Backend: Containerized AWS ECS Fargate",
        "    3.6. Dịch vụ Quản lý Khóa và Bảo mật: AWS KMS & IAM",
        "4. THIẾT KẾ SƠ ĐỒ KIẾN TRÚC TÍCH HỢP CLOUD VÀ MÔ TẢ LUỒNG DỮ LIỆU",
        "    4.1. Sơ đồ Kiến trúc So sánh Tổng quan: Truyền thống vs. Tích hợp Cloud",
        "    4.2. Sơ đồ Luồng Dữ liệu Chi tiết: Cơ chế Direct-to-S3 Upload qua Pre-signed URL",
        "    4.3. Mô tả Quy trình Tải lên (Upload Workflow) 6 Bước",
        "    4.4. Mô tả Quy trình Tải xuống và Xem trước (Download & Preview Workflow)",
        "5. ĐÁNH GIÁ TÁC ĐỘNG VỀ BẢO MẬT, CHI PHÍ VÀ HIỆU SUẤT SAU KHI TÍCH HỢP",
        "    5.1. Đánh giá Tác động về An toàn & Bảo mật (Security & Compliance)",
        "    5.2. Đánh giá Tác động về Chi phí Đầu tư & Vận hành (Cost Optimization & OpEx)",
        "    5.3. Đánh giá Tác động về Hiệu suất & Độ sẵn sàng (Performance, Latency & High Availability)",
        "    5.4. Bảng Tổng hợp Đối chiếu Trước và Sau Chuyển đổi",
        "KẾT LUẬN & KIẾN NGHỊ TRIỂN KHAI"
    ]
    for item in toc_items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(item)
        r.font.name = "Calibri"
        r.font.size = Pt(10)
        if item.startswith("1.") or item.startswith("2.") or item.startswith("3.") or item.startswith("4.") or item.startswith("5.") or item.startswith("KẾT"):
            r.font.bold = True
            r.font.color.rgb = RGBColor(30, 58, 138)
        else:
            r.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_page_break()

    # --------------------------------------------------------------------------
    # BƯỚC 1: LIỆT KÊ VÀ PHÂN TÍCH CÁC THÀNH PHẦN CỐT LÕI
    # --------------------------------------------------------------------------
    add_sec_h1(doc, "1. LIỆT KÊ VÀ PHÂN TÍCH CÁC THÀNH PHẦN CỐT LÕI CỦA ỨNG DỤNG QUẢN LÝ TÀI LIỆU")
    
    add_body_p(doc, 
        "Hệ thống Quản lý Tài liệu (Document Management System - DMS) là một giải pháp phần mềm chuyên biệt dùng để thu nhận, theo dõi, lưu trữ và quản lý tài liệu điện tử. Khi vận hành trên hạ tầng truyền thống (hoặc mô hình On-Premise / Client-Centric như ứng dụng CashewDoc hiện tại), kiến trúc hệ thống được phân tách thành 4 thành phần trụ cột cốt lõi:"
    )

    add_sec_h2(doc, "1.1. Tầng Giao diện Người dùng (Frontend / Client Layer)")
    add_bullet_p(doc, "Bản chất & Công nghệ: ", "Lớp phần mềm tương tác trực tiếp với người dùng (Sinh viên, Giảng viên). Công nghệ sử dụng phổ biến là Flutter Mobile App (Android/iOS), React.js/Vue.js (Web Browser) hoặc Windows Native Desktop.")
    add_bullet_p(doc, "Nhiệm vụ chính: ", 
        "Hiển thị danh sách tài liệu theo môn học; cung cấp form nhập liệu (Tiêu đề, Mô tả, Môn học); tích hợp bộ chọn tệp đính kèm (File Picker) hỗ trợ đa định dạng (.pdf, .docx, .pptx, ảnh); cung cấp thanh tìm kiếm từ khóa, bộ lọc trạng thái học tập (Chưa học, Đang học, Hoàn thành); hiển thị Dashboard thống kê và tích hợp trình xem trước tài liệu (PDF/Image Viewer)."
    )
    add_bullet_p(doc, "Quản lý trạng thái (State Management): ", "Duy trì Controller/Provider để cập nhật giao diện theo thời gian thực khi có sự kiện thay đổi dữ liệu từ tầng dưới.")

    add_sec_h2(doc, "1.2. Tầng Xử lý Nghiệp vụ (Backend / Application Logic Layer)")
    add_bullet_p(doc, "Bản chất & Công nghệ: ", "Bộ não điều phối toàn bộ quy tắc nghiệp vụ (Business Rules Engine). Có thể là Service Layer nội bộ (như DocumentService, SubjectService) hoặc RESTful API Server xây dựng bằng Node.js (Express), Python (FastAPI/Django), Java (Spring Boot) chạy trên máy chủ On-Premise.")
    add_bullet_p(doc, "Kiểm định dữ liệu (Validation): ", "Kiểm tra tính hợp lệ của tiêu đề (không rỗng), bắt buộc chọn môn học hợp lệ, kiểm tra phần mở rộng tệp nằm trong danh mục cho phép và kiểm soát dung lượng tệp tải lên.")
    add_bullet_p(doc, "Máy trạng thái & Định danh phân tán: ", "Tự động sinh chuỗi định danh độc bản UUID v4 (128-bit) cho mọi tài liệu; điều khiển chuyển đổi trạng thái học tập theo chu trình khép kín.")
    add_bullet_p(doc, "Xử lý tệp & An toàn dữ liệu: ", "Phân tích metadata của tệp đính kèm (kích thước KB/MB, MIME type, Checksum hash); điều phối quy trình Xóa mềm (Soft-Delete) bằng cách cập nhật cờ is_deleted = 1 và ghi vết vào bảng nhật ký xóa delete_logs.")

    add_sec_h2(doc, "1.3. Tầng Cơ sở Dữ liệu (Database Layer - Structured Metadata)")
    add_bullet_p(doc, "Bản chất & Công nghệ: ", "Lưu trữ dữ liệu có cấu trúc (Structured Data) và siêu dữ liệu (Metadata) mô tả tài liệu. Công nghệ sử dụng là SQLite Native FFI trên thiết bị hoặc hệ quản trị CSDL quan hệ RDBMS (PostgreSQL, MySQL, SQL Server) chạy trên máy chủ nội bộ.")
    add_bullet_p(doc, "Các bảng dữ liệu cốt lõi: ", 
        "• Bảng 'documents': Lưu UUID khóa chính, tiêu đề, mô tả, khóa ngoại môn học, trạng thái học tập, timestamps tạo/sửa, cờ xóa mềm.\n"
        "• Bảng 'subjects': Lưu danh mục môn học, mã màu nhận diện, thông tin giảng viên.\n"
        "• Bảng 'attachments': Lưu thông tin tệp đính kèm (attachment_pk, document_fk, file_name, file_path, file_size, file_type).\n"
        "• Bảng 'delete_logs': Lưu vết kiểm toán lịch sử xóa (log_pk, entity_pk, entity_type, deleted_at)."
    )
    add_bullet_p(doc, "Đặc tính kỹ thuật: ", "Tuân thủ chặt chẽ tính toàn vẹn dữ liệu (ACID Transactions), liên kết khóa ngoại (Foreign Keys) và lập chỉ mục (Indexes) trên các trường tra cứu nhanh.")

    add_sec_h2(doc, "1.4. Tầng Lưu trữ Tệp nhị phân (File Storage Layer - Unstructured Data)")
    add_bullet_p(doc, "Bản chất & Cơ chế lưu trữ: ", "Lưu trữ nội dung nhị phân thực tế của các tệp tài liệu số kích thước lớn (PDF, Slide bài giảng, file mã nguồn ZIP).")
    add_bullet_p(doc, "Hạ tầng truyền thống: ", "Sử dụng phân vùng ổ đĩa cục bộ của thiết bị (Local File System / App Storage) hoặc hệ thống lưu trữ mạng nội bộ (NAS - Network Attached Storage, SAN).")
    add_bullet_p(doc, "Nguyên lý tách biệt: ", "Hệ thống tuyệt đối không lưu trực tiếp nội dung nhị phân dạng BLOB vào trong Database để tránh làm phình to CSDL. CSDL chỉ lưu chuỗi đường dẫn tệp (file_path) trỏ tới tệp vật lý trên ổ đĩa.")

    add_sec_h2(doc, "1.5. Ma Trận Tương Tác Giữa 4 Thành Phần Hệ Thống")

    comp_tbl = doc.add_table(rows=5, cols=4)
    comp_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(comp_tbl, "CBD5E1")
    comp_w = [Inches(1.5), Inches(1.8), Inches(2.2), Inches(1.5)]
    format_table_header(comp_tbl.rows[0], comp_w, ["Thành Phần", "Dạng Dữ Liệu", "Trách Nhiệm Chính", "Vị Trí Vận Hành"])
    
    comp_data = [
        ("Frontend (Client)", "UI State, User Input, View Widgets", "Thu nhận biểu mẫu, hiển thị tài liệu, xem trước tệp, quản lý trạng thái UI", "Client (Mobile App, Trình duyệt Web)"),
        ("Backend (Logic)", "Business Logic, Request/Response", "Kiểm tra hợp lệ, sinh UUID v4, phân tích metadata, điều khiển xóa mềm", "Service nội bộ hoặc App Server On-Premise"),
        ("Database (Metadata)", "Structured Data (Bảng quan hệ)", "Lưu trữ metadata, liên kết khóa ngoại, chỉ mục tìm kiếm, audit logs", "SQLite cục bộ hoặc RDBMS On-Premise"),
        ("File Storage (Binary)", "Unstructured Data (Tệp nhị phân)", "Lưu trữ nội dung tệp vật lý kích thước lớn (.pdf, .docx, .zip)", "Ổ cứng cục bộ (Local Disk / NAS)"),
    ]
    for idx, r_vals in enumerate(comp_data):
        add_table_row(comp_tbl.rows[idx + 1], comp_w, r_vals, is_even=(idx % 2 == 1))

    doc.add_paragraph()

    # --------------------------------------------------------------------------
    # BƯỚC 2: ĐIỂM NGHẼN VÀ HẠN CHẾ CỦA HẠ TẦNG TRUYỀN THỐNG
    # --------------------------------------------------------------------------
    add_sec_h1(doc, "2. XÁC ĐỊNH CÁC ĐIỂM NGHẼN VÀ HẠN CHẾ KHI VẬN HÀNH TRÊN HẠ TẦNG TRUYỀN THỐNG")
    
    add_body_p(doc, 
        "Khi hệ thống Quản lý tài liệu mở rộng quy mô (từ vài chục sinh viên lên hàng nghìn người dùng, từ vài trăm MB lên hàng Terabytes tài liệu), mô hình hạ tầng truyền thống On-Premise / Local Storage bộc lộ 6 điểm nghẽn nghiêm trọng sau:"
    )

    add_sec_h2(doc, "2.1. Điểm nghẽn về khả năng mở rộng dung lượng (Storage Scalability Bottleneck)")
    add_bullet_p(doc, "Hạn chế mở rộng phần cứng (Vertical Scaling): ", "Dung lượng lưu trữ bị giới hạn cứng bởi kích thước vật lý của ổ đĩa (HDD/SSD) hoặc thiết bị NAS. Khi số lượng tài liệu số, bài giảng slide, file PDF tăng cao, ổ cứng sẽ rơi vào trạng thái 'Disk Full'.")
    add_bullet_p(doc, "Chi phí gián đoạn (Downtime): ", "Để nâng cấp dung lượng, quản trị viên buộc phải tắt máy chủ (Downtime), mua thêm ổ cứng, gắn vào khe RAID và mở rộng phân vùng. Quy trình này tốn nhiều ngày và làm gián đoạn việc học tập của sinh viên.")

    add_sec_h2(doc, "2.2. Điểm lỗi đơn lẻ (SPOF) và nguy cơ mất mát dữ liệu do sự cố")
    add_bullet_p(doc, "Rủi ro Single Point of Failure (SPOF): ", "Mọi tệp tài liệu đều tập trung trên một ổ đĩa hoặc một cụm máy chủ cục bộ. Nếu xảy ra sự cố hỏng hóc phần cứng (bad sector, cháy nổ mạch điều khiển), toàn bộ kho tài liệu sẽ biến mất vĩnh viễn.")
    add_bullet_p(doc, "Kém năng lực phục hồi thảm họa (Disaster Recovery): ", "Mô hình truyền thống hiếm khi có cơ chế sao lưu đa địa lý (Off-site Multi-region Backup). Khi xảy ra sự cố cháy nổ phòng server, ngập lụt, hoặc bị mã độc tống tiền (Ransomware) mã hóa ổ đĩa, mục tiêu thời gian phục hồi (RTO) và mục tiêu điểm phục hồi (RPO) đều không thể đảm bảo.")

    add_sec_h2(doc, "2.3. Tắc nghẽn băng thông và quá tải máy chủ ứng dụng (Bandwidth Chokepoint)")
    add_bullet_p(doc, "Nghẽn cổ chai luồng tệp (I/O Bottleneck): ", "Trong mô hình truyền thống, khi sinh viên tải lên hoặc tải xuống tệp bài giảng (dung lượng từ 20MB đến 100MB), toàn bộ luồng byte nhị phân phải đi xuyên qua máy chủ Backend (Client -> Backend -> File Storage).")
    add_bullet_p(doc, "Cạn kiệt tài nguyên CPU/RAM: ", "Máy chủ Backend phải mở socket giữ kết nối, cấp phát bộ đệm RAM để đọc ghi dữ liệu tệp nhị phân. Khi hàng trăm sinh viên cùng nộp bài thi hoặc tải tài liệu cùng lúc, máy chủ Backend sẽ bị cạn kiệt CPU, nghẽn card mạng, dẫn đến sập toàn bộ hệ thống API tra cứu thông tin.")

    add_sec_h2(doc, "2.4. Rào cản truy cập từ xa, độ trễ địa lý và thiếu cơ chế CDN")
    add_bullet_p(doc, "Truy cập từ xa phức tạp & kém ổn định: ", "Máy chủ nội bộ yêu cầu cấu hình IP tĩnh, mở Port Forwarding, NAT hoặc bắt buộc người dùng kết nối qua mạng riêng ảo VPN chậm chạp.")
    add_bullet_p(doc, "Độ trễ cao (High Latency): ", "Người dùng truy cập từ xa ngoài khuôn viên trường học hoặc ở các tỉnh thành khác phải chịu độ trễ mạng lớn do không có mạng lưới phân phối nội dung biên (Edge CDN).")

    add_sec_h2(doc, "2.5. Hạn chế về bảo mật, kiểm toán và mã hóa tĩnh (Encryption at Rest)")
    add_bullet_p(doc, "Dữ liệu không được mã hóa mặc định: ", "Tệp tài liệu lưu trên ổ đĩa On-Premise thường nằm dưới dạng tệp thô không được mã hóa. Bất kỳ ai có quyền truy cập vật lý vào ổ đĩa hoặc chiếm quyền root server đều có thể sao chép toàn bộ tài liệu mật.")
    add_bullet_p(doc, "Thiếu cơ chế phân quyền hạt mịn: ", "Hệ thống truyền thống khó kiểm soát cấp quyền truy cập tạm thời (Temporary access), thiếu chữ ký URL bảo mật và không có nhật ký kiểm toán (Audit Trail) toàn diện ghi vết mọi lượt đọc/ghi.")

    add_sec_h2(doc, "2.6. Gánh nặng chi phí đầu tư ban đầu (CapEx) và chi phí bảo trì")
    add_bullet_p(doc, "Chi phí vốn (CapEx) khổng lồ: ", "Tổ chức phải chi ngân sách lớn ngay từ đầu để mua sắm máy chủ vật lý, hệ thống lưu trữ SAN/NAS chuyên dụng, thiết bị mạng tường lửa, hệ thống điều hòa phòng server và nguồn điện dự phòng UPS.")
    add_bullet_p(doc, "Lãng phí tài nguyên: ", "Hạ tầng phải được đầu tư dự phòng cho mức tải cao nhất (Peak Traffic), nhưng phần lớn thời gian trong năm lại hoạt động dưới mức công suất, gây lãng phí lớn về điện năng và bảo trì.")

    # --------------------------------------------------------------------------
    # BƯỚC 3: LỰA CHỌN MÔ HÌNH CLOUD VÀ DỊCH VỤ CỤ THỂ
    # --------------------------------------------------------------------------
    add_sec_h1(doc, "3. LỰA CHỌN MÔ HÌNH TRIỂN KHAI CLOUD VÀ CÁC DỊCH VỤ CỤ THỂ")

    add_sec_h2(doc, "3.1. Đánh giá và Lựa chọn Mô hình Triển khai Cloud Phù hợp")
    add_body_p(doc, 
        "Để khắc phục triệt để các hạn chế của mô hình truyền thống, chúng tôi tiến hành phân tích và so sánh 3 mô hình triển khai điện toán đám mây chính:"
    )

    cloud_tbl = doc.add_table(rows=4, cols=4)
    cloud_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(cloud_tbl, "CBD5E1")
    c_w = [Inches(1.4), Inches(1.8), Inches(2.0), Inches(1.8)]
    format_table_header(cloud_tbl.rows[0], c_w, ["Mô Hình", "Đặc Điểm Kiến Trúc", "Ưu Điểm Nổi Bật", "Đánh Giá Áp Dụng"])
    
    cloud_model_data = [
        ("Public Cloud (AWS / GCP / Azure)", "Hạ tầng thuộc sở hữu của nhà cung cấp Cloud hàng đầu thế giới, chia sẻ tài nguyên an toàn đa khách hàng", "Mở rộng đàn hồi vô hạn, thanh toán theo dung lượng thực tế (Pay-as-you-go), độ bền 99.999999999%, không cần bảo trì phần cứng", "TỐI ƯU NHẤT: Rất phù hợp cho ứng dụng quản lý tài liệu học tập, tiết kiệm tối đa ngân sách và vận hành"),
        ("Private Cloud (On-Premises Cloud)", "Hạ tầng đám mây ảo hóa riêng biệt đặt tại trung tâm dữ liệu của tổ chức (OpenStack, VMware)", "Kiểm soát vật lý 100%, bảo mật nội bộ khép kín", "KHÔNG KHẢ THI: Chi phí đầu tư thiết bị và nhân sự vận hành quá đắt đỏ, không giải quyết được bài toán CapEx"),
        ("Hybrid Cloud (Lai ghép)", "Kết hợp giữa máy chủ nội bộ (lưu metadata mật) và Public Cloud (lưu trữ tệp nhị phân lớn)", "Tận dụng phần cứng cũ, tăng cường bảo mật dữ liệu nhạy cảm", "PHÙ HỢP GIAI ĐOẠN CHUYỂN TIẾP: Thích hợp nếu tổ chức đã có sẵn hạ tầng máy chủ CSDL nội bộ"),
    ]
    for idx, r_vals in enumerate(cloud_model_data):
        add_table_row(cloud_tbl.rows[idx + 1], c_w, r_vals, is_even=(idx % 2 == 1))

    doc.add_paragraph()
    add_callout(doc, 
        "QUYẾT ĐỊNH LỰA CHỌN: Lựa chọn mô hình PUBLIC CLOUD trên nền tảng Amazon Web Services (AWS) (kết hợp hỗ trợ Google Cloud Platform - GCP). Mô hình này mang lại tỷ suất lợi ích / chi phí tối đa, loại bỏ hoàn toàn chi phí bảo trì phần cứng, và sở hữu hệ sinh thái dịch vụ lưu trữ đối tượng dẫn đầu thị trường.",
        title="LỰA CHỌN MÔ HÌNH TRIỂN KHAI TỐI ƯU",
        border_color="16A34A", bg_color="F0FDF4"
    )

    add_sec_h2(doc, "3.2. Lựa chọn Dịch vụ Lưu trữ Đối tượng (Object Storage): Amazon S3")
    add_bullet_p(doc, "Dịch vụ: ", "Amazon Simple Storage Service (Amazon S3) (hoặc Google Cloud Storage - GCS).")
    add_bullet_p(doc, "Lý do lựa chọn: ", 
        "• Độ bền dữ liệu tuyệt đối (Durability): Đạt 99.999999999% (11 số 9) nhờ cơ chế tự động nhân bản dữ liệu qua tối thiểu 3 Trung tâm Dữ liệu độc lập (Availability Zones - AZs).\n"
        "• Khả năng mở rộng vô hạn: Tự động co giãn theo dung lượng thực tế từ vài Megabytes đến hàng Petabytes mà không cần cấu hình trước.\n"
        "• Tính năng Pre-signed URLs: Cho phép sinh đường dẫn có chữ ký số và thời hạn hết hạn tự động (TTL) để Client tải tệp trực tiếp lên S3 mà không cần đi qua máy chủ Backend.\n"
        "• Tối ưu chi phí bằng S3 Lifecycle Policies: Tự động chuyển tài liệu cũ ít truy cập sang S3 Standard-Infrequent Access và S3 Glacier Flexible Deep Archive giúp tiết kiệm tới 95% chi phí."
    )

    add_sec_h2(doc, "3.3. Lựa chọn Dịch vụ Mạng Phân phối Nội dung (CDN): Amazon CloudFront")
    add_bullet_p(doc, "Dịch vụ: ", "Amazon CloudFront (hoặc Cloudflare CDN).")
    add_bullet_p(doc, "Lý do lựa chọn: ", 
        "Mạng lưới hơn 600 điểm PoP (Points of Presence) biên trên toàn cầu. Các tệp tài liệu tĩnh phổ biến được lưu tạm (Cache) ngay tại điểm gần người dùng nhất, giảm độ trễ tải tệp xuống dưới 50ms, đồng thời giảm 80% lưu lượng truy cập trực tiếp vào S3 Bucket."
    )

    add_sec_h2(doc, "3.4. Lựa chọn Dịch vụ Cơ sở Dữ liệu Đám mây: Amazon RDS for PostgreSQL")
    add_bullet_p(doc, "Dịch vụ: ", "Amazon Relational Database Service (Amazon RDS PostgreSQL).")
    add_bullet_p(doc, "Lý do lựa chọn: ", 
        "Hệ quản trị CSDL quan hệ chuẩn mực, được AWS tự động quản lý sao lưu định kỳ hàng ngày (Automated Backups), nhân bản dự phòng nóng (Multi-AZ Standby Replica) đảm bảo RPO < 5 phút và tự động mở rộng dung lượng đĩa (Storage Auto-scaling)."
    )

    add_sec_h2(doc, "3.5. Lựa chọn Dịch vụ Tính toán Backend: Containerized AWS ECS Fargate")
    add_bullet_p(doc, "Dịch vụ: ", "AWS Elastic Container Service (ECS) với Serverless Launch Type (AWS Fargate).")
    add_bullet_p(doc, "Lý do lựa chọn: ", 
        "Vận hành API Backend dưới dạng Docker Container mà không cần quản lý máy chủ ảo EC2. Hệ thống tự động co giãn số lượng Container (Auto-scaling) theo lưu lượng truy cập thực tế của sinh viên, chỉ tính phí khi có Container đang chạy."
    )

    add_sec_h2(doc, "3.6. Dịch vụ Bảo mật & Quản lý Khóa: AWS KMS & IAM")
    add_bullet_p(doc, "Dịch vụ: ", "AWS Key Management Service (KMS) & AWS Identity and Access Management (IAM).")
    add_bullet_p(doc, "Lý do lựa chọn: ", 
        "Cung cấp khả năng mã hóa dữ liệu tĩnh chuẩn AES-256 (SSE-KMS) với khóa mã hóa được quản lý chặt chẽ. Áp dụng chính sách đặc quyền tối thiểu (Least Privilege Access) cho từng dịch vụ và người dùng."
    )

    # --------------------------------------------------------------------------
    # BƯỚC 4: THIẾT KẾ KIẾN TRÚC VÀ LUỒNG DỮ LIỆU
    # --------------------------------------------------------------------------
    add_sec_h1(doc, "4. THIẾT KẾ SƠ ĐỒ KIẾN TRÚC TÍCH HỢP CLOUD VÀ MÔ TẢ LUỒNG DỮ LIỆU")

    add_sec_h2(doc, "4.1. Sơ đồ Kiến trúc So sánh Tổng quan: Truyền thống vs. Tích hợp Cloud")
    add_body_p(doc, 
        "Hình 4.1 dưới đây thể hiện sự khác biệt mang tính bước ngoặt giữa mô hình truyền thống (bên trái) và mô hình hiện đại sau khi tích hợp toàn diện Điện toán Đám mây (bên phải):"
    )

    if os.path.exists(arch_img):
        doc.add_paragraph()
        doc.add_picture(arch_img, width=Inches(6.4))
        add_caption(doc, "Hình 4.1: So sánh Kiến trúc Hệ thống DMS Truyền thống (Trái) và Kiến trúc Tích hợp Cloud (Phải)")

    add_sec_h2(doc, "4.2. Sơ đồ Luồng Dữ liệu Chi tiết: Direct-to-S3 Upload qua Pre-signed URL")
    add_body_p(doc, 
        "Để loại bỏ hoàn toàn điểm nghẽn băng thông của máy chủ Backend, chúng tôi áp dụng mẫu thiết kế kiến trúc chuẩn mực thế giới: **Mô hình Tải lên Trực tiếp qua Pre-signed URL (Direct-to-S3 Upload)**:"
    )

    if os.path.exists(flow_img):
        doc.add_paragraph()
        doc.add_picture(flow_img, width=Inches(6.4))
        add_caption(doc, "Hình 4.2: Luồng Dữ liệu Tải lên Trực tiếp qua Pre-signed URL và Tải xuống Tốc độ cao qua CDN CloudFront")

    add_sec_h2(doc, "4.3. Mô tả Chi Tiết Quy Trình Tải Lên (Upload Workflow)")
    add_body_p(doc, "Quy trình tải một tài liệu học tập mới kèm tệp đính kèm lớn (PDF/Word/Slide) diễn ra qua 6 bước khép kín:")
    add_bullet_p(doc, "Bước 1 (Client -> Backend): ", "Người dùng chọn tệp tin từ thiết bị. Ứng dụng Client gửi một yêu cầu JSON nhẹ chứa metadata (tên tệp: 'BaiGiang_Chuong1.pdf', kích thước: 45MB, MIME: 'application/pdf') tới Backend API.")
    add_bullet_p(doc, "Bước 2 (Backend Authentication & Key Generation): ", "Backend API xác thực phiên đăng nhập (JWT), kiểm tra hạn ngạch của sinh viên, sinh chuỗi UUID v4 độc bản cho tài liệu và gọi AWS SDK để tạo một 'Pre-signed PUT URL' (với thời hạn hiệu lực 5 phút, gắn kèm khóa đối tượng 'uploads/2026/uuid_baigiang.pdf').")
    add_bullet_p(doc, "Bước 3 (Backend -> Client): ", "Backend trả mã Pre-signed URL về cho ứng dụng Client qua phản hồi HTTP 200.")
    add_bullet_p(doc, "Bước 4 (Client -> Amazon S3 - Direct Upload): ", "Ứng dụng Client sử dụng giao thức HTTP PUT truyền trực tiếp toàn bộ luồng byte nhị phân 45MB lên thẳng máy chủ Amazon S3. Dữ liệu KHÔNG đi qua Backend, giải phóng 100% băng thông của server!")
    add_bullet_p(doc, "Bước 5 (Client -> Backend Confirmation): ", "Sau khi S3 trả về mã HTTP 200 OK (xác nhận đã ghi file an toàn vào S3), Client gửi một request thông báo hoàn tất kèm ETag và UUID về cho Backend.")
    add_bullet_p(doc, "Bước 6 (Backend -> RDS Database): ", "Backend API ghi bản ghi metadata tài liệu (kèm đường dẫn S3 Key) vào Amazon RDS PostgreSQL bằng 1 Transaction an toàn và gửi thông báo thành công cho Client.")

    add_sec_h2(doc, "4.4. Mô tả Chi Tiết Quy Trình Tải Xuống và Xem Trước (Download & Preview Workflow)")
    add_bullet_p(doc, "Bước 1: ", "Sinh viên bấm vào một tài liệu để xem chi tiết hoặc tải về. Client gửi yêu cầu kèm Document ID tới Backend API.")
    add_bullet_p(doc, "Bước 2: ", "Backend kiểm tra quyền truy cập của sinh viên. Nếu hợp lệ, Backend sinh một CloudFront Signed URL (hoặc S3 Pre-signed GET URL) có thời hạn ngắn (ví dụ 15 phút).")
    add_bullet_p(doc, "Bước 3: ", "Client sử dụng đường dẫn có chữ ký số để tải nội dung tệp trực tiếp từ CloudFront CDN Edge Server gần nhất với độ trễ siêu thấp (<50ms) và băng thông không giới hạn.")

    # --------------------------------------------------------------------------
    # BƯỚC 5: ĐÁNH GIÁ TÁC ĐỘNG VỀ BẢO MẬT, CHI PHÍ VÀ HIỆU SUẤT
    # --------------------------------------------------------------------------
    add_sec_h1(doc, "5. ĐÁNH GIÁ TÁC ĐỘNG VỀ BẢO MẬT, CHI PHÍ VÀ HIỆU SUẤT SAU KHI TÍCH HỢP")

    add_sec_h2(doc, "5.1. Đánh giá Tác động về An toàn & Bảo mật (Security & Compliance)")
    add_bullet_p(doc, "Mã hóa toàn diện đa tầng: ", 
        "• Mã hóa luồng truyền (In-Transit): Bắt buộc giao thức HTTPS/TLS 1.3 cho toàn bộ giao tiếp giữa Client, Backend và S3.\n"
        "• Mã hóa tại chỗ (At-Rest): Toàn bộ tệp trên S3 Bucket được tự động mã hóa bằng chuẩn quân đội AES-256 thông qua dịch vụ AWS KMS."
    )
    add_bullet_p(doc, "Chính sách Private Bucket 100%: ", 
        "Kích hoạt cờ 'Block Public Access' tuyệt đối trên S3 Bucket. Không ai có thể truy cập tệp trực tiếp qua Internet nếu không có Pre-signed URL hợp lệ được cấp bởi Backend."
    )
    add_bullet_p(doc, "Kiểm toán & Truy vết (Audit Logging): ", 
        "Mọi hành động đọc, ghi, xóa trên S3 và RDS đều được ghi vết tự động vào AWS CloudTrail, đáp ứng các tiêu chuẩn an toàn thông tin quốc tế (ISO 27001, GDPR)."
    )

    add_sec_h2(doc, "5.2. Đánh giá Tác động về Chi phí Đầu tư & Vận hành (Cost Optimization & OpEx)")
    add_bullet_p(doc, "Chuyển dịch hoàn toàn từ CapEx sang OpEx: ", 
        "Loại bỏ chi phí đầu tư thiết bị phần cứng ban đầu trị giá hàng trăm triệu đồng. Thay vào đó, chi phí vận hành biến đổi linh hoạt theo mô hình 'Dùng bao nhiêu, trả bấy nhiêu' (Pay-as-you-go)."
    )
    add_bullet_p(doc, "Đơn giá lưu trữ siêu tiết kiệm: ", 
        "Amazon S3 Standard chỉ có giá ~0.023 USD / GB / tháng. Với 100 GB tài liệu học tập của sinh viên, chi phí lưu trữ chỉ vỏn vẹn ~2.3 USD / tháng (tương đương ~58.000 VNĐ)!"
    )
    add_bullet_p(doc, "Tiết kiệm 90% chi phí dài hạn bằng Lifecycle Policies: ", 
        "Các tài liệu học tập của các học kỳ trước (sau 90 ngày) tự động chuyển sang S3 Glacier Deep Archive với mức giá siêu rẻ: 0.00099 USD / GB / tháng (chỉ 2.500 VNĐ cho 100 GB mỗi tháng)."
    )
    add_bullet_p(doc, "Tối ưu chi phí máy chủ Backend: ", 
        "Nhờ mô hình Direct Upload, máy chủ Backend không phải xử lý tải file nặng, cho phép sử dụng các gói Container cấu hình nhỏ (0.5 vCPU, 1GB RAM), giảm tới 75% chi phí thuê máy chủ."
    )

    add_sec_h2(doc, "5.3. Đánh giá Tác động về Hiệu suất & Độ sẵn sàng (Performance & High Availability)")
    add_bullet_p(doc, "Độ bền dữ liệu bất hoại: ", "Đạt 99.999999999% (11 số 9), triệt tiêu hoàn toàn nỗi lo mất mát dữ liệu do hỏng ổ đĩa.")
    add_bullet_p(doc, "Độ sẵn sàng (High Availability): ", "Cam kết SLA 99.99% thời gian hoạt động liên tục (Uptime), phục hồi thảm họa tự động đa vùng chỉ trong vài phút.")
    add_bullet_p(doc, "Khả năng chịu tải đột biến (High Throughput): ", "Amazon S3 hỗ trợ tối thiểu 3.500 requests PUT và 5.500 requests GET mỗi giây trên mỗi prefix, cho phép hàng chục nghìn sinh viên nộp bài cùng lúc mà không xảy ra hiện tượng nghẽn mạng.")
    add_bullet_p(doc, "Tốc độ truy xuất toàn cầu: ", "Nhờ CDN CloudFront, tốc độ tải tài liệu nhanh hơn gấp 3 - 5 lần so với việc truy cập vào máy chủ On-Premise truyền thống.")

    add_sec_h2(doc, "5.4. Bảng Tổng Hợp Đối Chiếu Trước và Sau Khi Tích Hợp Cloud")

    eval_tbl = doc.add_table(rows=8, cols=4)
    eval_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(eval_tbl, "CBD5E1")
    e_w = [Inches(1.4), Inches(1.8), Inches(2.0), Inches(1.8)]
    format_table_header(eval_tbl.rows[0], e_w, ["Tiêu Chí Đánh Giá", "Mô Hình Truyền Thống", "Mô Hình Tích Hợp Cloud", "Mức Độ Cải Thiện"])
    
    eval_data = [
        ("1. Dung lượng lưu trữ", "Giới hạn bởi dung lượng ổ cứng vật lý (Disk Full)", "Mở rộng đàn hồi không giới hạn (Auto-scaling)", "Đột phá: Không lo thiếu đĩa"),
        ("2. Độ bền & An toàn", "Dễ mất dữ liệu khi hỏng ổ đĩa, thiên tai (SPOF)", "Độ bền 99.999999999% (11 số 9), sao lưu đa vùng", "Tuyệt đối: Triệt tiêu rủi ro"),
        ("3. Tải máy chủ Backend", "Nghẽn băng thông, cạn RAM do file đi qua server", "Giải phóng 100% nhờ Direct Upload qua Pre-signed URL", "Vượt trội: Server chỉ xử lý JSON"),
        ("4. Tốc độ tải tệp (Latency)", "Chậm, phụ thuộc đường truyền mạng nội bộ", "Siêu tốc (<50ms) nhờ mạng lưới CDN toàn cầu", "Nhanh hơn 300% - 500%"),
        ("5. Cơ chế Bảo mật", "Tệp lưu thô, khó kiểm soát cấp quyền tạm thời", "Mã hóa AES-256 tĩnh, Pre-signed URL có thời hạn (TTL)", "Chuẩn doanh nghiệp quốc tế"),
        ("6. Chi phí Đầu tư", "CapEx lớn (Mua sắm phần cứng, phòng server)", "OpEx linh hoạt (Chỉ trả tiền theo mức sử dụng thực tế)", "Tiết kiệm 70% - 85% chi phí"),
        ("7. Phục hồi sự cố (DR)", "Thủ công, mất nhiều ngày, nguy cơ mất dữ liệu cao", "Tự động hoàn toàn (RTO < 15 phút, RPO < 5 phút)", "Đảm bảo tính liên tục 24/7"),
    ]
    for idx, r_vals in enumerate(eval_data):
        add_table_row(eval_tbl.rows[idx + 1], e_w, r_vals, is_even=(idx % 2 == 1))

    doc.add_paragraph()

    # --------------------------------------------------------------------------
    # KẾT LUẬN & KIẾN NGHỊ
    # --------------------------------------------------------------------------
    add_sec_h1(doc, "KẾT LUẬN & KIẾN NGHỊ TRIỂN KHAI")
    
    add_body_p(doc, 
        "Việc chuyển đổi và tích hợp Điện toán Đám mây cho ứng dụng Quản lý Tài liệu (DMS) không chỉ là một nâng cấp kỹ thuật đơn thuần mà là một bước chuyển dịch chiến lược mang lại giá trị bền vững cho tổ chức giáo dục / doanh nghiệp:"
    )
    add_bullet_p(doc, "1. Giải quyết triệt để các điểm nghẽn cố hữu: ", "Mô hình Direct-to-S3 Upload qua Pre-signed URL loại bỏ hoàn toàn sự tắc nghẽn băng thông tại tầng Backend, trong khi dịch vụ Amazon S3 giải phóng áp lực về giới hạn dung lượng lưu trữ vật lý.")
    add_bullet_p(doc, "2. Nâng tầm bảo mật và trải nghiệm người dùng: ", "Dữ liệu được bảo vệ bằng mã hóa chuẩn AES-256 kết hợp mạng phân phối nội dung CDN CloudFront mang lại trải nghiệm tải tài liệu mượt mà, sẵn sàng 24/7 từ bất kỳ đâu.")
    add_bullet_p(doc, "3. Tối ưu hóa hiệu quả tài chính: ", "Mô hình Pay-as-you-go kết hợp S3 Lifecycle Policies giúp tiết kiệm ngân sách khổng lồ so với việc duy trì trung tâm dữ liệu On-Premise truyền thống.")
    add_bullet_p(doc, "4. Lộ trình triển khai kiến nghị: ", "Tổ chức nên áp dụng lộ trình chuyển đổi 3 giai đoạn: (1) Di chuyển tệp tĩnh sang Amazon S3 với Pre-signed URL; (2) Chuyển đổi CSDL sang Amazon RDS PostgreSQL; (3) Tích hợp các dịch vụ thông minh AI/ML trên Cloud như Amazon Textract để tự động trích xuất nội dung văn bản từ tệp scan phục vụ tìm kiếm thông minh.")

    doc.save(DOCX_OUT)
    print(f"-> Tạo thành công tài liệu Word hoàn chỉnh: {DOCX_OUT}")
    return DOCX_OUT

if __name__ == "__main__":
    build_cloud_doc()
