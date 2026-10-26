"""
Script tạo file tài liệu Word (.docx) đơn giản, tập trung và đúng trọng tâm:
BẢNG PHÂN CHIA CÔNG VIỆC BÀI TẬP LỚN
Lược bỏ các phần ngoài lề, chỉ tập trung vào: Tên thành viên, Vai trò, Nội dung chi tiết các phần việc làm và File sản phẩm phụ trách.
"""

import sys
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

PROJECT_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DOCX = PROJECT_ROOT / "BANG_PHAN_CONG_CONG_VIEC_INT1422.docx"

doc = Document()

# Thiết lập lề trang A4 chuẩn
for section in doc.sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)

# Màu sắc
COLOR_NAVY = RGBColor(15, 23, 42)     # #0f172a
COLOR_BLUE = RGBColor(37, 99, 235)     # #2563eb
COLOR_GRAY = RGBColor(100, 116, 139)   # #64748b

def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="CBD5E1", sz="6"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

# ─── TIÊU ĐỀ ĐƠN GIẢN, TRANG TRỌNG ──────────────────────────────────────────
p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sub.paragraph_format.space_after = Pt(2)
r = p_sub.add_run("HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG — KHOA CNTT 1\nBỘ MÔN HỆ THỐNG THÔNG TIN")
r.font.name = "Arial"
r.font.size = Pt(10)
r.font.bold = True
r.font.color.rgb = COLOR_GRAY

p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_title.paragraph_format.space_before = Pt(6)
p_title.paragraph_format.space_after = Pt(4)
r = p_title.add_run("BẢNG PHÂN CHIA CÔNG VIỆC THÀNH VIÊN")
r.font.name = "Arial"
r.font.size = Pt(16)
r.font.bold = True
r.font.color.rgb = COLOR_NAVY

p_info = doc.add_paragraph()
p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_info.paragraph_format.space_after = Pt(16)
r = p_info.add_run("Học phần: Kho Dữ Liệu & Khai Phá Dữ Liệu (INT1422)\n")
r.font.name = "Arial"
r.font.size = Pt(10)
r.font.italic = True
r = p_info.add_run("Đề tài: Xây Dựng Kho Dữ Liệu & Phân Tích Business Intelligence Thương Mại Điện Tử Đa Nguồn")
r.font.name = "Arial"
r.font.size = Pt(10.5)
r.font.bold = True

# ─── BẢNG PHÂN CHIA CÔNG VIỆC CHÍNH ───────────────────────────────────────────
tbl = doc.add_table(rows=4, cols=4)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl.autofit = False
set_table_borders(tbl, color="94A3B8", sz="6")

# Độ rộng cột (Tổng 6.8 inch)
col_widths = [Inches(0.5), Inches(1.8), Inches(3.2), Inches(1.3)]
for row in tbl.rows:
    for idx, w in enumerate(col_widths):
        row.cells[idx].width = w

# Header
headers = ["STT", "Thành viên", "Nội dung chi tiết các phần việc thực hiện", "File / Sản phẩm phụ trách"]
for idx, text in enumerate(headers):
    cell = tbl.cell(0, idx)
    set_cell_background(cell, "1E3A8A")  # Navy đậm
    set_cell_margins(cell, top=120, bottom=120, left=100, right=100)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor(255, 255, 255)

# Danh sách thành viên và công việc chi tiết
members_data = [
    (
        "1",
        "Trần Quang Minh\n(Trưởng nhóm)\nMSSV: [Điền MSSV]\n\nVai trò: Data Engineer &\nDWH Architect",
        "1. Thu thập dữ liệu đa nguồn (ETL):\n"
        "   • Trích xuất 10,000 dòng đánh giá từ Amazon Reviews 2023 qua HTTP Streaming JSONL.\n"
        "   • Gọi DummyJSON REST API lấy dữ liệu 100 người dùng và 50 giỏ hàng.\n"
        "   • Tích hợp Open Food Facts API làm giàu thông tin nguyên liệu sản phẩm.\n"
        "2. Tiền xử lý & Kiểm định toàn vẹn:\n"
        "   • Khử trùng lặp, xử lý giá trị thiếu, chuẩn hóa timestamp.\n"
        "   • Xây dựng bộ kiểm định tự động (src/etl/validate.py) đạt 100% Passed (0 bản ghi mồ côi).\n"
        "3. Thiết kế & Cài đặt Kho Dữ Liệu:\n"
        "   • Thiết kế Lược đồ hình sao (Star Schema Kimball) gồm 1 bảng Fact_Interactions và 4 bảng Dimension (Dim_User, Dim_Product, Dim_Category, Dim_Time).\n"
        "   • Cài đặt Docker PostgreSQL, viết DDL migrations và thiết lập kịch bản chạy tự động 1 lệnh run_all.py.",
        "• src/etl/extract/\n"
        "• src/etl/transform/\n"
        "• src/etl/validate.py\n"
        "• src/database.py\n"
        "• run_all.py"
    ),
    (
        "2",
        "Thành viên 2\n[Điền Họ và Tên]\nMSSV: [Điền MSSV]\n\nVai trò: BI & OLAP Engineer",
        "1. Tính toán trước Khối đa chiều (Precomputed Cubes):\n"
        "   • Xây dựng Khối 1 (vw_cube_sales_category_time) bằng cú pháp GROUP BY CUBE(category_name, year, quarter).\n"
        "   • Xây dựng Khối 2 (vw_cube_customer_segment_geo) bằng cú pháp GROUP BY ROLLUP(rfm_segment, country, city).\n"
        "   • Tối ưu hóa chỉ mục (Indexes) trên DWH giúp tăng tốc độ truy vấn từ O(N) -> O(1).\n"
        "2. Hiện thực hóa 5 thao tác OLAP:\n"
        "   • Cài đặt đủ 5 thao tác: Roll-up, Drill-down, Slice, Dice, Pivot.\n"
        "3. Xây dựng Backend REST API Microservice (FastAPI):\n"
        "   • Đóng gói các endpoints phục vụ tích hợp: /bi/segments, /bi/funnel, /recommend/user, /recommend/similar, /recommend/frequently-bought-together.\n"
        "4. Xây dựng Giao diện Web BI Dashboard:\n"
        "   • Tab 1: Executive KPI Dashboard (Doanh thu $199k, 9.9k đơn hàng, Phễu chuyển đổi tương tác).\n"
        "   • Tab 2: Bảng điều khiển trực tiếp 5 thao tác OLAP tương tác với người dùng.",
        "• sql/olap_cubes.sql\n"
        "• sql/bi_views.sql\n"
        "• src/api/\n"
        "  (main.py, templates/)\n"
        "• src/ui/dashboard.py\n"
        "  (Tab 1 & Tab 2)"
    ),
    (
        "3",
        "Thành viên 3\n[Điền Họ và Tên]\nMSSV: [Điền MSSV]\n\nVai trò: Data Scientist &\nML Engineer",
        "1. Phân cụm Khách hàng RFM (K-Means Clustering):\n"
        "   • Chuẩn hóa StandardScaler vector [Recency, Frequency, Monetary].\n"
        "   • Xác định K tối ưu bằng phương pháp Elbow và Silhouette Score (đạt đỉnh 0.6393 tại K=4).\n"
        "   • Phân tích 4 phân khúc (Champions, Potential Loyalists, Loyal, At Risk) và nạp nhãn ngược lại vào cột rfm_segment của bảng Dim_User.\n"
        "2. Khai phá Luật kết hợp Giỏ hàng (FP-Growth MBA):\n"
        "   • Khai phá giỏ hàng, trích xuất 7,674 luật kết hợp (độ nâng Lift tới 60.0x).\n"
        "   • Xây dựng giải thuật gợi ý Combo Bundle tự động chiết khấu 10%.\n"
        "3. Đánh giá Offline Benchmark đa mô hình:\n"
        "   • Chia tập dữ liệu theo thời gian (Temporal Split trên 874 test users), đo Precision@K, Recall@K, NDCG@K, HitRate, MAP.\n"
        "   • So sánh đối chiếu giữa Popularity Baseline, Content-Based và Collaborative Filtering (TruncatedSVD).\n"
        "4. Tích hợp giao diện Dashboard:\n"
        "   • Xây dựng Streamlit Tab 3 (RFM Segments) và Tab 4 (Cross-sell Combo).",
        "• src/mining/rfm_clustering.py\n"
        "• src/mining/association_rules.py\n"
        "• src/recsys/\n"
        "• src/ui/dashboard.py\n"
        "  (Tab 3 & Tab 4)"
    ),
]

for row_idx, data in enumerate(members_data, start=1):
    bg_color = "F8FAFC" if row_idx % 2 == 1 else "FFFFFF"
    for col_idx, text in enumerate(data):
        cell = tbl.cell(row_idx, col_idx)
        set_cell_background(cell, bg_color)
        set_cell_margins(cell, top=100, bottom=100, left=90, right=90)
        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(2)
        if col_idx == 0:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(9)
        if col_idx == 1:
            r.font.size = Pt(9.5)

# ─── PHẦN CHỮ KÝ XÁC NHẬN ĐƠN GIẢN ──────────────────────────────────────────
p_sp = doc.add_paragraph()
p_sp.paragraph_format.space_before = Pt(16)
p_sp.paragraph_format.space_after = Pt(4)

sign_tbl = doc.add_table(rows=1, cols=3)
sign_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
sign_tbl.autofit = False

for idx in range(3):
    sign_tbl.columns[idx].width = Inches(2.26)

signs = [
    "Thành viên 1 (Trưởng nhóm)\n(Ký và ghi rõ họ tên)\n\n\n\n\nTrần Quang Minh",
    "Thành viên 2\n(Ký và ghi rõ họ tên)\n\n\n\n\n[Họ và Tên]",
    "Thành viên 3\n(Ký và ghi rõ họ tên)\n\n\n\n\n[Họ và Tên]"
]

for idx, text in enumerate(signs):
    cell = sign_tbl.cell(0, idx)
    set_cell_margins(cell, top=60, bottom=60, left=40, right=40)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(9.5)
    r.font.bold = True

doc.save(str(OUTPUT_DOCX))
