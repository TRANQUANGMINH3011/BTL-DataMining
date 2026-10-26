"""
Script tạo PowerPoint 15 slide thuyết trình - Phiên bản nâng cấp hoàn toàn.
Thiết kế theo phong cách chuyên nghiệp: 2 layout chính
  - Slide tiêu đề (bìa): nền đậm gradient
  - Slide nội dung: Tiêu đề trên + chia đôi layout (Bullet trái | Hình ảnh phải)
"""

import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from lxml import etree

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ASSETS = PROJECT_ROOT / "assets"
OUTPUT = PROJECT_ROOT / "Bao_Cao_Thuyet_Trinh_INT1422.pptx"

# ─── MÀU SẮC ───────────────────────────────────────────────────────────────────
C_NAVY      = RGBColor(13, 27, 50)     # #0d1b32 - nền đậm chuyên nghiệp
C_DARK_BLUE = RGBColor(22, 55, 101)    # #163765 - tiêu đề nền
C_BLUE      = RGBColor(37, 99, 235)    # #2563eb - accent chính
C_SKY       = RGBColor(56, 145, 255)   # #3891ff - accent nhạt
C_EMERALD   = RGBColor(5, 150, 105)    # #059669 - green accent
C_AMBER     = RGBColor(217, 119, 6)    # #d97706 - amber accent
C_ROSE      = RGBColor(220, 38, 38)    # #dc2626 - red accent
C_WHITE     = RGBColor(255, 255, 255)
C_LIGHT     = RGBColor(241, 245, 249)  # #f1f5f9 - nền nhạt
C_MUTED     = RGBColor(148, 163, 184)  # #94a3b8 - text phụ
C_BODY      = RGBColor(30, 41, 59)     # #1e293b - body text
C_BORDER    = RGBColor(203, 213, 225)  # #cbd5e1
C_CARD      = RGBColor(248, 250, 252)  # #f8fafc

# Slide: 16:9 widescreen
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]   # blank layout


# ═══════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(int(hex_str[0:2],16), int(hex_str[2:4],16), int(hex_str[4:6],16))


def solid_rect(slide, x, y, w, h, fill_rgb, line_rgb=None, line_w=0):
    """Thêm hình chữ nhật đặc."""
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))  # MSO_SHAPE.RECTANGLE=1
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if line_rgb:
        shape.line.color.rgb = line_rgb
        shape.line.width = Pt(line_w)
    else:
        shape.line.fill.background()
    return shape


def rounded_rect(slide, x, y, w, h, fill_rgb, line_rgb, line_w=1.2, radius=0.12):
    """Hình chữ nhật bo góc bằng MSO_SHAPE."""
    from pptx.util import Inches
    shape = slide.shapes.add_shape(5, Inches(x), Inches(y), Inches(w), Inches(h))  # 5=ROUNDED_RECTANGLE
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    shape.line.color.rgb = line_rgb
    shape.line.width = Pt(line_w)
    # Điều chỉnh bán kính bo góc
    try:
        adj = shape.shape_element.find(qn("p:spPr")).find(qn("a:prstGeom")).find(qn("a:avLst")).find(qn("a:gd"))
        if adj is not None:
            pct = int(radius * 100000)
            adj.attrib["fmla"] = f"val {pct}"
    except Exception:
        pass
    return shape


def add_textbox(slide, x, y, w, h):
    return slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))


def para(tf, text, size, bold=False, color=C_BODY, space_before=0, align=PP_ALIGN.LEFT, italic=False):
    """Thêm paragraph vào text frame."""
    if tf.text == "":
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.italic = italic
    p.font.color.rgb = color
    p.space_before = Pt(space_before)
    p.alignment = align
    return p


def add_image(slide, img_path, x, y, w, h):
    """Nhúng ảnh với xử lý lỗi."""
    p = Path(img_path)
    if not p.exists():
        print(f"  [!] Không tìm thấy ảnh: {p}")
        return None
    pic = slide.shapes.add_picture(str(p), Inches(x), Inches(y), Inches(w), Inches(h))
    return pic


def slide_header_stripe(slide, title_line1, title_line2="", section_label=""):
    """
    Vẽ dải tiêu đề phía trên (dày ~1.5 in).
    """
    # Nền dải tiêu đề gradient-fake bằng 2 rect
    solid_rect(slide, 0, 0, 13.333, 1.55, C_DARK_BLUE)
    solid_rect(slide, 0, 1.45, 13.333, 0.1, C_BLUE)  # thin accent line

    # Label phần (ví dụ: "PHẦN 2: DỮ LIỆU")
    if section_label:
        lbox = add_textbox(slide, 0.5, 0.15, 12.333, 0.35)
        tf = lbox.text_frame
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = section_label.upper()
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = C_SKY
        p.font.letter_spacing = Pt(1)

    # Tiêu đề slide chính
    tbox = add_textbox(slide, 0.5, 0.38, 12.333, 1.0)
    tf = tbox.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    full_title = title_line1 + (" — " + title_line2 if title_line2 else "")
    p.text = full_title
    p.font.size = Pt(21) if len(full_title) < 80 else Pt(18)
    p.font.bold = True
    p.font.color.rgb = C_WHITE


def slide_footer(slide, slide_num, total=15):
    """Thanh chân trang."""
    solid_rect(slide, 0, 7.28, 13.333, 0.22, C_DARK_BLUE)
    fbox = add_textbox(slide, 0.5, 7.29, 12.333, 0.2)
    tf = fbox.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = f"INT1422 · Kho Dữ Liệu & Khai Phá Dữ Liệu  |  Nhóm đề tài Thương mại điện tử đa nguồn  |  Slide {slide_num}/{total}"
    p.font.size = Pt(9)
    p.font.color.rgb = C_MUTED
    p.alignment = PP_ALIGN.CENTER


def bullet_box(slide, x, y, w, h, items, header=None, header_color=C_BLUE, bg=C_CARD, border=C_BORDER):
    """
    Hộp nội dung: tùy chọn có tiêu đề header và danh sách bullets.
    items = list of (text, level, bold, color) 
    """
    shape = rounded_rect(slide, x, y, w, h, bg, border, line_w=1.0)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.22)
    tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.15)
    tf.margin_bottom = Inches(0.1)

    first = True
    if header:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.text = header
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = header_color
        p.space_before = Pt(0)

    for (text, size, bold, color, sp) in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.text = text
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = color
        p.space_before = Pt(sp)

    return shape


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — BÌA (Cover)
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)

# Background toàn trang
solid_rect(s, 0, 0, 13.333, 7.5, C_NAVY)

# Accent bar bên trái
solid_rect(s, 0, 0, 0.35, 7.5, C_BLUE)

# Dải màu ngang ở giữa
solid_rect(s, 0.35, 2.85, 12.983, 0.06, C_BLUE)

# Học viện
b0 = add_textbox(s, 1.0, 0.6, 11.333, 0.55)
tf = b0.text_frame
tf.word_wrap = True
para(tf, "HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG", 11.5, bold=True, color=C_MUTED)

b1 = add_textbox(s, 1.0, 1.05, 11.333, 0.45)
tf = b1.text_frame
para(tf, "KHOA CÔNG NGHỆ THÔNG TIN 1  ·  BỘ MÔN HỆ THỐNG THÔNG TIN", 10.5, color=C_MUTED)

# Tag môn học
tag = add_textbox(s, 1.0, 1.65, 4.5, 0.42)
tf = tag.text_frame
tf.margin_left = Inches(0.1)
tf.margin_top = Inches(0.05)
tf.margin_bottom = Inches(0.05)
solid_rect(s, 1.0, 1.65, 4.5, 0.42, C_BLUE)
para(tf, "  BÀI TẬP LỚN — INT1422  KHO DỮ LIỆU & KHAI PHÁ DỮ LIỆU", 10, bold=True, color=C_WHITE)

# Tiêu đề chính
b2 = add_textbox(s, 1.0, 2.15, 11.333, 0.45)
tf = b2.text_frame
para(tf, "XÂY DỰNG KHO DỮ LIỆU & PHÂN TÍCH BUSINESS INTELLIGENCE", 14, bold=True, color=C_SKY)

b3 = add_textbox(s, 1.0, 2.55, 11.333, 1.1)
tf = b3.text_frame
tf.word_wrap = True
para(tf, "THƯƠNG MẠI ĐIỆN TỬ ĐA NGUỒN", 34, bold=True, color=C_WHITE)

# Tag chuẩn đầu ra
b4 = add_textbox(s, 1.0, 3.65, 11.333, 0.45)
tf = b4.text_frame
para(tf, "Chuẩn đầu ra: CLO1 (45%) · CLO2 (45%) · CLO3 (10%)  |  Thời lượng thuyết trình: 15 phút  |  GitHub: TRANQUANGMINH3011/DataMining-Project", 10, color=C_MUTED)

# 3 Thẻ Thành viên Nhóm & Phân công công việc (Chuẩn Slide 24 & CLO3)
members_data = [
    ("Trần Quang Minh (Trưởng nhóm)", "Data Engineer & DWH Architect", [
        ("• Trích xuất đa nguồn ETL (Amazon, DummyJSON, OFF)", 9, False, C_BODY, 2),
        ("• Thiết kế Star Schema Kimball (1 Fact + 4 Dim)", 9, False, C_BODY, 2),
        ("• Hạ tầng PostgreSQL Docker & Pipeline run_all.py", 9, False, C_BODY, 2),
        ("Đóng góp: 34%  ·  Thuyết trình: 5 phút đầu (Phần 1 & 2)", 9, True, C_BLUE, 5),
    ], C_SKY),
    ("Thành viên 2", "BI & OLAP Engineer", [
        ("• Khối đa chiều CUBE & ROLLUP tăng tốc O(1)", 9, False, C_BODY, 2),
        ("• Hiện thực 5 thao tác OLAP (Roll-up, Drill, Slice...)", 9, False, C_BODY, 2),
        ("• Xây dựng Streamlit BI Dashboard (Tab 1 & Tab 2)", 9, False, C_BODY, 2),
        ("Đóng góp: 33%  ·  Thuyết trình: 5 phút giữa (Phần 3 & 4)", 9, True, C_EMERALD, 5),
    ], C_EMERALD),
    ("Thành viên 3", "Data Scientist & ML Engineer", [
        ("• Phân cụm RFM K-Means (Silhouette 0.6393, K=4)", 9, False, C_BODY, 2),
        ("• FP-Growth Market Basket (7,674 luật, Lift 60.0x)", 9, False, C_BODY, 2),
        ("• Đánh giá Offline đa mô hình & Combo Cross-sell", 9, False, C_BODY, 2),
        ("Đóng góp: 33%  ·  Thuyết trình: 5 phút cuối (Phần 5 & 6)", 9, True, C_AMBER, 5),
    ], C_AMBER),
]

for idx, (m_name, m_role, m_bullets, m_col) in enumerate(members_data):
    mx = 1.0 + idx * 3.85
    my = 4.25
    mw = 3.65
    mh = 2.85
    rounded_rect(s, mx, my, mw, mh, C_WHITE, m_col, line_w=1.2)
    
    # Title & Role box
    tb = add_textbox(s, mx + 0.15, my + 0.12, mw - 0.3, 0.65)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    para(tf, m_name, 11, bold=True, color=C_NAVY)
    para(tf, m_role, 9.5, italic=True, color=m_col, space_before=2)
    
    # Bullets box
    bb = add_textbox(s, mx + 0.15, my + 0.8, mw - 0.3, 1.9)
    tf = bb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    for b_text, b_size, b_bold, b_color, b_sp in m_bullets:
        para(tf, b_text, b_size, bold=b_bold, color=b_color, space_before=b_sp)

slide_footer(s, 1)

print("Slide 1 (Bìa & Phân công) xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — CÂU HỎI NGHIỆP VỤ & MỤC TIÊU
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "1. Câu Hỏi Nghiệp Vụ & Bối Cảnh Thực Tế Sàn TMĐT", section_label="Phần 1: Bối cảnh & Mục tiêu (2 phút)")
slide_footer(s, 2)

# --- Cột trái: OLTP vs DWH
bullet_box(s, 0.35, 1.65, 5.9, 5.5,
    items=[
        ("Tại sao CSDL Giao dịch (OLTP) không đủ?", 12.5, True, C_BLUE, 0),
        ("• OLTP thiết kế chuẩn hóa 3NF tối ưu GHI dữ liệu giao dịch theo thời gian thực; không phù hợp với phân tích tổng hợp lớn vì phải scan toàn bảng (O(N)), gây nghẽn cổ chai hệ thống bán hàng.", 10.5, False, C_BODY, 5),
        ("Kho Dữ Liệu (Data Warehouse) giải quyết:", 12.5, True, C_EMERALD, 10),
        ("• Mô hình Kimball Star Schema tách biệt phân tích khỏi vận hành.", 10.5, False, C_BODY, 5),
        ("• Tổng hợp đa chiều (OLAP): Roll-up / Drill-down / Slice / Dice / Pivot.", 10.5, False, C_BODY, 3),
        ("• Tăng tốc truy vấn báo cáo từ O(N) → O(1) nhờ Khối tính trước (CUBE & ROLLUP).", 10.5, False, C_BODY, 3),
    ],
    header=None, bg=C_WHITE, border=C_BLUE
)

# --- Cột phải: 3 câu hỏi nghiệp vụ
bullet_box(s, 6.55, 1.65, 6.45, 5.5,
    items=[
        ("3 Câu hỏi nghiệp vụ cần giải quyết bằng DWH", 12.5, True, C_NAVY, 0),
        ("📊 Q1 — Kinh doanh & Ngành hàng:", 11.5, True, C_BLUE, 10),
        ("Doanh thu biến động thế nào theo Năm/Quý/Tháng? Nhóm sản phẩm nào đóng góp chủ đạo?", 10.5, False, C_BODY, 3),
        ("👥 Q2 — Quản trị Khách hàng (CRM):", 11.5, True, C_EMERALD, 10),
        ("Nhóm VIP (Champions) nào đóng góp trên 50% doanh thu? Nhóm nào có nguy cơ rời bỏ cần tái kích hoạt?", 10.5, False, C_BODY, 3),
        ("🛒 Q3 — Tối ưu Giá trị Đơn hàng (AOV):", 11.5, True, C_AMBER, 10),
        ("Khi khách chọn SP A, họ thường mua kèm SP nào? Thiết kế Combo Bundle thế nào để tăng doanh số bán chéo?", 10.5, False, C_BODY, 3),
    ],
    header=None, bg=C_WHITE, border=C_EMERALD
)

print("Slide 2 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — KIẾN TRÚC ĐA NGUỒN
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "2. Kiến Trúc Thu Thập Dữ Liệu Đa Nguồn & Luồng ETL", section_label="Phần 2: Dữ liệu & Tiền xử lý (3 phút)")
slide_footer(s, 3)

# Trái: 3 nguồn
bullet_box(s, 0.35, 1.65, 5.9, 5.5,
    items=[
        ("Nguồn 1 — Amazon Reviews 2023", 12, True, C_BLUE, 0),
        ("• Giao thức: HTTP Streaming JSONL từ HuggingFace McAuley-Lab.\n• Quy mô: 10,000 đánh giá & 7,067 sản phẩm thực tế.\n• Cung cấp: user_id, rating, timestamp, review_text, helpful_vote.", 10, False, C_BODY, 3),
        ("Nguồn 2 — DummyJSON REST API", 12, True, C_EMERALD, 10),
        ("• Giao thức: RESTful HTTP GET phân trang + Exponential Backoff.\n• Quy mô: 100 hồ sơ người dùng & 50 giỏ hàng (194 mặt hàng).\n• Cung cấp: Tuổi, giới tính, thành phố, quốc gia, giao dịch.", 10, False, C_BODY, 3),
        ("Nguồn 3 — Open Food Facts REST API", 12, True, C_AMBER, 10),
        ("• Giao thức: REST API v2 kèm định danh User-Agent chuẩn.\n• Quy mô: Làm giàu top 50–100 sản phẩm phổ biến.\n• Cung cấp: ingredients_text, nhãn dinh dưỡng cho gợi ý nội dung.", 10, False, C_BODY, 3),
    ],
    header=None, bg=C_WHITE, border=C_BLUE
)

# Phải: Sơ đồ kiến trúc thực tế
solid_rect(s, 6.55, 1.65, 6.45, 5.5, C_WHITE, C_BORDER, 1.0)
add_image(s, ASSETS / "architecture_diagram.png", 6.6, 1.7, 6.35, 5.35)
cap = add_textbox(s, 6.6, 7.05, 6.35, 0.22)
tf = cap.text_frame
para(tf, "Hình 1: Sơ đồ kiến trúc tổng thể 4 tầng (Multi-source → ETL → DWH → BI & ML)", 8.5, italic=True, color=C_MUTED, align=PP_ALIGN.CENTER)

print("Slide 3 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — TIỀN XỬ LÝ & CHẤT LƯỢNG DỮ LIỆU
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "3. Quy Trình Tiền Xử Lý & Kiểm Định Toàn Vẹn Dữ Liệu (100% Passed)", section_label="Phần 2: Dữ liệu & Tiền xử lý (3 phút)")
slide_footer(s, 4)

# 4 thẻ bước ETL xếp dọc
etl_steps = [
    ("✓  1. Làm sạch & Chuẩn hóa Amazon Reviews", "Loại bỏ bản ghi rỗng user_id/product_id; chuẩn hóa rating ∈ [1.0–5.0]; chuyển đổi millisecond Unix timestamp → full_timestamp; khử trùng lặp theo cặp (user_id, product_id).", C_BLUE),
    ("✓  2. Entity Resolution — Ghép nối nhân khẩu học DummyJSON", "Tập hợp user_id duy nhất từ Amazon Reviews; ánh xạ xoay vòng (round-robin) hồ sơ nhân khẩu học (tuổi, giới tính, thành phố, quốc gia) từ DummyJSON API để tạo hồ sơ KH đầy đủ.", C_BLUE),
    ("✓  3. Tích hợp đặc trưng SP & Bóc tách Dim_Time", "Hợp nhất Metadata Amazon với ingredients_text từ Open Food Facts; bóc tách full_timestamp → time_id (YYYYMMDDHH), day, month, quarter, year, day_of_week để điền Dim_Time.", C_EMERALD),
    ("✓  4. Kiểm định tự động — Data Quality Gate (src/etl/validate.py)", "100% Khóa chính duy nhất, không NULL. 0 bản ghi mồ côi (Zero Orphan Records) giữa Fact_Interactions và cả 4 bảng Dim. → GATE PASSED: Hệ thống tự động nạp vào PostgreSQL DWH.", C_EMERALD),
]

for i, (title, desc, col) in enumerate(etl_steps):
    top = 1.65 + i * 1.35
    bullet_box(s, 0.35, top, 12.633, 1.22,
        items=[
            (title, 11.5, True, col, 0),
            (desc, 10, False, C_BODY, 3),
        ],
        header=None, bg=C_WHITE, border=col
    )

print("Slide 4 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — LƯỢC ĐỒ HÌNH SAO (STAR SCHEMA ERD)
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "4. Thiết Kế Lược Đồ Hình Sao (Ralph Kimball Star Schema)", section_label="Phần 3: Thiết kế & Cài đặt Kho Dữ Liệu (4 phút)")
slide_footer(s, 5)

# Trái: Mô tả Fact & Dim
bullet_box(s, 0.35, 1.65, 5.9, 5.5,
    items=[
        ("Bảng Sự kiện — Fact_Interactions", 12.5, True, C_NAVY, 0),
        ("• Quy mô: 9,971 bản ghi tương tác thực tế.", 10.5, True, C_BLUE, 5),
        ("• Granularity: Mỗi dòng = 1 tương tác (review/cart/purchase) của 1 user với 1 product tại 1 mốc giờ.", 10, False, C_BODY, 2),
        ("• FK: user_id → Dim_User | product_id → Dim_Product | time_id → Dim_Time.", 10, False, C_BODY, 2),
        ("• Measures: rating_score, helpful_vote, event_type.", 10, False, C_BODY, 2),
        ("4 Bảng Chiều (Dimensions)", 12.5, True, C_NAVY, 12),
        ("• Dim_Time (8,955): Phân cấp Day → Month → Quarter → Year.", 10.5, False, C_BLUE, 5),
        ("• Dim_Category (3): All Beauty / Appliances / Electronics.", 10.5, False, C_EMERALD, 3),
        ("• Dim_Product (7,067): product_name, price, brand, ingredients_text (OFF).", 10.5, False, C_AMBER, 3),
        ("• Dim_User (6,585): age, gender, city, country, rfm_segment.", 10.5, False, C_ROSE, 3),
    ],
    header=None, bg=C_WHITE, border=C_BORDER
)

# Phải: ERD ảnh thực tế
solid_rect(s, 6.55, 1.65, 6.45, 5.5, C_WHITE, C_BORDER, 1.0)
add_image(s, ASSETS / "star_schema_erd.png", 6.6, 1.7, 6.35, 5.35)
cap = add_textbox(s, 6.6, 7.05, 6.35, 0.22)
tf = cap.text_frame
para(tf, "Hình 2: Sơ đồ ERD Lược đồ hình sao — Fact_Interactions kết nối 4 Dimensions với PK/FK chuẩn Kimball", 8.5, italic=True, color=C_MUTED, align=PP_ALIGN.CENTER)

print("Slide 5 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — OLAP CUBE (TÍNH TRƯỚC ĐA CHIỀU)
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "5. Tính Toán Trước Khối Đa Chiều — GROUP BY CUBE & ROLLUP", section_label="Phần 3: Thiết kế & Cài đặt Kho Dữ Liệu (4 phút)")
slide_footer(s, 6)

# Cột trái: Mô tả 2 cube
bullet_box(s, 0.35, 1.65, 5.9, 5.5,
    items=[
        ("Khối 1 — vw_cube_sales_category_time", 12.5, True, C_BLUE, 0),
        ("Cú pháp SQL: GROUP BY CUBE(category_name, year, quarter)", 10, True, C_MUTED, 4),
        ("• Tính toán trước doanh số ở mọi tổ hợp cấp độ: (Danh mục, Năm, Quý), (Danh mục, Năm), (Danh mục), (Năm) và Grand Total (Tổng toàn sàn).", 10, False, C_BODY, 4),
        ("• Tăng tốc: O(N) quét Fact → O(1) đọc trực tiếp từ Khối tính trước.", 10, False, C_BODY, 3),
        ("Khối 2 — vw_cube_customer_segment_geo", 12.5, True, C_EMERALD, 12),
        ("Cú pháp SQL: GROUP BY ROLLUP(rfm_segment, country, city)", 10, True, C_MUTED, 4),
        ("• Phân cấp địa lý: Phân khúc RFM → Quốc gia → Thành phố.", 10, False, C_BODY, 4),
        ("• Cung cấp chỉ số ARPU (Chi tiêu TB/KH) và số KH VIP theo từng cấp địa lý.", 10, False, C_BODY, 3),
        ("• Phục vụ Marketing phân bổ ngân sách khuyến mại chính xác.", 10, False, C_BODY, 3),
    ],
    header=None, bg=C_WHITE, border=C_BORDER
)

# Phải: Hình biểu đồ OLAP thực tế
solid_rect(s, 6.55, 1.65, 6.45, 5.5, C_WHITE, C_BORDER, 1.0)
add_image(s, ASSETS / "olap_pivot_heatmap.png", 6.6, 1.7, 6.35, 5.35)
cap = add_textbox(s, 6.6, 7.05, 6.35, 0.22)
tf = cap.text_frame
para(tf, "Hình 3: Kết quả từ Khối đa chiều — Roll-up Doanh thu theo Quý & Pivot Heatmap (Danh mục × Năm)", 8.5, italic=True, color=C_MUTED, align=PP_ALIGN.CENTER)

print("Slide 6 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — BI DASHBOARD (EXECUTIVE KPI)
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "6. Báo Cáo Hiệu Năng Doanh Nghiệp (Executive BI Dashboard)", section_label="Phần 4: Dịch vụ OLAP & Trực quan hóa (4 phút)")
slide_footer(s, 7)

# 4 KPI cards ở trên
kpis = [
    ("$199,320.29", "Tổng Doanh Thu DWH", C_EMERALD),
    ("9,971", "Tổng Tương Tác / Đơn Hàng", C_BLUE),
    ("6,585", "Tổng Khách Hàng", C_AMBER),
    ("4.25 / 5.0 ★", "Đánh Giá Trung Bình", C_ROSE),
]
kpi_w = 2.9
for i, (val, lbl, col) in enumerate(kpis):
    kx = 0.35 + i * (kpi_w + 0.23)
    # Card
    solid_rect(s, kx, 1.65, kpi_w, 1.05, C_WHITE, col, 1.5)
    vb = add_textbox(s, kx + 0.15, 1.7, kpi_w - 0.3, 0.55)
    tf = vb.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    para(tf, val, 20, bold=True, color=col, align=PP_ALIGN.CENTER)
    lb = add_textbox(s, kx + 0.1, 2.2, kpi_w - 0.2, 0.4)
    tf = lb.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    para(tf, lbl, 9.5, color=C_MUTED, align=PP_ALIGN.CENTER)

# Phần dưới: Trái bullet, Phải ảnh giao diện thực
bullet_box(s, 0.35, 2.85, 5.9, 4.2,
    items=[
        ("Phân tích Xu hướng Doanh thu theo Tháng", 12.5, True, C_NAVY, 0),
        ("• View vw_sales_monthly_summary cho góc nhìn MoM Growth (%) và giá trị đơn hàng TB (AOV).", 10.5, False, C_BODY, 5),
        ("• Giúp ban điều hành nắm bắt mùa vụ mua sắm cao điểm trong năm.", 10.5, False, C_BODY, 3),
        ("Phễu Chuyển Đổi Tương Tác Khách Hàng", 12.5, True, C_NAVY, 10),
        ("• 4 tầng: Tiếp cận (6,585 Users) → Xem SP (7,067 Products) → Thêm giỏ (1,394 Carts) → Đánh giá & Chốt đơn (9,971 Reviews).", 10.5, False, C_BODY, 5),
        ("• Tỷ lệ đánh giá và chốt đơn tích cực cao so với mặt bằng TMĐT.", 10.5, False, C_BODY, 3),
    ],
    header=None, bg=C_WHITE, border=C_BORDER
)
solid_rect(s, 6.55, 2.85, 6.45, 4.2, C_WHITE, C_BORDER, 1.0)
add_image(s, ASSETS / "ui_tab1_bi_overview.png", 6.6, 2.9, 6.35, 4.0)
cap = add_textbox(s, 6.6, 6.9, 6.35, 0.22)
tf = cap.text_frame
para(tf, "Hình 4: Ảnh chụp thực tế Streamlit Tab 1 — KPI Dashboard, Doanh thu theo tháng & Phễu tương tác", 8.5, italic=True, color=C_MUTED, align=PP_ALIGN.CENTER)

print("Slide 7 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — 5 THAO TÁC OLAP
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "7. Trình Diễn Đủ 5 Thao Tác OLAP Trực Tiếp Trên Dashboard", section_label="Phần 4: Dịch vụ OLAP & Trực quan hóa (4 phút)")
slide_footer(s, 8)

olap_ops = [
    ("1. ROLL-UP", "Thu gọn theo trục thời gian: Tháng → Quý → Năm hoặc từ Sản phẩm → Danh mục để xem tổng doanh số toàn sàn.", C_BLUE),
    ("2. DRILL-DOWN", "Đi sâu từ Danh mục (All Beauty) xuống 15 sản phẩm bán chạy nhất kèm giá bán, rating trung bình và thương hiệu.", C_EMERALD),
    ("3. SLICE", "Cố định 1 chiều: Cắt lát theo Mức Đánh Giá = 5.0 Sao hoặc cố định Quý 3 để phân tích đơn hàng chất lượng hoàn hảo.", C_AMBER),
    ("4. DICE", "Lọc đồng thời 2 chiều: Chiều Thời gian (Năm 2022–2023) VÀ Chiều Sản phẩm (Khoảng giá $10 đến $60).", C_ROSE),
    ("5. PIVOT", "Xoay trục: Danh mục thành Hàng, Năm thành Cột → Ma trận chéo Cross-table kèm Heatmap doanh thu trực quan.", hex_to_rgb("7c3aed")),
]

for i, (op_name, op_desc, col) in enumerate(olap_ops):
    top = 1.65 + i * 1.05
    bullet_box(s, 0.35, top, 5.9, 0.94,
        items=[
            (op_name, 12, True, col, 0),
            (op_desc, 9.5, False, C_BODY, 3),
        ],
        header=None, bg=C_WHITE, border=col
    )

solid_rect(s, 6.55, 1.65, 6.45, 5.5, C_WHITE, C_BORDER, 1.0)
add_image(s, ASSETS / "ui_tab2_olap_ops.png", 6.6, 1.7, 6.35, 5.35)
cap = add_textbox(s, 6.6, 7.05, 6.35, 0.22)
tf = cap.text_frame
para(tf, "Hình 5: Ảnh chụp thực tế Streamlit Tab 2 — Radio chọn thao tác OLAP & bảng biểu đồ tương tác trực tiếp", 8.5, italic=True, color=C_MUTED, align=PP_ALIGN.CENTER)

print("Slide 8 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — RFM K-MEANS CLUSTERING
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "8. Khai Phá Dữ Liệu: Phân Cụm Khách Hàng RFM (K-Means)", section_label="Phần 5: Chức năng Khai phá Dữ liệu (4 phút)")
slide_footer(s, 9)

bullet_box(s, 0.35, 1.65, 5.9, 5.5,
    items=[
        ("Vector đặc trưng RFM", 12.5, True, C_BLUE, 0),
        ("• R (Recency): Số ngày kể từ lần tương tác gần nhất.", 10.5, False, C_BODY, 5),
        ("• F (Frequency): Tổng số lần mua / đánh giá sản phẩm.", 10.5, False, C_BODY, 3),
        ("• M (Monetary): Tổng chi tiêu tích lũy (tính từ giá SP).", 10.5, False, C_BODY, 3),
        ("Chuẩn hóa & Lựa chọn K tối ưu", 12.5, True, C_EMERALD, 12),
        ("• StandardScaler đưa R, F, M về phân phối chuẩn N(0, 1).", 10.5, False, C_BODY, 5),
        ("• Đo Elbow (WCSS) + Silhouette Score cho K = 2 → 7.", 10.5, False, C_BODY, 3),
        ("• Đạt Silhouette = 0.6393 tại K = 4 ← Tối ưu!", 11, True, C_ROSE, 4),
        ("Đóng gói kết quả vào DWH (Closed-loop)", 12.5, True, C_NAVY, 12),
        ("• Nhãn phân cụm tự động cập nhật vào cột rfm_segment của Dim_User trong PostgreSQL để phục vụ phân tích OLAP tức thì.", 10.5, False, C_BODY, 5),
    ],
    header=None, bg=C_WHITE, border=C_BORDER
)

solid_rect(s, 6.55, 1.65, 6.45, 5.5, C_WHITE, C_BORDER, 1.0)
add_image(s, ASSETS / "rfm_clusters_vis.png", 6.6, 1.7, 6.35, 5.35)
cap = add_textbox(s, 6.6, 7.05, 6.35, 0.22)
tf = cap.text_frame
para(tf, "Hình 6: Tỷ trọng Doanh thu 4 cụm RFM (Donut) & Đường cong Silhouette Score — Đỉnh K=4 = 0.6393", 8.5, italic=True, color=C_MUTED, align=PP_ALIGN.CENTER)

print("Slide 9 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — 4 PHÂN KHÚC KHÁCH HÀNG
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "9. Chân Dung 4 Phân Khúc Khách Hàng & Chiến Lược Hành Động", section_label="Phần 5: Chức năng Khai phá Dữ liệu (4 phút)")
slide_footer(s, 10)

bullet_box(s, 0.35, 1.65, 5.9, 5.5,
    items=[
        ("Champions (VIP) — ARPU $286.65", 12, True, RGBColor(217, 119, 6), 0),
        ("• Đóng góp 30.2% tổng doanh thu. Tần suất mua cao nhất.", 10, False, C_BODY, 4),
        ("→ Chăm sóc đặc quyền VIP, quyền mua sớm, quà tri ân.", 10, True, C_AMBER, 2),
        ("Potential Loyalists — Tiềm năng mới", 12, True, RGBColor(37, 99, 235), 12),
        ("• Mới tương tác gần đây, giá trị đơn triển vọng.", 10, False, C_BODY, 4),
        ("→ Gửi voucher đơn thứ 2, email gợi ý SP cá nhân hóa.", 10, True, C_BLUE, 2),
        ("Loyal Customers — Trung thành đều đặn", 12, True, RGBColor(5, 150, 105), 12),
        ("• Mua sắm định kỳ, gắn kết cao với thương hiệu.", 10, False, C_BODY, 4),
        ("→ Tích điểm loyalty, gợi ý ngành hàng mới qua Cross-sell.", 10, True, C_EMERALD, 2),
        ("At Risk / Inactive — Nguy cơ rời bỏ", 12, True, RGBColor(220, 38, 38), 12),
        ("• Đã lâu không quay lại (Recency lớn), 47% tổng số KH.", 10, False, C_BODY, 4),
        ("→ Chiến dịch Re-engagement, giảm giá sâu tái kích hoạt.", 10, True, C_ROSE, 2),
    ],
    header=None, bg=C_WHITE, border=C_BORDER
)

solid_rect(s, 6.55, 1.65, 6.45, 5.5, C_WHITE, C_BORDER, 1.0)
add_image(s, ASSETS / "ui_tab3_rfm_segments.png", 6.6, 1.7, 6.35, 5.35)
cap = add_textbox(s, 6.6, 7.05, 6.35, 0.22)
tf = cap.text_frame
para(tf, "Hình 7: Ảnh chụp thực tế Tab 3 Streamlit — Donut Doanh thu, Bar ARPU và bảng số liệu 4 phân khúc", 8.5, italic=True, color=C_MUTED, align=PP_ALIGN.CENTER)

print("Slide 10 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 — FP-GROWTH ASSOCIATION RULES
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "10. Khai Phá Luật Kết Hợp Giỏ Hàng — FP-Growth (7,674 Luật)", section_label="Phần 5: Chức năng Khai phá Dữ liệu (4 phút)")
slide_footer(s, 11)

bullet_box(s, 0.35, 1.65, 5.9, 5.5,
    items=[
        ("Dữ liệu giỏ hàng (Market Basket Data)", 12.5, True, C_NAVY, 0),
        ("• DummyJSON Carts: 194 mặt hàng trong 50 giỏ hàng thực tế + nhóm SP được cùng user đánh giá trên sàn Amazon.", 10.5, False, C_BODY, 5),
        ("Thuật toán FP-Growth (Frequent Pattern Growth)", 12.5, True, C_EMERALD, 12),
        ("• Nén dữ liệu thành cây FP-Tree; khai phá nhanh vượt trội so với Apriori truyền thống (không phải quét lại CSDL nhiều lần).", 10.5, False, C_BODY, 5),
        ("Kết quả khai phá — Kiểm tra bắt buộc Lift (Slide 25)", 12.5, True, C_AMBER, 12),
        ("• Ngưỡng lọc: min_support = 0.005, min_confidence = 0.1, min_lift ≥ 1.1.", 10.5, False, C_BODY, 5),
        ("• Trích xuất thành công: 7,674 Luật kết hợp!", 11, True, C_EMERALD, 3),
        ("• Lift tối đa đạt 60.0x (Xác suất mua kèm cao gấp 60 lần so với ngẫu nhiên).", 10.5, False, C_BODY, 3),
        ("Ứng dụng tạo Combo Bundle (Cross-Sell -10%)", 12.5, True, C_BLUE, 12),
        ("• Luật [Cat Food, Cricket Ball] ⟹ [Tropical Earring]: Lift = 60.0x → Chiết khấu Combo -10% tự động.", 10.5, False, C_BODY, 5),
    ],
    header=None, bg=C_WHITE, border=C_BORDER
)

solid_rect(s, 6.55, 1.65, 6.45, 5.5, C_WHITE, C_BORDER, 1.0)
add_image(s, ASSETS / "ui_tab4_cross_sell.png", 6.6, 1.7, 6.35, 5.35)
cap = add_textbox(s, 6.6, 7.05, 6.35, 0.22)
tf = cap.text_frame
para(tf, "Hình 8: Ảnh chụp thực tế Tab 4 Streamlit — 7,674 luật kết hợp, lọc theo Min Lift, Lift tới 60.0x", 8.5, italic=True, color=C_MUTED, align=PP_ALIGN.CENTER)

print("Slide 11 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 12 — ĐÁNH GIÁ OFFLINE & BENCHMARK
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "11. Đánh Giá Offline Hệ Thống Gợi Ý & So Sánh Mô Hình Cơ Sở", section_label="Phần 5: Chức năng Khai phá Dữ liệu (4 phút)")
slide_footer(s, 12)

bullet_box(s, 0.35, 1.65, 5.9, 5.5,
    items=[
        ("Temporal Split — Tránh rò rỉ dữ liệu", 12.5, True, C_NAVY, 0),
        ("• Dữ liệu quá khứ dùng huấn luyện; 20% giao dịch gần nhất kiểm thử trên 874 người dùng.", 10.5, False, C_BODY, 5),
        ("So sánh 3 mô hình tại K = 10 (Slide 25)", 12.5, True, C_BLUE, 12),
        ("• Popularity Baseline (best overall):", 10.5, True, C_BLUE, 5),
        ("  Recall@10 = 0.0160 | NDCG@10 = 0.0074 | HitRate = 1.60% | MAP = 0.0049", 10, False, C_BODY, 2),
        ("• Content-Based (TF-IDF + Open Food Facts):", 10.5, True, C_EMERALD, 5),
        ("  Recall@10 = 0.0023 | NDCG@10 = 0.0015 | HitRate = 0.34% | MAP = 0.0013", 10, False, C_BODY, 2),
        ("• CF — TruncatedSVD (Collaborative Filtering):", 10.5, True, C_AMBER, 5),
        ("  Recall@10 = 0.0034 | NDCG@10 = 0.0015 | HitRate = 0.34% | MAP = 0.0009", 10, False, C_BODY, 2),
        ("Nhận định về Sparsity 99.98%", 12.5, True, C_NAVY, 10),
        ("• Dữ liệu TMĐT thực tế cực thưa (TB 1.30 tương tác/user). SVD gặp Cold-Start nặng; Content-Based có lợi thế khi tích hợp đặc trưng từ Open Food Facts.", 10, False, C_BODY, 5),
    ],
    header=None, bg=C_WHITE, border=C_BORDER
)

solid_rect(s, 6.55, 1.65, 6.45, 5.5, C_WHITE, C_BORDER, 1.0)
add_image(s, ASSETS / "recsys_benchmark_vis.png", 6.6, 1.85, 6.35, 4.65)
cap = add_textbox(s, 6.6, 7.05, 6.35, 0.22)
tf = cap.text_frame
para(tf, "Hình 9: So sánh Offline — Popularity Baseline vs Content-Based vs CF-SVD tại K=10 (Recall, NDCG, HitRate, MAP)", 8.5, italic=True, color=C_MUTED, align=PP_ALIGN.CENTER)

print("Slide 12 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 13 — ỨNG DỤNG COMBO BUNDLE & CROSS-SELL
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "12. Ứng Dụng Thực Tế: Gói Ưu Đãi Mua Kèm & Gợi Ý Cá Nhân Hóa", section_label="Phần 6: Khuyến nghị & Hành động (2 phút)")
slide_footer(s, 13)

bullet_box(s, 0.35, 1.65, 5.9, 5.5,
    items=[
        ("Tích hợp Luật Kết Hợp → Gợi ý Mua kèm", 12.5, True, C_BLUE, 0),
        ("• Khi KH thêm SP vào giỏ, hệ thống tra cứu tức thì các luật Lift cao nhất để gợi ý SP mua kèm phù hợp nhất.", 10.5, False, C_BODY, 5),
        ("Gói Combo Bundle — Chiết khấu 10%", 12.5, True, C_EMERALD, 12),
        ("• Tự động gom Combo các SP liên quan và hiển thị mức giá chiết khấu -10% để kích thích hoàn tất đơn hàng giá trị cao (Tăng AOV).", 10.5, False, C_BODY, 5),
        ("Gợi ý cá nhân hóa theo Phân khúc RFM", 12.5, True, C_AMBER, 12),
        ("• Champions (VIP): Ưu tiên gợi ý SP cao cấp, giá cao.", 10.5, False, C_BODY, 5),
        ("• Potential Loyalists: Gợi ý các mặt hàng phổ biến đang trending.", 10.5, False, C_BODY, 3),
        ("Trải nghiệm trực quan trên Web Dashboard", 12.5, True, C_NAVY, 12),
        ("• Giao diện Streamlit cho phép chọn thử bất kỳ SP nào → xem gợi ý Combo tính toán thời gian thực.", 10.5, False, C_BODY, 5),
    ],
    header=None, bg=C_WHITE, border=C_BORDER
)

solid_rect(s, 6.55, 1.65, 6.45, 5.5, C_WHITE, C_BORDER, 1.0)
add_image(s, ASSETS / "ui_tab4_bundle_cards.png", 6.6, 1.7, 6.35, 5.35)
cap = add_textbox(s, 6.6, 7.05, 6.35, 0.22)
tf = cap.text_frame
para(tf, "Hình 10: Ảnh chụp thực tế — Giao diện thẻ Gói Sản phẩm Mua kèm Combo (Bundle Recommendation) trên Streamlit", 8.5, italic=True, color=C_MUTED, align=PP_ALIGN.CENTER)

print("Slide 13 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 14 — TỰ ĐỘNG HÓA 1 LỆNH & HẠN CHẾ
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "13. Tự Động Hóa Vận Hành Bằng Một Lệnh & Hạn Chế Của Đề Tài", section_label="Phần 6: Khuyến nghị & Hành động (2 phút)")
slide_footer(s, 14)

# Cột trái: 1 lệnh duy nhất
bullet_box(s, 0.35, 1.65, 5.9, 5.5,
    items=[
        ("✓  Tái lập 100% bằng 1 lệnh duy nhất (Slide 12)", 13, True, C_EMERALD, 0),
        ("    python run_all.py", 14, True, C_NAVY, 6),
        ("Toàn bộ 6 giai đoạn chạy tự động khép kín:", 11, True, C_BLUE, 10),
        ("  1.  Khởi tạo cấu trúc Star Schema trong PostgreSQL", 10, False, C_BODY, 4),
        ("  2.  Trích xuất đa nguồn (Amazon + DummyJSON + Open Food Facts)", 10, False, C_BODY, 3),
        ("  3.  Tiền xử lý, khử trùng lặp & Kiểm định Quality Gate", 10, False, C_BODY, 3),
        ("  4.  Tính toán trước 2 Khối đa chiều (CUBE & ROLLUP)", 10, False, C_BODY, 3),
        ("  5.  Huấn luyện K-Means RFM & Khai phá FP-Growth MBA", 10, False, C_BODY, 3),
        ("  6.  Tự động khởi chạy Web BI Dashboard trên cổng 8501", 10, False, C_BODY, 3),
        ("→ Giúp Hội đồng kiểm chứng kết quả minh bạch tuyệt đối!", 11, True, C_EMERALD, 8),
    ],
    header=None, bg=C_WHITE, border=C_EMERALD
)

# Cột phải: Hạn chế
bullet_box(s, 6.55, 1.65, 6.45, 5.5,
    items=[
        ("⚠️  Các Hạn Chế Trung Thực Của Đề Tài (Slide 24)", 13, True, C_NAVY, 0),
        ("Hạn chế 1 — Giới hạn Rate Limiting REST API:", 12, True, C_AMBER, 10),
        ("• Gọi qua API công cộng miễn phí (DummyJSON, Open Food Facts) có giới hạn tần suất. Nhóm giới hạn quy mô ~10,000 bản ghi để đảm bảo tốc độ demo.", 10, False, C_BODY, 4),
        ("Hạn chế 2 — Độ thưa dữ liệu tương tác cao (Sparsity = 99.98%):", 12, True, C_AMBER, 10),
        ("• Phần lớn KH Amazon chỉ có 1 tương tác → gây Cold-Start nặng cho mô hình SVD. Content-Based bù lại nhờ đặc trưng sản phẩm.", 10, False, C_BODY, 4),
        ("Hạn chế 3 — Môi trường tính toán cục bộ:", 12, True, C_AMBER, 10),
        ("• DWH hiện chạy trên PostgreSQL Docker container cá nhân, chưa triển khai trên Cloud (AWS Redshift / Snowflake / Google BigQuery).", 10, False, C_BODY, 4),
    ],
    header=None, bg=C_WHITE, border=C_AMBER
)

print("Slide 14 xong")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 15 — HƯỚNG PHÁT TRIỂN & TỔNG KẾT Q&A
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
solid_rect(s, 0, 0, 13.333, 7.5, C_LIGHT)
slide_header_stripe(s, "14. Hướng Phát Triển Mở Rộng & Tổng Kết Đồ Án", section_label="Phần 6: Khuyến nghị & Hành động (2 phút)")
slide_footer(s, 15)

# Cột trái: Hướng phát triển
bullet_box(s, 0.35, 1.65, 5.9, 5.5,
    items=[
        ("Hướng Phát Triển Tiếp Theo", 13, True, C_BLUE, 0),
        ("🔄  Streaming Ingestion thời gian thực:", 12, True, C_BLUE, 10),
        ("Tích hợp Apache Kafka + Apache Flink để bắt sự kiện tương tác Clickstream và cập nhật DWH theo thời gian thực (Near-realtime ETL).", 10, False, C_BODY, 4),
        ("🤖  Mô hình Deep Learning Recommender:", 12, True, C_EMERALD, 10),
        ("Thử nghiệm kiến trúc Two-Tower DNN hoặc Graph Neural Networks (GNN) để nắm bắt mối quan hệ phi tuyến tính giữa user-item phức tạp.", 10, False, C_BODY, 4),
        ("⚙️  MLOps & CI/CD Pipeline tự động:", 12, True, C_AMBER, 10),
        ("Sử dụng Apache Airflow để lập lịch huấn luyện lại mô hình định kỳ hàng tuần và tự động hóa triển khai (CD).", 10, False, C_BODY, 4),
    ],
    header=None, bg=C_WHITE, border=C_BLUE
)

# Cột phải: Đối chiếu CLO & Cảm ơn
bullet_box(s, 6.55, 1.65, 6.45, 5.5,
    items=[
        ("✓  Tổng Kết — Đối Chiếu Chuẩn Đầu Ra CLO", 13, True, C_NAVY, 0),
        ("CLO1 (45%) — Kho Dữ Liệu & OLAP:", 12, True, C_EMERALD, 10),
        ("Hoàn thành Star Schema (1 Fact + 4 Dim); tính trước 2 Khối (CUBE & ROLLUP); trình diễn đủ 5 thao tác OLAP trực tiếp trên giao diện Web.", 10, False, C_BODY, 4),
        ("CLO2 (45%) — Khai Phá Dữ Liệu & Đánh Giá:", 12, True, C_EMERALD, 10),
        ("K-Means RFM đạt Silhouette = 0.6393 (K=4); FP-Growth: 7,674 luật (Lift 60x); Đánh giá Offline đa mô hình so với Baseline chuẩn.", 10, False, C_BODY, 4),
        ("CLO3 (10%) — Kỹ Năng Dự Án & Tự Động Hóa:", 12, True, C_EMERALD, 10),
        ("1 lệnh duy nhất: python run_all.py; Quản trị Git minh bạch; Báo cáo slide 15 trang chuẩn mực.", 10, False, C_BODY, 4),
        ("", 8, False, C_BODY, 5),
        ("NHÓM XIN TRÂN TRỌNG CẢM ƠN THẦY CÔ & HỘI ĐỒNG!", 13, True, C_NAVY, 8),
        ("Sẵn sàng cho phần Hỏi & Đáp — Q&A", 11.5, False, C_MUTED, 3),
    ],
    header=None, bg=C_WHITE, border=C_EMERALD
)

print("Slide 15 xong")


# ─── LƯU FILE ─────────────────────────────────────────────────────────────────
prs.save(str(OUTPUT))
print(f"\n HOÀN TẤT! Đã lưu: {OUTPUT}")
print(f"  Kích thước: {OUTPUT.stat().st_size // 1024} KB")
print(f"  Số slide: {len(prs.slides)}")
