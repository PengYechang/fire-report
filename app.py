import streamlit as st
import sqlite3
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO

# --- 数据库配置与工具函数 ---

DB_FILE = "fire_inspections.db"

def init_db():
    """初始化数据库表"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # 创建表：包含自增ID、项目名、相关字段和图片二进制数据
    c.execute('''
              CREATE TABLE IF NOT EXISTS inspections (
                                                         id INTEGER PRIMARY KEY AUTOINCREMENT,
                                                         project_name TEXT NOT NULL,
                                                         category TEXT,
                                                         loc TEXT,
                                                         desc TEXT,
                                                         remark TEXT,
                                                         img_bytes BLOB,
                                                         created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
              )
              ''')
    conn.commit()
    conn.close()

def get_all_projects():
    """获取所有唯一的项目名称"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT DISTINCT project_name FROM inspections ORDER BY created_at DESC")
    projects = [row[0] for row in c.fetchall()]
    conn.close()
    if not projects:
        return ["默认项目"]
    return projects

def add_item_to_db(project, category, loc, desc, remark, img_bytes):
    """添加一条记录"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
              INSERT INTO inspections (project_name, category, loc, desc, remark, img_bytes)
              VALUES (?, ?, ?, ?, ?, ?)
              ''', (project, category, loc, desc, remark, img_bytes))
    conn.commit()
    conn.close()

def get_items_by_project(project_name):
    """获取指定项目的所有记录 (按时间倒序，最新的在最前)"""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row # 让结果可以通过列名访问
    c = conn.cursor()
    c.execute("SELECT * FROM inspections WHERE project_name = ? ORDER BY id DESC", (project_name,))
    rows = c.fetchall()
    conn.close()

    # 将 sqlite3.Row 对象转换为字典列表，兼容之前的逻辑
    data_list = []
    for row in rows:
        data_list.append({
            "id": row["id"], # 用于删除
            "category": row["category"],
            "desc": row["desc"],
            "loc": row["loc"],
            "remark": row["remark"],
            "img_bytes": row["img_bytes"]
        })
    return data_list

def delete_item_from_db(item_id):
    """根据ID删除记录"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM inspections WHERE id = ?", (item_id,))
    conn.commit()
    conn.close()

# 初始化数据库
init_db()

# --- 核心逻辑：生成 Word 文档 (保持不变) ---
def set_font(run, font_name='宋体', size=10, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold

def create_word_file(report_name, data_list):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run(f'{report_name} - 消防检查问题清单')
    set_font(run_title, '黑体', 18, bold=True)

    categories = ["建筑防火问题清单", "消防设施问题清单"]

    for cat_name in categories:
        current_items = [item for item in data_list if item['category'] == cat_name]
        doc.add_paragraph("")
        prefix = "一、" if cat_name == "建筑防火问题清单" else "二、"
        h_p = doc.add_paragraph()
        run_h = h_p.add_run(f"{prefix}{cat_name}")
        set_font(run_h, '黑体', 14, bold=True)

        if not current_items:
            p_none = doc.add_paragraph("（该项无问题）")
            p_none.alignment = WD_ALIGN_PARAGRAPH.LEFT
            continue

        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        table.autofit = False
        widths = [Cm(1.5), Cm(7), Cm(6), Cm(2.5)]

        headers = ["序号", "问题描述", "相关照片", "备注"]
        hdr_cells = table.rows[0].cells
        for i, text in enumerate(headers):
            hdr_cells[i].width = widths[i]
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            set_font(run, '宋体', 12, bold=True)

        for idx, item in enumerate(current_items, 1):
            row_cells = table.add_row().cells

            p1 = row_cells[0].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(p1.add_run(str(idx)))

            p2 = row_cells[1].paragraphs[0]
            run_desc = p2.add_run(f"问题描述：{item['desc']}\n")
            set_font(run_desc)
            run_loc = p2.add_run(f"问题位置：{item['loc']}")
            set_font(run_loc)

            cell_img = row_cells[2]
            p3 = cell_img.paragraphs[0]
            p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if item['img_bytes']:
                try:
                    run_img = p3.add_run()
                    run_img.add_picture(BytesIO(item['img_bytes']), width=Inches(2.0))
                except:
                    set_font(p3.add_run("[图片格式错误]"))
            else:
                set_font(p3.add_run("/"))

            p4 = row_cells[3].paragraphs[0]
            p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(p4.add_run(item['remark']))

    return doc

# --- 页面 UI 逻辑 ---
st.set_page_config(page_title="消防检查助手", layout="centered")

# --- 状态管理 ---
if 'current_report_name' not in st.session_state:
    st.session_state.current_report_name = "默认项目"

# 获取数据库中的项目列表
db_projects = get_all_projects()
# 确保当前选中的项目在列表中，否则重置为第一个
if st.session_state.current_report_name not in db_projects:
    if "默认项目" not in db_projects:
        # 如果是刚开始没有任何项目，列表中至少有默认项目
        pass
    else:
        st.session_state.current_report_name = db_projects[0]

current_name = st.session_state.current_report_name

# --- 顶部：项目切换 ---
with st.expander(f"📂 当前项目：{current_name} (点击切换)", expanded=False):
    # 选择框直接使用数据库里的项目名
    selected_report = st.selectbox("选择已有项目", db_projects, index=db_projects.index(current_name) if current_name in db_projects else 0)

    if selected_report != current_name:
        st.session_state.current_report_name = selected_report
        st.rerun()

    new_report_name = st.text_input("新建项目名称", placeholder="输入新项目名 (如：万达广场)")
    if st.button("新建并切换"):
        if new_report_name:
            # 新建时，我们不需要立刻往数据库建表，
            # 只要切换了名字，下次添加问题时就会自动关联这个新名字
            st.session_state.current_report_name = new_report_name
            st.rerun()

# --- 从数据库加载当前项目的数据 ---
current_list = get_items_by_project(current_name)

# --- 核心区域：添加问题 ---
st.markdown("### 📸 现场录入")

with st.container(border=True):
    with st.form("mobile_add_form", clear_on_submit=True):
        location = st.text_input("📍 问题位置", placeholder="如：8楼楼梯间")
        category = st.radio("⚠️ 问题类别", ["建筑防火问题清单", "消防设施问题清单"], horizontal=True)
        desc = st.text_area("📝 问题描述", placeholder="描述具体隐患...", height=100)

        st.markdown("**📷 添加照片 (任选一种)**")
        col_cam, col_upl = st.tabs(["调用摄像头", "从相册上传"])
        with col_cam:
            camera_file = st.camera_input("点击拍照", label_visibility="collapsed")
        with col_upl:
            uploaded_file = st.file_uploader("选择文件", type=['png', 'jpg', 'jpeg'], label_visibility="collapsed")

        remark = st.text_input("💡 备注 (选填)", placeholder="整改人/建议")

        submitted = st.form_submit_button("✅ 确认添加", use_container_width=True, type="primary")

        if submitted:
            if not desc or not location:
                st.error("位置和描述必填！")
            else:
                final_img = camera_file if camera_file else uploaded_file
                img_data = final_img.getvalue() if final_img else None

                # --- 修改点：写入数据库 ---
                add_item_to_db(current_name, category, location, desc, remark, img_data if img_data else b'')

                st.success("已保存到数据库！")
                st.rerun()

# --- 列表展示区 ---
st.markdown("---")
st.markdown(f"### 📋 已记录 ({len(current_list)})")

if not current_list:
    st.info("当前项目暂无记录，请在上方添加。")
else:
    for item in current_list:
        with st.container(border=True):
            col_top_1, col_top_2 = st.columns([3, 1])
            with col_top_1:
                st.markdown(f"**📍 {item['loc']}**")
            with col_top_2:
                tag_color = "red" if "防火" in item['category'] else "orange"
                st.caption(f":{tag_color}[{item['category'][:4]}]")

            st.text(item['desc'])

            if item['img_bytes']:
                st.image(item['img_bytes'], width=150)

            col_foot_1, col_foot_2 = st.columns([3, 1])
            with col_foot_1:
                if item['remark']:
                    st.caption(f"备注: {item['remark']}")
            with col_foot_2:
                # --- 修改点：删除时使用数据库ID ---
                # 使用 key 防止按钮ID重复
                if st.button("🗑️", key=f"del_{item['id']}"):
                    delete_item_from_db(item['id'])
                    st.rerun()

# --- 底部：下载区域 ---
st.markdown("---")
# 生成 Word 时需要反转列表，因为数据库查出来是 "最新在最前"，
# 但 Word 报告里通常希望序号 1 对应 "最早发现的问题"
doc_object = create_word_file(current_name, current_list[::-1])
output_buffer = BytesIO()
doc_object.save(output_buffer)
output_buffer.seek(0)

st.download_button(
    label="📥 生成并下载 Word 报告",
    data=output_buffer,
    file_name=f"{current_name}_消防问题清单.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True,
    type="primary"
)

st.write("")
st.write("")