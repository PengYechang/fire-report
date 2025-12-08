import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO

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

# --- 页面 UI 逻辑 (手机端优化版) ---
# layout="centered" 更适合手机竖屏阅读
st.set_page_config(page_title="消防检查助手", layout="centered")

# --- 状态初始化 ---
if 'all_reports' not in st.session_state:
    st.session_state.all_reports = {"默认项目": []}
if 'current_report_name' not in st.session_state:
    st.session_state.current_report_name = "默认项目"

# 获取当前数据引用
current_name = st.session_state.current_report_name
current_list = st.session_state.all_reports[current_name]

# --- 顶部：项目切换 (收纳在折叠栏中，节省空间) ---
with st.expander(f"📂 当前项目：{current_name} (点击切换)", expanded=False):
    report_names = list(st.session_state.all_reports.keys())
    selected_report = st.selectbox("选择或新建项目", report_names, index=report_names.index(current_name))

    if selected_report != st.session_state.current_report_name:
        st.session_state.current_report_name = selected_report
        st.rerun()

    new_report_name = st.text_input("新建项目名称", placeholder="输入新项目名")
    if st.button("新建并切换"):
        if new_report_name and new_report_name not in st.session_state.all_reports:
            st.session_state.all_reports[new_report_name] = []
            st.session_state.current_report_name = new_report_name
            st.rerun()

# --- 核心区域：添加问题 (默认展开) ---
st.markdown("### 📸 现场录入")
# 使用 container 包裹，稍微区分背景
with st.container(border=True):
    with st.form("mobile_add_form", clear_on_submit=True):
        # 第一行：位置 (手机打字慢，位置通常比较短，放前面)
        location = st.text_input("📍 问题位置", placeholder="如：8楼楼梯间")

        # 第二行：类别
        category = st.radio("⚠️ 问题类别", ["建筑防火问题清单", "消防设施问题清单"], horizontal=True)

        # 第三行：描述 (大文本框)
        desc = st.text_area("📝 问题描述", placeholder="描述具体隐患...", height=100)

        # 第四行：图片 (支持拍照 OR 上传)
        st.markdown("**📷 添加照片 (任选一种)**")
        col_cam, col_upl = st.tabs(["调用摄像头", "从相册上传"])

        with col_cam:
            camera_file = st.camera_input("点击拍照", label_visibility="collapsed")
        with col_upl:
            uploaded_file = st.file_uploader("选择文件", type=['png', 'jpg', 'jpeg'], label_visibility="collapsed")

        remark = st.text_input("💡 备注 (选填)", placeholder="整改人/建议")

        # 提交按钮
        submitted = st.form_submit_button("✅ 确认添加", use_container_width=True, type="primary")

        if submitted:
            if not desc or not location:
                st.error("位置和描述必填！")
            else:
                # 优先使用摄像头图片，如果没有则使用上传的图片
                final_img = camera_file if camera_file else uploaded_file
                img_data = final_img.getvalue() if final_img else None

                # 插入到列表的最前面 (inset 0)，这样手机端看列表时，最新的在最上面，不用滑到底部
                st.session_state.all_reports[current_name].insert(0, {
                    "category": category,
                    "desc": desc,
                    "loc": location,
                    "remark": remark,
                    "img_bytes": img_data,
                    "img_name": final_img.name if final_img else "无图片"
                })
                st.success("已添加！")
                st.rerun()

# --- 列表展示区 (卡片式) ---
st.markdown("---")
st.markdown(f"### 📋 已记录 ({len(current_list)})")

if not current_list:
    st.info("暂无记录，请在上方添加。")
else:
    # 遍历显示（虽然数据已经是倒序插入了，但为了保险还是用索引定位方便删除）
    # 为了删除方便，我们需要保留原始索引。这里稍微处理一下展示逻辑。
    # 实际显示时，我们直接显示 current_list，因为它已经是“最新在最前”了。

    for i, item in enumerate(current_list):
        # 使用 border=True 创建卡片感
        with st.container(border=True):
            # 第一行：位置 + 类别标签
            col_top_1, col_top_2 = st.columns([3, 1])
            with col_top_1:
                st.markdown(f"**📍 {item['loc']}**")
            with col_top_2:
                # 简单的颜色区分
                tag_color = "red" if "防火" in item['category'] else "orange"
                st.caption(f":{tag_color}[{item['category'][:4]}]")

            # 第二行：描述
            st.text(item['desc'])

            # 第三行：如果有图，显示缩略图
            if item['img_bytes']:
                st.image(item['img_bytes'], width=150) # 限制宽度，防止手机刷屏

            # 第四行：备注 + 删除按钮
            col_foot_1, col_foot_2 = st.columns([3, 1])
            with col_foot_1:
                if item['remark']:
                    st.caption(f"备注: {item['remark']}")
            with col_foot_2:
                # 删除按钮
                if st.button("🗑️", key=f"del_{i}"):
                    current_list.pop(i)
                    st.rerun()

# --- 底部：下载区域 ---
st.markdown("---")
# 将生成逻辑预先处理
doc_object = create_word_file(current_name, current_list[::-1]) # 生成时反转回去，让序号1对应最早录入的
output_buffer = BytesIO()
doc_object.save(output_buffer)
output_buffer.seek(0)

st.download_button(
    label="📥 生成并下载 Word 报告",
    data=output_buffer,
    file_name=f"{current_name}_消防问题清单.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True, # 按钮撑满宽度，方便手机点击
    type="primary"
)

# 留一点底部空白，防止按钮贴底
st.write("")
st.write("")