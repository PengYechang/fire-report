import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO

# --- 核心逻辑：生成 Word 文档 ---
def set_font(run, font_name='宋体', size=10, bold=False):
    """辅助函数：设置字体样式"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold

def create_word_file(report_name, data_list):
    """
    根据传入的数据列表生成 Word 文档对象
    data_list 结构: [{'category': 'Building Fire', 'desc': 'xxx', 'loc': 'xxx', 'remark': 'xxx', 'img_bytes': binary}, ...]
    """
    doc = Document()

    # 设置页边距
    section = doc.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)

    # 1. 大标题 (使用报告名称)
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # 如果用户没改名，默认叫这个，否则可以用 report_name 拼接到标题里
    run_title = title_p.add_run(f'{report_name} - 消防检查问题清单')
    set_font(run_title, '黑体', 18, bold=True)

    # 定义两类问题
    categories = ["建筑防火问题清单", "消防设施问题清单"]

    for cat_name in categories:
        # 筛选当前类别的数据
        current_items = [item for item in data_list if item['category'] == cat_name]

        doc.add_paragraph("") # 空行

        # 类别标题
        prefix = "一、" if cat_name == "建筑防火问题清单" else "二、"
        h_p = doc.add_paragraph()
        run_h = h_p.add_run(f"{prefix}{cat_name}")
        set_font(run_h, '黑体', 14, bold=True)

        if not current_items:
            p_none = doc.add_paragraph("（该项无问题）")
            p_none.alignment = WD_ALIGN_PARAGRAPH.LEFT
            continue

        # 创建表格
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        table.autofit = False

        # 设置列宽
        widths = [Cm(1.5), Cm(7), Cm(6), Cm(2.5)]

        # 表头
        headers = ["序号", "问题描述", "相关照片", "备注"]
        hdr_cells = table.rows[0].cells
        for i, text in enumerate(headers):
            hdr_cells[i].width = widths[i]
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            set_font(run, '宋体', 12, bold=True)

        # 填充内容
        for idx, item in enumerate(current_items, 1):
            row_cells = table.add_row().cells

            # 1. 序号
            p1 = row_cells[0].paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(p1.add_run(str(idx)))

            # 2. 问题描述与位置
            p2 = row_cells[1].paragraphs[0]
            run_desc = p2.add_run(f"问题描述：{item['desc']}\n")
            set_font(run_desc)
            run_loc = p2.add_run(f"问题位置：{item['loc']}")
            set_font(run_loc)

            # 3. 图片
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

            # 4. 备注
            p4 = row_cells[3].paragraphs[0]
            p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(p4.add_run(item['remark']))

    return doc

# --- 页面 UI 逻辑 ---
st.set_page_config(page_title="多项目消防检查工具", layout="wide")

# --- 状态初始化 ---
# all_reports 结构: { "项目A": [item1, item2], "项目B": [] }
if 'all_reports' not in st.session_state:
    st.session_state.all_reports = {"默认项目": []}

if 'current_report_name' not in st.session_state:
    st.session_state.current_report_name = "默认项目"

# 获取当前选中的列表（引用）
current_name = st.session_state.current_report_name
current_list = st.session_state.all_reports[current_name]

# --- 侧边栏：报告管理 & 添加数据 ---
with st.sidebar:
    st.title("📂 报告管理")

    # 1. 切换/新建报告
    with st.expander("切换或新建项目", expanded=True):
        # 切换报告
        report_names = list(st.session_state.all_reports.keys())
        # 确保当前选中项在列表里
        if current_name not in report_names:
            current_name = report_names[0]
            st.session_state.current_report_name = current_name

        selected_report = st.selectbox(
            "选择当前操作的项目：",
            report_names,
            index=report_names.index(current_name)
        )

        # 如果用户切换了下拉框，更新状态并刷新
        if selected_report != st.session_state.current_report_name:
            st.session_state.current_report_name = selected_report
            st.rerun()

        st.markdown("---")
        # 新建报告
        new_report_name = st.text_input("新建项目名称", placeholder="例如：汉口分店检查")
        if st.button("➕ 创建新项目"):
            if new_report_name and new_report_name not in st.session_state.all_reports:
                st.session_state.all_reports[new_report_name] = []
                st.session_state.current_report_name = new_report_name
                st.success(f"已创建并切换至：{new_report_name}")
                st.rerun()
            elif new_report_name in st.session_state.all_reports:
                st.warning("该项目名称已存在！")
            else:
                st.warning("请输入名称")

        # 删除当前报告
        if st.button("🗑️ 删除当前项目", type="primary"):
            if len(st.session_state.all_reports) <= 1:
                st.error("至少保留一个项目！")
            else:
                del st.session_state.all_reports[current_name]
                # 删除后默认切回第一个
                st.session_state.current_report_name = list(st.session_state.all_reports.keys())[0]
                st.rerun()

    st.markdown("---")

    # 2. 添加数据表单
    st.header(f"📝 添加记录到: {st.session_state.current_report_name}")

    with st.form("add_form", clear_on_submit=True):
        category = st.radio("问题类别", ["建筑防火问题清单", "消防设施问题清单"])
        desc = st.text_area("问题描述", placeholder="描述具体隐患...")
        location = st.text_input("问题位置", placeholder="具体楼层/区域")
        remark = st.text_input("备注", placeholder="整改建议或责任人")
        uploaded_file = st.file_uploader("现场照片", type=['png', 'jpg', 'jpeg'])

        submitted = st.form_submit_button("添加条目")

        if submitted:
            if not desc or not location:
                st.error("【问题描述】和【问题位置】必填！")
            else:
                img_data = uploaded_file.getvalue() if uploaded_file else None

                # 直接添加到当前选中的 report list 中
                st.session_state.all_reports[st.session_state.current_report_name].append({
                    "category": category,
                    "desc": desc,
                    "loc": location,
                    "remark": remark,
                    "img_bytes": img_data,
                    "img_name": uploaded_file.name if uploaded_file else "无图片"
                })
                st.success("添加成功！")
                st.rerun() # 强制刷新以立即在主界面显示

# --- 主区域 ---
st.title(f"📊 {st.session_state.current_report_name} - 检查清单")

# 1. 列表预览
if len(current_list) == 0:
    st.info(f"项目【{st.session_state.current_report_name}】暂无数据，请在左侧侧边栏添加。")
else:
    # 构造显示的表格
    display_data = []
    for i, item in enumerate(current_list):
        display_data.append({
            "序号": i + 1,
            "类别": item['category'],
            "位置": item['loc'],
            "描述": item['desc'],
            "备注": item['remark'],
            "照片": "✅ 有" if item['img_bytes'] else ""
        })
    st.dataframe(display_data, use_container_width=True)

    # 列表操作区
    col_del_idx, col_del_btn = st.columns([1, 4])
    with col_del_idx:
        del_idx = st.number_input("条目序号", min_value=1, max_value=len(current_list), step=1, key="del_idx")
    with col_del_btn:
        st.write("")
        st.write("")
        if st.button("删除该条目", type="secondary"):
            st.session_state.all_reports[st.session_state.current_report_name].pop(del_idx - 1)
            st.rerun()

    st.markdown("---")

    # 2. 生成下载区
    st.subheader("📥 生成文档")

    # 生成文档
    doc_object = create_word_file(st.session_state.current_report_name, current_list)
    output_buffer = BytesIO()
    doc_object.save(output_buffer)
    output_buffer.seek(0)

    file_name = f"{st.session_state.current_report_name}_消防检查清单.docx"

    col1, col2 = st.columns([1, 1])
    with col1:
        st.download_button(
            label=f"⬇️ 下载 {file_name}",
            data=output_buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
    with col2:
        if st.button("⚠️ 清空当前项目所有数据"):
            st.session_state.all_reports[st.session_state.current_report_name] = []
            st.rerun()