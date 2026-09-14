import streamlit as st
import pandas as pd
import math
import io

# ================= 核心算法模块 =================

@st.cache_data
def clean_and_parse_data(file_buffer):
    """【数据解析】拉平表格，并处理接力队合并"""
    xls = pd.ExcelFile(file_buffer)
    parsed_data = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(file_buffer, sheet_name=sheet, header=None)
        
        current_class, current_gender, current_events = "", "", []
        
        for i in range(len(df)):
            row = df.iloc[i].values
            row_head = str(row[0]).strip() if pd.notna(row[0]) else ""
            
            if row_head == "班级":
                current_class = str(row[1]).strip()
            elif row_head in ["女生", "男生"]:
                current_gender = row_head
                current_events = row[1:] 
            elif row_head == "姓名" and i + 1 < len(df):
                next_row = df.iloc[i+1].values
                next_row_head = str(next_row[0]).strip() if pd.notna(next_row[0]) else ""
                
                if "号码" in next_row_head or next_row_head == "号码":
                    names, numbers = row[1:], next_row[1:]
                    for j in range(len(names)):
                        name = str(names[j]).strip()
                        num = numbers[j]
                        event = str(current_events[j]).strip() if j < len(current_events) and pd.notna(current_events[j]) else ""
                        
                        if pd.notna(names[j]) and name and name != "nan" and "裁判" not in name and "注" not in name:
                            if pd.notna(num) and str(num).strip() not in ["", "nan"] and event and event != "nan":
                                if "接力" in event:
                                    short_class = sheet + current_class.replace(sheet, "").replace("班", "")
                                    name, num = f"{short_class}队", short_class
                                    
                                parsed_data.append({
                                    "Grade": sheet, "Class": current_class,
                                    "Name": name, "Number": num,
                                    "Gender": current_gender, "Event": event
                                })
                                        
    df_all = pd.DataFrame(parsed_data)
    df_relays = df_all[df_all['Event'].str.contains('接力')].drop_duplicates(subset=["Grade", "Class", "Gender", "Event"])
    df_individuals = df_all[~df_all['Event'].str.contains('接力')]
    
    return pd.concat([df_individuals, df_relays], ignore_index=True)

def generate_schedule(all_data, max_p_input, start_lane, end_lane, random_seed, scatter_class):
    """【算法排表】接入 UI 参数的动态排表引擎"""
    final_schedule = []
    grouped = all_data.groupby(["Grade", "Gender", "Event"])
    
    lane_list = [f"第{i}道" for i in range(start_lane, end_lane + 1)]
    
    for (grade, gender, event), group in grouped:
        participants = group.sample(frac=1, random_state=random_seed).to_dict('records')
        is_lane = any(x in event for x in ["100", "200", "400", "接力"])
        max_p = max_p_input if is_lane else len(participants)
        heats = []
        
        for p in participants:
            placed = False
            p_class = p["Class"]
            
            if scatter_class:
                for heat in heats:
                    if len(heat) < max_p and not any(member["Class"] == p_class for member in heat):
                        heat.append(p)
                        placed = True
                        break
            
            if not placed:
                for heat in heats:
                    if len(heat) < max_p:
                        heat.append(p)
                        placed = True
                        break
                        
            if not placed:
                heats.append([p])
                
        for h_idx, heat in enumerate(heats):
            for l_idx, member in enumerate(heat):
                member["Heat"] = "第一组" if h_idx == 0 and not is_lane else f"第{h_idx + 1}组"
                if is_lane:
                    member["Lane"] = lane_list[l_idx] if l_idx < len(lane_list) else ""
                else:
                    member["Lane"] = f"{l_idx + 1}号"
                final_schedule.append(member)

    return pd.DataFrame(final_schedule)

def export_pdf_style_format(result_df):
    """【排版引擎】将一维数据重塑为 PDF 视觉网格"""
    output_lines = []
    track_events = ["100M", "200M", "400M", "800M", "1000M", "4*100M接力"]
    field_events = ["跳高", "跳远", "铅球", "铅球5kg", "侧向推实心球（女2kg)"]
    
    for gender in ["女生", "男生"]:
        gender_prefix = "女子" if gender == "女生" else "男子"
        
        output_lines.append(["====== 【径赛】 " + gender_prefix + " ======"])
        for event in track_events:
            event_df = result_df[(result_df['性别'] == gender) & (result_df['比赛项目'] == event)]
            if event_df.empty: continue
            
            output_lines.append([""])
            output_lines.append([f"{gender_prefix}{event}"])
            
            for grade in ["高一", "高二", "高三"]:
                df_g = event_df[event_df['年级'] == grade]
                if df_g.empty: continue
                
                total_people = len(df_g)
                num_heats = df_g['组别'].nunique()
                
                stage = "预赛" if "100" in event and "接力" not in event else "预决赛"
                unit = "队" if "接力" in event else "人"
                output_lines.append([f"{grade}组{stage} ({total_people}{unit}分{num_heats}组)"])
                
                is_lane = any(x in event for x in ["100", "200", "400", "接力"])
                if is_lane:
                    active_lanes = sorted(df_g[df_g["道次/序号"] != ""]["道次/序号"].unique())
                    output_lines.append(["道次", "号码"] + [""] * (len(active_lanes) - 1))
                    output_lines.append(["组别"] + active_lanes)
                    
                    pivot = df_g.pivot_table(index="组别", columns="道次/序号", values="号码", aggfunc=lambda x: ' '.join(str(v) for v in x)).fillna("")
                    for index, row in pivot.iterrows():
                        row_data = [index] + [row[col] if col in pivot.columns else "" for col in active_lanes]
                        output_lines.append(row_data)
                else:
                    output_lines.append(["组别", "", "", "号码", "", ""])
                    numbers = df_g['号码'].tolist()
                    for i in range(0, len(numbers), 6):
                        chunk = numbers[i:i+6]
                        output_lines.append(["第一组" if i == 0 else ""] + chunk + [""] * (6 - len(chunk)))
        
        output_lines.append([""])
        output_lines.append(["====== 【田赛】 " + gender_prefix + " ======"])
        for event in field_events:
            event_df = result_df[(result_df['性别'] == gender) & (result_df['比赛项目'] == event)]
            if event_df.empty: continue
            
            output_lines.append([""])
            output_lines.append([f"{gender_prefix}{event}"])
            for grade in ["高一", "高二", "高三"]:
                df_g = event_df[event_df['年级'] == grade]
                if df_g.empty: continue
                output_lines.append([f"{grade}组预决赛 ({len(df_g)}人一组)"])
                numbers = df_g['号码'].tolist()
                for i in range(0, len(numbers), 10):
                    output_lines.append([" ".join(str(v) for v in numbers[i:i+10])])

    return pd.DataFrame(output_lines)

# ================= 网页 UI 构建 =================

st.set_page_config(page_title="北外田园运动会自动分组系统", page_icon="🏃", layout="wide")

# 注入 CSS 放大全局字体，优化视觉排版
st.markdown("""
<style>
.stMarkdown p, .stText { font-size: 18px !important; }
.stButton>button { font-size: 18px !important; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

st.sidebar.markdown("### ⚙️ 分组设置")
st.sidebar.markdown("<br>", unsafe_allow_html=True)
max_p_input = st.sidebar.number_input("每组人数", min_value=1, max_value=12, value=6, step=1)
start_lane = st.sidebar.number_input("起始道次", min_value=1, max_value=9, value=1, step=1)
end_lane = st.sidebar.number_input("结束道次", min_value=1, max_value=9, value=8, step=1)
random_seed = st.sidebar.number_input("随机种子", value=2025, step=1, help="相同的种子值能保证每次运算出的排表结果完全一致。")

st.sidebar.markdown("<hr>", unsafe_allow_html=True)
balance_gender = st.sidebar.checkbox("尽量平衡男女比例", value=True)
scatter_class = st.sidebar.checkbox("尽量分散同班学生", value=True)
balance_grade = st.sidebar.checkbox("尽量平衡年级", value=True)

st.title("🏃 北外田园运动会自动分组系统")
st.markdown("<span style='color:gray; font-size:16px;'>适用于“班级报名表 -> 项目名单 -> 自动分组 -> Excel导出”的流程</span>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 上传报名表 Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    clean_df = clean_and_parse_data(uploaded_file)
    st.success(f"成功读取 {len(clean_df)} 条“学生-项目”报名记录。")
    
    unique_events = sorted(clean_df['Event'].unique().tolist())
    
    tab1, tab2, tab3 = st.tabs(["📋 报名数据", "🏃 自动分组", "📊 数据统计"])
    
    with tab1:
        st.markdown("##### 数据筛选检索")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            grade_filter = st.multiselect("📌 按年级筛选", options=sorted(clean_df['Grade'].unique()))
        with col2:
            class_filter = st.multiselect("🏫 按班级筛选", options=sorted(clean_df['Class'].unique()))
        with col3:
            gender_filter = st.multiselect("🚻 按性别筛选", options=sorted(clean_df['Gender'].unique()))
        with col4:
            event_filter = st.multiselect("🏅 按项目筛选", options=unique_events)
            
        filtered_df = clean_df.copy()
        if grade_filter:
            filtered_df = filtered_df[filtered_df['Grade'].isin(grade_filter)]
        if class_filter:
            filtered_df = filtered_df[filtered_df['Class'].isin(class_filter)]
        if gender_filter:
            filtered_df = filtered_df[filtered_df['Gender'].isin(gender_filter)]
        if event_filter:
            filtered_df = filtered_df[filtered_df['Event'].isin(event_filter)]
            
        display_df = filtered_df.copy()
        display_df.columns = ["年级", "班级", "姓名", "号码", "性别", "报名项目"]
        
        st.markdown(f"当前筛选条件共匹配到 **{len(display_df)}** 条记录。")
        st.dataframe(display_df, use_container_width=True, hide_index=True)
        
    with tab2:
        st.markdown("##### 选择需要编排的项目")
        selected_events = st.multiselect("", unique_events, default=unique_events, label_visibility="collapsed")
        
        if st.button("🚀 开始自动分组", type="primary"):
            if not selected_events:
                st.warning("请至少选择一个项目！")
            else:
                with st.spinner("系统正在结合侧边栏参数进行运算..."):
                    target_df = clean_df[clean_df['Event'].isin(selected_events)]
                    result_df = generate_schedule(target_df, max_p_input, start_lane, end_lane, random_seed, scatter_class)
                    
                    cols_order = ["Grade", "Gender", "Event", "Heat", "Lane", "Class", "Name", "Number"]
                    result_df = result_df[cols_order]
                    result_df.columns = ["年级", "性别", "比赛项目", "组别", "道次/序号", "班级", "姓名", "号码"]
                    
                    visual_df = export_pdf_style_format(result_df)
                    
                st.success("分组编排完成！")
                st.dataframe(visual_df, use_container_width=True, hide_index=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    visual_df.to_excel(writer, index=False, header=False, sheet_name='分组检录表(排版版)')
                    result_df.to_excel(writer, index=False, sheet_name='后台数据源')
                
                st.download_button(
                    label="💾 下载 Excel",
                    data=output.getvalue(),
                    file_name="田径运动会_自动分组结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    with tab3:
        st.markdown("##### 各项目报名人数概览")
        stats_df = clean_df.groupby(['Grade', 'Gender', 'Event']).size().reset_index(name='报名人数/队数')
        stats_df.columns = ['年级', '性别', '项目', '报名人数/队数']
        st.dataframe(stats_df, use_container_width=True, hide_index=True)
