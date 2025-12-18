import streamlit as st
import pandas as pd
import sqlite3
import os
import base64
import io
import plotly.express as px
from datetime import datetime
from zhipuai import ZhipuAI
import hashlib # 用于基础密码加密

# --- 1. 配置与数据库初始化 ---
ZHIPU_API_KEY = "Your API_KEY" 
client = ZhipuAI(api_key=ZHIPU_API_KEY)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
db_path = os.path.join(BASE_DIR, "medical_records_multi.db")

conn = sqlite3.connect(db_path, check_same_thread=False)
c = conn.cursor()

def init_db():
    # 用户表
    c.execute('''CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password TEXT)''')
    # 病人表
    c.execute('''CREATE TABLE IF NOT EXISTS patient_info (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, gender TEXT, age INTEGER, disease TEXT, diagnosis_date TEXT, allergies TEXT)''')
    # 报告表
    c.execute('''CREATE TABLE IF NOT EXISTS medical_reports (id INTEGER PRIMARY KEY AUTOINCREMENT, patient_id INTEGER, date TEXT, category TEXT, summary TEXT, full_text TEXT)''')
    # 记录表
    c.execute('''CREATE TABLE IF NOT EXISTS treatment_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, patient_id INTEGER, date TEXT, treat_type TEXT, hospital TEXT, details TEXT, side_effects TEXT)''')
    # 用药表
    c.execute('''CREATE TABLE IF NOT EXISTS medication_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, patient_id INTEGER, drug_name TEXT, dosage TEXT, frequency TEXT, status TEXT)''')
    conn.commit()

init_db()

# --- 2. 登录与身份验证逻辑 ---
def make_hashes(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

def check_hashes(password, hashed_text):
    if make_hashes(password) == hashed_text:
        return hashed_text
    return False

def login_user(username, password):
    c.execute('SELECT * FROM users WHERE username =? AND password =?', (username, password))
    data = c.fetchall()
    return data

def add_user(username, password):
    c.execute('INSERT INTO users(username,password) VALUES (?,?)', (username, password))
    conn.commit()

# --- 3. 界面逻辑控制 ---
st.set_page_config(page_title="智能健康管理系统", layout="wide")

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

# 登录/注册界面
if not st.session_state.logged_in:
    st.header("🔐 智能健康管理系统 - 身份验证")
    auth_mode = st.tabs(["用户登录", "新用户注册"])
    
    with auth_mode[0]:
        user = st.text_input("用户名", key="login_user")
        pwd = st.text_input("密码", type="password", key="login_pwd")
        if st.button("立即登录"):
            hashed_pswd = make_hashes(pwd)
            result = login_user(user, check_hashes(pwd, hashed_pswd))
            if result:
                st.session_state.logged_in = True
                st.session_state.username = user
                st.success(f"欢迎回来, {user}")
                st.rerun()
            else:
                st.error("用户名或密码错误")
                
    with auth_mode[1]:
        new_user = st.text_input("设置用户名", key="reg_user")
        new_pwd = st.text_input("设置密码", type="password", key="reg_pwd")
        if st.button("提交注册"):
            try:
                add_user(new_name, make_hashes(new_pwd))
                st.success("注册成功，请切换到登录页。")
            except:
                st.error("该用户名已存在")
else:
    # --- 已登录后的主体程序 ---
    st.sidebar.title(f"👤 当前用户: {st.session_state.username}")
    if st.sidebar.button("安全退出"):
        st.session_state.logged_in = False
        st.rerun()
    
    st.sidebar.divider()
    st.sidebar.subheader("👥 成员档案选择")

    # 获取病人列表
    def get_patient_list():
        return pd.read_sql_query("SELECT id, name FROM patient_info", conn)

    patient_df = get_patient_list()
    if patient_df.empty:
        st.sidebar.warning("请新建成员档案")
        current_patient_id = None
    else:
        patient_options = {row['name']: row['id'] for _, row in patient_df.iterrows()}
        selected_patient_name = st.sidebar.selectbox("当前操作对象：", list(patient_options.keys()))
        current_patient_id = patient_options[selected_patient_name]

    st.sidebar.divider()
    menu = st.sidebar.radio("功能导航", ["➕ 管理成员", "👤 个人资料", "📂 报告识别", "💉 治疗记录", "💊 服药管理", "🔍 检索修改", "📊 健康分析看板"])

    # 底部版权标识
    st.sidebar.markdown("---")
    st.sidebar.markdown(
        """
        <div style='text-align: center; color: grey; font-size: 0.8em;'>
            © 2024 Eumenes Studios<br>
            All Rights Reserved
        </div>
        """, 
        unsafe_allow_html=True
    )

    # --- 辅助函数 ---
    def to_excel(p_id):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.read_sql_query(f"SELECT * FROM patient_info WHERE id={p_id}", conn).to_excel(writer, sheet_name='资料', index=False)
            pd.read_sql_query(f"SELECT * FROM medical_reports WHERE patient_id={p_id}", conn).to_excel(writer, sheet_name='报告', index=False)
            pd.read_sql_query(f"SELECT * FROM treatment_logs WHERE patient_id={p_id}", conn).to_excel(writer, sheet_name='治疗', index=False)
            pd.read_sql_query(f"SELECT * FROM medication_logs WHERE patient_id={p_id}", conn).to_excel(writer, sheet_name='用药', index=False)
        return output.getvalue()

    # --- 业务逻辑模块 ---
    if menu == "➕ 管理成员":
        st.header("👥 家庭成员空间")
        new_name = st.text_input("新增成员姓名")
        if st.button("立即创建"):
            if new_name:
                c.execute("INSERT INTO patient_info (name) VALUES (?)", (new_name,))
                conn.commit()
                st.rerun()
        st.dataframe(patient_df, use_container_width=True)

    elif not current_patient_id:
        st.warning("请在侧边栏选择成员")

    elif menu == "👤 个人资料":
        st.header(f"👤 {selected_patient_name} 的资料")
        c.execute(f"SELECT * FROM patient_info WHERE id={current_patient_id}")
        d = c.fetchone()
        c1, c2 = st.columns(2)
        u_name = c1.text_input("姓名", value=d[1] if d[1] else "")
        u_gender = c1.selectbox("性别", ["男", "女"], index=0 if d[2]=="男" else 1)
        u_age = c1.number_input("年龄", value=d[3] if d[3] else 0)
        u_disease = c2.text_input("主诊病史", value=d[4] if d[4] else "")
        u_diag_date = c2.text_input("诊断日期", value=d[5] if d[5] else "")
        u_allergies = st.text_area("🚨 药物过敏史", value=d[6] if d[6] else "无")
        if st.button("更新保存"):
            c.execute("UPDATE patient_info SET name=?, gender=?, age=?, disease=?, diagnosis_date=?, allergies=? WHERE id=?", (u_name, u_gender, u_age, u_disease, u_diag_date, u_allergies, current_patient_id))
            conn.commit()
            st.success("已保存")

    elif menu == "📂 报告识别":
        st.header(f"📂 AI 全文提取 - {selected_patient_name}")
        if 'ai_res' not in st.session_state: st.session_state.ai_res = {"category": "其他", "summary": "", "full_text": "", "status": "等待上传"}
        file = st.file_uploader("上传图片", type=["jpg", "png", "jpeg"])
        if file:
            c1, c2 = st.columns([1, 1.2])
            with c1:
                st.image(file, use_container_width=True)
                if st.button("✨ 启动识别"):
                    img_str = base64.b64encode(file.getvalue()).decode('utf-8')
                    resp = client.chat.completions.create(model="glm-4v-flash", messages=[{"role":"user","content":[{"type":"text","text":"识别文字。格式：分类\n===\n总结\n===\n全文"},{"type":"image_url","image_url":{"url":img_str}}]}])
                    res = resp.choices[0].message.content
                    if "===" in res:
                        p = res.split("===")
                        st.session_state.ai_res = {"category":p[0].strip(),"summary":p[1].strip(),"full_text":p[2].strip(),"status":"识别成功"}
                    else: st.session_state.ai_res["full_text"] = res
                st.info(f"状态：{st.session_state.ai_res['status']}")
            with c2:
                f_cat = st.selectbox("类别确认", ["化验检查", "影像报告", "病理报告", "超声检查", "其他"])
                f_date = st.date_input("记录日期")
                f_sum = st.text_input("总结结论", value=st.session_state.ai_res["summary"])
                f_txt = st.text_area("原文详情", value=st.session_state.ai_res["full_text"], height=350)
                if st.button("存入档案"):
                    c.execute("INSERT INTO medical_reports (patient_id, date, category, summary, full_text) VALUES (?,?,?,?,?)", (current_patient_id, str(f_date), f_cat, f_sum, f_txt))
                    conn.commit()
                    st.success("已存入档案")

    elif menu == "💉 治疗记录":
        st.header(f"💉 治疗/就诊记录 - {selected_patient_name}")
        with st.expander("➕ 添加新纪录"):
            c1, c2 = st.columns(2)
            d, t = c1.date_input("日期"), c2.selectbox("类型", ["门诊", "住院", "手术", "理疗", "其他"])
            h, dt = st.text_input("医院"), st.text_area("详情内容")
            s = st.text_area("反馈/副作用")
            if st.button("保存记录"):
                c.execute("INSERT INTO treatment_logs (patient_id, date, treat_type, hospital, details, side_effects) VALUES (?,?,?,?,?,?)", (current_patient_id, str(d), t, h, dt, s))
                conn.commit()
                st.success("已保存")

    elif menu == "💊 服药管理":
        st.header(f"💊 服药清单 - {selected_patient_name}")
        with st.expander("➕ 登记药品"):
            c1, c2 = st.columns(2)
            n, d = c1.text_input("药名"), c2.text_input("单次剂量")
            f, s = c1.text_input("频次"), c2.radio("当前状态", ["正在服用", "已停药"], horizontal=True)
            if st.button("确认录入"):
                c.execute("INSERT INTO medication_logs (patient_id, drug_name, dosage, frequency, status) VALUES (?,?,?,?,?)", (current_patient_id, n, d, f, s))
                conn.commit()
                st.success("已登记")

    elif menu == "🔍 检索修改":
        st.header(f"🔍 综合检索与修改 - {selected_patient_name}")
        st.download_button("📥 导出该成员 Excel", data=to_excel(current_patient_id), file_name=f"{selected_patient_name}_档案.xlsx")
        tab1, tab2, tab3 = st.tabs(["医疗报告", "就诊历程", "用药历史"])
        def manage_data(table, df, fields):
            st.dataframe(df, use_container_width=True)
            if df.empty: return
            tid = st.number_input("输入 ID 进行操作", min_value=0, step=1, key=f"t_{table}")
            act = st.radio("选择动作", ["修改", "删除"], horizontal=True, key=f"a_{table}")
            if act == "修改":
                row = df[df['id'] == tid]
                if not row.empty:
                    with st.form(f"f_{table}_{tid}"):
                        upd = {f: st.text_area(f, value=str(row.iloc[0][f])) for f in fields}
                        if st.form_submit_button("保存修改"):
                            sql = f"UPDATE {table} SET " + ", ".join([f"{f}=?" for f in fields]) + " WHERE id=?"
                            c.execute(sql, list(upd.values()) + [tid])
                            conn.commit()
                            st.rerun()
            else:
                if st.button(f"确认永久删除 {tid}", key=f"d_{table}"):
                    c.execute(f"DELETE FROM {table} WHERE id=?", (tid,))
                    conn.commit()
                    st.rerun()
        with tab1: manage_data("medical_reports", pd.read_sql_query(f"SELECT id, date, category, summary, full_text FROM medical_reports WHERE patient_id={current_patient_id}", conn), ["date", "category", "summary", "full_text"])
        with tab2: manage_data("treatment_logs", pd.read_sql_query(f"SELECT id, date, treat_type, hospital, details FROM treatment_logs WHERE patient_id={current_patient_id}", conn), ["date", "treat_type", "hospital", "details"])
        with tab3: manage_data("medication_logs", pd.read_sql_query(f"SELECT id, drug_name, dosage, frequency, status FROM medication_logs WHERE patient_id={current_patient_id}", conn), ["drug_name", "dosage", "frequency", "status"])

    elif menu == "📊 健康分析看板":
        st.header(f"📊 {selected_patient_name} 健康多维分析")
        df_rep = pd.read_sql_query(f"SELECT date, category, summary FROM medical_reports WHERE patient_id={current_patient_id}", conn)
        df_tre = pd.read_sql_query(f"SELECT date, treat_type FROM treatment_logs WHERE patient_id={current_patient_id}", conn)
        df_med = pd.read_sql_query(f"SELECT drug_name, status FROM medication_logs WHERE patient_id={current_patient_id}", conn)

        if df_tre.empty and df_med.empty and df_rep.empty:
            st.info("数据量不足以生成分析。")
        else:
            c1, c2, c3 = st.columns(3)
            now_meds = df_med[df_med['status']=='正在服用']['drug_name'].tolist()
            c1.metric("累计就诊", len(df_tre))
            c2.metric("当前服药", len(now_meds))
            c3.metric("档案总数", len(df_rep) + len(df_tre))

            st.divider()
            g1, g2 = st.columns(2)
            with g1:
                st.subheader("💊 药物图谱概览")
                if not df_med.empty:
                    df_med['count'] = 1
                    fig_med = px.treemap(df_med, path=['status', 'drug_name'], values='count', color='status', color_discrete_map={'(?)':'#DDDDDD', '正在服用':'#2ECC71', '已停药':'#E74C3C'})
                    st.plotly_chart(fig_med, use_container_width=True)
            with g2:
                st.subheader("🏥 就诊类型分布")
                if not df_tre.empty:
                    tre_data = df_tre['treat_type'].value_counts().reset_index()
                    tre_data.columns = ['type_name', 'count_val']
                    fig_tre = px.bar(tre_data, x='type_name', y='count_val', color='type_name', text_auto=True)
                    st.plotly_chart(fig_tre, use_container_width=True)

            st.divider()
            st.subheader("🤖 AI 健康深度复盘")
            if st.button("🧬 生成智能报告"):
                rep_text = "、".join(df_rep['summary'].astype(str).tolist()) if not df_rep.empty else "无"
                tre_text = "、".join(df_tre['treat_type'].astype(str).tolist()) if not df_tre.empty else "无"
                med_text = "、".join(now_meds) if now_meds else "无"
                c.execute(f"SELECT disease, allergies FROM patient_info WHERE id={current_patient_id}")
                base = c.fetchone()
                prompt = f"分析病人【{selected_patient_name}】: 病史{base[0]}, 过敏史{base[1]}, 用药{med_text}, 报告汇总{rep_text}, 记录{tre_text}。评估稳定性及复查建议。"
                with st.spinner("AI 正在工作..."):
                    try:
                        res = client.chat.completions.create(model="glm-4", messages=[{"role":"user","content":prompt}])
                        st.write(res.choices[0].message.content)
                    except Exception as e: st.error(f"分析失败：{e}")