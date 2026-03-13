import streamlit as st
import pandas as pd
import openpyxl
import random
import io
import zipfile

# 页面配置
st.set_page_config(page_title="RandomSCC 自动化工具", layout="wide")
st.title("📦 RandomSC 数据填充工具")
st.markdown("""
### 使用说明：
1. 上传三个 Excel 文件。
2. 程序将**保留模板原始格式**，仅进行数据填充。
3. 处理完成后，点击下载生成的 ZIP 压缩包。
""")

# --- 第一部分：文件上传界面 ---
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. 整柜数据")
    ci_file = st.file_uploader("上传 containerinformation.xlsx", type=["xlsx"], key="ci")

with col2:
    st.subheader("2. ICS 表头")
    tmpl_file = st.file_uploader("上传 icstemplate.xlsx", type=["xlsx"], key="tmpl")

with col3:
    st.subheader("3. RSC 源数据")
    sc_file = st.file_uploader("上传 realsc.xlsx", type=["xlsx"], key="sc")

# --- 第二部分：核心处理逻辑 ---
if ci_file and tmpl_file and sc_file:
    st.success("✅ 三个文件已全部上传，可以开始处理！")
    
    if st.button("🚀 开始生成并打包下载"):
        try:
            # 1. 处理 realsc 数据池 (每4行为一组)
            df_realsc = pd.read_excel(sc_file, header=None)
            df_realsc.dropna(how='all', inplace=True)
            realsc_data = df_realsc.values.tolist()
            realsc_groups = [realsc_data[i:i+4] for i in range(0, len(realsc_data), 4) if len(realsc_data[i:i+4]) == 4]

            # 2. 处理 containerinformation 并处理空行填充单号
            df_ci = pd.read_excel(ci_file)
            if '单号' not in df_ci.columns:
                st.error("错误：containerinformation 文件中未找到 '单号' 列，请检查文件！")
                st.stop()
            
            df_ci['单号'] = df_ci['单号'].ffill()
            grouped = df_ci.groupby('单号')

            # 3. 准备 ZIP 容器
            zip_buffer = io.BytesIO()
            template_bytes = tmpl_file.read()

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                progress_bar = st.progress(0)
                status_text = st.empty()
                total_groups = len(grouped)

                for idx, (order_no, group) in enumerate(grouped):
                    status_text.text(f"正在处理单号: {order_no} ({idx+1}/{total_groups})")
                    
                    # 从内存加载模板（保留原始格式）
                    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
                    ws = wb.active 
                    
                    # --- A. 填充基础信息 ---
                    ws['B5'] = order_no
                    f130_val = ws['F130'].value

                    if len(group) == 1:
                        # 单行逻辑
                        row_data = group.iloc[0]
                        ws['B8'] = row_data['件数']
                        ws['B9'] = row_data['重量(KGS)']
                        ws['A130'] = row_data['HS CODE']
                        ws['B130'] = row_data['品名']
                        ws['C130'] = row_data['件数']
                        ws['E130'] = row_data['重量(KGS)']
                    else:
                        # 多行逻辑
                        ws['B8'] = group['件数'].sum()
                        ws['B9'] = group['重量(KGS)'].sum()
                        for i, (_, row_data) in enumerate(group.iterrows()):
                            curr_row = 130 + i
                            ws.cell(row=curr_row, column=1, value=row_data['HS CODE'])
                            ws.cell(row=curr_row, column=2, value=row_data['品名'])
                            ws.cell(row=curr_row, column=3, value=row_data['件数'])
                            ws.cell(row=curr_row, column=4, value="PK-Package")
                            ws.cell(row=curr_row, column=5, value=row_data['重量(KGS)'])
                            ws.cell(row=curr_row, column=6, value=f130_val)

                    # --- B. 随机填充 realsc 数据 ---
                    if realsc_groups:
                        chosen_sc = random.choice(realsc_groups)
                        target_rows = [14, 15, 18, 19]
                        for r_offset, data_row in enumerate(chosen_sc):
                            t_row = target_rows[r_offset]
                            # 数据填入 C, D, E, F, G, H 列 (3到8列)
                            for c_offset, value in enumerate(data_row, start=3):
                                if c_offset <= 8:
                                    ws.cell(row=t_row, column=c_offset, value=value)

                    # --- C. 保存结果（无额外格式操作） ---
                    file_stream = io.BytesIO()
                    wb.save(file_stream)
                    zip_file.writestr(f"{order_no}.xlsx", file_stream.getvalue())
                    
                    progress_bar.progress((idx + 1) / total_groups)

            status_text.text("✅ 处理完成！已保持模板原始格式。")
            
            # 提供下载
            st.download_button(
                label="📥 点击下载生成的文件包 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="processed_files_no_format.zip",
                mime="application/x-zip-compressed"
            )
        except Exception as e:
            st.error(f"运行出错: {e}")
else:
    st.info("💡 请在上方上传全部三个 Excel 文件以激活处理程序。")
