import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
import random
import io
import zipfile
import os  # <--- 关键修复：一定要加上这一行

# 页面配置
st.set_page_config(page_title="RandomSCC 自动化工具", layout="wide")
st.title("📦 RandomSCC 数据填充工具")
st.markdown("上传 `containerinformation.xlsx` 文件，点击开始即可生成处理后的文件包。")

# 1. 配置检查：确保模板和数据源已在 GitHub 仓库中
st.sidebar.header("配置检查")
template_file = 'icstemplate.xlsx'
realsc_file = 'realsc.xlsx'

# 检查文件是否在当前目录下
if not os.path.exists(template_file) or not os.path.exists(realsc_file):
    st.error(f"⚠️ 仓库中缺少必要文件！请确保 {template_file} 和 {realsc_file} 已上传到 GitHub 根目录。")
    st.stop()
else:
    st.sidebar.success("✅ 静态资源（模板与数据源）已就绪")

# 2. 上传文件界面
uploaded_file = st.file_uploader("第一步：请选择 containerinformation.xlsx 文件", type=["xlsx"])

if uploaded_file:
    st.info("文件已上传，准备就绪。")
    if st.button("🚀 第二步：开始批量生成并打包"):
        try:
            # 读取 realsc
            df_realsc = pd.read_excel(realsc_file, header=None)
            df_realsc.dropna(how='all', inplace=True)
            realsc_data = df_realsc.values.tolist()
            # 每4行一组
            realsc_groups = [realsc_data[i:i+4] for i in range(0, len(realsc_data), 4) if len(realsc_data[i:i+4]) == 4]

            # 读取上传的 containerinformation
            df_ci = pd.read_excel(uploaded_file)
            df_ci['单号'] = df_ci['单号'].ffill()
            grouped = df_ci.groupby('单号')

            # 准备内存中的 ZIP 压缩包
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                # 样式定义
                custom_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                custom_font = Font(name='微软雅黑', size=11)

                progress_bar = st.progress(0)
                status_text = st.empty()
                total_groups = len(grouped)

                for idx, (order_no, group) in enumerate(grouped):
                    status_text.text(f"正在处理单号: {order_no} ({idx+1}/{total_groups})")
                    
                    # 加载模板
                    wb = openpyxl.load_workbook(template_file)
                    ws = wb.active 
                    
                    # 填充逻辑
                    ws['B5'] = str(order_no).upper()
                    f130_val = ws['F130'].value

                    if len(group) == 1:
                        row_data = group.iloc[0]
                        ws['B8'], ws['B9'] = row_data['件数'], row_data['重量(KGS)']
                        ws['A130'], ws['B130'] = str(row_data['HS CODE']).upper(), str(row_data['品名']).upper()
                        ws['C130'], ws['E130'] = row_data['件数'], row_data['重量(KGS)']
                    else:
                        ws['B8'], ws['B9'] = group['件数'].sum(), group['重量(KGS)'].sum()
                        for i, (_, row_data) in enumerate(group.iterrows()):
                            curr_row = 130 + i
                            ws.cell(row=curr_row, column=1, value=str(row_data['HS CODE']).upper())
                            ws.cell(row=curr_row, column=2, value=str(row_data['品名']).upper())
                            ws.cell(row=curr_row, column=3, value=row_data['件数'])
                            ws.cell(row=curr_row, column=4, value="PK-PACKAGE")
                            ws.cell(row=curr_row, column=5, value=row_data['重量(KGS)'])
                            ws.cell(row=curr_row, column=6, value=f130_val)

                    # 随机 realsc 填充
                    if realsc_groups:
                        chosen_sc = random.choice(realsc_groups)
                        target_rows = [14, 15, 18, 19]
                        for r_idx, data_row in enumerate(chosen_sc):
                            for c_idx, value in enumerate(data_row, start=3):
                                if c_idx <= 8:
                                    final_val = str(value).upper() if isinstance(value, str) else value
                                    ws.cell(row=target_rows[r_idx], column=c_idx, value=final_val)

                    # 全表格式化
                    for row in ws.iter_rows():
                        for cell in row:
                            cell.alignment = custom_alignment
                            cell.font = custom_font
                            if isinstance(cell.value, str):
                                cell.value = cell.value.upper()

                    # 将文件保存到内存流
                    file_stream = io.BytesIO()
                    wb.save(file_stream)
                    zip_file.writestr(f"{order_no}.xlsx", file_stream.getvalue())
                    
                    # 更新进度
                    progress_bar.progress((idx + 1) / total_groups)

            status_text.text("✅ 所有文件处理完成！")
            
            # 提供 ZIP 下载
            st.download_button(
                label="📥 点击下载生成的文件包 (ZIP格式)",
                data=zip_buffer.getvalue(),
                file_name="processed_files.zip",
                mime="application/x-zip-compressed"
            )
        except Exception as e:
            st.error(f"处理过程中出错: {e}")
