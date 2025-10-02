import streamlit as st
import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder
import base64
import pickle
from IPython.display import display
import datetime
import io
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# --- ฟังก์ชันแปลงภาพเป็น base64 สำหรับใช้ใน CSS ---
def img_to_base64(img_file_path):
    with open(img_file_path, "rb") as f:
        encoded_img = base64.b64encode(f.read()).decode()
    return encoded_img

# --- ตั้งค่า ---
st.set_page_config(page_title="Potential CUI Locations", layout="wide", # Use the full page width >> centered
    initial_sidebar_state="expanded",)

# --- โหลดรูปภาพและแปลงเป็น base64 ---
logo_path = "Logo.png"
header_path = "Header SCT.png"
logo_base64 = img_to_base64(logo_path)
bg_base64 = img_to_base64(header_path)

# --- CSS และส่วนหัวเว็บ ---
st.markdown(f"""
    <style>
    .banner {{
        background-image: url("data:image/png;base64,{bg_base64}");
        background-size: cover;
        padding: 40px 40px 20px 40px;
        border-radius: 0px 0px 20px 20px;
        margin-bottom: 30px;
    }}
    .logo {{
        margin-bottom: 10px;
    }}
    .headline {{
        font-size: 28px;
        color: white;
        font-weight: bold;
        margin-bottom: 0px;
    }}
    .subheadline {{
        font-size: 20px;
        color: white;
        margin-bottom: 18px;
        margin-top: 5px;
        font-weight: 300;
    }}
    .desc {{
        color: #d0e6fa;
        font-size: 15px;
        margin-top: 0;
    }}
    </style>

    <div class='banner'>
        <div class='logo'>
            <img src="data:image/png;base64,{logo_base64}" height="48">
        </div>
        <div class='headline'>
            <span style='color:#222;font-weight:bold;'>POTENTIAL&nbsp;|&nbsp;</span>
            <span style='color:#fff;'>CUI Locations</span>
        </div>
        <div class='subheadline'>Smart CUI Troubleshooting Project</div>
        <div class='desc'>
            Potential CUI Locations can be predicted using AI technology. 
            By leveraging parameters that influence CUI as input, a machine learning model can forecast potential area of CUI. 
            Users simply import the collected data into the model, and the system predicts the likely locations of CUI. 
            This approach enables accurate and efficient assessments, enhancing maintenance planning and prioritization.    
        </div>
    </div>
""", unsafe_allow_html=True)

# --- ข้อความแนะนำการอัปโหลด ---
st.markdown(
    "<div style='text-align:center; font-size:20px; color:#3498db; margin-bottom:15px;'>📤 Upload your Excel file to predict potential CUI locations.</div>",
    unsafe_allow_html=True
)

# --- โหลดโมเดลจาก .pkl ---
@st.cache_resource
def load_model_from_pickle():
    with open("model_and_encoders_latest.pkl", "rb") as f:
        clf, encoders, target_encoder, feature_columns = pickle.load(f)
    return clf, encoders, target_encoder, feature_columns

clf, encoders, target_encoder, feature_columns = load_model_from_pickle()

# --- 

uploaded_file = st.file_uploader("Upload Excel (.xlsx) ตาม Master Template", type=["xlsx"])


if uploaded_file is not None:
    file_name = uploaded_file.name  # <-- เก็บชื่อไฟล์เดิม (รวม .xlsx)
#st.write(f"ชื่อไฟล์ที่อัปโหลด: {file_name}")

# ตัวอย่าง ตัดเอาแค่ชื่อไฟล์ (ไม่เอานามสกุล)
    ori_file_name = file_name.rsplit('.', 1)[0]
#st.write(f"ชื่อไฟล์ (ไม่เอานามสกุล): {file_stem}")
    try:
        # อ่านไฟล์: แถวแรกเป็น header (Excel row 1)
        df_raw = pd.read_excel(uploaded_file, header=0)
        
        # Replace Double Space with Space, Trim, Drop null row
        df_raw = pd.read_excel(uploaded_file, header=0)
        df_raw = df_raw.replace(r'\s+', ' ', regex=True).apply(lambda x: x.str.strip() if x.dtype == "object" else x).dropna(how='all')
        #df_raw = pd.read_excel(uploaded_file, header=0).replace('  ', ' ', regex=False).apply(lambda x: x.str.strip() if x.dtype == "object" else x).dropna(how='all')

        # ใช้ตั้งแต่แถว 2 เป็นต้นไป (ใน pandas index 0 = Excel row 2)
        manual_input = df_raw.copy().reset_index(drop=True)
        excel_data = df_raw.iloc[:, 0:10].copy().reset_index(drop=True)  # columns 0-9 (A-J)
        #st.write("📄 ข้อมูลที่นำมาคำนวณ (ตั้งแต่แถวที่ 2 เป็นต้นไป):")
        #st.dataframe(manual_input)
        
        # เลือกเฉพาะคอลัมน์ฟีเจอร์ตามที่โมเดลต้องการ
        manual_input = manual_input[feature_columns].copy()
        manual_input_drop = manual_input[feature_columns].copy()

        unknown_details = []  # เก็บ detailed information
        # ตรวจ unknown values และเข้ารหัส
        unknown_warning = []        
        undata = pd.DataFrame()
        rows_with_unknown = []
        unknown_details_grouped = pd.DataFrame()


        import pandas as pd

# สร้าง list เพื่อเก็บข้อมูลสำหรับ DataFrame
        unknown_data = []

        for col in manual_input.columns:
            ##print('col>>>',col)

            if col in encoders:
                test_values = manual_input[col].astype(str)
                known_values = set(encoders[col].classes_.astype(str))
                unknown_values = set(test_values.unique()) - known_values
                

                try:
                    manual_input[col] = encoders[col].transform(test_values)
                    ##print(manual_input[col])                           
                except:
                    pass
                
                if unknown_values:
                    value_index_mapping = {}
                    all_col_indices = []
                    
                    
                    for unknown_val in unknown_values:
                        val_mask = manual_input[col].astype(str) == unknown_val
                        val_indices = manual_input[val_mask].index.tolist()
                        value_index_mapping[unknown_val] = val_indices
                        all_col_indices.extend(val_indices)
                        
                        # เพิ่มแต่ละ unknown_value เป็นแถวแยกใน DataFrame
                        for idx in val_indices:
                            unknown_data.append({
                                'colname': col,
                                'unknown_values': unknown_val,
                                'colindex': idx
                            })
                    
                    rows_with_unknown.extend(all_col_indices)

        # สร้าง DataFrame จากข้อมูลที่รวบรวม
        unknown_details = pd.DataFrame(unknown_data)
        



        # ลบแถวที่ซ้ำกัน (unique)
        if not unknown_details.empty:
            unknown_details = unknown_details.drop_duplicates().reset_index(drop=True)

        #print("unknown_details DataFrame:")
        #print(unknown_details)

                

        # เก็บและลบ rows ที่มี unknown values
        if rows_with_unknown:
            unique_rows = list(set(rows_with_unknown))  # ลบ duplicate index
            undata = manual_input.loc[unique_rows].copy()

            # เพิ่ม 2 columns ใหม่
            undata['Prediction'] = "Unknown"
            undata['Unknown Parameter'] = ", ".join(unknown_warning)

            #print("before if manual_input_drop",manual_input_drop)
            manual_input_drop = manual_input_drop.drop(index=unique_rows).reset_index(drop=False)
            ##print ('manual_input_drop_1  :', manual_input_drop)
        

        if not manual_input_drop.empty:  # เปลี่ยนจาก if manual_input_drop:
            ##print("manual_input_drop", manual_input_drop)
            #manual_input = manual_input_drop[feature_columns].copy()
            #print("if manual_input",manual_input)
            manual_input = manual_input_drop[feature_columns].copy()
           

            ##print("manual_input1", manual_input)

            for col in manual_input.columns:
                if col in encoders:
                    test_values = manual_input[col].astype(str)
                    manual_input[col] = encoders[col].transform(test_values) 
            #print("if manual_input_drop",manual_input_drop)
        else:
            for col in manual_input.columns:
                if col in encoders:
                    test_values = manual_input[col].astype(str)
                    manual_input[col] = encoders[col].transform(test_values) 
            
        
            
        
       
        y_pred = clf.predict(manual_input)  
        ##print (y_pred)     
        predicted_label = target_encoder.inverse_transform(y_pred)

        # จัดผลลัพธ์กลับไปแสดง (ใส่คอลัมน์ Prediction + เลขแถว Excel)
        result_df = manual_input_drop.copy()
        #print('##########result_df before', result_df)
        result_df.insert(0, "ExcelRow", result_df.index + 2)  # index 0 = Excel row 2
        result_df["Prediction"] = predicted_label
        



        result_df["Unknown Parameter"] = ""
        from collections import defaultdict
        index2txt = defaultdict(list)

        if not unknown_details.empty:
        # วิธีที่ง่ายที่สุด
            for _, row in unknown_details.iterrows():
                col = row['colname']
                val = row['unknown_values'] 
                idx = row['colindex']
                
                pretty_col = col.strip().replace(' group', '').replace('.', '.')
                index2txt[idx].append(f"{pretty_col} : {val}")
        

        # for detail in unknown_details:
        #     col = detail['col']
        #     vmap = detail['value_index_mapping']
        #     for val, index_list in vmap.items():
        #         for idx in index_list:
        #             pretty_col = col.strip().replace(' group','').replace('.', '.')
        #             index2txt[idx].append(f"{pretty_col} : {val}")

            print("unknown_details DataFrame:")
            print(unknown_details)
            print("\nData types:")
            print(unknown_details.dtypes)

            unknown_details["colval_str"] = unknown_details["colname"] + ": " + unknown_details["unknown_values"]

# group by index แล้วรวมข้อความแต่ละกลุ่ม
            unknown_details_grouped = (
                unknown_details
                .groupby('colindex')["colval_str"]
                .apply(lambda x: ', '.join(x))
                .reset_index()
                .rename(columns={"colval_str": "colname_unknown_values"})
            )

            print(unknown_details_grouped)


        print("rows_with_unknown", rows_with_unknown)

        # step 2: ใส่ค่าใน result_df ใหม่ ตาม index
        def value_for_unknown_parameter(idx):
            return ', '.join(index2txt[idx]) if idx in index2txt else "" #None
        
        print("result_df",result_df)
        

# (ถ้าอยากให้ค่า default สำหรับแถวที่ไม่มี error ใช้ fillna เพิ่ม)
# result_df["Unknown Parameter"].fillna("-", inplace=True)


        #print('##########result_df after', result_df)

        

        st.write("ข้อมูลจาก Excel Template (Column Name: 1-8) และ ผลการทำนาย (Column Name: Predictive)")

       
        
        # 1. เพิ่ม columns ใน excel_data
        excel_data['Prediction'] = ""  # หรือ None, np.nan
        excel_data['Unknown Parameter'] = ""

        #result_df["Unknown Parameter"] = result_df.index.map(value_for_unknown_parameter)
        if 'result_df' in locals() and not result_df.empty:
            # ตรวจสอบว่า result_df มี column 'index' หรือไม่
            if 'index' in result_df.columns:
                # Map ข้อมูลโดยใช้ค่าใน column 'index' ของ result_df
                for idx, row in result_df.iterrows():
                    target_index = row['index']  # ค่า index ที่จะ map
                    if target_index in excel_data.index:
                        excel_data.loc[target_index, 'Prediction'] = row['Prediction']
                        if 'Unknown Parameter' in result_df.columns:
                            excel_data.loc[target_index, 'Unknown Parameter'] = row['Unknown Parameter']
                
                #print(f"Mapped {len(result_df)} rows from result_df to excel_data using 'index' column")
                #print("#####resultdfif", result_df)

            else:
                result_df['index'] = result_df.index
                if 'index' in result_df.columns:
                # Map ข้อมูลโดยใช้ค่าใน column 'index' ของ result_df
                    for idx, row in result_df.iterrows():
                        target_index = row['index']  # ค่า index ที่จะ map
                        if target_index in excel_data.index:
                            excel_data.loc[target_index, 'Prediction'] = row['Prediction']
                            if 'Unknown Parameter' in result_df.columns:
                                excel_data.loc[target_index, 'Unknown Parameter'] = row['Unknown Parameter']
                
                
        if not unknown_details_grouped.empty:
            # Map ข้อมูล Prediction และ Unknown Parameter จาก undata
            excel_data.loc[undata.index, 'Prediction'] = undata['Prediction']
            #print(excel_data.loc[unknown_details_grouped.colindex, 'Unknown Parameter'])
            #excel_data.loc[unknown_details_grouped.colindex, 'Unknown Parameter'] = unknown_details_grouped['colname_unknown_values']

            mapping_series = unknown_details_grouped.set_index('colindex')['colname_unknown_values']

            # map ค่าใหม่ไปยังคอลัมน์ Unknown Parameter ของ excel_data
            excel_data['Unknown Parameter'] = excel_data.index.map(mapping_series)
            
            #print(f"Mapped {len(undata)} rows from undata to excel_data")
        
       

        # 3. สร้าง manual_input_result
        manual_input_result = excel_data.copy()

        # สร้างซีรีส์หมายเลขจาก index + 1
        no_series = manual_input_result.index.to_series().add(1).astype(int)

        # แทรกเป็นคอลัมน์แรก (ตำแหน่ง 0)
        manual_input_result.insert(0, 'No.', no_series.values)
        #manual_input_result['Unknown Parameter'] = manual_input_result['Unknown Parameter'].replace(["nan", "NaN", "Nan", "NAN"], "")
        #manual_input_result['Unknown Parameter'] = manual_input_result['Unknown Parameter'].fillna("")
        # แทนทั้ง actual NaN และ string "NaN"
        manual_input_result['Unknown Parameter'] = manual_input_result['Unknown Parameter'].fillna("").replace("NaN", "")

        print("manual_input_result: ", manual_input_result)
        #print("########################################################################")
        df_display = manual_input_result.astype(str)

        styler = (
            df_display
            .style
            .set_table_styles([
                {"selector": "thead th", "props": "text-align: center;"},
                {"selector": "th.col_heading", "props": "text-align: center;"},
                {"selector": "th.col_heading.level0", "props": "text-align: center;"}
                
            ])
            .set_properties(**{"text-align": "left"})
            # ตั้งค่าสีสำหรับคอลัมน์ Prediction (พื้นหลังเขียว, ตัวอักษรขาว)
            .set_properties(
                subset=["Prediction"],
                **{
                    "background-color": "white",  # สีเขียว
                    "color": "#098F0E"
                }
            )
            # ตั้งค่าสีสำหรับคอลัมน์ Unknown Parameter (พื้นหลังแดงอ่อน, ตัวอักษรดำ)
            .set_properties(
                subset=["Unknown Parameter"],
                **{
                    "background-color": "white",  # สีแดงอ่อน
                    "color": "#8D1606"
                
                }
            )
        )
        
        st.dataframe(styler, hide_index=True)
                
    
        ##print (result_df)
        #st.dataframe(result_df.Prediction)
        
        # สรุปจำนวนแต่ละคลาส
        st.write("📊 สรุปผลการทำนาย:")
        if 'Prediction' in manual_input_result.columns:
            counts = manual_input_result['Prediction'].value_counts().astype(str)
            df_counts = counts.rename_axis("Prediction Results").reset_index(name="Counts")
            df_counts = df_counts.astype(str)

            # ใช้ Styler เพื่อตั้งค่า alignment
            styled_df = (
                df_counts
                .style
                .set_table_styles([
                    # จัดกึ่งกลางหัวตาราง
                    {"selector": "thead th", "props": "text-align: center;"},
                    {"selector": "th.col_heading", "props": "text-align: center;"},
                    {"selector": "th.col_heading.level0", "props": "text-align: center;"}
                ])
                # จัดกึ่งกลางข้อมูลทั้งหมด
                .set_properties(**{"text-align": "center"})
                .set_properties(subset=["Prediction Results", "Counts"], **{
                    "width": "60px"
                })
            )

            # แสดงใน Streamlit
            st.dataframe(styled_df, use_container_width=False, hide_index=True)
        else:
            st.warning("Column 'Prediction' not found in manual_input_result")

    except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}")


   
    if st.button("Generate result report"):
        # 3. Export to Excel (user specifies filename/location)
        
        dtstr = datetime.datetime.now().strftime("%Y%m%d")
        #export_filename = f"{dtstr}_{ori_file_name}_results.xlsx"
        export_filename = f"{ori_file_name}_results.xlsx"

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            manual_input_result.to_excel(writer, index=False, sheet_name='Result Data')
            df_counts.to_excel(writer, index=False, sheet_name='Summary')

        st.success(f"Export file ready: {export_filename}")

        output.seek(0)

        st.download_button(
            label="📥 Download Excel file",
            data=output.getvalue(),
            file_name=export_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        
        )
       
    





