import streamlit as st # ใช้สร้าง web UI แบบ interactive เช่น ปุ่ม อัปโหลดไฟล์ แสดงผลลัพธ์โมเดล ฯลฯ
import pandas as pd # ช่วยอ่านไฟล์ข้อมูล (เช่น Excel, CSV) และจัดการข้อมูลในรูปแบบตาราง DataFrame
import base64 # ใช้ในการเข้ารหัสไฟล์ (เช่น รูปภาพ) ให้อยู่ในรูปข้อความ เพื่อสามารถฝังใน HTML/CSS ได้
import pickle # ใช้โหลดโมเดล Machine Learning ที่บันทึกไว้ เช่น model.pkl
import io #	ใช้จัดการข้อมูลแบบ “stream” เช่น การสร้างไฟล์ Excel ชั่วคราวใน memory เพื่อให้ดาวน์โหลด
import numpy as np #ใช้คำนวณค่าตัวเลขขนาดใหญ่ เช่น array, matrix อย่างมีประสิทธิภาพ

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
    "<div style='text-align:center; font-size:20px; color:#3498db; margin-bottom:15px;'>" \
    "📤 Upload your Excel file to predict potential CUI locations.</div>",
    unsafe_allow_html=True
)

# --- โหลดโมเดลจาก .pkl ---
@st.cache_resource
def load_model_from_pickle():
    with open("20251003_model_and_encoders_newpercent.pkl", "rb") as f:
        clf, encoders, target_encoder, feature_columns, X_train, y_train = pickle.load(f)
    return clf, encoders, target_encoder, feature_columns, X_train, y_train 

clf, encoders, target_encoder, feature_columns, X_train, y_train  = load_model_from_pickle()


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
       # 1) แทนที่ whitespace หลายตัว (รวม space, tab, newline) ให้เป็น space เดียว
        df_raw = df_raw.replace(r'\s+', ' ', regex=True)

        # 2) Trim เฉพาะคอลัมน์ที่เป็นข้อความจริงๆ และเฉพาะค่าที่เป็น string
        text_cols = df_raw.select_dtypes(include=['object']).columns
        for c in text_cols:
            # ใช้ .where เพื่อเลี่ยงการแปลง NaN และคงค่าที่ไม่ใช่ string
            df_raw[c] = df_raw[c].where(df_raw[c].isna(), df_raw[c].map(lambda v: v.strip() if isinstance(v, str) else v))

        # 3) ลบแถวที่ทุกคอลัมน์เป็น NaN
        df_raw = df_raw.dropna(how='all')  

          
        #Copy โดยเลือก Data ทั้งหมด 
        manual_input= df_raw.copy().reset_index(drop=True)
        
        excel_data = df_raw.copy().reset_index(drop=True)
        
        # เลือกเฉพาะคอลัมน์ฟีเจอร์ตามที่โมเดลต้องการ Column 1.Substrate to 8.Water Get In from Jacket Damage   
        manual_input = manual_input[feature_columns].copy()
        manual_input_drop = manual_input[feature_columns].copy()

        unknown_details = []  # เก็บ detailed information
        unknown_data = []
        # ตรวจ unknown values และเข้ารหัส
        unknown_warning = []        
        undata = pd.DataFrame()
        rows_with_unknown = []
        unknown_details_grouped = pd.DataFrame()               
        

        for col in manual_input.columns:      

            if col in encoders:
                test_values = manual_input[col].astype(str)
                known_values = set(encoders[col].classes_.astype(str))
                unknown_values = set(test_values.unique()) - known_values

                try:
                    manual_input[col] = encoders[col].transform(test_values)
                    print("manual_input[col]",manual_input[col])                           
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
        
        # เก็บและลบ rows ที่มี unknown values
        if rows_with_unknown:
            unique_rows = list(set(rows_with_unknown))  # ลบ duplicate index
            undata = manual_input.loc[unique_rows].copy()
            print("########################################################")
                            
            print("unique_rows : ", unique_rows)
            print("undata : ", undata)

            # เพิ่ม 2 columns ใหม่
            undata['Prediction'] = "Unknown"
            print("undata Prediction: ", undata)

            undata['Unknown Parameter'] = ", ".join(unknown_warning)         
            print("undata Unknown Parameter: ", undata)  

            manual_input_drop = manual_input_drop.drop(index=unique_rows).reset_index(drop=False)
            print("manual_input_dropr: ", manual_input_drop) 
            
        if not manual_input_drop.empty:            
            manual_input = manual_input_drop[feature_columns].copy()
           
            for col in manual_input.columns:
                if col in encoders:
                    test_values = manual_input[col].astype(str)
                    manual_input[col] = encoders[col].transform(test_values) 
            
        else:
            for col in manual_input.columns:
                if col in encoders:
                    test_values = manual_input[col].astype(str)
                    manual_input[col] = encoders[col].transform(test_values) 
            
        def predict_with_confidence_exact(clf, manual_input, X_train):
        #"""Version 2: Check for exact match (all features must be identical)"""
            y_pred = clf.predict(manual_input)
            y_proba = clf.predict_proba(manual_input)
            
            confidences = []
            
            # Convert to numpy arrays
            if not isinstance(X_train, np.ndarray):
                X_train_array = X_train.values
            else:
                X_train_array = X_train
                
            if not isinstance(manual_input, np.ndarray):
                manual_array = manual_input.values
            else:
                manual_array = manual_input
            
            for i in range(len(manual_array)):
                # Check if exact match exists
                matches = np.all(X_train_array == manual_array[i], axis=1)
                
                if np.any(matches):
                    # Found exact match in training data
                    confidence = 100.0
                else:
                    # Use Random Forest probability
                    confidence = np.max(y_proba[i]) * 100
                
                confidences.append(confidence)
            
            return y_pred, np.array(confidences)
    
        y_pred, confidences = predict_with_confidence_exact(clf, manual_input, X_train)
       
        predicted_label = target_encoder.inverse_transform(y_pred)

        # จัดผลลัพธ์กลับไปแสดง (ใส่คอลัมน์ Prediction + เลขแถว Excel)
        result_df = manual_input_drop.copy()
        
        result_df.insert(0, "ExcelRow", result_df.index + 2)  # index 0 = Excel row 2
        result_df["Prediction"] = predicted_label
        result_df["confidences"] = confidences
        
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

            unknown_details["colval_str"] = unknown_details["colname"] + ": " + unknown_details["unknown_values"]
            
            # group by index แล้วรวมข้อความแต่ละกลุ่ม
            unknown_details_grouped = (
                unknown_details
                .groupby('colindex')["colval_str"]
                .apply(lambda x: ', '.join(x))
                .reset_index()
                .rename(columns={"colval_str": "colname_unknown_values"})
            )

            
        # step 2: ใส่ค่าใน result_df ใหม่ ตาม index
        def value_for_unknown_parameter(idx):
            return ', '.join(index2txt[idx]) if idx in index2txt else "" #None
                
        st.write("ข้อมูลจาก Excel Template (Column Name: 1-8) และ ผลการทำนาย (Column Name: Predictive)")
            
        # 1. เพิ่ม columns ใน excel_data
        excel_data['Prediction'] = ""
        excel_data['Unknown Parameter'] = ""  # หรือ None, np.nan
        excel_data['confidences'] = ""

        #result_df["Unknown Parameter"] = result_df.index.map(value_for_unknown_parameter)
        if 'result_df' in locals() and not result_df.empty:
            # ตรวจสอบว่า result_df มี column 'index' หรือไม่
            if 'index' in result_df.columns:
                # Map ข้อมูลโดยใช้ค่าใน column 'index' ของ result_df
                for idx, row in result_df.iterrows():
                    target_index = row['index']  # ค่า index ที่จะ map
                    if target_index in excel_data.index:
                        excel_data.loc[target_index, 'Prediction'] = row['Prediction']
                        excel_data.loc[target_index, 'confidences'] = row['confidences']
                        if 'Unknown Parameter' in result_df.columns:
                            excel_data.loc[target_index, 'Unknown Parameter'] = row['Unknown Parameter']
                           
            else:
                result_df['index'] = result_df.index
                if 'index' in result_df.columns:
                # Map ข้อมูลโดยใช้ค่าใน column 'index' ของ result_df
                    for idx, row in result_df.iterrows():
                        target_index = row['index']  # ค่า index ที่จะ map
                        if target_index in excel_data.index:
                            excel_data.loc[target_index, 'Prediction'] = row['Prediction']
                            excel_data.loc[target_index, 'confidences'] = row['confidences']
                            if 'Unknown Parameter' in result_df.columns:
                                excel_data.loc[target_index, 'Unknown Parameter'] = row['Unknown Parameter']
                            
        if not unknown_details_grouped.empty:
            # Map ข้อมูล Prediction และ Unknown Parameter จาก undata
            excel_data.loc[undata.index, 'Prediction'] = undata['Prediction']           
            mapping_series = unknown_details_grouped.set_index('colindex')['colname_unknown_values']

            # map ค่าใหม่ไปยังคอลัมน์ Unknown Parameter ของ excel_data
            excel_data['Unknown Parameter'] = excel_data.index.map(mapping_series)
            
        # 3. สร้าง manual_input_result
        manual_input_result = excel_data.copy()

        # สร้างซีรีส์หมายเลขจาก index + 1
        no_series = manual_input_result.index.to_series().add(1).astype(int)

        # แทรกเป็นคอลัมน์แรก (ตำแหน่ง 0)
        manual_input_result.insert(0, 'No.', no_series.values)
        # แทนทั้ง actual NaN และ string "NaN"
        manual_input_result['Unknown Parameter'] = manual_input_result['Unknown Parameter'].fillna("").replace("NaN", "")
        manual_input_result = manual_input_result.rename(columns={'confidences': '% Confidence'})
        # 1. เปลี่ยนชื่อคอลัมน์
        manual_input_result = manual_input_result.rename(columns={'confidences': '% Confidence'})

        # 1) แปลงเป็นตัวเลข
        manual_input_result['% Confidence'] = pd.to_numeric(
            manual_input_result['% Confidence'], errors='coerce'
        )
        # 2) ตัดทศนิยมแบบไม่ปัดให้เหลือ 2 ตำแหน่ง
        manual_input_result['% Confidence'] = np.floor(
            manual_input_result['% Confidence'] * 100) / 100

        # 3) จัดรูปแบบการแสดงผล:
        #    - ถ้าเป็นจำนวนเต็มพอดี -> ตัด .00 ออก
        #    - ถ้าเป็นทศนิยม -> แสดงค่าทศนิยมตามจริง (เท่าที่มีหลัง step 2)
        manual_input_result['% Confidence'] = manual_input_result['% Confidence'].apply(
            lambda x: (str(int(x)) if x == np.floor(x) else ('{:.2f}'.format(x).rstrip('0').rstrip('.')))
            if pd.notna(x) else np.nan
        )

        manual_input_result['Unknown Parameter'] = manual_input_result['Unknown Parameter'].fillna("").replace("NaN", "")
        manual_input_result['% Confidence'] = manual_input_result['% Confidence'].fillna("").replace("NaN", "")
               
        #print("########################################################################")
        df_display = manual_input_result.astype(str)

        styler = (
            df_display
            .style
            .set_table_styles([
                {"selector": "thead th", "props": "text-align: center;"},
                {"selector": "th.col_heading", "props": "text-align: center;"},
                {"selector": "th.col_heading.level0", "props": "text-align: center;"},

                # เพิ่มการ wrap text สำหรับ header
                {"selector": "th.col_heading", 
                "props": "white-space: normal; word-wrap: break-word; overflow-wrap: break-word; max-width: 150px;"},
                {"selector": "thead th", 
                "props": "white-space: normal; word-wrap: break-word; overflow-wrap: break-word;"}
            
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
    # 1. Create export filename
        #dtstr = datetime.datetime.now().strftime("%Y%m%d")
        export_filename = f"{ori_file_name}_results.xlsx"

        # 2. Create in-memory Excel file
        output = io.BytesIO()

        # 3. Write Excel with borders applied to all cells in data range
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write both DataFrames
            manual_input_result.to_excel(writer, index=False, sheet_name='Result Data')
            df_counts.to_excel(writer, index=False, sheet_name='Summary')

            workbook = writer.book
            border_fmt = workbook.add_format({'border': 1})

            # ===== Apply to "Result Data" sheet =====
            worksheet1 = writer.sheets['Result Data']
            rows1, cols1 = manual_input_result.shape

            # Apply borders to all data cells + header
            worksheet1.conditional_format(
                0, 0, rows1, cols1 - 1,
                {'type': 'no_errors', 'format': border_fmt}
            )

            # ===== Apply to "Summary" sheet =====
            worksheet2 = writer.sheets['Summary']
            rows2, cols2 = df_counts.shape

            worksheet2.conditional_format(
                0, 0, rows2, cols2 - 1,
                {'type': 'no_errors', 'format': border_fmt}
            )

        # 4. Enable downloading
        st.success(f"Export file ready: {export_filename}")
        output.seek(0)
        st.download_button(
            label="📥 Download Excel file",
            data=output.getvalue(),
            file_name=export_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
   

       
    

