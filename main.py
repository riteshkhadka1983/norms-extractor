import io
import re
import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException
from sqlalchemy import create_engine, text

app = FastAPI(title="Water Supply & DOR Rate Analysis Norms Inserter 2081")

# Database Connection URL (Update credentials as needed)
DATABASE_URL = "mysql+pymysql://root:@localhost:3306/mahuwa_normsentry"
engine = create_engine(DATABASE_URL)


def combine_serialization(col0, col2):
    """
    Combines row 1 (col0) and row 3 (col2) serialization patterns if alphanumeric serials are present.
    Example: col0='A', col2='I' => 'A.I'
    """
    col0_str = str(col0).strip() if pd.notna(col0) else ""
    col2_str = str(col2).strip() if pd.notna(col2) else ""

    clean_col0 = re.sub(r'[\(\)]', '', col0_str)
    clean_col2 = re.sub(r'[\(\)]', '', col2_str)

    if clean_col0 and clean_col2:
        return f"{clean_col0}.{clean_col2}"
    elif clean_col0:
        return clean_col0
    elif clean_col2:
        return clean_col2
    return ""


def extract_section_info(df):
    """Reads Section Number and Name around rows 1-5 of the Excel sheet."""
    section_no = ""
    section_name = ""
    for r in range(min(6, len(df))):
        for c in range(min(5, len(df.columns))):
            val = str(df.iloc[r, c])
            if "SECTION" in val.upper():
                match = re.search(r'SECTION\s*(\d+)[\:\s\-\_]*(.*)', val, re.IGNORECASE)
                if match:
                    section_no = match.group(1).strip()
                    section_name = match.group(2).strip()
                    return section_no, section_name
    return section_no, section_name


@app.post("/upload-norms/")
def process_norms_excel(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload an Excel sheet.")

    try:
        # Read file contents into BytesIO stream
        contents = file.file.read()
        buffer = io.BytesIO(contents)
        xls = pd.ExcelFile(buffer, engine='openpyxl')
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not read Excel file: {str(e)}")

    total_norms_inserted = 0
    total_details_inserted = 0
    total_resources_inserted = 0

    with engine.begin() as conn:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            section_no, section_name = extract_section_info(df)

            current_norms_detail_id = None
            current_work_type_detail_id = None

            for row_idx in range(len(df)):
                # Ignore header rows (first 6 rows usually contain titles & table headers)
                if row_idx < 6:
                    continue

                col0 = df.iloc[row_idx, 0] if len(df.columns) > 0 else None
                col1 = df.iloc[row_idx, 1] if len(df.columns) > 1 else None
                col2 = df.iloc[row_idx, 2] if len(df.columns) > 2 else None
                col3 = df.iloc[row_idx, 3] if len(df.columns) > 3 else None

                is_col0_numeric = pd.notna(col0) and str(col0).replace('.', '').isdigit()
                is_col1_numeric = pd.notna(col1) and str(col1).replace('.', '').replace(',', '').strip().isdigit()

                # =========================================================
                # CONDITION 1: Main Norm Item Header (e.g. 1.1 / 111)
                # =========================================================
                if is_col0_numeric and is_col1_numeric:
                    norms_type = "Water Supply"
                    norms_no = str(col0).strip()
                    reference_no = str(col1).strip()
                    raw_title = str(col2).strip() if pd.notna(col2) else ""
                    norms_date = "2081"

                    remark = "-"
                    if "Remark:" in raw_title or "Remarks:" in raw_title:
                        parts = re.split(r'Remarks?:', raw_title, flags=re.IGNORECASE)
                        norms_title_final = parts[0].strip()
                        remark = parts[1].strip()
                    else:
                        norms_title_final = raw_title

                    # Determine next position in norms_details
                    pos_query = text("SELECT COALESCE(MAX(CAST(position AS UNSIGNED)), 0) + 1 FROM norms_details")
                    next_pos = conn.execute(pos_query).scalar()

                    insert_norms_sql = text("""
                        INSERT INTO norms_details 
                        (norms_type, section_no, section_name, norms_no, norms_date, reference_no, norms_title, norms_desc, remark, position)
                        VALUES (:norms_type, :section_no, :section_name, :norms_no, :norms_date, :reference_no, :norms_title, :norms_desc, :remark, :position)
                    """)
                    
                    res = conn.execute(insert_norms_sql, {
                        "norms_type": norms_type,
                        "section_no": section_no,
                        "section_name": section_name,
                        "norms_no": norms_no,
                        "norms_date": norms_date,
                        "reference_no": reference_no,
                        "norms_title": norms_title_final,
                        "norms_desc": norms_title_final,
                        "remark": remark,
                        "position": str(next_pos)
                    })
                    
                    current_norms_detail_id = res.lastrowid
                    current_work_type_detail_id = None
                    total_norms_inserted += 1

                # =========================================================
                # CONDITION 2: Sub-Work Type Detail or Resource Row
                # =========================================================
                elif current_norms_detail_id is not None:
                    sn_pattern = combine_serialization(col0, col2)
                    detail_text = str(col2).strip() if pd.notna(col2) else ""
                    unit_val = str(col3).strip() if pd.notna(col3) else ""

                    # Create new norms_work_type_detail if serial or work description exists
                    if detail_text or sn_pattern or pd.notna(col0):
                        pos_wt_query = text("""
                            SELECT COALESCE(MAX(CAST(position AS UNSIGNED)), 0) + 1 
                            FROM norms_work_type_details 
                            WHERE norms_detail_id = :norms_detail_id
                        """)
                        wt_pos = conn.execute(pos_wt_query, {"norms_detail_id": current_norms_detail_id}).scalar()

                        insert_wt_sql = text("""
                            INSERT INTO norms_work_type_details
                            (norms_detail_id, sn, header, detail, unit_detail, unit_quantity, unit, position, remark)
                            VALUES (:norms_detail_id, :sn, :header, :detail, :unit_detail, :unit_quantity, :unit, :position, :remark)
                        """)

                        res_wt = conn.execute(insert_wt_sql, {
                            "norms_detail_id": current_norms_detail_id,
                            "sn": sn_pattern,
                            "header": detail_text,
                            "detail": detail_text,
                            "unit_detail": f"Unit = {unit_val}" if unit_val else "",
                            "unit_quantity": "1",
                            "unit": unit_val,
                            "position": str(wt_pos),
                            "remark": "-"
                        })
                        current_work_type_detail_id = res_wt.lastrowid
                        total_details_inserted += 1

                    # =========================================================
                    # EXTRACT RESOURCES (Labour, Material, Equipment)
                    # =========================================================
                    if current_work_type_detail_id is not None:
                        # 1. LABOUR (Columns 4=Class, 5=Unit, 6=Qty)
                        labour_class = df.iloc[row_idx, 4] if len(df.columns) > 4 else None
                        labour_unit = df.iloc[row_idx, 5] if len(df.columns) > 5 else None
                        labour_qty = df.iloc[row_idx, 6] if len(df.columns) > 6 else None

                        if pd.notna(labour_class) and pd.notna(labour_qty):
                            _insert_quantity_resource(
                                conn, 
                                norms_detail_id=current_norms_detail_id,
                                work_type_detail_id=current_work_type_detail_id,
                                category="Basic_Rate_Labour",
                                resource_type="Labour",
                                item_name=str(labour_class).strip(),
                                unit=str(labour_unit).strip() if pd.notna(labour_unit) else "m-day",
                                quantity=float(labour_qty)
                            )
                            total_resources_inserted += 1

                        # 2. CONSTRUCTION MATERIAL (Columns 7=Type, 8=Unit, 9=Qty)
                        mat_type = df.iloc[row_idx, 7] if len(df.columns) > 7 else None
                        mat_unit = df.iloc[row_idx, 8] if len(df.columns) > 8 else None
                        mat_qty = df.iloc[row_idx, 9] if len(df.columns) > 9 else None

                        if pd.notna(mat_type) and pd.notna(mat_qty):
                            _insert_quantity_resource(
                                conn, 
                                norms_detail_id=current_norms_detail_id,
                                work_type_detail_id=current_work_type_detail_id,
                                category="Basic_Rate_Material",
                                resource_type="Material",
                                item_name=str(mat_type).strip(),
                                unit=str(mat_unit).strip() if pd.notna(mat_unit) else "",
                                quantity=float(mat_qty)
                            )
                            total_resources_inserted += 1

                        # 3. MACHINERY / EQUIPMENT (Columns 10=Type, 11=Unit, 12=Qty)
                        eq_type = df.iloc[row_idx, 10] if len(df.columns) > 10 else None
                        eq_unit = df.iloc[row_idx, 11] if len(df.columns) > 11 else None
                        eq_qty = df.iloc[row_idx, 12] if len(df.columns) > 12 else None

                        if pd.notna(eq_type) and pd.notna(eq_qty):
                            _insert_quantity_resource(
                                conn, 
                                norms_detail_id=current_norms_detail_id,
                                work_type_detail_id=current_work_type_detail_id,
                                category="Equipment_Usage",
                                resource_type="Equipment",
                                item_name=str(eq_type).strip(),
                                unit=str(eq_unit).strip() if pd.notna(eq_unit) else "hr",
                                quantity=float(eq_qty)
                            )
                            total_resources_inserted += 1

    return {
        "status": "success",
        "message": "Norms and resource analysis extracted successfully.",
        "norms_inserted": total_norms_inserted,
        "work_type_details_inserted": total_details_inserted,
        "resources_inserted": total_resources_inserted
    }


def _insert_quantity_resource(conn, norms_detail_id, work_type_detail_id, category, resource_type, item_name, unit, quantity):
    """
    Helper function to insert resource quantities into `norms_quantities` 
    and linking table records `norms_resources_types`.
    """
    # 1. Insert into norms_resources_types
    res_type_sql = text("""
        INSERT INTO norms_resources_types (norms_work_type_detail_id, header, resource_type, resource_type_sn)
        VALUES (:wt_id, :header, :res_type, 'a')
    """)
    res_type_res = conn.execute(res_type_sql, {
        "wt_id": work_type_detail_id,
        "header": item_name,
        "res_type": resource_type
    })
    res_type_id = res_type_res.lastrowid

    # 2. Insert into norms_quantities
    quant_sql = text("""
        INSERT INTO norms_quantities (norms_detail_id, norms_work_type_detail_id, category_id, category, quantity)
        VALUES (:norms_detail_id, :wt_id, :category_id, :category, :quantity)
    """)
    conn.execute(quant_sql, {
        "norms_detail_id": norms_detail_id,
        "wt_id": work_type_detail_id,
        "category_id": res_type_id,
        "category": category,
        "quantity": quantity
    })