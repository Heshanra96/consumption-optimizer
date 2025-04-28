
import streamlit as st
import pandas as pd
import io
from collections import Counter
from datetime import datetime

st.set_page_config(page_title="Consumption Optimizer", layout="centered")
st.title("Consumption Optimizer Web App")

def process_buy(file):
    df_buy = pd.read_excel(file, header=1)
    columns = df_buy.columns + " " + df_buy.iloc[0].values
    pattern = "Unnamed: \d+ "
    columns = [
        col.strip() if not isinstance(col, str) else col.replace(pattern, "", 1).strip()
        for col in columns
    ]
    df_buy.columns = columns
    df_buy.drop(index=0, inplace=True)
    df_buy = df_buy.reset_index(drop=True)
    select_columns = ["Style", "CW", "Vendor"] + [col for col in columns if "Total Item Qty" in col]
    df_buy = df_buy[select_columns]
    df_buy.columns = [col.replace("Total Item Qty", "").strip().upper() for col in df_buy.columns]
    df_buy = df_buy.fillna(0)
    df_buy["CW"] = df_buy["CW"].astype(str)
    df_buy["STYLE"] = df_buy["STYLE"].astype(str).str.strip()
    df_buy["BUY_KEY"] = df_buy["STYLE"] + "-" + df_buy["CW"] + "-" + df_buy["VENDOR"]
    df_buy.drop(columns=["STYLE", "CW", "VENDOR"], inplace=True)
    df_buy = df_buy.groupby(by="BUY_KEY").sum(numeric_only=True)
    return df_buy

def get_required_columns(df_columns):
    required_cols = []
    for col in df_columns:
        if col.lower().startswith("current"):
            break
        required_cols.append(col)
    return required_cols

def process_editable(file):
    df_editable = pd.read_excel(file, sheet_name="Editable")
    cols_needed = get_required_columns(df_editable.columns)
    df_editable = df_editable.loc[:, cols_needed]
    df_editable = df_editable.fillna(0)
    df_editable["Style No"] = df_editable["Style No"].astype(str).str.strip()
    df_editable["IM"] = df_editable["IM"].astype(str).str.strip()
    df_editable["EDIT_KEY"] = df_editable["Style No"] + "-" + df_editable["IM"]
    df_editable.set_index("EDIT_KEY", inplace=True)
    df_editable.drop_duplicates(subset=["Style No"], inplace=True)
    return df_editable

def calculate_actual_yy(df_editable, df_buy):
    data_buy = []
    data_edit = []
    data_missed = []
    def is_subset(list1, list2):
        return set(list1).issubset(set(list2))
    def add_style_to_buy(buy_key, size_matrix):
        row = {"buy_key": buy_key}
        row.update(size_matrix)
        data_buy.append(row)
    def add_style_edit(style, material, size_matrix):
        row = {"style": style, "material": material}
        row.update(size_matrix)
        data_edit.append(row)
    def add_missed_ones(style, style_cw_vendor, buy_size_matrix, edit_size_matrix, sizes_missing_in_editable_sheet):
        row = {"style": style, "style_cw_vendor": style_cw_vendor,
               "buy_size_matrix": buy_size_matrix, "edit_size_matrix": edit_size_matrix,
               "sizes_missing_in_editable_sheet": sizes_missing_in_editable_sheet}
        data_missed.append(row)
    for style_im in df_editable.index:
        style = style_im.split("-")[0]
        for row_id, row in df_buy.iterrows():
            if style in row_id:
                sizes_in_buy_sheet = row[:-1].index[row[:-1] != 0.0].to_list()
                style_yy = df_editable.loc[style_im][2:]
                sizes_in_editable = style_yy[style_yy != 0.0].index.to_list()
                sizes_missing_in_editable_sheet = set(sizes_in_buy_sheet).difference(set(sizes_in_editable))
                if is_subset(sizes_in_buy_sheet, sizes_in_editable):
                    add_style_to_buy(row_id, row[:-1][row[:-1].apply(bool)].to_dict())
                    add_style_edit(style, df_editable.loc[style_im][1], style_yy[sizes_in_buy_sheet].to_dict())
                else:
                    add_missed_ones(style, row_id, sizes_in_buy_sheet, sizes_in_editable, sizes_missing_in_editable_sheet)
    buy = pd.DataFrame(data_buy)
    buy["total_qty"] = buy.sum(axis=1, numeric_only=True)
    buy.fillna(0.0, inplace=True)
    edit = pd.DataFrame(data_edit)
    edit.fillna(0.0, inplace=True)
    final_df = buy.join(edit, lsuffix="_buy", rsuffix="_edit")
    pcs_cols = final_df.filter(like="_buy")
    yy_cols = final_df.filter(like="_edit")
    style_yy = (pcs_cols.values * yy_cols.values).sum(axis=1) / pcs_cols.sum(axis=1)
    final_df["actual_yy"] = style_yy
    final_df = final_df[["buy_key", "style", "material", "total_qty", "actual_yy"]]
    final_df["final_unique"] = final_df.loc[:, ["buy_key", "material"]].agg("-".join, axis=1)
    missed_df = pd.DataFrame(data_missed)
    return final_df, missed_df

def process_costing_yy(file):
    plant_map = {"C050": "LMS", "C150": "LMJ", "C300": "LSB"}
    df_costing = pd.read_excel(file)
    df_costing = df_costing[["Plant", "Customer Style", "Customer Color Code", "RM Customer Reference", "Consumption", "Standard Price", "Stand. Prc Unit"]]
    df_costing["RM Customer Reference"] = df_costing["RM Customer Reference"].str.replace("#", "")
    df_costing["Customer Color Code"] = df_costing["Customer Color Code"].astype(str)
    df_costing["Plant"] = df_costing["Plant"].map(plant_map)
    df_costing.columns = [col.lower().replace(" ", "_").replace(".", "") for col in df_costing.columns]
    df_costing["costing_unique"] = df_costing[["customer_style", "customer_color_code", "plant", "rm_customer_reference"]].agg("-".join, axis=1)
    df_costing.drop_duplicates(inplace=True)
    return df_costing

def compare_calculate_savings(final_df, df_costing):
    df_compare = final_df[["final_unique", "style", "material", "actual_yy", "total_qty"]].merge(df_costing, how="left", left_on="final_unique", right_on="costing_unique")
    df_compare = df_compare[["final_unique", "style", "material", "actual_yy", "total_qty", "consumption", "standard_price", "stand_prc_unit"]]
    df_compare["adjusted_yy"] = df_compare["actual_yy"] * 1.01
    df_compare["costed-actual_yy"] = df_compare["consumption"] - df_compare["actual_yy"]
    df_compare["costed-adjusted_yy"] = df_compare["consumption"] - df_compare["adjusted_yy"]
    def get_final_yy(row):
        if row["costed-actual_yy"] > 0 and row["costed-adjusted_yy"] > 0:
            return row["adjusted_yy"]
        elif row["costed-actual_yy"] > 0 and row["costed-adjusted_yy"] < 0:
            return row["consumption"]
        elif row["costed-actual_yy"] < 0 and row["costed-adjusted_yy"] < 0:
            return row["actual_yy"]
    df_compare["final_yy"] = df_compare.apply(get_final_yy, axis=1)
    def get_savings(row):
        try:
            return round((row["consumption"] - row["final_yy"]) * row["total_qty"], 2)
        except:
            return 0
    df_compare["savings"] = df_compare.apply(get_savings, axis=1)
    return df_compare

uploaded_buy = st.file_uploader("Upload Standard Buy Sheet", type=["xlsx"])
uploaded_editable = st.file_uploader("Upload Editable Sheet", type=["xlsx"])
uploaded_costing = st.file_uploader("Upload Costing YY Sheet", type=["xlsx"])

if st.button("Run Optimization"):
    if uploaded_buy and uploaded_editable and uploaded_costing:
        with st.spinner("Processing your files..."):
            df_buy = process_buy(uploaded_buy)
            df_editable = process_editable(uploaded_editable)
            final_df, missed_df = calculate_actual_yy(df_editable, df_buy)
            df_costing = process_costing_yy(uploaded_costing)
            df_compare = compare_calculate_savings(final_df, df_costing)
            excel_output = io.BytesIO()
            with pd.ExcelWriter(excel_output, engine="xlsxwriter") as writer:
                df_compare.to_excel(writer, sheet_name="consumption_optimizer_results", index=False)
            st.success("Optimization complete!")
            st.download_button(
                label="Download Optimized Consumption Report",
                data=excel_output.getvalue(),
                file_name=f"consumption_optimizer_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            excel_missed = io.BytesIO()
            with pd.ExcelWriter(excel_missed, engine="xlsxwriter") as writer2:
                missed_df.to_excel(writer2, sheet_name="missing_sizes_editable", index=False)
            st.download_button(
                label="Download Missing Sizes Report",
                data=excel_missed.getvalue(),
                file_name=f"missing_sizes_editable_{datetime.now().strftime('%Y%m%dT%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.error("Please upload all three required Excel files.")
