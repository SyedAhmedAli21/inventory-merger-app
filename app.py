import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Inventory & Orders Merger", layout="wide")

st.title("📦 Multi-File Inventory & Orders Merger")
st.write("Upload your 3 Excel files below to aggregate On-Hand and WIP inventory into your Orders Master report.")

# File Uploaders
col1, col2, col3 = st.columns(3)

with col1:
    orders_file = st.file_uploader("1. Upload ORDERS File", type=["xlsx", "xls"])

with col2:
    onhand_file = st.file_uploader("2. Upload ON-HAND File", type=["xlsx", "xls"])

with col3:
    wip_file = st.file_uploader("3. Upload WIP File", type=["xlsx", "xls"])

# Helper functions for WIP processing
def join_unique_strings(series):
    cleaned = series.dropna().astype(str).str.strip()
    cleaned = [val for val in cleaned if val != '' and val.lower() != 'nan']
    return ", ".join(sorted(set(cleaned)))

def format_dates(series):
    formatted_dates = []
    for val in series.dropna():
        if pd.notna(val):
            try:
                formatted_dates.append(pd.to_datetime(val).strftime('%Y-%m-%d'))
            except Exception:
                formatted_dates.append(str(val).strip())
    return ", ".join(sorted(set(formatted_dates)))

# Process Button
if st.button("🚀 Process & Combine Files", type="primary"):
    if orders_file and onhand_file and wip_file:
        try:
            with st.spinner("Processing files..."):
                # Load Data
                df_orders = pd.read_excel(orders_file)
                df_onhand = pd.read_excel(onhand_file)
                df_wip = pd.read_excel(wip_file)

                # Clean Column Names
                df_orders.columns = df_orders.columns.str.strip()
                df_onhand.columns = df_onhand.columns.str.strip()
                df_wip.columns = df_wip.columns.str.strip()

                UPC_COL = 'Upc No'
                ONHAND_QTY_COL = 'On Hand'
                PO_COL = 'Po'
                WIP_QTY_COL = 'Po Qty'
                ETA_COL = 'Eta Date'

                # Standardize UPCs
                for df in [df_orders, df_onhand, df_wip]:
                    if UPC_COL in df.columns:
                        df[UPC_COL] = df[UPC_COL].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

                # Process On-Hand
                if ONHAND_QTY_COL in df_onhand.columns:
                    onhand_grouped = df_onhand.groupby(UPC_COL)[ONHAND_QTY_COL].sum().reset_index()
                else:
                    onhand_grouped = pd.DataFrame(columns=[UPC_COL, ONHAND_QTY_COL])

                # Process WIP
                wip_grouped = df_wip.groupby(UPC_COL).agg({
                    PO_COL: join_unique_strings,
                    WIP_QTY_COL: 'sum',
                    ETA_COL: format_dates
                }).reset_index()

                wip_grouped.rename(columns={
                    PO_COL: 'WIP Po Numbers',
                    WIP_QTY_COL: 'Total WIP Po Qty',
                    ETA_COL: 'WIP Eta Dates'
                }, inplace=True)

                # Merge Data
                df_final = pd.merge(df_orders, onhand_grouped, on=UPC_COL, how='left')
                df_final = pd.merge(df_final, wip_grouped, on=UPC_COL, how='left')

                if ONHAND_QTY_COL in df_final.columns:
                    df_final[ONHAND_QTY_COL] = df_final[ONHAND_QTY_COL].fillna(0)
                df_final['Total WIP Po Qty'] = df_final['Total WIP Po Qty'].fillna(0)
                df_final['WIP Po Numbers'] = df_final['WIP Po Numbers'].fillna('')
                df_final['WIP Eta Dates'] = df_final['WIP Eta Dates'].fillna('')

                # Convert output to Excel in memory
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False)
                processed_data = output.getvalue()

            st.success("✅ Files successfully combined!")
            
            # Download Button
            st.download_button(
                label="📥 Download Consolidated Excel File",
                data=processed_data,
                file_name="Consolidated_Orders_Inventory.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.warning("Please upload all 3 Excel files before processing.")