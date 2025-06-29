import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Invoice Projection Tool", layout="centered")

st.title("📊 SPR Invoice Projection Tool")
st.markdown("Upload your SPR and AR files, and get projection results instantly!")

# --- File Upload ---
spr_file = st.file_uploader("📄 Upload SPR File", type=["xlsx"])

ar_files = {}
for company in ['CAA', 'CCA', 'CSM']:
    ar_file = st.file_uploader(f"📄 Upload AR {company} File", type=["xlsx"], key=company)
    if ar_file:
        ar_files[company] = ar_file

if spr_file and len(ar_files) == 3:
    spr = pd.read_excel(spr_file)

    # Clean SPR
    spr['Code'] = spr['Code'].astype(str).str.strip().str.upper()
    spr['NO. INV 1'] = spr['NO. INV 1'].astype(str).str.strip()
    spr['STATUS SM OPERASIONAL'] = spr['STATUS SM OPERASIONAL'].astype(str).str.strip().str.upper()

    def extract_invoices(text):
        if pd.isna(text):
            return []
        return re.findall(r'(P.*?V|P\d+\/INV)', str(text))

    spr['Extra_Invoices'] = spr['NO. INV 2 & SETERUSNYA'].apply(extract_invoices)
    spr['All_Invoices'] = spr.apply(
        lambda row: [row['NO. INV 1']] + row['Extra_Invoices'], axis=1
    )

    spr['Formatted_Invoices'] = spr.apply(
        lambda row: [re.sub(r'\s+', '', inv.strip().upper()) + '/' + re.sub(r'\s+', '', row['Code']) for inv in row['All_Invoices']],
        axis=1
    )

    # Process AR files
    ar_combined = pd.DataFrame()

    for company, file in ar_files.items():
        df_ar = pd.read_excel(file, header=2)
        df_ar.columns = df_ar.columns.str.strip().str.lower()
        df_ar['no. invoice'] = df_ar['no. invoice'].astype(str).str.strip() + f'/INV/{company}'

        if 'proyeksi dari customer' in df_ar.columns:
            df_ar['proyeksi dari customer'] = df_ar['proyeksi dari customer'].apply(
                lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) and hasattr(x, 'strftime') else str(x).strip()
            )

        ar_combined = pd.concat([ar_combined, df_ar], ignore_index=True)

    ar_outstanding = ar_combined[ar_combined['unnamed: 14'].isna()].copy()

    invoice_proj_map = {}
    for _, row in ar_outstanding.iterrows():
        invoice_key = row['no. invoice'].strip().upper()
        proj_date = row['proyeksi dari customer']
        if proj_date and str(proj_date).lower() != 'nan':
            invoice_proj_map[invoice_key] = str(proj_date).strip()

    max_inv = spr['Formatted_Invoices'].apply(len).max()

    for i in range(max_inv):
        spr[f'INV_{i+1}'] = spr['Formatted_Invoices'].apply(lambda x: x[i] if i < len(x) else None)

    def build_summary(row):
        if row['STATUS SM OPERASIONAL'] in ['O14', 'O20']:
            return None
        summary = []
        for i in range(max_inv):
            inv = row.get(f'INV_{i+1}')
            if inv:
                date = invoice_proj_map.get(inv)
                if pd.notna(date) and date != '' and str(date).lower() != 'nan':
                    summary.append(f"INV {i+1}: {date}")
        return '; '.join(summary) if summary else None

    spr['Projection_Summary'] = spr.apply(build_summary, axis=1)

    # Download button
    output = BytesIO()
    spr.to_excel(output, index=False)
    output.seek(0)

    st.success("✅ Projection summary successfully generated!")
    st.download_button("📥 Download Final SPR File", data=output, file_name="spr_projection_result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("ℹ️ Please upload 1 SPR file and all 3 AR files (CA, CC, CS) to begin.")
