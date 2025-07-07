import pandas as pd
import streamlit as st
from io import BytesIO

st.title("Excel Processor for Asset Data")

# Upload files
file_8223 = st.file_uploader("Upload Excel file for 8223", type=["xlsx"])
file_8224 = st.file_uploader("Upload Excel file for 8224", type=["xlsx"])
file_8225 = st.file_uploader("Upload Excel file for 8225", type=["xlsx"])
file_8226 = st.file_uploader("Upload Excel file for 8226", type=["xlsx"])
file_8229 = st.file_uploader("Upload Excel file for 8229", type=["xlsx"])
file_8235 = st.file_uploader("Upload Excel file for 8235", type=["xlsx"])
file_8236 = st.file_uploader("Upload Excel file for 8236", type=["xlsx"])
file_8297 = st.file_uploader("Upload Excel file for 8297", type=["xlsx"])

# User input
selected_date = st.date_input("Enter the date to check against 'Date Acquired'")
remark = st.text_input("Enter Remark")

# Continue only when all files are uploaded
if all([file_8223, file_8224, file_8225, file_8226, file_8229, file_8235, file_8236, file_8297]):

    # Read Excel files
    df_8223 = pd.read_excel(file_8223)
    df_8224 = pd.read_excel(file_8224, skiprows=24)
    df_8225 = pd.read_excel(file_8225, skiprows=24)
    df_8226 = pd.read_excel(file_8226, skiprows=13)
    df_8229 = pd.read_excel(file_8229, skiprows=24)
    df_8235 = pd.read_excel(file_8235, skiprows=1)
    df_8236 = pd.read_excel(file_8236)
    df_8297 = pd.read_excel(file_8297, skiprows=13)

    ###############################################################################################################
    df_8224 = df_8224[~df_8224.apply(lambda row: row.astype(str).str.contains('Grand Total').any(), axis=1)]
    df_8224 = df_8224[~df_8224.apply(lambda row: row.astype(str).str.contains('GRAND TOTAL').any(), axis=1)]
    df_8224 = df_8224[~df_8224.apply(lambda row: row.astype(str).str.contains('Total', case=False).any(), axis=1)]
    df_8224 = df_8224[~df_8224.apply(lambda row: row.astype(str).str.contains('128300', case=False).any(), axis=1)]
    df_8235 = df_8235[~df_8235.apply(lambda row: row.astype(str).str.contains('12810900', case=False).any(), axis=1)]
    df_8235['Cum.acq.value'] = df_8235['Cum.acq.value'].where(df_8235['Trans.acq.val'] == 0, df_8235['Trans.acq.val'])

    df_8225 = df_8225[~df_8225.apply(lambda row: row.astype(str).str.contains('Grand Total').any(), axis=1)]
    df_8225 = df_8225[~df_8225.apply(lambda row: row.astype(str).str.contains('GRAND TOTAL').any(), axis=1)]
    df_8225 = df_8225[~df_8225.apply(lambda row: row.astype(str).str.contains('Total', case=False).any(), axis=1)]
    df_8225 = df_8225[~df_8225.apply(lambda row: row.astype(str).str.contains('128300', case=False).any(), axis=1)]

    df_8229 = df_8229[~df_8229.apply(lambda row: row.astype(str).str.contains('Grand Total').any(), axis=1)]
    df_8229 = df_8229[~df_8229.apply(lambda row: row.astype(str).str.contains('GRAND TOTAL').any(), axis=1)]
    df_8229 = df_8229[~df_8229.apply(lambda row: row.astype(str).str.contains('Total', case=False).any(), axis=1)]
    df_8229 = df_8229[~df_8229.apply(lambda row: row.astype(str).str.contains('128300', case=False).any(), axis=1)]

    df_8223.columns = df_8223.columns.str.strip()
    df_8236.columns = df_8236.columns.str.strip()
    # if df_8226['Location'].isnull().all() or (df_8226['Location'] == '').all():
    #     df_8226 = df_8226.drop(columns=['Location'])
    # if df_8297['Location'].isnull().all() or (df_8297['Location'] == '').all():
    #     df_8297 = df_8297.drop(columns=['Location'])
    df_8226.rename(columns={'Location':'Location1'},inplace=True)
    df_8297.rename(columns={'Location':'Location1'},inplace=True)
    df_8224['Capitalized On.'] = df_8224["Capitalized On"]
    df_8224['Capitalized On..'] = df_8224["Capitalized On"]
    df_8224['Capitalized On...'] = df_8224["Capitalized On"]
    df_8225['Capitalized On.'] = df_8225["Capitalized On"]
    df_8225['Capitalized On..'] = df_8225["Capitalized On"]
    df_8225['Capitalized On...'] = df_8225["Capitalized On"]
    df_8229['Capitalized On.'] = df_8229["Capitalized On"]
    df_8229['Capitalized On..'] = df_8229["Capitalized On"]
    df_8229['Capitalized On...'] = df_8229["Capitalized On"]
    df_8223['Fxd asset.'] = df_8223["Fxd asset"]
    df_8223['Activ.'] = df_8223["Activ"]
    df_8223['Acquis.val.'] = df_8223["Acquis.val"]
    df_8223['Acquis.val..'] = df_8223["Acquis.val"]
    df_8223['Acquis.val..'] = df_8223["Acquis.val"]
    df_8236['FA tp.'] = df_8236["FA tp"]
    df_8236['Activ.'] = df_8236["Activ"]
    df_8226['ATP Asset Type.'] = df_8226["ATP Asset Type"]
    df_8297['ATP Asset Type.'] = df_8297["ATP Asset Type"]


    ###############################################################
    df_8223['Entity'] = '8223'
    df_8224['Entity'] = '8224'
    df_8225['Entity'] = '8225'
    df_8226['Entity'] = '8226'
    df_8229['Entity'] = '8229'
    df_8236['Entity'] = '8236'
    df_8297['Entity'] = '8297'
    df_8235['Entity'] = '8235'
    # Define mappings for renaming columns
    mapping_8235 = {
        "Asset":"Asset Number",
        "Capitalized On":"Date Acquired",
        "Asset Description":"Description (Name of Assets)",
        "Cum.acq.value":"Gross block",
        "Bal.Sh.Acct APC":"Asset Nature",
        "Trans.acq.val":"Gross block-2",
        "PlndDep":"Total Depreciation",
        "End book val":"Net Book Value"
    }
    mapping_8224 = {
        "Asset Number": "Asset Number",
        "Asset": "Description (Name of Assets)",
        "Capitalized On.": "Date Acquired",
        "Capitalized On..": "Date Capitalised",
        "Balance Amount.12": "Total Depreciation",
        "Balance Amount.14": "Net Book Value",
        "Account Determ..1"	: "Asset Nature",
        "Balance Amount.5" : "Gross block",
        "Account Determ.": "Gross block Account"
    }
    mapping_8225 = {
        "Asset Number": "Asset Number",
        "Asset": "Description (Name of Assets)",
        "Capitalized On.": "Date Acquired",
        "Capitalized On..": "Date Capitalised",
        "Balance Amount.12": "Total Depreciation",
        "Balance Amount.14": "Net Book Value",
        "Account Determ..1"	: "Asset Nature",
        "Balance Amount.5" : "Gross block",
        "Account Determ.": "Gross block Account"
    }

    mapping_8229 = {
        "Asset Number": "Asset Number",
        "Asset": "Description (Name of Assets)",
        "Capitalized On.": "Date Acquired",
        "Capitalized On..": "Date Capitalised",
        "Balance Amount.12": "Total Depreciation",
        "Balance Amount.14": "Net Book Value",
        "Account Determ..1"	: "Asset Nature",
        "Balance Amount.5" : "Gross block",
        "Account Determ.": "Gross block Account"
    }

    mapping_8223 = {
        "Fxd asset": "Asset Number",
        "Name": "Description (Name of Assets)",
        "Acquis.val": "Gross block",
        "Depreciat": "Total Depreciation",
        "Net book V": "Net Book Value",
        "Activ": "Date Acquired",
        "Activ.": "Date Capitalised",
        "FA tp": "Asset Nature",
        "Location": "Location"
    }

    mapping_8236 = {
        "Fxd asset": "Asset Number",
        "Name": "Description (Name of Assets)",
        "Acq Cost": "Gross block",
        "Depreciat": "Total Depreciation",
        "NBV": "Net Book Value",
        "Activ": "Date Acquired",
        "Activ.": "Date Capitalised",
        "FA tp.": "Asset Nature",
        "Location": "Location"
    }
    mapping_8226= {
        "Description": "Description (Name of Assets)",
        "Cost": "Gross block",
        "Depreciation Life": "Life of an asset",
        "ATP Asset Type": "Asset Nature",
        "ATP Asset Type.": "Gross block Account",
        "DPA BS Acc Depreciation": "Accumulated Depre Account",
        "EXA EXP PL1 Acc Depreciation": "Depreciation Account",
        "LOC Location": "Location"
    }
    mapping_8297 = {
        "Description": "Description (Name of Assets)",
        "Cost": "Gross block",
        "Depreciation Life": "Life of an asset",
        "ATP Asset Type": "Asset Nature",
        "ATP Asset Type.": "Gross block Account",   
        "DPA BS Acc Depreciation": "Accumulated Depre Account",
        "EXA EXP PL1 Acc Depreciation": "Depreciation Account",
        "LOC Location": "Location"
    }


    df_8223.rename(columns=mapping_8223, inplace=True)
    df_8224.rename(columns=mapping_8224, inplace=True)
    df_8225.rename(columns=mapping_8225, inplace=True)
    df_8226.rename(columns=mapping_8226, inplace=True)
    df_8229.rename(columns=mapping_8229, inplace=True)
    df_8235.rename(columns=mapping_8235, inplace=True)
    df_8236.rename(columns=mapping_8236, inplace=True)
    df_8297.rename(columns=mapping_8297, inplace=True)

    # List of required columns
    required_columns = [
        "Asset Number",
        "Entity",
        "Description (Name of Assets)",
        "Date Acquired",
        "Date Capitalised",
        "Location",
        "Gross block",
        "Depreciation Actual Year",
        "Total Depreciation",
        "Net Book Value",
        "Asset Nature",
        "Life of an asset",
        "Gross block Account",
        "Accumulated Depre Account",
        "Depreciation Account"
    ]
    dataframes = [df_8223, df_8224, df_8225, df_8226, df_8229, df_8236, df_8297, df_8235 ]

    processed_dfs = []
    for df in dataframes:
        # Ensure columns match required_columns
        columns_to_drop = [col for col in df.columns if col not in required_columns]
        if columns_to_drop:
            df.drop(columns=columns_to_drop, inplace=True)
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            for col in missing_columns:
                df[col] = pd.NA  # Or another default value if preferred

        # Reorder columns to match required_columns order
        df = df[required_columns]
        
        # Append processed DataFrame to the list
        processed_dfs.append(df)

    # Concatenate all DataFrames
    df_8223 = processed_dfs[0]
    df_8224 = processed_dfs[1]
    df_8225 = processed_dfs[2]
    df_8226 = processed_dfs[3]
    df_8229 = processed_dfs[4]
    df_8235 = processed_dfs[5]
    df_8236 = processed_dfs[6]
    df_8297 = processed_dfs[7]
    df_8224['Depreciation Actual Year'] = df_8224['Depreciation Actual Year'].replace('Total', None)
    df_8225['Depreciation Actual Year'] = df_8225['Depreciation Actual Year'].replace('Total', None)
    df_8229['Depreciation Actual Year'] = df_8229['Depreciation Actual Year'].replace('Total', None)
    # Build list of processed dataframes
    dfs = [df_8223, df_8224, df_8225, df_8226, df_8229, df_8236, df_8297, df_8235]
    sheet_names = ['8223', '8224', '8225', '8226', '8229', '8236', '8297', '8235']

    # Consolidate
    consolidated_df = pd.concat(dfs, ignore_index=True)
    consolidated_df = consolidated_df[~consolidated_df.apply(lambda row: row.astype(str).str.contains('Grand Total', case=False).any(), axis=1)]
    consolidated_df = consolidated_df[~consolidated_df.apply(lambda row: row.astype(str).str.contains('128300', case=False).any(), axis=1)]

    # Map asset nature
    asset_data = {
        20000: "Buildings at cost", 30000: "Plant and Machinery", 50000: "Fixtures and fittings",
        40000: "Cars and other transport equipment", 70000: "IT Software, Acquired",
        11100000: "Buildings at cost", 12100000: "Plant and Machinery", 12200000: "Fixtures",
        12400000: "Cars", 10306000: "IT Software", 12700000: "Rental Equipments",
        10307000: "Intangible assets", 12201000: "Fixtures", 11300000: "Land"
    }
    consolidated_df['Asset Nature'] = consolidated_df['Asset Nature'].map(asset_data).fillna(consolidated_df['Asset Nature'])

    consolidated_df['Date Acquired'] = pd.to_datetime(consolidated_df['Date Acquired'], errors='coerce')
    consolidated_df['Date Capitalised'] = pd.to_datetime(consolidated_df['Date Capitalised'], errors='coerce')

    # Add Remark
    consolidated_df['Remark'] = ''
    consolidated_df.loc[consolidated_df['Date Acquired'] > pd.to_datetime(selected_date), 'Remark'] = remark

    # Save to Excel in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, name in zip(dfs, sheet_names):
            df.to_excel(writer, sheet_name=name, index=False)
        consolidated_df.to_excel(writer, sheet_name='Consolidated', index=False)
    output.seek(0)

    st.success("Processing complete! Download the result below:")

    # Download button
    st.download_button(
        label="📥 Download Processed Excel",
        data=output,
        file_name='FAR_Consolidated.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
else:
    st.warning("Please upload all 8 required Excel files to proceed.")

import streamlit as st
import requests
from docx import Document
from io import BytesIO

# GitHub raw URL
docx_url = "https://github.com/test-nexease/FAR_Conso/blob/main/%F0%9F%93%84%20SOP%20Note.docx"

# Download and read the .docx
response = requests.get(docx_url)
doc = Document(BytesIO(response.content))

# Extract and show content
content = "\n".join([para.text for para in doc.paragraphs])
st.text_area("📘 Preview of Word Note", content, height=400)
