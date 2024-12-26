import numpy as np
import pandas as pd
import json
import gspread
from flask import Flask, render_template, request, jsonify
from google.oauth2 import service_account

# Initialize Flask app
app = Flask(__name__)
application = app

# # Initialize global variables
# list_tables = []
# hc_table = None
# npv1, npv2, npv3, npv4 = 0, 0, 0, 0

def is_float(s):
    """Check if a string can be converted to a float."""
    s = s.replace('.', '', 1)
    return s.isdigit()

@app.route('/')
def index():
    global list_tables
    list_tables = []
    
    global hc_table
    hc_table = None
    
    return render_template('index.html')

@app.route('/submit_form', methods=['POST'])
def submit_form():
    global list_tables
    data = request.get_json()
    
    form_data = data['formData']
    form_df = pd.DataFrame([form_data])
    df = form_df.tail(1)

    document_id = "1BRjiwdrDe0Q1vLFywMZwUALwNMOBYAPLrS_CWeiFilc"
    tab_name = "UserDataEntry"
    full_url = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet={tab_name}"
    
    harvest_carbon_url = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet=HarvestCarbonCalculator"

    # Load Google Sheets API credentials
    with open('keys.json') as file:
        file_content = json.load(file, strict=False)

    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_info(file_content, scopes=scope)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(document_id)
    worksheet = sh.get_worksheet(3)

    conversions = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet=Smith_TableD6").iloc[:38, :13]

    for i in range(df.shape[0]):
        area = df.iloc[i]['area']
        region = df.iloc[i]['region']
        grp = df.iloc[i]['forestTypeGroup']
        origin = df.iloc[i]['origin']
        age = df.iloc[i]['age']
        hyb = df.iloc[i]['harvestYearsBusiness']
        hye = df.iloc[i]['harvestYearsER']

        temp = [int(area), region, grp, origin, age, int(hyb), int(hye)]
        for i in range(len(temp)):
            cell = worksheet.acell(f'C{i+3}')
            cell.value = temp[i]
            worksheet.update_cells([cell])

        list_table_temp = pd.read_csv(full_url).iloc[7:12, 6:18]
        list_table_temp.columns = ['Attributes', 'Year_0', 'Year_5', 'Year_10', 'Year_15', 'Year_20', 'Year_25', 'Year_30', 'Year_35', 'Year_40', 'Year_45', 'Year_50']
        list_table_temp.Attributes = [' '.join(i.split('\n')) for i in list_table_temp.Attributes]
        list_table_temp = list_table_temp.iloc[1:, :].fillna(0)
        list_table_temp.reset_index(drop=True, inplace=True)

        ind = len(list_table_temp.columns) - 1
        for j in range(len(list_table_temp.columns) - 1, -1, -1):
            un = list_table_temp.iloc[:, j].unique()
            if len(un) == 1 and un[0] == 0 and j < ind:
                ind = j
        list_table_temp = list_table_temp.iloc[:, :ind].reset_index(drop=True)

        list_tables.append(list_table_temp)

        # Harvest Carbon
        hc = pd.read_csv(harvest_carbon_url).iloc[0:6, 17:24].fillna('-')
        unique_vals = list(set(hc.iloc[:, -4:].values.flatten()))
        if sum([is_float(str(i).strip()) for i in unique_vals]):
            con = conversions[conversions.TD6RegionTool.str.contains(region)]
            vals = {'Softwood Sawlog': [], 'Softwood Pulpwood': [], 'Softwood Fuelwood': [],
                    'Hardwood Sawlog': [], 'Hardwood Pulpwood': [], 'Hardwood Fuelwood': []}
            for _, i in con.iterrows():
                if i.TD6WoodType == 'Softwood':
                    if i.TD6LogType == 'Sawlog' or i.TD6LogType == 'All':
                        vals[f'{i.TD6WoodType} Sawlog'].append(i['TD6Softwood lumber'])
                        vals[f'{i.TD6WoodType} Fuelwood'].append(i['TD6Fuel and other_emissions'])
                    if i.TD6LogType == 'Pulpwood' or i.TD6LogType == 'All':
                        vals[f'{i.TD6WoodType} Pulpwood'].append(i['TD6Wood pulp'])
                elif i.TD6WoodType == 'Hardwood':
                    if i.TD6LogType == 'Sawlog' or i.TD6LogType == 'All':
                        vals[f'{i.TD6WoodType} Sawlog'].append(i['TD6Hardwood lumber'])
                        vals[f'{i.TD6WoodType} Fuelwood'].append(i['TD6Fuel and other_emissions'])
                    if i.TD6LogType == 'Pulpwood' or i.TD6LogType == 'All':
                        vals[f'{i.TD6WoodType} Pulpwood'].append(i['TD6Wood pulp'])
            vals = pd.DataFrame(vals).T.reset_index().iloc[:, -1]
            hc.iloc[:, -1] = round(hc.iloc[:, -1].str.strip().str.replace(',', '').replace('-', 0).astype(float) * vals, 2)
            list_tables.append(hc)
    return jsonify({"status": "processing_complete"})

@app.route('/delete_row', methods=['POST'])
def delete_row():
    data = request.get_json()
    deleted_row_idx = data.get('deletedRowIdx')
    if deleted_row_idx is not None:
        global list_tables
        list_tables.pop((deleted_row_idx + 1) * 2)
        next_ind = (deleted_row_idx + 1) * 2
        if next_ind < len(list_tables) and list_tables[next_ind].columns[0] == 'Timber Type':
            list_tables.pop(next_ind)
    return "row deleted"

@app.route('/output')
def output():
    global list_tables
    ind_list = []
    for i in range(len(list_tables)):
        if list_tables[i].columns[0] == 'Attributes':
            ind_list.append(i)
    return render_template('output.html', list_tables=list_tables, ind_list=ind_list)

@app.route('/submit_final', methods=['POST'])
def submit_final():
    data = request.get_json()
    ec_data = data.get('economicData', {})

    global hc_table
    hc_table = None

    global list_tables
    for i in range(len(list_tables)):
        if list_tables[i].columns[0] == 'Timber Type':
            temp = list_tables[i]
            temp.columns = temp.columns.str.strip()

            # Convert the necessary columns to numeric values
            temp[temp.columns[3]] = pd.to_numeric(temp[temp.columns[3]].astype(str).str.replace(",", "").str.strip(), errors='coerce')
            temp[temp.columns[4]] = pd.to_numeric(temp[temp.columns[4]].astype(str).str.replace(",", "").str.strip(), errors='coerce')
            temp[temp.columns[5]] = pd.to_numeric(temp[temp.columns[5]].astype(str).str.replace(",", "").str.strip(), errors='coerce')
            temp[temp.columns[6]] = pd.to_numeric(temp[temp.columns[6]].astype(str).str.replace(",", "").str.strip(), errors='coerce')

            # Combine with hc_table if it is not None
            if hc_table is not None:
                hc_table.iloc[:, [3, 4, 5, 6]] += temp.iloc[:, [3, 4, 5, 6]]
            else:
                hc_table = temp.copy()

    # Ensure hc_table is not None before proceeding
    if hc_table is None:
        return jsonify({"error": "No data available to perform economic analysis."}), 400

    # Define the prices dictionary
    prices = {
        "Softwood Sawlog": float(ec_data.get("p1", 50)),
        "Softwood Pulpwood": float(ec_data.get("p2", 30)),
        "Softwood Fuelwood": float(ec_data.get("p3", 20)),
        "Hardwood Sawlog": float(ec_data.get("p4", 60)),
        "Hardwood Pulpwood": float(ec_data.get("p5", 40)),
        "Hardwood Fuelwood": float(ec_data.get("p6", 25))
    }

    i = float(ec_data.get("interestRate", 5))/100
    carbon_price = float(ec_data.get("carbonPrice", 30))
    n1 = 10
    n2 = 20
    
    print(hc_table.values)
    
    
    # Calculate NPV and other values
    hc_table[hc_table.columns[3]] = (
        hc_table[hc_table.columns[3]] *
        ((hc_table['Timber Type'] + ' ' + hc_table['Roundwood Category']).map(prices))
    ) / ((1 + i) ** n1)
    
    global npv1, npv2, npv21, npv3, npv4, npv5
    
    npv1 = hc_table[hc_table.columns[3]].sum()
    
    hc_table[hc_table.columns[4]] = (
        hc_table[hc_table.columns[4]] *
        ((hc_table['Timber Type'] + ' ' + hc_table['Roundwood Category']).map(prices))
    ) / ((1 + i) ** n2)
    
    npv2 = hc_table[hc_table.columns[4]].sum()
    
    npv3 = ((
        hc_table[hc_table.columns[5]] *
        carbon_price
    ))
    
    npv4 = npv2 + npv3
    npv5 = npv4 - npv1
    npv21 = npv2 - npv1
    
    print("NPV1", npv1)
    print("NPV2", npv2)
    print("NPV21", npv21)
    print("NPV3", npv3)
    print("NPV4", npv4)
    print("NPV5", npv5)
    
    npv_values = {
        "npv1": npv1,
        "npv2": npv2,
        "npv21": npv21,
        "npv3": npv3,
        "npv4": npv4,
        "npv5": npv5
    }
    
    return jsonify({"status": "processing_complete"})

@app.route('/finaloutput')
def finaloutput():
    return render_template('finaloutput.html', npv1 = round(npv1,2), npv2 = round(npv2,2), npv21 = round(npv21,2), npv3=round(npv3,2), npv4=round(npv4,2), npv5 = round(npv5,2),)

if __name__ == '__main__':
    app.run(debug=True,port=5000)
