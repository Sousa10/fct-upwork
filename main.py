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
list_tables = []
hc_table = None
npv1, npv2, npv3, npv4 = 0, 0, 0, 0

def is_float(s):
    """Check if a string can be converted to a float."""
    s = s.replace('.', '', 1)
    return s.isdigit()

# function to filter numbers larger than the last number
def filter_larger_than_last(numbers):
    if not numbers:
        return None
    last_number = numbers[-1]
    filtered_list = [num for num in numbers if num > last_number]
    return filtered_list if filtered_list else [last_number]

def convert_to_numeric(value):
    try:
        return float(str(value).strip().replace(",", ""))
    except ValueError:
        return 0

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

    document_id = "17l0BupZ9jf5mORvW-1Ak7qooLqKs8IgsRfmTiLIeI4o"
    tab_name = "UserDataEntry"
    full_url = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet={tab_name}"
    
    harvest_carbon_url = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet=HarvestCarbonCalculator"
    harvest_carbon_bau_url = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet=Harvest%20Carbon%20Calculator%20(BAU)"

    # Load Google Sheets API credentials
    with open('keys.json') as file:
        file_content = json.load(file, strict=False)

    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_info(file_content, scopes=scope)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(document_id)
    worksheet = sh.get_worksheet(3)

    conversions = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet=Smith_TableD6").iloc[:38, :13]
    global hyb, hye, eco_carbon, ccf, carbon_seq, total_hwp, total_afolu, benefit, g_area

    for i in range(df.shape[0]):
        # forest_mgmt_trmt = df.iloc[i]['forest_mgmt_trmt']
        area = df.iloc[i]['area']
        g_area = int(area)
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

        # Harvest Carbon
        hc = pd.read_csv(harvest_carbon_url).iloc[0:6, 17:24]
        hc_bau = pd.read_csv(harvest_carbon_bau_url).iloc[0:6, 17:24]
        col1_col2 = hc.iloc[:, :2]
        hc_bau_last = hc_bau.iloc[:, -1].apply(convert_to_numeric).fillna(0)
        hc_last = hc.iloc[:, -1].apply(convert_to_numeric).fillna(0)
        difference = hc_bau_last - hc_last
        new_df = pd.concat([col1_col2, hc_bau_last, hc_last, difference], axis=1)
        new_df.columns = ["Timber type", "Roundwood category", "Wood Volume at BAU (CCF/Cunit)", "Wood Volume with Extended Rotation (CCF/Cunit)", "Difference Between Extended Rotation and BAU (CCF/Cunit)"]
        ## CCF table
        temp_df = pd.DataFrame({
            "combined": new_df.iloc[:, 0] + " " + new_df.iloc[:, 1],
            "a_values": hc_bau_last,
            "b_values": hc_last,
            "a_minus_b_values": difference
        })
        filtered_df = temp_df[~((temp_df["a_values"] == 0) & (temp_df["b_values"] == 0) & (temp_df["a_minus_b_values"] == 0))]
        ccf = {
            "combined": filtered_df["combined"].to_list(),
            "a_values": filtered_df["a_values"].to_list(),
            "b_values": filtered_df["b_values"].to_list(),
            "a_minus_b_values": filtered_df["a_minus_b_values"].to_list(),
        }
        list_tables.append(new_df)

        list_table_temp = pd.read_csv(full_url).iloc[6:12, 6:18]
        list_table_temp.columns = ['Attributes', 'Year_0', 'Year_5', 'Year_10', 'Year_15', 'Year_20', 'Year_25', 'Year_30', 'Year_35', 'Year_40', 'Year_45', 'Year_50']
        list_table_temp.Attributes = [' '.join(i.split('\n')) for i in list_table_temp.Attributes] 
        list_table_temp = list_table_temp.iloc[1:, :].fillna(0)
        list_table_temp.reset_index(drop=True, inplace=True)

        ind = len(list_table_temp.columns)
        for j in range(len(list_table_temp.columns) - 1, -1, -1):
            un = list_table_temp.iloc[:, j].unique()
            if len(un) == 1 and un[0] == 0 and j < ind:
                ind = j
        list_table_temp = list_table_temp.iloc[:, :ind].reset_index(drop=True)
        filtered_numbers = [
            int(x.replace(',', '')) if isinstance(x, str) and x.replace(',', '').lstrip('-').isdigit() and x.replace(',', '').isdigit()
            else float(x.replace(',', '')) if isinstance(x, str) and x.replace(',', '').lstrip('-').isdigit()
            else x
            for x in list_table_temp.iloc[3, 1:].tolist()
        ]
        formatted_numbers = [x for x in filtered_numbers if isinstance(x, (int, float)) and x > 0]
        eco_carbon = len(formatted_numbers) and ", ".join(map(str, filter_larger_than_last(formatted_numbers)))
        
        list_tables.append(list_table_temp)

        # Forest Management Results
        forest_mgmt_sheet_name = "Forest%20Mgmt%20%26%20HWP%20Results"
        forest_mgmt_url = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet={forest_mgmt_sheet_name}"
        forest_mgmt = pd.read_csv(forest_mgmt_url).iloc[0:19, 1:5].fillna('-')

        carbon_seq = forest_mgmt.iloc[0, 1]
        total_hwp = forest_mgmt.iloc[16, 1]
        total_afolu = forest_mgmt.iloc[17, 1]
        benefit = forest_mgmt.iloc[18, 1]
        list_tables.append(forest_mgmt)

    return jsonify({"status": "processing_complete"})

@app.route('/calculate_factor', methods=['POST'])
def calculate_factor():
    global prices
    data = request.get_json()
    ec_data = data.get('economicData', {})
    x1 = 2.012072
    z1 = 1 / 1.05
    x2 = 2.012072
    z2 = 1 / 1.35
    x9 = 1.006519343
    x10 = 0.817685468
    multiple = 1
    prices = {
        "Softwood Sawlog": {
            "value": float(ec_data.get("p1", 50)),
            "unit": ec_data.get("p1_unit")
        },
        "Softwood Pulpwood": {
            "value": float(ec_data.get("p2", 50)),
            "unit": ec_data.get("p2_unit")
        },
        "Softwood Fuelwood": {
            "value": float(ec_data.get("p3", 50)),
            "unit": ec_data.get("p3_unit")
        },
        "Hardwood Sawlog": {
            "value": float(ec_data.get("p4", 50)),
            "unit": ec_data.get("p4_unit")
        },
        "Hardwood Pulpwood": {
            "value": float(ec_data.get("p5", 50)),
            "unit": ec_data.get("p5_unit")
        },
        "Hardwood Fuelwood": {
            "value": float(ec_data.get("p6", 50)),
            "unit": ec_data.get("p6_unit")
        }
    }
    carbon_unit = ec_data.get("p7_unit")
    multiple_a_values = []
    multiple_b_values= []
    sum_a_values = []
    sum_b_values = []
    for idx, value in enumerate(ccf["combined"]):
        a_value = ccf["a_values"][idx]
        b_value = ccf["b_values"][idx]
        unit = prices[value]["unit"]
        val = prices[value]["value"]
        if value == "Softwood Sawlog":
            if unit == "mbf-international":
                multiple = x1
            elif unit == "mbf-scribner":
                multiple = x1 * z1
            elif unit == "mbf-doyle":
                multiple = x1 * z2
            else:
                multiple = 1            
        elif value == "Softwood Pulpwood":
            if unit == "mbf-other":
                multiple = x9
            else:
                multiple = 1
        elif value == "Softwood Fuelwood":
            multiple = 1
        elif value == "Hardwood Sawlog":
            if unit == "mbf-international":
                multiple = x2
            elif unit == "mbf-scribner":
                multiple = x2 * z1
            elif unit == "mbf-doyle":
                multiple = x2 * z2
            else:
                multiple = 1
        elif value == "Hardwood Pulpwood":
            if unit == "mbf-other":
                multiple = x10
            else:
                multiple = 1
        elif value == "Hardwood Fuelwood":
            multiple = 1
        multiple_a_values.append(a_value * multiple)
        multiple_b_values.append(b_value * multiple)
        sum_a_values.append(a_value * multiple * val)
        sum_b_values.append(b_value * multiple * val)
    ccf["multiple_a_values"] = multiple_a_values
    ccf["multiple_b_values"] = multiple_b_values
    # Calculate np values
    global npv1, npv2, npv21, npv3, npv4, npv5
    i = float(ec_data.get("interestRate", 5))/100
    carbon_price = float(ec_data.get("carbonPrice"))
    n1 = int(hyb)
    n2 = int(hye)
    npv1 = sum(sum_a_values) / ((1 + i) ** n1)
    npv2 = sum(sum_b_values) / ((1 + i) ** n2)
    npv21 = npv2 - npv1
    npv3 = 0
    if carbon_unit == "tonne-co2eq":
        formatted_benefit = float(benefit.replace(",", "")) if float(benefit.replace(",", "")) < 0 else 0
        npv3 = formatted_benefit * carbon_price
        print(f"formatted_benefit: {formatted_benefit}")
    else:
        f = float(hye) - float(hyb)
        npv3 = g_area * carbon_price * f
        print(f"else entered\n g_area: {g_area}\n carbon_price: {carbon_price}\n f: {f}")
    print(f"npv3: {npv3}")
    # print(f"formatted_benefit: {formatted_benefit}")
    print(f"benefit: {benefit}")
    npv4 = npv2 + npv3
    npv5 = npv4 - npv1
    npv_values = {
        "npv1": round(npv1,2),
        "npv2": round(npv2,2),
        "npv21": round(npv21,2),
        "npv3": round(npv3,2),
        "npv4": round(npv4,2),
        "npv5": round(npv5,2)
    }    
    return jsonify({"data": ccf, "npv_values": npv_values})
        

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
    list_tables[2].columns = [list_tables[2].columns[0], 'Value', 'Value', list_tables[2].columns[3]]
    return render_template('output.html', list_tables=list_tables[1:], ind_list=ind_list[1:])

@app.route('/finaloutput')
def finaloutput():
    return render_template('finaloutput.html', npv1 = round(npv1,2), npv2 = round(npv2,2), npv21 = round(npv21,2), npv3=round(npv3,2), npv4=round(npv4,2), npv5 = round(npv5,2), hyb=hyb, hye=hye, eco_carbon=eco_carbon, ccf=ccf, carbon_seq=carbon_seq, total_hwp=total_hwp, total_afolu=total_afolu, benefit=benefit, area=g_area)

@app.route('/summary')
def summary():
    return render_template('summary.html', hyb=hyb, hye=hye, eco_carbon=eco_carbon, ccf=ccf, carbon_seq=carbon_seq, total_hwp=total_hwp, total_afolu=total_afolu, benefit=benefit, area=g_area)

if __name__ == '__main__':
    app.run()
