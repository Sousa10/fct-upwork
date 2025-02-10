import pandas as pd
from flask import Flask, render_template, request, jsonify
from onedrive_helper import OneDriveHelper
from config import FILE_PATH
import requests
import logging
import time

# Initialize Flask app
app = Flask(__name__)
application = app

# Initialize logging
# logging.basicConfig(level=logging.DEBUG)

# # Initialize global variables
list_tables = []
hc_table = None
npv1, npv2, npv3, npv4 = 0, 0, 0, 0

pd.set_option('display.max_columns', None)

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

def wait_for_data_update(onedrive, file_path, expected_values, sheet_name, range_address, max_wait=10):
    """
    Wait until the updated values are actually reflected in OneDrive.
    """
    start_time = time.time()
    
    while time.time() - start_time < max_wait:
        print("⏳ Checking if OneDrive has updated values...")
        response = onedrive.read_excel(file_path, sheet_name, range_address)
        data = response.get('values')

        if data == expected_values:
            print("✅ OneDrive has updated values!")
            return True

        time.sleep(1)

    print("⚠️ Timed out waiting for OneDrive update to propagate.")
    return False


@app.route('/submit_form', methods=['POST'])
def submit_form():
    global list_tables
    data = request.get_json()
    form_data = data['formData']    
    form_df = pd.DataFrame([form_data])
    df = form_df.tail(1)

    try:
        # Initialize OneDrive helper
        onedrive = OneDriveHelper()
        # print(onedrive.get_file_metadata("Copy of Excel Workbook V1.1 - Copy.xlsx"))
        root_files = onedrive.list_root_files()
        # for item in root_files.get('value', []):
        #     print(item['name'], item['id'], item.get('parentReference', {}).get('path', 'Unknown Path'))
    except Exception as e:
        logging.error(f"OneDrive initialization error: {str(e)}")
        return jsonify({"error": str(e)}), 500

    # document_id = "17l0BupZ9jf5mORvW-1Ak7qooLqKs8IgsRfmTiLIeI4o"
    # tab_name = "UserDataEntry"
    # full_url = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet={tab_name}"
    
    # harvest_carbon_url = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet=HarvestCarbonCalculator"
    # harvest_carbon_bau_url = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet=Harvest%20Carbon%20Calculator%20(BAU)"

    # Load Google Sheets API credentials
    # with open('keys.json') as file:
    #     file_content = json.load(file, strict=False)

    # scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    # creds = service_account.Credentials.from_service_account_info(file_content, scopes=scope)
    # gc = gspread.authorize(creds)
    # sh = gc.open_by_key(document_id)
    # worksheet = sh.get_worksheet(3)

    # conversions = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet=Smith_TableD6").iloc[:38, :13]
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

        # Update UserDataEntry sheet
        values = [[int(area)], [region], [grp], [origin], [age], [int(hyb)], [int(hye)]]
        try:
            onedrive.update_excel(
                FILE_PATH,
                # EXCEL_FILE_ID,
                'User Data Entry',
                'C3:C9',
                values
            )
        except requests.exceptions.RequestException as e:
            logging.error(f"Request failed: {e.response}")  # Log response details
            return jsonify({"error": str(e)}), 500
        
        # try:
        #     recalculation_result = onedrive.recalculate_workbook(EXCEL_FILE_ID)
        #     logging.info(f"Recalculation result: {recalculation_result}")
        # except requests.exceptions.RequestException as e:
        #     logging.error(f"Recalculation failed: {str(e)}")
        #     return jsonify({"error": str(e)}), 500

        # temp = [int(area), region, grp, origin, age, int(hyb), int(hye)]
        # for i in range(len(temp)):
        #     cell = worksheet.acell(f'C{i+3}')
        #     cell.value = temp[i]
        #     worksheet.update_cells([cell])

        # Read other sheets using Microsoft Graph API
        # Read other sheets using Microsoft Graph API with specified range
         # Wait for update to complete
        print("Waiting for Excel update to complete...")

        # Force OneDrive to refresh the data
        wait_for_data_update(onedrive, FILE_PATH, values, 'User Data Entry', 'C3:C9')
        
        try:
            harvest_carbon_response = onedrive.read_excel(FILE_PATH, 'Harvest Carbon Calculator', 'A1:Z100')
            harvest_carbon_bau_response = onedrive.read_excel(FILE_PATH, 'Harvest Carbon Calculator (BAU)', 'A1:Z100')
            
            # Extract values from the response
            harvest_carbon = harvest_carbon_response.get('values')
            # print(f"harvest_carbon: {harvest_carbon}")
            harvest_carbon_bau = harvest_carbon_bau_response.get('values')
            
            if harvest_carbon is None or harvest_carbon_bau is None:
                raise ValueError("No data found in the specified range")
            
        except Exception as e:
            logging.error(f"Error reading Excel data: {str(e)}")
            return jsonify({"error": str(e)}), 500
            # Harvest Carbon

        # hc = pd.read_csv(harvest_carbon_url).iloc[0:6, 17:24]
        # hc_bau = pd.read_csv(harvest_carbon_bau_url).iloc[0:6, 17:24]
        # Debug print to check the content of harvest_carbon and harvest_carbon_bau
        # print(f"harvest_carbon content: {harvest_carbon}")
        # print(f"harvest_carbon_bau content: {harvest_carbon_bau}")

        # Ensure harvest_carbon and harvest_carbon_bau are valid inputs for pd.DataFrame
        try:
            hc = pd.DataFrame(harvest_carbon).iloc[0:6, 17:24]
            # print(f"hc: {hc}")
            hc_bau = pd.DataFrame(harvest_carbon_bau).iloc[0:6, 17:24]
            # print(f"hc_bau: {hc_bau}")
        except Exception as e:
            print(f"Error creating DataFrame: {e}")
            return jsonify({"error": "Invalid data format for harvest_carbon or harvest_carbon_bau"}), 400
        col1_col2 = hc.iloc[:, :2]
        hc_bau_last = hc_bau.iloc[:, 3].apply(convert_to_numeric).fillna(0)
        hc_last = hc.iloc[:, 3].apply(convert_to_numeric).fillna(0)
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

        # list_table_temp = pd.read_csv(full_url).iloc[6:12, 6:18]
        try:
            list_table_temp = onedrive.read_excel(FILE_PATH, 'User Data Entry', 'G9:R14')['values']
            logging.debug(f"UserDataEntry Data: {list_table_temp}")
        except Exception as e:
            logging.error(f"Error reading UserDataEntry data: {str(e)}")
            return jsonify({"error": str(e)}), 500

        list_table_temp = pd.DataFrame(list_table_temp)

        # Inspect the extracted DataFrame
        # print(list_table_temp)  # Check the extracted data
        # print(list_table_temp.columns)  # See the number of columns
        # Update column names dynamically
        expected_columns = ['Attributes', 'Year_0', 'Year_5', 'Year_10', 'Year_15', 'Year_20', 'Year_25', 'Year_30', 'Year_35', 'Year_40', 'Year_45', 'Year_50']

        # Adjust column length to match the DataFrame
        if len(list_table_temp.columns) != len(expected_columns):
            expected_columns = expected_columns[:len(list_table_temp.columns)]

        list_table_temp.columns = expected_columns

        # Process attributes column and clean up
        list_table_temp['Attributes'] = [' '.join(str(i).split('\n')) for i in list_table_temp['Attributes']]
        list_table_temp = list_table_temp.fillna(0)
        list_table_temp.reset_index(drop=True, inplace=True)

        # print(list_table_temp)

        ind = len(list_table_temp.columns)
        for j in range(len(list_table_temp.columns) - 1, -1, -1):
            un = list_table_temp.iloc[:, j].unique()
            if len(un) == 1 and un[0] == 0 and j < ind:
                ind = j
        list_table_temp = list_table_temp.iloc[1:, :ind].reset_index(drop=True)
        filtered_numbers = [
            int(x.replace(',', '')) if isinstance(x, str) and x.replace(',', '').lstrip('-').isdigit() and x.replace(',', '').isdigit()
            else float(x.replace(',', '')) if isinstance(x, str) and x.replace(',', '').lstrip('-').isdigit()
            else x
            for x in list_table_temp.iloc[3, 1:].tolist()
        ]
        # formatted_numbers = [x for x in filtered_numbers if isinstance(x, (int, float)) and x > 0]

        # Find the first non-negative value in the last row, iterating from right to left
        last_row = list_table_temp.iloc[-1, 1:].astype(str).reset_index(drop=True)  # Exclude the 'Attributes' column
        # print(f"Last row: {last_row}")
        last_row = last_row.str.replace(',', '')  # Remove commas
        last_row = pd.to_numeric(last_row, errors='coerce')  # Convert to numeric, setting errors to NaN
        eco_carbon = last_row.iloc[-1]  # Get the rightmost value
        # eco_carbon = next((x for x in reversed(last_row) if x >= 0), None)
        
        list_tables.append(list_table_temp)

        # Forest Management Results
        forest_mgmt_sheet_name = "Forest%20Mgmt%20&%20HWP%20Results"

        try:
            # Fetch data from the OneDrive Excel sheet
            forest_mgmt_response = onedrive.read_excel(FILE_PATH, forest_mgmt_sheet_name, 'A4:G26')['values']
            logging.debug(f"Forest Mgmt Data: {forest_mgmt_response}")
            
            # Convert the data into a DataFrame
            forest_mgmt = pd.DataFrame(forest_mgmt_response)
            # print(forest_mgmt)

            # Extract relevant values
            carbon_seq = forest_mgmt.iloc[3, 2]
            total_hwp = forest_mgmt.iloc[20, 2]
            total_afolu = forest_mgmt.iloc[21, 2]
            benefit = forest_mgmt.iloc[22, 2]
            print(f"carbon_seq: {carbon_seq}, total_hwp: {total_hwp}, total_afolu: {total_afolu}, benefit: {benefit}")
            
            # Inspect the shape and columns of the DataFrame
            # print(forest_mgmt)  # Debugging: Check the extracted data
            # print(forest_mgmt.shape)  # Debugging: See the dimensions
            
            # Dynamically adjust the range to pick the correct rows and columns
            if forest_mgmt.shape[1] < 6:
                raise ValueError(f"Expected at least 6 columns, but got {forest_mgmt.shape[1]}")
            
            if forest_mgmt.shape[0] < 19:
                raise ValueError(f"Expected at least 19 rows, but got {forest_mgmt.shape[0]}")
            
            # Trim the DataFrame to the appropriate range
            forest_mgmt = forest_mgmt.iloc[0:23, 1:6].fillna('-')
            # print(forest_mgmt)
            
            # Append the trimmed DataFrame to the list
            list_tables.append(forest_mgmt)
            
        except Exception as e:
            logging.error(f"Error reading Forest Mgmt data: {str(e)}")
            return jsonify({"error": str(e)}), 500


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
        if isinstance(benefit, str):
            formatted_benefit = float(benefit.replace(",", "")) if float(benefit.replace(",", "")) < 0 else 0
        else:
            formatted_benefit = benefit if benefit < 0 else 0
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

def safe_round(value):
    try:
        if isinstance(value, str):
            value = float(value.replace('.', ''))
        return int(round(float(value)))
    except (ValueError, TypeError) as e:
        print(f"Failed to round value: {value}, Error: {e}")
        return value

def format_number_with_commas(x):
    if isinstance(x, (int, float)):
        return "{:,}".format(x)
    return x

@app.route('/output')
def output():
    global list_tables
    ind_list = []
    new_tables = []  # To hold the modified tables
    
    for i, table in enumerate(list_tables):
        table = list_tables[i]
        # print(f"Processing Table {i} - Columns: {table.columns.tolist()}")
        # print(f"Table columns: {table.columns.tolist()}")
        # print(f"Table first column values: {table[table.columns[0]].tolist()}")
        
        # Round numeric values in the DataFrame
        split_column = None
        for col in table.columns:
            table[col] = table[col].apply(
                lambda x: format_number_with_commas(safe_round(x)) if isinstance(x, (int, float)) or (
                    isinstance(x, str) and x.replace('-', '').replace('.', '').isdigit()) else x)
            
            if any("POST HARVEST CARBON IMPACTS" in str(val) for val in table[col]):
                split_column = col
        
        # Debug print to check column values
        # print(f"Checking column values: {table[table.columns[1]].tolist()}")

        # Split the table containing "POST HARVEST CARBON IMPACTS"
        if split_column:
            split_index = table[table[split_column] == "POST HARVEST CARBON IMPACTS"].index
            if not split_index.empty:
                split_index = split_index[0]
                print(f"Splitting Table {i} at index {split_index} based on column '{split_column}'")

                # Split the table
                table1 = table.iloc[:split_index]
                table2 = table.iloc[split_index:]

                # Reset and format table2
                table2.reset_index(drop=True, inplace=True)
                table2.loc[0] = table2.columns  # Promote column names to the first row
                table2.columns = ['Attributes', 'Year 0 post-harvest', 'By Year 100 Post-Harvest',
                                'Year 0 post-harvest', 'By Year 100 Post-Harvest']
                table2 = table2.iloc[1:]  # Drop the header row

                # Round and format numeric values in both tables
                # for col in table1.columns:
                #     table1[col] = table1[col].apply(
                #         lambda x: format_number_with_commas(safe_round(x)) if isinstance(x, (int, float)) or (
                #             isinstance(x, str) and x.replace(',', '').replace('-', '').replace('.', '').isdigit()) else x)
                
                for col in table2.columns:
                    table2[col] = table2[col].apply(
                        lambda x: format_number_with_commas(safe_round(x)) if isinstance(x, (int, float)) or (
                            isinstance(x, str) and x.replace(',', '').replace('-', '').replace('.', '').isdigit()) else x)
                
                if "ChiSquare Decay Function" in table2['Attributes'].values:
                    table2.loc[table2['Attributes'] == "ChiSquare Decay Function", :] = table2.loc[table2['Attributes'] == "ChiSquare Decay Function", :].map(
                        lambda x: '-' if x != "ChiSquare Decay Function" else x
                    )

                # Append both tables
                new_tables.append(table1)
                new_tables.append(table2)
            else:
                new_tables.append(table)
        else:
            new_tables.append(table)
    
    # Update the global list_tables with the modified list
    list_tables = new_tables
    # print(list_tables)
     # Get original column names
    temp_df = list_tables[2].copy()
    original_cols = temp_df.columns.tolist()
    
    # Create new column names, preserving original first and fourth columns
    new_cols = [
        'ECOSYSTEM CARBON IMPACTS From Forest Growth',  # First column (unchanged)
        'Value',           # Second column
        'Value',           # Third column
        'BAU scenario (only calculated for Extended Rotation activity)',  # Fourth column (unchanged)
        'Value'            # Fifth column (to be dropped)
    ]
    
    # Set new column names
    temp_df.columns = new_cols
    
    # Drop only the last column
    temp_df = temp_df.iloc[2:, [0, 1, 3]]
    
    # Update the table
    list_tables[2] = temp_df
    # print(list_tables[2])

    return render_template('output.html', list_tables=list_tables[1:], ind_list=ind_list[1:])

# @app.route('/finaloutput')
# def finaloutput():
#     return render_template('finaloutput.html', npv1 = round(npv1,2), npv2 = round(npv2,2), npv21 = round(npv21,2), npv3=round(npv3,2), npv4=round(npv4,2), npv5 = round(npv5,2), hyb=hyb, hye=hye, eco_carbon=eco_carbon, ccf=ccf, carbon_seq=carbon_seq, total_hwp=total_hwp, total_afolu=total_afolu, benefit=benefit, area=g_area)

# @app.route('/summary')
# def summary():
#     return render_template('summary.html', hyb=hyb, hye=hye, eco_carbon=eco_carbon, ccf=ccf, carbon_seq=carbon_seq, total_hwp=total_hwp, total_afolu=total_afolu, benefit=benefit, area=g_area)

@app.route('/finaloutput')
def finaloutput():
    rounded_ccf = {
        'combined': ccf['combined'],
        'a_values': [format_number_with_commas(safe_round(x)) for x in ccf['a_values']],
        'b_values': [format_number_with_commas(safe_round(x)) for x in ccf['b_values']],
        'a_minus_b_values': [format_number_with_commas(safe_round(x)) for x in ccf['a_minus_b_values']]
    }
    
    return render_template('finaloutput.html', 
        npv1=format_number_with_commas(safe_round(npv1)),
        npv2=format_number_with_commas(safe_round(npv2)),
        npv21=format_number_with_commas(safe_round(npv21)),
        npv3=format_number_with_commas(safe_round(npv3)),
        npv4=format_number_with_commas(safe_round(npv4)),
        npv5=format_number_with_commas(safe_round(npv5)),
        hyb=format_number_with_commas(safe_round(hyb)),
        hye=format_number_with_commas(safe_round(hye)),
        eco_carbon=format_number_with_commas(safe_round(eco_carbon)),
        ccf=rounded_ccf,
        carbon_seq=format_number_with_commas(safe_round(carbon_seq)),
        total_hwp=format_number_with_commas(safe_round(total_hwp)),
        total_afolu=format_number_with_commas(safe_round(total_afolu)),
        benefit=format_number_with_commas(safe_round(benefit)),
        area=format_number_with_commas(safe_round(g_area))
    )

@app.route('/summary')
def summary():
    rounded_ccf = {
        'combined': ccf['combined'],
        'a_values': [format_number_with_commas(safe_round(x)) for x in ccf['a_values']],
        'b_values': [format_number_with_commas(safe_round(x)) for x in ccf['b_values']],
        'a_minus_b_values': [format_number_with_commas(safe_round(x)) for x in ccf['a_minus_b_values']]
    }

    print(f"ccf: {ccf['b_values']}")

    # onedrive = OneDriveHelper()
    # onedrive.recalculate_workbook(FILE_PATH)
    
    return render_template('summary.html',
        hyb=format_number_with_commas(safe_round(hyb)),
        hye=format_number_with_commas(safe_round(hye)),
        eco_carbon=format_number_with_commas(safe_round(eco_carbon)),
        ccf=rounded_ccf,
        carbon_seq=format_number_with_commas(safe_round(carbon_seq)),
        total_hwp=format_number_with_commas(safe_round(total_hwp)),
        total_afolu=format_number_with_commas(safe_round(total_afolu)),
        benefit=format_number_with_commas(safe_round(benefit)),
        area=format_number_with_commas(safe_round(g_area))
    )

if __name__ == '__main__':
    app.run()
