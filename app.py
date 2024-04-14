from flask import Flask, request, send_file
import pandas as pd
import re

app = Flask(__name__)

def convert_string_num(num):
    num = str(num)
    num = num.replace(',', '')
    num = num.replace('nan', '0')
    num = num.replace('NaN', '0')
    return int(num)

def execute_script(df_issue_file, df_stock_bal_file, startDate, endDate, download_sheet_name):
    # Read Excel files
    df_issue = pd.read_excel(df_issue_file)
    df_stockBal = pd.read_excel(df_stock_bal_file)
    month_start = startDate
    month_end = endDate

    # Your existing script here
    # Helper Functions
    # 1A. Convert String Numbers 
    def convertStringNum (num):
        num = str(num)
        num = num.replace(',', '')
        num = num.replace('nan', '0')
        num = num.replace('NaN', '0')
        return int(num)
    
    # 2. Compile Data (Cumulative) based on Item
    drugs = {}
    for i in range(len(df_issue['Item Code'])):
        try:
            ref_drug = df_issue['Item Description'].iloc[i]
            ref_qty = convertStringNum(df_issue['Quantity Issued'].iloc[i])

            if drugs.get(ref_drug) is not None:
                drugs[ref_drug] = drugs[ref_drug] + ref_qty
            else:
                drugs[ref_drug] = ref_qty
        except:
            print(ref_drug)
            
    # 3. Create Data Frame - Simplify Table
    df = pd.DataFrame()
    df['Item Name'] = drugs.keys()
    df['Usage'] = drugs.values()
    df['Item Code'] = df.shape[0]*''
    df['Purchase Type'] = df.shape[0]*''
    df['Calculated Buffer'] = df.shape[0]*''
    df['Stock Balance'] = df.shape[0]*''
    df['Top up Max'] = df.shape[0]*''
    df['Top up Buffer'] = df.shape[0]*''
    df['Purchase Status'] = df.shape[0]*''

    # 4. Remove last 2 rows 
    df = df.iloc[:-2]   

    

# 5a. Match Data - Item Code
    for x in range(len(df['Item Name'])):
        drug_name = df['Item Name'].iloc[x]
        for y in range(len(df_issue['Item Description'])):
            ref_name = df_issue['Item Description'].iloc[y]
            ref_code = df_issue['Item Code'].iloc[y]
            
            if drug_name == ref_name:
                df['Item Code'].iloc[x] = str(ref_code)
                break
            else:
                pass
            
    # 5b. Regex Code - Purcahse Type 
    code_pattern = re.compile(r'^([D]?\d{2})\.\d{4}\.\d{2}$')
    purchase_type = ['APPL' if re.match(code_pattern, item) else 'LP/Contract' for item in df['Item Code']]
    df['Purchase Type'] = purchase_type

    # 5c. Compute Calc Buffer 
    if (int(month_end.split('/')[-1]) - int(month_start.split('/')[-1]) == 0):
        time_frame = int(month_end.split('/')[1]) - int(month_start.split('/')[1])
    else:
        time_frame = 12 + int(month_end.split('/')[1]) - int(month_start.split('/')[1])
        
    calc_buffer = [round(2 * qty / time_frame) for qty in df['Usage']]
    df['Calculated Buffer'] = calc_buffer

    # 5d. Match Stock Balance
    for x in range(len(df['Item Name'])):
        drug_name = df['Item Name'].iloc[x]
        drug_code = df['Item Code'].iloc[x]
        cur_bal = 0
        for y in range(len(df_stockBal['Item Description'])):
            ref_name = df_stockBal['Item Description'].iloc[y]
            ref_code = df_stockBal['Item Code'].iloc[y]
            ref_bal = df_stockBal['Total Stock (SKU)'].iloc[y]
            ref_bal = convertStringNum(ref_bal)
            
            if drug_name == ref_name or drug_code == ref_code:
                cur_bal = cur_bal + ref_bal
            
        df['Stock Balance'].iloc[x] = cur_bal

    # 5e. Compute Amount to Top Up & Purchase Status
    for x in range(len(df['Item Name'])):
        bal = df['Stock Balance'].iloc[x]
        buffer = df['Calculated Buffer'].iloc[x]
    
        try:
            maxQty = buffer * 1.5 
            to_max = maxQty - bal
            to_buffer = buffer - bal
            df['Top up Max'].iloc[x] = to_max
            df['Top up Buffer'].iloc[x] = to_buffer
            
            if (bal <= (buffer * 0.75)):
                df['Purchase Status'].iloc[x] = 'Alert'
                
        except:
            print('Check problem with: {}'.format(df['Item Name'].iloc[x]))
        
    # 5f. Sort Dataframe
    df = df.sort_values(by = ['Purchase Type', 'Item Name'], ascending=[True, True])


    # Example of saving the final_df to a file
    final_df = pd.DataFrame(df)
    filename = download_sheet_name + '.xlsx'
    final_df.to_excel(filename, sheet_name="Sheet1", index=False)
    return filename

@app.route('/')
def index():
    return open('index.html').read()

@app.route('/submit', methods=['POST'])
def submit():
    df_issue_file = request.files['dfIssue']
    df_stock_bal_file = request.files['dfStockBal']
    start_date = request.form['startDate']
    end_date = request.form['endDate']
    download_sheet_name = request.form['downloadSheetName']
    
    filename = execute_script(df_issue_file, df_stock_bal_file, start_date, end_date, download_sheet_name)
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
