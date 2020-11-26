from tkinter import *
import pandas as pd
import json
import datetime
from tkinter.filedialog import askopenfilename
import os
from currency_converter import CurrencyConverter
import subprocess
from sqlalchemy import create_engine
from credentials import *

# root = Tk()
# root.frame = Frame(root)
# root.title("Choose daily revenue file.")
# root.update()

filename = askopenfilename(title='Daily revenue file:')

df = pd.read_excel(filename)
start_time = datetime.datetime.now()

# # # Define the code from Campaign name # # #
df['Code'] = ''
list_of_8 = ['BET', 'FIN', 'UNI', 'WEB', 'DAT', 'PSY', 'BGC', 'POK', 'OFB']
list_of_7 = ['MD', 'PL']
list_of_5 = ['other']
test_dict = {
    "03": "CA", "04": "FI", "05": "DE", "06": "FI", "07": "SE", "08": "NO", "09": "CA", "10": "DE", "11": "DE",
    "12": "CA", "13": "DE", "14": "CA", "15": "AU", "16": "DE", "17": "NO", "18": "FI", "19": "FI", "24": "FI",
    "25": "UK", "26": "DE", "27": "GL", "29": "DE", "32": "FI", "33": "UK", "34": "NO", "35": "GL", "36": "GL",
    "37": "SE", "38": "SE", "39": "SE", "43": "NO", "47": "CA", "50": "FI", "51": "SE", "52": "FI", "57": "CA",
    "58": "CA", "64": "NO", "63": "NO", "62": "NO",
}
out_dict = {
    '01': 'CA', '02': 'DE', '04': 'DE', '05': 'AU', '06': 'FI', '09': 'DE', '10': 'DE', '11': 'NO', '15': 'DE',
    '19': 'CH', '21': 'CA', '23': 'DE', '24': 'DE', '25': 'DE', '29': 'FI', '27': 'NO', '31': 'FI', '32': 'CA',
    '35': 'DE',
    '36': 'CA', '37': 'NO', '40': 'JP', '41': 'JP', "42": "DE", "43": "CA", "45": "FI", "48": "DE", "52": "FI"
}

obt_dict = {
    '01': 'GL', '02': 'GL', '03': 'GL', '04': 'GL', '05': 'GL', '06': 'GL', '07': 'GL', '09': 'DE', '16': 'DE',
    '22': 'CA',
    '38': 'FI', '39': 'FI', '50': 'DE'
}
bng_dict = {'01': 'GL', '12': 'DE', '22': 'FI'

            }
ofb_dict = {
    '21': 'CA', '22': 'DE'
}

df['Campaign'] = df['Campaign'].str.upper()
# Define the Code
for _, code_row in df.iterrows():
    if code_row['Campaign'][:3] in list_of_8:
        df.loc[_, 'Code'] = code_row['Campaign'][:8]
    elif code_row['Campaign'][:2] in list_of_7:
        df.loc[_, 'Code'] = code_row['Campaign'][:7]
    elif code_row['Campaign'][:4] == 'TEST':
        if code_row['Campaign'][-2:].isalpha():
            df.loc[_, 'Code'] = code_row['Campaign'][:6] + '-' + code_row['Campaign'][-2:]
        else:
            test_last = test_dict.get(code_row['Campaign'][4:6])
            df.loc[_, 'Code'] = code_row['Campaign'][:6] + "-" + test_last

    elif code_row['Campaign'][:3] == 'OUT':
        if code_row['Campaign'][-2:].isalpha():
            df.loc[_, 'Code'] = code_row['Campaign'][:5] + '-' + code_row['Campaign'][-2:]
        else:
            out_last = out_dict.get(code_row['Campaign'][3:5])
            df.loc[_, 'Code'] = code_row['Campaign'][:5] + "-" + out_last

    elif code_row['Campaign'][:3] == 'OBT':
        if code_row['Campaign'][-2:].isalpha():
            df.loc[_, 'Code'] = code_row['Campaign'][:5] + '-' + code_row['Campaign'][-2:]
        else:
            obt_last = obt_dict.get(code_row['Campaign'][3:5])
            df.loc[_, 'Code'] = code_row['Campaign'][:5] + "-" + obt_last

    elif code_row['Campaign'][:3] == 'OFB':
        if code_row['Campaign'][-2:].isalpha():
            df.loc[_, 'Code'] = code_row['Campaign'][:5] + '-' + code_row['Campaign'][-2:]
        else:
            ofb_last = ofb_dict.get(code_row['Campaign'][3:5])
            df.loc[_, 'Code'] = code_row['Campaign'][:5] + "-" + ofb_last

    elif code_row['Campaign'][:4] == 'BING':
        if code_row['Campaign'][-2:].isalpha():
            df.loc[_, 'Code'] = code_row['Campaign'][:6] + '-' + code_row['Campaign'][-2:]
        else:
            bng_last = bng_dict.get(code_row['Campaign'][3:5])
            df.loc[_, 'Code'] = code_row['Campaign'][:5] + "-" + bng_last

    elif code_row['Campaign'][:3] == 'BNG':
        if code_row['Campaign'][-2:].isalpha():
            df.loc[_, 'Code'] = code_row['Campaign'][:5] + '-' + code_row['Campaign'][-2:]
        else:
            bng_last = bng_dict.get(code_row['Campaign'][3:5])
            df.loc[_, 'Code'] = code_row['Campaign'][:5] + "-" + bng_last

    elif 'BR' in code_row['Campaign'] and 'BET' in code_row['Campaign']:
        df.loc[_, 'Code'] = code_row['Campaign'][:8] + 'BR'
    elif 'BR' in code_row['Campaign'] and 'BET' not in code_row['Campaign'] and 'CA61' not in code_row['Campaign']:
        df.loc[_, 'Code'] = code_row['Campaign'][:4] + 'BR'

    elif '888' in code_row['Campaign']:
        df.loc[_, 'Code'] = code_row['Campaign'][:4]

    elif '10T' in code_row['Campaign']:
        df.loc[_, 'Code'] = code_row['Campaign'][:4]

    elif code_row['Campaign'][:5] in list_of_5:
        df.loc[_, 'Code'] = code_row['Campaign'][:5]
    elif code_row['Campaign'][2:5].isnumeric():
        df.loc[_, 'Code'] = code_row['Campaign'][:5]
    elif code_row['Campaign'][:5] == 'OTHER':
        df.loc[_, 'Code'] = "OTHER"
    else:
        df.loc[_, 'Code'] = code_row['Campaign'][:4]

# TESTBET29-DE
for _, test_row in df.iterrows():
    if test_row['Code'] == "TEST29-DE":
        df.loc[_, 'Code'] = "TESTBET29-DE"
    elif test_row['Code'][:6] == "TEST36":
        df.loc[_, 'Code'] = "TEST36-GL"
    elif test_row['Code'][:8] == "OBT04-DE":
        df.loc[_, 'Code'] = "OBT04-GL"
    elif test_row['Code'][:5] == "OUT10":
        df.loc[_, 'Code'] = "DE73"

# # # Define Country # # #
df['Country'] = ''
for _, state_row in df.iterrows():
    if state_row['Code'][:3] in list_of_8:
        df.loc[_, 'Country'] = state_row['Code'][4:8]
    elif state_row['Code'][:3] == 'OUT':
        df.loc[_, 'Country'] = state_row['Code'][-2:]
    elif state_row['Code'][:3] == 'OBT':
        df.loc[_, 'Country'] = state_row['Code'][-2:]
    elif state_row['Code'][:3] == 'OFB':
        df.loc[_, 'Country'] = state_row['Code'][-2:]
    elif state_row['Code'][:2] in list_of_7:
        df.loc[_, 'Country'] = state_row['Code'][3:7]
    elif state_row['Code'][:5] == "OTHER":
        df.loc[_, 'Country'] = 'Other'
    elif state_row['Code'][:4] == "TEST":
        df.loc[_, 'Country'] = state_row['Code'][-2:]
    elif state_row['Code'][:4] == "DC01":
        df.loc[_, 'Country'] = state_row['Campaign'][-2:]
    elif state_row['Code'][:4] == "BING":
        df.loc[_, 'Country'] = state_row['Campaign'][-2:]
    elif state_row['Code'][:3] == "BNG":
        df.loc[_, 'Country'] = state_row['Code'][-2:]
    else:
        df.loc[_, 'Country'] = state_row['Code'][:4]

# Some countries don't show country code in front of the Campaign, instead in the last strings.
global_countries = ['GW', 'MB', 'GL', 'DH', 'IE', 'SC', 'DC']
for _, glob_row in df.iterrows():
    if glob_row['Country'][:2] in global_countries:
        df.loc[_, 'Country'] = glob_row['Campaign'][-2:]


# Define Countries with country code.
def load_schema(path):
    with open(path) as f:
        return json.load(f)


def get_country(key):
    countries_list = load_schema('countries.json')
    countries = {i['alpha-2']: i['name'] for i in countries_list}
    return countries.get(key)


custom_countries = {'GL': 'GLOBAL'}
for _, country_row in df.iterrows():
    if country_row['Code'] == 'TEST36-GL':
        df.loc[_, 'Country'] = 'Canada'
    elif country_row['Country'][:5] == "Other":
        df.loc[_, 'Country'] = 'Other'
    elif country_row['Country'] in custom_countries:
        df.at[_, 'Country'] = custom_countries.get(country_row['Country'])
    else:
        country = get_country(country_row['Country'][:2])
        df.at[_, 'Country'] = country

df['Country'] = df['Country'].fillna('Other')

# # # Region # # #
df['Region'] = ''
region_scandy = ['Finland', 'Norway', 'Sweden']
region_dach = ['Austria', 'Germany', 'Switzerland']

for _, region_row in df.iterrows():
    if region_row['Country'] in region_scandy:
        df.loc[_, 'Region'] = 'SCANDY'
    elif region_row['Country'] in region_dach:
        df.loc[_, 'Region'] = 'DACH'
    elif region_row['Code'] == 'OBT04-DE':
        df.loc[_, 'Code'] = 'OBT04-GL'
    elif region_row['Code'] == 'OBT03-CA':
        df.loc[_, 'Code'] = 'OBT03-GL'
    elif 'BR' in region_row['Campaign'] and 'IT21' in region_row['Campaign'] and 'BET' in region_row['Campaign']:
        df.loc[_, 'Code'] = 'BET-IT21BR'
        df.loc[_, 'Region'] = 'OTHERS'
    else:
        df.loc[_, 'Region'] = "OTHERS"

# # # Verticals # # #
df['Vertical'] = ''
verticals = {
    'BET': "BETTING", 'FIN': 'FINANCE', 'WEB': 'WEB DEVELOPMENT', 'DAT': 'DATING', 'PSY': 'PSYCHIC READINGS',
    'MD': 'MEAL DELIVERY', 'POP': 'POPUP', 'PL': 'PERSONAL LOANS', 'DK01': 'WHITE CASINO', 'DK02': 'WHITE CASINO',
    'NJ04': 'WHITE CASINO', 'UK01': 'WHITE CASINO'
}

for _, vertical_row in df.iterrows():
    if vertical_row['Code'][:3] in list_of_8:
        df.loc[_, 'Vertical'] = verticals.get(vertical_row['Code'][:3])
    elif vertical_row['Code'][:2] in list_of_7:
        df.loc[_, 'Vertical'] = verticals.get(vertical_row['Code'][:2])
    elif vertical_row['Code'][:4] == 'DK02':
        df.at[_, 'Vertical'] = 'WHITE CASINO'
    elif vertical_row['Code'][:4] == 'DK01':
        df.at[_, 'Vertical'] = 'WHITE CASINO'
    elif vertical_row['Code'][:4] == 'NJ04':
        df.at[_, 'Vertical'] = 'WHITE CASINO'
    elif vertical_row['Code'][:4] == 'UK01':
        df.at[_, 'Vertical'] = 'WHITE CASINO'
        df.at[_, 'Country'] = 'UK'
    elif vertical_row['Code'][:6] == 'TEST36':
        df.at[_, 'Vertical'] = 'BETTING'
    elif vertical_row['Code'] == 'TESTBET29-DE':
        df.at[_, 'Vertical'] = 'BETTING'
    else:
        df.at[_, 'Vertical'] = 'GREY CASINO'

# # # Traffic Channel # # #
bing = ["AT37", "AU14", "AU17", "BET-AT06", "BET-CA13", "BET-DE21", "DE96", "IT21", "NJ02", "BET-AU03", "BET-CH04",
        "BET-DE24", "BET-NL02", "BET-NZ05", "DE97", "IE01", "IE06", "IE08", "BET-FI09", "BET-IT08", "BET-NO03",
        "BET-SE18", "SE64", "IE13", "IE16", "IT19", "NL16", "NZ20", "CA61", "DE102", "AU25", "ES02", "FR02", "JP02",
        "IN02", "BNG12-DE", "NJ002", "CA84", "CA86", "BET-IT21BR", "IT21BR", "DE12BR", "FI01BR", "DE12", "AT01BR", "DE102BR", "DK03", "NL16BR", "SE01BR", "FR02BR", ]
facebook = ['SE01', 'DE75']

df['Traffic Channel'] = ''
for _, traffic_row_channel in df.iterrows():
    if traffic_row_channel['Code'] in bing:
        df.at[_, 'Traffic Channel'] = 'Bing'
    elif traffic_row_channel['Code'] in facebook:
        df.at[_, 'Traffic Channel'] = 'Facebook'
    else:
        df.at[_, 'Traffic Channel'] = 'Google Ads'

df['POPUP'] = 'STANDARD'
for _, popup in df.iterrows():
    if 'POP' in popup['Campaign']:
        df.at[_, 'POPUP'] = 'POPUP'

# # # Revenue (LC) # # #
df['Revenue (LC)'] = 0
for _, revenue_lc in df.iterrows():
    if revenue_lc['QP'] > revenue_lc['Registers'] and revenue_lc['Type'] == 'CPA':
        minus = revenue_lc['QP'] - revenue_lc['Registers']
        df.loc[_, 'Revenue (LC)'] = revenue_lc['Revenue'] / revenue_lc['QP'] * minus
        # df['Revenue (LC)'] = df['Revenue'] / df['QP'] * minus

# c_lc = CurrencyConverter()
c_lc = CurrencyConverter('http://www.ecb.europa.eu/stats/eurofxref/eurofxref.zip')

for _, conver in df.iterrows():
    df.loc[_, 'Revenue'] = c_lc.convert(conver['Revenue'], conver['CCY'])
    df.loc[_, 'Revenue (LC)'] = c_lc.convert(conver['Revenue (LC)'], conver['CCY'])

# # # Adwords Expense # # #
df['AdwordsExpense'] = 0
# # # Traffic # # #
df['Traffic'] = 0
# # # POD # # #
df['POD'] = ''
philipp_accounts = ["NL16", "IE06", "DE96", "CA61", "DE102", "IE01", "IE16", "AU14", "DE97", "DE82",
                    "BET-CA13", "AU25", "BET-DE21", "IE08", "BET-DE24", "IT19", "BET-NL02",
                    "BET-AT06", "NL32", "BET-NZ05", "AU17", "SC07", "OUT42-DE", "OUT43-CA", "OBT07-GL", "FR02",
                    "JP02", "IN02", "OBT03-GL"]

farid_accounts = ["TEST50-FI", "TEST35-GL", "TEST58-CA", "DE84", "FI79", "TEST57-CA", "OUT13-CA", "TEST63-NO",
                  "TEST60-CA", "TEST61-NO", "TEST54-DE", "TEST66-FI", "OUT07-FI", "DE26", "DE73", "BET-DE36",
                  "OUT23-DE", "OUT26-NO", "OUT15-DE", "DE123", "OBT16-DE", "OBT04-GL", "IE13", "OBT01-DE", "AT37",
                  "ES02", "NZ20", "BNG12-DE", "DE12", "DE12BR", "AT36"]

islam_accounts = ["OUT02-DE", "SC12", "AU17", "OUT04-DE", "BET-NO13", "BET-DE30", "TEST32-CA", "OUT06-FI", "BET-DE35",
                  "TEST64-NO", "BET-FI06", "BET-CA23", "DC01", "TEST62-NO", "JP04", "BET-DE43", "CA20", "BET-FI16",
                  "BET-FI17", "OUT08-NO", "IT05", "OUT05-AU", "DE121", "CA76", "CA74", "FI83"]
emir_accounts = ["DK02", "DK01", "BET-DK02", "BET-DK01", "TEST36-GL", "NJ04", "DK03", "PA01"]
outsource_accounts = ["NZ17", "DE75", "DE133", "OBT50-DE"]
fakhri_accounts = ['OUT52-NO']
turgut_accounts = ['NJ02', 'NJ002']
pasha_accounts = ["IT21", "CA84", "CA86", "BET-IT21BR", "IT21BR", "FI01BR", "AT01BR", "DE102BR", "NL16BR", "FR02BR", "SE01BR"]

for _, pod_acc in df.iterrows():
    if pod_acc['Code'] in philipp_accounts:
        df.loc[_, 'POD'] = 'Philipp'
    elif pod_acc['Code'] in farid_accounts:
        df.loc[_, 'POD'] = 'Farid M'
    elif pod_acc['Code'] in islam_accounts:
        df.loc[_, 'POD'] = 'Islam'
    elif pod_acc['Code'] in emir_accounts:
        df.loc[_, 'POD'] = 'Emir'
    elif pod_acc['Code'] in outsource_accounts:
        df.loc[_, 'POD'] = 'Outsource'
    elif pod_acc['Code'] in fakhri_accounts:
        df.loc[_, 'POD'] = 'Fakhri'
    elif pod_acc['Code'] in turgut_accounts:
        df.loc[_, 'POD'] = 'Turgut'
    elif pod_acc['Code'] in pasha_accounts:
        df.loc[_, 'POD'] = 'Pasha'
    else:
        df.loc[_, 'POD'] = 'Other'

# Revenue Share
df['RS'] = 0
for _, revshare in df.iterrows():
    if revshare['Type'] == 'RS':
        df.loc[_, 'RS'] = revshare['Revenue']
    elif revshare['Type'] == 'CPA+RS':
        camp_num = 0
        for i in revshare['Campaign'][-7:]:
            if i.isnumeric():
                camp_num += 1
        if camp_num > 4:
            if revshare['Campaign'][-2:].isnumeric():
                deal_num = revshare['Campaign'][-5:]
                if str(deal_num[0]) == 0:
                    df.loc[_, 'RS'] = revshare['Revenue'] - (float(deal_num[1:3]) * revshare['QP'])
                else:
                    df.loc[_, 'RS'] = revshare['Revenue'] - (float(deal_num[0:3]) * revshare['QP'])
            elif revshare['Campaign'][-2:].isalpha():
                deal_num = revshare['Campaign'][-7:-2]
                if str(deal_num[0]) == 0:
                    df.loc[_, 'RS'] = revshare['Revenue'] - (float(deal_num[1:3]) * revshare['QP'])
                else:
                    df.loc[_, 'RS'] = revshare['Revenue'] - (float(deal_num[0:3]) * revshare['QP'])
        else:
            df.loc[_, 'RS'] = revshare['Revenue']

# df['Revenue Share'] = df['RS']
df = df.sort_values(by='Code')

# Reorder columns.
df = df[['Date', 'Partner', 'Code', 'Country', 'Vertical', 'POD', 'Revenue', 'Revenue (LC)', 'AdwordsExpense',
         'Traffic', 'Visits', 'Registers', 'FTD', 'QP', 'Traffic Channel', 'Region', 'POPUP', 'Campaign', 'RS']]

# Change column names for mysql
df.columns = ['date', 'partner', 'code', 'country', 'vertical', 'pod', 'revenue', 'revenuelc', 'adwordsExpense',
              'traffic',
              'visits', 'registers', 'ftd', 'qp', 'trafficchannel', 'region', 'popup', 'campaign', 'revenueshare']

# Write df to mysql... Both server and local.
df_wait = df.copy()
# Checking file here for accurateness ...

# Leave unnecessary columns out of df for by code report.
df_un = df[['date', 'code', 'country', 'vertical', 'pod', 'revenue', 'revenuelc', 'adwordsExpense', 'traffic',
            'visits', 'registers', 'ftd', 'qp', 'trafficchannel', 'region', 'popup', 'campaign', 'revenueshare']]

# Group by Unique Date and Code.
df_code = df_un.groupby(['date', 'code']).agg(
    country=pd.NamedAgg(column='country', aggfunc=min),
    vertical=pd.NamedAgg(column='vertical', aggfunc=min),
    pod=pd.NamedAgg(column='pod', aggfunc=min),
    revenue=pd.NamedAgg(column='revenue', aggfunc=sum),
    revenuelc=pd.NamedAgg(column='revenuelc', aggfunc=sum),
    adwords=pd.NamedAgg(column='adwordsExpense', aggfunc=sum),
    traffic=pd.NamedAgg(column='traffic', aggfunc=sum),
    visits=pd.NamedAgg(column='visits', aggfunc=sum),
    registers=pd.NamedAgg(column='registers', aggfunc=sum),
    ftd=pd.NamedAgg(column='ftd', aggfunc=sum),
    qp=pd.NamedAgg(column='qp', aggfunc=sum),
    trafficchannel=pd.NamedAgg(column='trafficchannel', aggfunc=min),
    region=pd.NamedAgg(column='region', aggfunc=min),
    popup=pd.NamedAgg(column='popup', aggfunc=min),
    revenueshare=pd.NamedAgg(column='revenueshare', aggfunc=sum), ).reset_index()

# Change column names for mysql "by Code"
df_code.columns = ['date', 'code', 'country', 'vertical', 'pod', 'revenue', 'revenuelc', 'adwordsExpense', 'traffic',
                   'visits', 'registers', 'ftd', 'qp', 'trafficchannel', 'region', 'popup', 'revenueshare']

# df['date'] = datetime.datetime.strftime(df['date'].iloc[0].date(), '%m/%d/%Y')
df['date'] = df['date'].dt.strftime('%m/%d/%Y')
# df['Date'] = datetime.datetime.strftime(df['Date'].iloc[0].date(), '%d %B %Y')
df = df.set_index('date')

end_time = datetime.datetime.now()
save_path = '//Users/admin/OneDrive - Ante Technologies/Desktop/Daily Revenues/MTD/2020/April'
name_of_mtd = input("Name your MTD file: ")
complete_name = os.path.join(save_path, name_of_mtd + ".xlsx")
df.to_excel(complete_name)

elapsed_time = end_time - start_time
print(end_time - start_time)

exc = str(input('Do you want to open the excel file? yes/no ... '))
if exc == 'Yes' or exc == 'yes':
    subprocess.run(['open', complete_name], check=True)
else:
    pass
    
# root.destroy()

allowing = str(input('Do you want to write file to MySQL Server? yes/no ... '))
if allowing == 'Yes' or allowing == 'yes':
    # sshtunnel.SSH_TIMEOUT = 5.0
    # sshtunnel.TUNNEL_TIMEOUT = 5.0
    try:
        # with sshtunnel.SSHTunnelForwarder(
        #         ('94.20.248.56', 22),
        #         ssh_username='root',
        #         ssh_password='Extr@w3b!',
        #         remote_bind_address=(
        #         '127.0.0.1', 3306)
        # ) as tunnel:
        #     port = tunnel.local_bind_port
            engine = create_engine(
                'mysql+mysqlconnector://' + user + ':' + password + '@' + host + ':3306/mtd_db')
            connection = engine.connect()
    except Exception as e:
        print(e)
    df_wait.to_sql(name='mtd_revenue', con=connection, schema='mtd_db', if_exists='append', index=False)
    print("mtd_table is written successfully!")
    # Write "by Code" df to mysql.
    # df_code.to_sql(name='mtd_code', con=connection1, schema='mtd_report', if_exists='append', index=False)
    # df_code.to_sql(name='mtd_code', con=connection2, schema='mtd_report', if_exists='append', index=False)
    # print("File successfully written to MySQL Server.")
else:
    print("File is not written to MySQL Server.")

sys.exit()
