import os
import win32com.client
from datetime import datetime
import pandas as pd
import pyperclip

def save_attachments_from_subfolder(save_path: str, subfolder_name: str) -> None:
    # Initialize Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Access the Inbox folder
    inbox = namespace.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

    # Access the specified subfolder within the Inbox
    try:
        subfolder = inbox.Folders.Item(subfolder_name)
    except Exception as e:
        print(f"Error accessing subfolder: {e}")
        return

    # Get today's date
    today = datetime.today().date()

    # Loop through each item in the subfolder
    for item in subfolder.Items:
        # Check if the item is a mail item and if it was received today
        if item.Class == 43 and item.ReceivedTime.date() == today:
            # Loop through each attachment in the mail item
            for attachment in item.Attachments:
                attachment.SaveAsFile(os.path.join(save_path, attachment.FileName))

def account_type(budget: float,account_name: str,expenditure: float,account: str) -> str:
    if 'budgetentry' in account_name.lower().replace(' ',''):
        return 'Budget Account'
    elif (budget != 0.0) and ('budgetentry' not in account_name.lower().replace(' ','')):
        return 'Parent Account'
    elif (expenditure != 0.0) and (budget == 0.0):
        return 'Expense Account'
    else:
        if account in ['520049','520389','520485','520609','520825','521200','530005','530170','530600','540129','540165','540345','550005','560220','560226','560240']:
            return 'Parent Account'
        else:
            return 'Expense Account'

def currency_to_float(currency_str: str) -> float:
    currency_str = currency_str.replace(',','')
    if currency_str.startswith('(') and currency_str.endswith(')'):
        return -float(currency_str.replace('(', '').replace(')', '').replace('$', ''))
    else:
        return float(currency_str.replace('$', ''))

def clean_account(account_str: str) -> str:
    return account_str.replace('‬﻿﻿','').replace('﻿‭','').replace(' ','')

def clean_reports(folder: str) -> pd.DataFrame:
    hold = pd.DataFrame()
    for i in os.listdir(folder):
        f = os.path.join(folder,i)
        if not i.endswith('.xls'):
            os.remove(f)
            continue
        df = pd.read_html(f)
        period = df[0][1].at[4]

        df = df[1]
        df.columns = ['Office','Office Name','Account Tree','Account Name','GL Account','Original Budget','Current Budget','Expenditures','Committments','Obligations','Other Encumbrances','Total Expenditure Encumbrances','Expenditure Percentage','Funds Available','Fund','Fund Name','Progam','Program Name',
                      'Expenditures (MTD)','Committments (MTD)','Obligations (MTD)','Other Encumbrances (MTD)','Total Expenditure Encumbrances (MTD)']
        df = df.iloc[2:].reset_index(drop = True)
        df = df.drop(['Expenditures (MTD)','Committments (MTD)','Obligations (MTD)','Other Encumbrances (MTD)','Total Expenditure Encumbrances (MTD)'],axis=1)
        hold = pd.concat([hold,df],join='outer')
    df = hold.sort_values('GL Account').reset_index(drop=True)
    df['Period'] = [period for _ in df.index]

    df['Account Tree'] = [clean_account(acct) for acct in df['Account Tree']]

    for col in ['Original Budget','Current Budget','Expenditures','Committments','Obligations','Other Encumbrances','Total Expenditure Encumbrances','Funds Available']:
        df[col] = [currency_to_float(i) for i in df[col]]

    df['acctType'] = [account_type(bud,name,expend,num) for bud,name,expend,num in zip(df['Original Budget'],df['Account Name'],df['Expenditures'],df['Account Tree'])]
    
    return df

def export_report(dataframe: pd.DataFrame,export_path: str,hold_path: str) -> None:
    filepath = os.path.join(export_path,f'fundsAvailable{datetime.today().date()}.xlsx')
    dataframe.to_excel(filepath,index=False)

    for i in os.listdir(hold_path):
        f = os.path.join(hold_path,i)
        os.remove(f)

def main() -> None:
    # Define the save path and subfolder name
    hold_path = r'hold-reports'
    export_path = r'c:\Users\nathansmalley\OneDrive - Cook County Government\1 - Reports\FundsAvailable'
    subfolder_name = 'Funds Available'

    # Create the save directory if it doesn't exist
    os.makedirs(hold_path, exist_ok=True)

    # Call the function to save attachments from the specified subfolder
    save_attachments_from_subfolder(hold_path, subfolder_name)

    # Clean reports
    df = clean_reports(hold_path)

    # Export report
    export_report(df,export_path,hold_path)

if __name__ == '__main__':
    main()