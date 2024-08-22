import os
import win32com.client
from datetime import datetime,timedelta
import pandas as pd
import pyperclip

def initialize_outlook(subfolder_name: str) -> tuple[object | None, object | None]:
    '''Initialize Outlook and access the specified subfolder and Deleted Items folder.'''
    outlook = win32com.client.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')

    # Access the Inbox folder
    inbox = namespace.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

    # Access the specified subfolder within the Inbox
    try:
        subfolder = inbox.Folders.Item(subfolder_name)
    except Exception as e:
        print(f'Error accessing subfolder: {e}')
        return None, None
    
    # Access the Deleted Items folder
    deleted_items = namespace.GetDefaultFolder(3)  # 3 corresponds to the Deleted Items folder

    return subfolder, deleted_items

def process_emails(subfolder: object, report_one_time: str, report_two_time: str) -> list[object]:
    '''Process emails and determine which to keep and delete.'''
    today = datetime.today().date()
    # Convert report times to offset-naive datetime
    report_one_time = datetime.combine(today, datetime.strptime(report_one_time, '%H:%M').time())
    report_two_time = datetime.combine(today, datetime.strptime(report_two_time, '%H:%M').time())

    emails_to_delete = []
    report_one_email = None
    report_two_email = None

    for item in subfolder.Items:
        # Ensure the item is a MailItem
        if item.Class == 43:  # 43 corresponds to a MailItem
            item.UnRead = False  # Mark email as read
            received_time = item.ReceivedTime
            received_date = received_time.date()

            # Convert received_time to offset-naive datetime if necessary
            if received_time.tzinfo is not None:
                received_time = received_time.replace(tzinfo=None)

            if received_date == today:
                if received_time < report_one_time:
                    if report_one_email is None or received_time > report_one_email.ReceivedTime.replace(tzinfo=None):
                        if report_one_email:
                            emails_to_delete.append(report_one_email)
                        report_one_email = item
                    else:
                        emails_to_delete.append(item)
                elif report_one_time <= received_time <= report_two_time:
                    if report_two_email is None or received_time > report_two_email.ReceivedTime.replace(tzinfo=None):
                        if report_two_email:
                            emails_to_delete.append(report_two_email)
                        report_two_email = item
                    else:
                        emails_to_delete.append(item)
                else:
                    emails_to_delete.append(item)

    return emails_to_delete

def move_emails_to_deleted(emails_to_delete: list[object], deleted_items: object) -> None:
    '''Move the specified emails to the Deleted Items folder.'''
    for email in emails_to_delete:
        email.Move(deleted_items)

def clean_outlook_folder(subfolder_name: str, report_one_time: str, report_two_time: str) -> None:
    '''Clean the Outlook folder by keeping specific emails and deleting others.'''
    subfolder, deleted_items = initialize_outlook(subfolder_name)
    if subfolder is None or deleted_items is None:
        return
    
    emails_to_delete = process_emails(subfolder, report_one_time, report_two_time)
    move_emails_to_deleted(emails_to_delete, deleted_items)

def save_attachments_from_subfolder(save_path: str, subfolder_name: str) -> None:
    subfolder,_ = initialize_outlook(subfolder_name)

    # Get today's date
    today = datetime.today().date()

    # Loop through each item in the subfolder
    for item in subfolder.Items:
        # Check if the item is a mail item and if it was received today
        if item.Class == 43 and item.ReceivedTime.date() == today:
            # Loop through each attachment in the mail item
            for attachment in item.Attachments:
                attachment.SaveAsFile(os.path.join(save_path, attachment.FileName))

def get_account_type(budget: float,account_name: str,expenditure: float,account: str) -> str:
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

def match_accounts(account_types: pd.Series,account_numbers: pd.Series) -> list:
    matched_accounts = []
    last_parent = ''
    for typ,num in zip(account_types,account_numbers):
        if typ == 'Parent Account':
            last_parent = num
        matched_accounts.append(last_parent)
    
    return matched_accounts


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

    df['acctType'] = [get_account_type(bud,name,expend,num) for bud,name,expend,num in zip(df['Original Budget'],df['Account Name'],df['Expenditures'],df['Account Tree'])]

    df['parentAcct'] = match_accounts(df['acctType'],df['GL Account'])
    
    return df

def export_report(dataframe: pd.DataFrame,export_path: str,hold_path: str) -> None:
    filepath = os.path.join(export_path,f'fundsAvailable{datetime.today().date()}.xlsx')
    dataframe.to_excel(filepath,index=False)

    for i in os.listdir(hold_path):
        f = os.path.join(hold_path,i)
        os.remove(f)

def main() -> None:
    # Define the save path and subfolder name
    hold_path = r'C:\Users\nathansmalley\OneDrive - Cook County Government\2 - Coding\funds-available-reporting\hold-reports'
    export_path = r'c:\Users\nathansmalley\OneDrive - Cook County Government\1 - Reports\FundsAvailable'
    subfolder_name = 'Funds Available'

    # Create the save directory if it doesn't exist
    os.makedirs(hold_path, exist_ok=True)

    # clean exra reports from folder
    clean_outlook_folder(subfolder_name, '04:25', '04:45')

    # Call the function to save attachments from the specified subfolder
    save_attachments_from_subfolder(hold_path, subfolder_name)

    # Clean reports
    df = clean_reports(hold_path)

    # Export report
    export_report(df,export_path,hold_path)

if __name__ == '__main__':
    print('COMPILING FUNDS AVAILABLE')
    main()
    print(' Success')