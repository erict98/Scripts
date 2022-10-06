import win32com.client
import datetime as dt
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

##################################################################
BEGIN = dt.datetime(2022,8,1)
END = dt.datetime(2022,12,30)
WAVE = 'Wave3'
PATH = '.\\'
#PATH = 'C:\\Users\\erictran\\OneDrive - Deloitte (O365D)\\Meeting Notes - Infra\\'
DATA = 'Wave 3 Key Dates Tracker.xlsx'
OUTPUT = 'example.xlsx'
##################################################################

##################################################################
data = pd.read_excel(PATH + DATA)
gathering = 'Cloud: Server Data Gathering - '
sd = 'Cloud: Solution Design - '
testing = 'Cloud Migration: Testing Complete; Sign Off Needed | '
hypercare = ''
##################################################################

# Returns a collection of appointments from START to END
# Note: the DATETIME is based on user, need to fix to get the user's timezone
def get_Outlook(begin, end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    
    emails = outlook.getDefaultFolder(6).Items
    emails.Sort('[Start]')
    restriction = "[ReceivedTime] >= '" + begin.strftime('%m/%d/%Y') + "' AND [ReceivedTime] <= '" + end.strftime('%m/%d/%Y') + "'"
    emails = emails.Restrict(restriction)

    return calendar, emails

# Creates and returns a directory of appointments with the given title
# Works even if there are multiple apps in the message's body
# Note: Error when there are both prod and non-prod of the same app, need to explicitly have (<env>) in title
def get_appointments(appointments, title):
    dir = {}

    for appointment in appointments:
        if title in appointment.subject:
            
            if 'Canceled' in appointment.subject:
                continue
            
            if WAVE not in appointment.body:
                continue
            
            apps = set()
                    
            for line in appointment.body.split('APM')[1:]:
                try:
                    name = line.split('\r\n\r\n')[1]
                    if name[-1] == ' ':
                        name = name[:-1]
                    apps.add(name)
                except:
                    pass
            
            for name in apps:
                dir[name] = appointment
    return dir

# Creates and returns a directory of email messages with the given title
# Only reads the inital message sent and ignores any replies
def get_messages(emails, title):
    dir = {}
    
    for message in emails:
        if title in message.subject[:len(title)]:
            name = message.subject[len(title):]
            if name[-1] == ' ':
                name = name[:-1]
            dir[name] = message
    
    return dir

# Updates the XLSX file from the created directories
def update_data(sd_dir, testing_dir):
    update_sd(sd_dir)
    update_Testing(testing_dir)
    
# Updates the solution design columns
def update_sd(dir):
    error = []
    skipped = []
    
    for index in data.index:
        app = data.iloc[index]['App']
    
        if str(data.at[index, 'Solution Design Scheduled']).upper() == 'CANCELED':
            skipped.append(app)
            continue
        
        if not pd.isnull(data.at[index, 'Solution Design Scheduled']):
            continue
            
        try:
            data.at[index, 'Solution Design Scheduled'] = dir[app].Start.strftime('%m/%d/%Y') + ' ' + dir[app].Start.strftime('%H:%M')
            fw_date = dir[app].Start + dt.timedelta(days=14)
            fw_date = fw_date.strftime('%m/%d/%Y')
            data.at[index, 'Firewall Request Due'] = fw_date
            if data.at[index, 'Prod/Non-Prod'] == 'Prod':
                data.at[index, 'Change Requests Due'] = fw_date
        except:
            error.append(app)
    
    print('Missing apps for Solution Design:', error)
    print()
    print('Canceled apps for Solution Design:', error)
    print()

# Updates the testing columns
# Note: these emails do not indicate which Wave the app belongs too, be cautious with the date
def update_Testing(dir):
    error = []

    for index in data.index:
        app = data.iloc[index]['App']
        
        if pd.isnull(data.at[index, 'Testing Completed']):
            try:
                data.at[index, 'Testing Completed'] = dir[app].ReceivedTime.strftime('%m/%d/%Y')
            except:
                error.append(app)
        else:
            error.append(app)
        
    print('Testing completed missing:', error)
    print()
            
app_list = data[data.columns[0]]
appointments, messages = get_Outlook(BEGIN, END)
server_dir = get_messages(messages, gathering)
sd_dir = get_appointments(appointments, sd)
testing_dir = get_messages(messages, testing)

# Saves and exports the updated excel
def export_excel():
    workbook = openpyxl.load_workbook(PATH + DATA)
    worksheet = workbook.active
    rows = dataframe_to_rows(data, index=False)
    
    for r_idx, row in enumerate(rows, 1):
        
        for c_idx, value in enumerate(row, 1):
            worksheet.cell(row=r_idx, column=c_idx, value=value)
            
    workbook.save(PATH + OUTPUT)

if __name__ == "__main__":
    update_data(sd_dir, testing_dir)
    export_excel()