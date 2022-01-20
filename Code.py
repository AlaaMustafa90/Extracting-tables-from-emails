python

import win32com.client as wc
import os
import pandas as pd

#Assigning diroctories to variables
msg_location = r'C:\Downloads\Mails' # the place of the saved emails
output_location = r'C:\Downloads\Output' # Where we will save the CSV files

#Getting list of all files in msg_location
files = os.listdir(msg_location)

for file in files:
    if file.endswith('.msg'):
        outlook = wc.Dispatch('Outlook.Application').GetNamespace('MAPI')
        msg = outlook.OpenSharedItem(msg_location + '/' + file)
        html_str = msg.HTMLBody
        try:
            pd.read_html(html_str)[1].to_csv(output_location + '\\' + file[:-4] + '.csv', index=False)
        except ValueError:
            continue