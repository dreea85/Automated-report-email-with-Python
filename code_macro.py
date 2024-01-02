
import pandas as pd
import os
import win32com.client as win32
from datetime import date, time
import openpyxl
import traceback

import outlook

outlook_mail = win32.Dispatch("Outlook.Application")

df_associates = pd.read_excel(r"\\path\file.xlsx", sheet_name='text1', engine='openpyxl')
reviewerid_to_firstname = dict(zip(df_associates['user_id'], df_associates['first_name']))

current_date = (date.today()).strftime("%b-%d-%Y")

data = pd.read_excel(r"\\path\file.xlsx", sheet_name='text2', engine='openpyxl')

data['reviewer_timestamp'] = data['reviewer_timestamp'].astype(str).str.split().str[0]
#convert the 'reviewer_timestamp' column to a datetime object:
data['reviewer_timestamp'] = pd.to_datetime(data['reviewer_timestamp'])
#format the datetime object to the desired string format
data['reviewer_timestamp'] = data['reviewer_timestamp'].dt.strftime("%b-%d-%Y")


#retrieve last processed index, reads the txt file
def get_last_processed_index(filename= "C:\\path\last_processed.txt"):    #last_processed.txt will save the row from which it will start to generate emails once new entries are added
    try:
        if os.path.exists(filename):
            with open(filename, 'r') as file:
                return int(file.read().strip())
    except Exception as e:
        print(f"Failed to read the last processed index from {filename}: {e}")
        #return 1 to skip the header row in case of an error
        #return -1 to indicate failure
    return 0


#save last processed index into txt file
def save_last_processed_index(index, filename= "C:\\path\last_processed.txt"):
    try:
        with open(filename, 'w') as file:
            file.write(str(index))
    except Exception as e:
        print(f"Failed to save the last processed index to {filename}: {e}")


    # Construct email body
msg_body_beginning = '''
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>Hi {actual_name},</p>
    <br>
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>text here</p><p style='color:black;font-size:16px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>text here</p>
    <br>
    '''

msg_body_table1 = '''
    <table border = "1" style="border-collapse: collapse;width:70%;font-size:15px;font-family:'Times New Roman'">
        <tr style="font-weight: bold;">
            <td style="width:50%;padding: 10px; background-color: lightcoral">text here</td>
            <td style="padding: 10px; background-color: lightcoral">text here</td>
        </tr>
        <tr>
            <td style="width:60%;padding: 10px;">Date</td>
            <td style="padding: 10px;">{current_date}</td>
        </tr>
        <tr>
            <td style="width:60%;padding: 10px;">Reviewer username</td>
            <td style="padding: 10px;">{reviewer}</td>
        </tr>
        <tr>
            <td style="width:60%;padding: 10px;">Captured number</td>
            <td style="padding: 10px;">{number}</td>
        </tr>
        <tr>
            <td style="width:60%;padding: 10px;">Reviewer classification/date</td>
            <td style="padding: 10px;">{reviewer_timestamp}</td>
        </tr>
        <tr>
            <td style="width:60%;padding: 10px;">Investigation results</td>
            <td style="padding: 10px;">{Comment}</td>
        </tr>
    </table>
    '''

msg_body_table2 = '''
        </tbody>
   </table>
    '''
msg_body_end = '''
    <br>
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>
        <strong><span style="color:black;">text here</span></strong>
        <br>
        <span style="color:black;">&bull;text here</span>
        <br>   
        <span style="color:black;">&bull;text here</span>
        <br>
        <span style="color:black;">&bull;text here</span>
    </p>
    <br>
    <p><strong><span style='font-size:15px;font-family:"Times New Roman",serif;color:black;'>Thank you in advance for your understanding and cooperation!</span></strong></p>
    <br>
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'><br></p>
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>Best regards,<br></p>
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>Reporting Team</p>
    '''


def sendMailtoReviewer(number, reviewer, reviewer_timestamp, Comment, current_date):
    mail = outlook_mail.CreateItem(0)
    actual_name = reviewerid_to_firstname.get(reviewer, ' ')

    recipients = [f'{reviewer}@domain.com', 'mail1@@domain.com']
    mail.To = ';'.join(recipients)
    # mail.To = f'{reviewer}@domain.com'
    # mail.To =  reviewer + "@domain.com"
    # mail.To = "mail@domain.com"
    mail.SentOnBehalfOfName = "automation_team@domain.com"
    mail.Subject = f'Audit: {current_date}'
    mail.HTMLBody = (msg_body_beginning.format(actual_name=actual_name) +
                     msg_body_table1.format(current_date=current_date,
                                            reviewer=reviewer,
                                            number=number,
                                            reviewer_timestamp=reviewer_timestamp,
                                            Comment=Comment) +
                     msg_body_table2 +
                     msg_body_end)

    # mail.Display()
    mail.Send()


def send_error_email(reviewer, error_message):
    mail = outlook_mail.CreateItem(0)
    actual_name = reviewerid_to_firstname.get(reviewer, ' ')
    mail.To = "email@domain.com"
    mail.SentOnBehalfOfName = "automation_team@domain.com"
    mail.Subject = "Script failure"
    mail.Body = f"An error occurred in the script: \n {error_message}"
    mail.Send()

def process_new_entries(data, current_date):
    # get last processed index
    last_index = get_last_processed_index()
    new_last_index = last_index

    # print(f'Last processed index: {last_index}')   #debugging line
    # print(f'Current data length: {len(data)}')     #debugging line

    if last_index < len(data):
        for index, row in data.iloc[last_index:].iterrows():  # process rows after the last processed index
            # add a condition to check if the node is what we need
            if row['Node'] == 'text2':
                sendMailtoReviewer(row['number'], row['reviewer'], row['reviewer_timestamp'], row['Comment'],
                                       current_date)
                print(f'currently procesing index: {index}')

                new_last_index = index
        # save_last_processed_index(index + 1)
        save_last_processed_index(new_last_index + 1)

    else:
        print('No new entries to process')
        # save_last_processed_index(index + 1)


try:
    process_new_entries(data, current_date)

except Exception as e:
    error_traceback = traceback.format_exc()
    send_error_email(error_traceback)
