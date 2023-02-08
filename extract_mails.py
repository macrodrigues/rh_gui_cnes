""" Script to extract information from emails and write data in a excel file"""
import win32com.client
import xlsxwriter

#outlook object
outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")

#check how many outlook accounts there are
def get_email_accounts():
    """get all the email accounts in outlook"""
    accounts = [folder.Name for folder in outlook.Folders]
    return accounts

def extract(account, subfolder):
    """Extract body text from a subfolder inside the email account"""
    messages = outlook.Folders(account).Folders(subfolder) # 2 for "inbox"
    separators = []
    for message in messages.Items:
        if ('Demande DEX:' in message.Subject)\
            or ('Demande DERC:' in message.Subject)\
            or ('Demande DARC:' in message.Subject):
            message_body = message.Body
            items= message_body.split('• ')
            cols= [item.split(':')[0] for item in items if ':' in item]
            separators.append(cols)
            break

    all_values = []
    for message in messages.Items:
        if ('Demande DEX:' in message.Subject)\
            or ('Demande DERC:' in message.Subject)\
            or ('Demande DARC:' in message.Subject):
            disp = message_body.split('.')[0].split(' ')[-1]
            vals = [item.split(': ')[1].split('\r')[0] for item in items if ':' in item]
            vals[2] = vals[2].split(' ')[0] #remove hyperlink from mail
            vals.append(disp)
            all_values.append(vals)

    return separators[0], all_values

def write(headers, values, output):
    """ This function creates an excel file and writes de data gathered from
    'extract()' """
    workbook = xlsxwriter.Workbook(output) #Create a workbook
    worksheets = [ # add worksheets for each dispositif
        workbook.add_worksheet('DERC'),
        workbook.add_worksheet('DEX'),
        workbook.add_worksheet('DARC')]
    for ws in worksheets: #write headers for each worksheets
        row = 0
        col = 0
        for header in headers:
            ws.write(row, col, header)
            col+=1
    # write values in excel
    rows = 1
    col = 0
    for item in values:
        if item[-1] == 'DERC':
            for value in item:
                if value != 'DERC': # do not pass dispositif into excel
                    worksheets[0].write(rows, col, value)
                    col+=1
            rows+=1
            col=0
        if item[-1] == 'DEX':
            for value in item:
                if value != 'DEX': # do not pass dispositif into excel
                    worksheets[1].write(rows, col, value)
                    col+=1
            rows+=1
            col=0
        if item[-1] == 'DARC':
            for value in item:
                if value != 'DARC': # do not pass dispositif into excel
                    worksheets[2].write(rows, col, value)
                    col+=1
            rows+=1
            col=0

    workbook.close()
