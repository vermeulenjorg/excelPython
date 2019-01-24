import win32com
from win32com.client import Dispatch
import os

def returnString(name):
    printable = 'Hi ' + name
    print(printable)
    return printable


def openWorkbook(file):
    # Start an instance of Excel
    xlapp = win32com.client.DispatchEx("Excel.Application")
    # Open the workbook in said instance of Excel
    wb = xlapp.workbooks.open(file)
    # Optional, e.g. if you want to debug
    # xlapp.Visible = debug
    # run excel on background
    xlapp.DisplayAlerts = False










    def convertToXlsx2(file, debug, fileName):





        # wb.RefreshAll()

        # wb.saveAs(fileName)
        wb.SaveAs(os.path.abspath(r'C:\Repositories\kpn_zm_financial_forecasting\python_simple\helpers\email\output.xlsx'),
            FileFormat="51")
        wb.Save()
        # Quit
        xlapp.Quit()

    # def convertToDate(dateString):
    #     date_time_obj = date.datetime.strptime(dateString, '%d-%m-%Y').date()
    #     return date_time_obj
    tempLocation = os.getcwd() + '\\temp.xlsb'
    # Open Outlook
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    # Get Inbox for Financial Forecasting
    inbox = outlook.folders('financial forecasting').Folders('Postvak IN')
    donebox = outlook.folders('financial forecasting').Folders("Done")
    data = pd.DataFrame()
    # val_date = date.date.today()
    for msg in inbox.Items:
        if msg.Sender.GetExchangeUser().PrimarySmtpAddress == 'hans.wijnbergen@kpn.com':
            for att in msg.Attachments:
                if (re.search('xlsb', att.FileName)):
                    att.SaveAsFile(tempLocation)

                    # ex.convertToXlsx(tempLocation, True , os.getcwd() + '\\temp.xlsx')
                    # data = pd.read_excel(os.getcwd() + '\\temp.xlsx', 'Sheet1', header=0, encoding='utf-8',engine='xlrd')
                    # data = ex.readExcel(tempLocation,'Sheet1',0,False)

    # print(os.path.abspath(tempLocation))
    convertToXlsx2(tempLocation, True, "\output.xlsx")

    #