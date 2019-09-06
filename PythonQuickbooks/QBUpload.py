from intuitlib.client import AuthClient
from quickbooks import QuickBooks
from quickbooks.objects import *
from quickbooks.batch import *
from tkinter import filedialog
import datetime
import openpyxl

auth_client = AuthClient(
        client_id='ABD2a9rkkwbLIK1xc4yG1OpmxlbFpTjRtVQJONAcdKThNnHC4a',
        client_secret='Njox3cMmbFLovPdpIcJ5XEyP5yURFaxEoQi6pmaG',
        environment='sandbox',
        redirect_uri='https://Quickbooks-API--loganmr.repl.co/sampleappoauth2/authCodeHandler',
    )

client = QuickBooks(
         auth_client=auth_client,
         refresh_token='AB11576357896MjOYQSwRMUMDFb8wOl9sSOLrfNtYGBsYFoUXA',
         company_id='4611809164062348683',
    )

root = tk.Tk()
root.withdraw()
FNameIn = filedialog.askopenfilename(title = "Open upload sheet", defaultextension = ".xlsx", filetypes = [("Excel workbook (.xlsx)",".xlsx")])
root.destroy()
