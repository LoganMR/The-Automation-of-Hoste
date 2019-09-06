from requests.auth import HTTPBasicAuth
from tkcalendar import DateEntry
import tkinter as tk
import json
import requests
import datetime
import openpyxl
import calendar

# Create calendar interface

def dateselector():


    def s():
        global G
        G = [e1.get_date(),e3.get_date(),e2.get_date(),e4.get_date()]
        master.destroy()

    master = tk.Tk()
    tk.Label(master, text="Check-In Dates").grid(row=0,column=0)
    tk.Label(master, text="Check-Out Dates").grid(row=1,column=0)

    e1 = DateEntry(master,width=12,year=2019,month=8,day=1,background='darkblue', foreground='white', borderwidth=2)
    e2 = DateEntry(master,width=12,year=2019,month=8,day=14,background='darkblue', foreground='white', borderwidth=2)
    e3 = DateEntry(master,width=12,year=2019,month=8,day=31,background='darkblue', foreground='white', borderwidth=2)
    e4 = DateEntry(master,width=12,year=2019,month=8,day=31,background='darkblue', foreground='white', borderwidth=2)

    e1.grid(row=0, column=1)
    e2.grid(row=1, column=1)
    e3.grid(row=0, column=2)
    e4.grid(row=1, column=2)

    tk.Button(master, text='Get Reservations', command= s).grid(row=3, column=0, sticky=tk.W, pady=4)

    tk.mainloop()
    return(G)

# Access the guesty api

headers = {
  'Content-Type': "application/json",
  'cache-control': "no-cache",
  'Postman-Token': % Omitted
  }

key = % Omitted
secret = '% Omitted
url = 'https://api.guesty.com/api/v2/reservations'

def getRes(fil):
    skp = 0
    sng = {'limit' : 100, 'skip' : skp, 'filters' : fil, 'fields' : 'checkIn checkOut confirmationCode listing.nickname money guest.fullName integration.platform confirmedAt nightsCount source'}
    r = reses = requests.request("GET", url, headers=headers, params=sng, auth=HTTPBasicAuth(key, secret)).json()['results']
    while r != []:
        skp = skp + 100
        sng = {'limit' : 100, 'skip' : skp, 'filters' : fil, 'fields' : 'checkIn checkOut confirmationCode listing.nickname money guest.fullName integration.platform confirmedAt nightsCount source'}
        r = requests.request("GET", url, headers=headers, params=sng, auth=HTTPBasicAuth(key, secret)).json()['results']
        reses = reses + r
    return(reses)


def resdats(date1,date2,date3,date4,con): #YYYY-MM-DD format, con = 'confirmed' or 'cancelled'
    try:
        for k in range(1,5):
            datetime.datetime.strptime(eval('date'+str(k)), '%Y-%m-%d')
    except:
        return('Dates must have YYYY-MM-DD format.')
    fil = '[{"field": "checkInDateLocalized","operator": "$gte","value": "' + date1 + '"},{"field": "checkInDateLocalized","operator": "$lte","value": "' + date2 + '"},{"field": "checkOutDateLocalized","operator": "$gte","value": "' + date3 + '"},{"field": "checkOutDateLocalized","operator": "$lte","value": "' + date4 + '"},{"field": "status","operator": "$eq","value": "' + con + '"}]'
    return(getRes(fil))
