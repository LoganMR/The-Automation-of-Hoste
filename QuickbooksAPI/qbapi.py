from intuitlib.client import AuthClient
from quickbooks import QuickBooks
from quickbooks.objects import *
from quickbooks.batch import *
import tkinter as tk
import os
from inspect import getmembers as gm

def NewOrGetDe(name,s,client):
    try:
        q = Department.filter(Active=True, Name=name, qb=client)[0]
    except:
        De = Department()
        De.Name = name
        if s:
            q = De.save(qb=client)
    return(q)

def NewOrGetCu(name,s,client):
    try:
        q = Customer.filter(Active=True, DisplayName=name, qb=client)[0]
    except:
        Cu = Customer()
        Cu.DisplayName = name
        if s:
            q = Cu.save(qb=client)
    return(q)

def NewOrGetCl(name,s,client):
    try:
        q = Class.filter(Active=True, Name=name, qb=client)[0]
    except:
        Cl = Class()
        Cl.Name = name
        if s:
            q = Cl.save(qb=client)
    return(q)

def GetIt(name,client):
    q = Item.filter(Active=True, Name=name, qb=client)
    try:
        return(q[0])
    except:
        return(False)

def NewSR(Cus,Dat,Loc,LINES,s,client):
    SR = SalesReceipt()
    C = NewOrGetCu(Cus,True,client)
    SR.CustomerRef = C.to_ref()
    SR.TxnDate = Dat
    D = NewOrGetDe(Loc,True,client)
    REF = Ref()
    REF.value = 4
    REF.name = "Undeposited Funds"
    SR.DepositToAccountRef = REF
    SR.DepartmentRef = D.to_ref()
    for k in LINES:
        line_detail = SalesItemLineDetail()
        line_detail.ClassRef = NewOrGetCl(k['class'],True,client).to_ref()
        line_detail.ItemRef = GetIt(k['account'],client).to_ref()
        line_detail.ServiceDate = Dat
        line = SalesItemLine()
        line.Amount = k['amount']
        line.Description = k['description']
        line.SalesItemLineDetail = line_detail
        SR.Line.append(line)
    if s:
        SR = SR.save(qb=client)
    return(SR)
    


