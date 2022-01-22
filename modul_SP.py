import pyodbc
import time
from datetime import datetime
import pandas as pd
# from pandas.io.data import DataReader
# from pandas import DataFrame

def Product():
    
    cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                      "Server=Server;"
                      "Database=DB;"
                      "uid=User;pwd=Pass;"
                      "Trusted_Connection=no;")


    query1="select TransID,[Mom Name],MemberID,City,Province,	Channel,	SubChannel,	TradeChannel,	OutletName,	OutletLocation, dbo.getdateonly(TransDate) as TransDate,	ProductName,	Grammage,	Qty,	TotalPoint,	dbo.getdateonly(InputDate) as InputDate,	dbo.getdateonly(ReceivedDate) as ReceivedDate, PeriodName,	PricePerUnit, Qty*IsNull(PricePerUnit,0) as TotalPrice from vl_point_trans_detail where outletname = 'Yogya group' and dbo.getdateonly(transdate) > = '2022-01-01' and dbo.getdateonly(transdate) < = '2022-02-28' and dbo.getdateonly(inputdate) <  '2022-02-28' and productname like 'S-26%' and productname not like '%nutrisure%' and periodname = 'regular program'"  
    print('Exec TransDetail')
    query2="select distinct a.Memberid, ProductName from vl_point_trans_detail a where outletname = 'yogya group' and dbo.getdateonly(transdate) > = '2022-01-01' and dbo.getdateonly(transdate) < = '2022-02-28' and dbo.getdateonly(inputdate) <  '2022-02-28' and productname not like '%nutrisure%' and productname like 'S-26%'"
    print('Exec Product')
    query3="select [Mom name] MemberName, a.MemberID,a.City, dbo.getdateonly(a.JoinDate) as JoinDate,dbo.getgoldororiginal(a.MemberId, 'Regular Program') as GoldOriorUltima, (select top 1 dbo.getdateonly(receiveddate) from  vl_point_trans_detail c where outletname = 'yogya group' and dbo.getdateonly(transdate) > = '2022-01-01' and dbo.getdateonly(transdate) < = '2022-02-28' and dbo.getdateonly(inputdate) <  '2022-02-28' and productname like 'S-26%'	 and productname not like '%nutrisure%' and c.memberid = 	a.memberid) 	FirstStrukReceived, 	Sum(TotalPoint) TotalPointEarning, 	  	Sum(Qty*Isnull(PricePerUnit,0)) TotalSpend from vl_point_trans_detail a left join _RPT_LOYALTY_POINT_BALANCE_REGULAR b on a.memberid = b.memberid  where outletname = 'Yogya group' and dbo.getdateonly(transdate) > = '2022-01-01' and dbo.getdateonly(transdate) < = '2022-02-28' and dbo.getdateonly(inputdate) <  '2022-02-28' and productname not like '%nutrisure%' and productname like 'S-26%' group by a.MemberID, a.JoinDate,[Mom name],a.City,GoldorOri, dbo._getGoldOrOriginal_cek(b.memberid, 'regular program')"  
    print('Exec TotalSpend')

    df1 = pd.read_sql_query(query1, cnxn)
    df2 = pd.read_sql_query(query2, cnxn)
    df3 = pd.read_sql_query(query3, cnxn)
    print('Get All Data')


    #print (df)


    path = r"WN S-26 LP SAB Yogya Gebyar Hebat.xlsx"
    writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
    df1.to_excel(writer, sheet_name = 'TransDetail')
    df2.to_excel(writer, sheet_name = 'Product')
    df3.to_excel(writer, sheet_name = 'TotalSpend')


    writer.save()
    print ('Save To Excel')
    writer.close()
    print ('Save Done')


  
