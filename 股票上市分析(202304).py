#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Apr  4 17:22:07 2023

@author: newshuang
"""
def stock_time_aGroup(stock_id):
    
    open_year = df_a["民國年"][df_a["證券代號"] == stock_id]
    open_year = open_year.iloc[0]

    open_month = df_a["西元月"][df_a["證券代號"] == stock_id]
    open_month = open_month.iloc[0]

    if int(open_month)<10:
        open_month = "0"+str(open_month)
    else:
        open_month = open_month
     
    return {
        "stock":str(stock_id),
        "year": str(open_year),
        "month": str(open_month)
    }

def stock_time_bGroup(stock_id):
    
    open_year = df_b["民國年"][df_b["證券代號"] == stock_id]
    open_year = open_year.iloc[0]

    open_month = df_b["西元月"][df_b["證券代號"] == stock_id]
    open_month = open_month.iloc[0]

    if int(open_month)<10:
        open_month = "0"+str(open_month)
    else:
        open_month = open_month
     
    return {
        "stock":str(stock_id),
        "year": str(open_year),
        "month": str(open_month)
    }

def stock_time_cGroup(stock_id):
    
    open_year = df_c["民國年"][df_c["證券代號"] == stock_id]
    open_year = open_year.iloc[0]

    open_month = df_c["西元月"][df_c["證券代號"] == stock_id]
    open_month = open_month.iloc[0]

    if int(open_month)<10:
        open_month = "0"+str(open_month)
    else:
        open_month = open_month
     
    return {
        "stock":str(stock_id),
        "year": str(open_year),
        "month": str(open_month)
    }


import pandas as pd
df = pd.read_excel("/Users/newshuang/Desktop/Python/Grace/stocklist.xlsx")
df2 = df[:][["證券代號","公司名稱","生效日期","案件類別"]]
df2["西元年月日"]=df2["生效日期"]+19110000

df2["西元年"]=round(df2["西元年月日"]/10000,0)
df2["民國年"]=df2["西元年"]-1911
df2["西元月"]=round(df2["西元年月日"]/10000,2)
df2["西元月"]=df2["西元月"]*100-df2["西元年"]*100
df2["西元日"]=df2["西元年月日"]-df2["西元年"]*10000-df2["西元月"]*100

df2["西元日"] = df2["西元日"].astype(int)
df2["西元日"] = df2["西元日"].astype(str)

df2["西元年"] = df2["西元年"].astype(int)
df2["西元年"] = df2["西元年"].astype(str)

df2["民國年"] = df2["民國年"].astype(int)
df2["民國年"] = df2["民國年"].astype(str)

df2["西元月"] = df2["西元月"].astype(int)
df2["西元月"] = df2["西元月"].astype(str)

df2["西元年月日合併"] = df2["西元年"].map(str)+"/"+df2["西元月"].map(str)+"/"+df2["西元日"].map(str)

df_a = df2[df2["案件類別"]=='現金增資']
df_b = df2[df2["案件類別"]=='轉換公司債(有擔保)']
df_c = df2[df2["案件類別"]=='轉換公司債(無擔保)']

stock_a_list = []
stock_a_set = []
stock_a_list_time=[]



for i in range (len(df_a[:]['證券代號'])):
    stock_a_list.append(df_a["證券代號"].iloc[i])
    stock_a_set = set(stock_a_list)
    stock_a_list = list(stock_a_set)

for namelist in stock_a_list:
    stock_a_list_time.append(stock_time_aGroup(namelist))

stock_b_list = []
stock_b_set = []
stock_b_list_time=[]

i=0
for i in range (len(df_b[:]['證券代號'])):
    stock_b_list.append(df_b["證券代號"].iloc[i])
    stock_b_set = set(stock_b_list)
    stock_b_list = list(stock_b_set)

for namelist in stock_b_list:
    stock_b_list_time.append(stock_time_bGroup(namelist))
    
    
    
stock_c_list = []
stock_c_set = []
stock_c_list_time=[]

i=0
for i in range (len(df_c[:]['證券代號'])):
    stock_c_list.append(df_c["證券代號"].iloc[i])
    stock_c_set = set(stock_c_list)
    stock_c_list = list(stock_c_set)

for namelist in stock_c_list:
    stock_c_list_time.append(stock_time_cGroup(namelist))
    
    
# stock_a_list_time[0]["stock"]
# stock_a_list_time[0]["year"]
# stock_a_list_time[0]["month"]
# print(stock_a_list_time[0]["stock"])


#開始處理網頁抓資料

# from bs4 import BeautifulSoup
# import requests
# import pandas as pd

# headers ={'user-agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36'}
# r = requests.get('https://stock.wearn.com/cdata.asp?Year=112&month=03&kind=2330',headers = headers)
# r.encoding = 'big5'

# html = BeautifulSoup(r.text, "html.parser")


#現金增資
#p是計數器

p=0
# dfs.append(crawlist)
for j in range (len(stock_a_list_time)):   
    stock = stock_a_list_time[j]["stock"]
    year = stock_a_list_time[j]["year"]
    month = stock_a_list_time[j]["month"]  

    crawlist=[]
    crawlist.append([str(year),str(month)])
    
    for z in (-2,-1):
        if int(month)+z ==0 :
            year_z = str(int(year)-1)  
            month_z = "12"
            if year_z == "112" and int(month_z) >= 4:
                continue       
            crawlist.append([year_z,month_z])
        elif int(month)+z == -1 :
            year_z = str(int(year)-1)  
            month_z = "11"
            if year_z == "112" and int(month_z) >= 4:
                continue
            crawlist.append([year_z,month_z])
        else:
            if int(month)+z >=10:
                year_z = str(int(year))
                month_z = str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])
            if int(month)+z <10:
                year_z = str(int(year))
                month_z = "0"+str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])    
    
    for z in (1,2,3,4,5,6): 
        if int(month)+z >12 :
            year_z = str(int(year)+1)
            month_z = "0"+str((int(month)+z) % 12)
            if year_z == "112" and int(month_z) >= 4:
                continue
            crawlist.append([year_z,month_z])
        if int(month)+z <=12 :
            if int(month)+z >=10:
                year_z = str(int(year))
                month_z = str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])
            if int(month)+z <10:
                year_z = str(int(year))
                month_z = "0"+str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])         
    s=0
    dfs=[]
    for s in range (len(crawlist)):       
        stock_c = stock_a_list_time[j]["stock"]
        year_c = crawlist[s][0]
        month_c = crawlist[s][1]   
       
        url = f'https://stock.wearn.com/cdata.asp?Year={year_c}&month={month_c}&kind={stock_c}'  
        # 要抓取的網址
        try:
            getdata=pd.read_html(
                url, #想爬的網址
                encoding='big5', # 如何編碼爬下來的資料
                header=0, # 資料取代標題列
                )
            df_stock = getdata [0]
# ------------------------------------------------------------------------------            
            df_stock = df_stock.transpose()
            df_stock = df_stock.reset_index(drop = True,inplace=False)
            df_stock = df_stock.transpose()
            df_stock.columns=[df_stock[0][0],df_stock[1][0],df_stock[2][0],df_stock[3][0],df_stock[4][0],df_stock[5][0]]
            df_stock=df_stock[:][1:]        
            effective_date = df_a[df_a["證券代號"]==int(stock_c)]["西元年月日合併"]
            df_stock["上市日期"] = effective_date.iloc[0]
            df_stock["股票代號"] = int(stock_c)     
            company_name = df_a[df_a["證券代號"] == int(stock_c)]["公司名稱"]
            df_stock["公司名稱"] = company_name.iloc[0]
# ------------------------------------------------------------------------------             
            dfs.append(df_stock)   #當月的資料
        except:
            getdata=pd.read_html(
                url, #想爬的網址
                encoding='cp950', # 如何編碼爬下來的資料
                header=0, # 資料取代標題列
                )
            df_stock = getdata [0]
# ------------------------------------------------------------------------------            
            df_stock = df_stock.transpose()
            df_stock = df_stock.reset_index(drop = True,inplace=False)
            df_stock = df_stock.transpose()
            df_stock.columns=[df_stock[0][0],df_stock[1][0],df_stock[2][0],df_stock[3][0],df_stock[4][0],df_stock[5][0]]
            df_stock=df_stock[:][1:]       
            effective_date = df_a[df_a["證券代號"]==int(stock_c)]["西元年月日合併"]
            df_stock["上市日期"] = effective_date.iloc[0]
            df_stock["股票代號"] = int(stock_c)     
            company_name = df_a[df_a["證券代號"] == int(stock_c)]["公司名稱"]
            df_stock["公司名稱"] = company_name.iloc[0]
# ------------------------------------------------------------------------------  
            dfs.append(df_stock)   #當月的資料    
    dfs = pd.concat(dfs,ignore_index=True)
    dfs = dfs.drop_duplicates()
    p=p+1
    dfs.to_csv((f"/Users/newshuang/Desktop/Python/Grace/股票資料/現金增資/第{p}筆資料_{stock_c}-{year}-{month}.csv"))
    print("爬第",str(p),"筆資料","股票代號：",stock_c)



#有擔保
p=0
for j in range (len(stock_b_list_time)):   
    stock = stock_b_list_time[j]["stock"]
    year = stock_b_list_time[j]["year"]
    month = stock_b_list_time[j]["month"]  

    crawlist=[]
    crawlist.append([str(year),str(month)])
    
    for z in (-2,-1):
        if int(month)+z ==0 :
            year_z = str(int(year)-1)  
            month_z = "12"
            if year_z == "112" and int(month_z) >= 4:
                continue       
            crawlist.append([year_z,month_z])
        elif int(month)+z == -1 :
            year_z = str(int(year)-1)  
            month_z = "11"
            if year_z == "112" and int(month_z) >= 4:
                continue
            crawlist.append([year_z,month_z])
        else:
            if int(month)+z >=10:
                year_z = str(int(year))
                month_z = str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])
            if int(month)+z <10:
                year_z = str(int(year))
                month_z = "0"+str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])    
    
    for z in (1,2,3,4,5,6): 
        if int(month)+z >12 :
            year_z = str(int(year)+1)
            month_z = "0"+str((int(month)+z) % 12)
            if year_z == "112" and int(month_z) >= 4:
                continue
            crawlist.append([year_z,month_z])
        if int(month)+z <=12 :
            if int(month)+z >=10:
                year_z = str(int(year))
                month_z = str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])
            if int(month)+z <10:
                year_z = str(int(year))
                month_z = "0"+str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])
    s=0
    dfs=[]
    for s in range (len(crawlist)):       
        stock_c = stock_b_list_time[j]["stock"]
        year_c = crawlist[s][0]
        month_c = crawlist[s][1]   
       
        url = f'https://stock.wearn.com/cdata.asp?Year={year_c}&month={month_c}&kind={stock_c}'  
        # 要抓取的網址
        try:
            getdata=pd.read_html(
                url, #想爬的網址
                encoding='big5', # 如何編碼爬下來的資料
                header=0, # 資料取代標題列
                )
            df_stock = getdata [0]           
# ------------------------------------------------------------------------------            
            df_stock = df_stock.transpose()
            df_stock = df_stock.reset_index(drop = True,inplace=False)
            df_stock = df_stock.transpose()
            df_stock.columns=[df_stock[0][0],df_stock[1][0],df_stock[2][0],df_stock[3][0],df_stock[4][0],df_stock[5][0]]
            df_stock=df_stock[:][1:]         
            effective_date = df_b[df_b["證券代號"]==int(stock_c)]["西元年月日合併"]
            df_stock["上市日期"] = effective_date.iloc[0]
            df_stock["股票代號"] = int(stock_c)     
            company_name = df_b[df_b["證券代號"] == int(stock_c)]["公司名稱"]
            df_stock["公司名稱"] = company_name.iloc[0]
# ------------------------------------------------------------------------------                         
            dfs.append(df_stock)   #當月的資料
        except:
            getdata=pd.read_html(
                url, #想爬的網址
                encoding='cp950', # 如何編碼爬下來的資料
                header=0, # 資料取代標題列
                )
            df_stock = getdata [0]          
# ------------------------------------------------------------------------------            
            df_stock = df_stock.transpose()
            df_stock = df_stock.reset_index(drop = True,inplace=False)
            df_stock = df_stock.transpose()
            df_stock.columns=[df_stock[0][0],df_stock[1][0],df_stock[2][0],df_stock[3][0],df_stock[4][0],df_stock[5][0]]
            df_stock=df_stock[:][1:]        
            effective_date = df_b[df_b["證券代號"]==int(stock_c)]["西元年月日合併"]
            df_stock["上市日期"] = effective_date.iloc[0]
            df_stock["股票代號"] = int(stock_c)     
            company_name = df_b[df_b["證券代號"] == int(stock_c)]["公司名稱"]
            df_stock["公司名稱"] = company_name.iloc[0]
# ------------------------------------------------------------------------------                           
            dfs.append(df_stock)   #當月的資料    
    dfs = pd.concat(dfs,ignore_index=True)
    dfs = dfs.drop_duplicates()
    p=p+1
    dfs.to_csv((f"/Users/newshuang/Desktop/Python/Grace/股票資料/轉換公司債(有擔保)/第{p}筆資料_{stock_c}-{year}-{month}.csv"))
    print("爬第",str(p),"筆資料","股票代號：",stock_c)

#無擔保
j=0
p=0
for j in range (len(stock_c_list_time)):   
    stock = stock_c_list_time[j]["stock"]
    year = stock_c_list_time[j]["year"]
    month = stock_c_list_time[j]["month"]  

    crawlist=[]
    crawlist.append([str(year),str(month)])
    
    for z in (-2,-1):
        if int(month)+z ==0 :
            year_z = str(int(year)-1)  
            month_z = "12"
            if year_z == "112" and int(month_z) >= 4:
                continue       
            crawlist.append([year_z,month_z])
        elif int(month)+z == -1 :
            year_z = str(int(year)-1)  
            month_z = "11"
            if year_z == "112" and int(month_z) >= 4:
                continue
            crawlist.append([year_z,month_z])
        else:
            if int(month)+z >=10:
                year_z = str(int(year))
                month_z = str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])
            if int(month)+z <10:
                year_z = str(int(year))
                month_z = "0"+str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])    
    
    for z in (1,2,3,4,5,6): 
        if int(month)+z >12 :
            year_z = str(int(year)+1)
            month_z = "0"+str((int(month)+z) % 12)
            if year_z == "112" and int(month_z) >= 4:
                continue
            crawlist.append([year_z,month_z])
        if int(month)+z <=12 :
            if int(month)+z >=10:
                year_z = str(int(year))
                month_z = str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])
            if int(month)+z <10:
                year_z = str(int(year))
                month_z = "0"+str((int(month)+z))
                if year_z == "112" and int(month_z) >= 4:
                    continue
                crawlist.append([year_z,month_z])
    s=0
    dfs=[]
    for s in range (len(crawlist)):       
        stock_c = stock_c_list_time[j]["stock"]
        year_c = crawlist[s][0]
        month_c = crawlist[s][1]   
       
        url = f'https://stock.wearn.com/cdata.asp?Year={year_c}&month={month_c}&kind={stock_c}'  
        # 要抓取的網址
        try:
            getdata=pd.read_html(
                url, #想爬的網址
                encoding='big5', # 如何編碼爬下來的資料
                header=0, # 資料取代標題列
                )
            df_stock = getdata [0]
# ------------------------------------------------------------------------------            
            df_stock = df_stock.transpose()
            df_stock = df_stock.reset_index(drop = True,inplace=False)
            df_stock = df_stock.transpose()
            df_stock.columns=[df_stock[0][0],df_stock[1][0],df_stock[2][0],df_stock[3][0],df_stock[4][0],df_stock[5][0]]
            df_stock=df_stock[:][1:]      
            effective_date = df_c[df_c["證券代號"]==int(stock_c)]["西元年月日合併"]
            df_stock["上市日期"] = effective_date.iloc[0]
            df_stock["股票代號"] = int(stock_c)     
            company_name = df_c[df_c["證券代號"] == int(stock_c)]["公司名稱"]
            df_stock["公司名稱"] = company_name.iloc[0]
# ------------------------------------------------------------------------------                            
            dfs.append(df_stock)   #當月的資料
        except:
            getdata=pd.read_html(
                url, #想爬的網址
                encoding='cp950', # 如何編碼爬下來的資料
                header=0, # 資料取代標題列
                )
            df_stock = getdata [0]
# ------------------------------------------------------------------------------            
            df_stock = df_stock.transpose()
            df_stock = df_stock.reset_index(drop = True,inplace=False)
            df_stock = df_stock.transpose()
            df_stock.columns=[df_stock[0][0],df_stock[1][0],df_stock[2][0],df_stock[3][0],df_stock[4][0],df_stock[5][0]]
            df_stock=df_stock[:][1:]      
            effective_date = df_c[df_c["證券代號"]==int(stock_c)]["西元年月日合併"]
            df_stock["上市日期"] = effective_date.iloc[0]
            df_stock["股票代號"] = int(stock_c)     
            company_name = df_c[df_c["證券代號"] == int(stock_c)]["公司名稱"]
            df_stock["公司名稱"] = company_name.iloc[0]
# ------------------------------------------------------------------------------                 
            dfs.append(df_stock)   #當月的資料    
    dfs = pd.concat(dfs,ignore_index=True)
    dfs = dfs.drop_duplicates()
    p=p+1
    dfs.to_csv((f"/Users/newshuang/Desktop/Python/Grace/股票資料/轉換公司債(無擔保)/第{p}筆資料_{stock_c}-{year}-{month}.csv"))
    print("爬第",str(p),"筆資料","股票代號：",stock_c)




# getdata=pd.read_html(
#     url, #想爬的網址
#     encoding='big5', # 如何編碼爬下來的資料
#     header=0, # 資料取代標題列
#     attrs={'border':'2'}
#     )
