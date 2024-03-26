#!/usr/bin/env python
# coding: utf-8
import pandas as pd
import numpy as np
import math
from calendar import monthrange
import warnings
warnings.filterwarnings("ignore")
from datetime import date

import sys
from os.path import join, abspath
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return join(sys._MEIPASS, relative_path)
    return join(abspath("."), relative_path)
 
import PySimpleGUI as sg

layout = [[sg.Text('Enter year   '),sg.Input(key='-year-'), sg.Button('Run')],
          [sg.Text('Enter month'), sg.Input(key='-month-'), sg.Button('Exit')],
          [sg.Output(size=(61,15))],
          [sg.Checkbox("Generate Full Year View", key='FY_v')]]

window = sg.Window('IceStar | Commercial Tool', layout, icon=resource_path('icon.ico'))

while True:  # Event Loop
    warnings.filterwarnings("ignore")
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Run':
        # Update the "output" text element to be the value of "input" element
        currentYear = int(values['-year-'])
        currentMonth = int(values['-month-'])
        #-------------------------------------------------------------------- Test Code -------------------------------------------------------
        # Base de dados completa
        arq= pd.ExcelFile('Databases/'+ str(currentYear) + "_Commercial_Tool_Database.xlsm")
        # General Input Database 
        print("Reading Commercial Inputs Database")
        #window.refresh() if window else None
        sheets= ["Budget","Forecast3+9","Forecast6+6","Forecast9+3","ACT_"+str(currentYear),"LY_"+str(currentYear-1)]
        dbs=[]
        for s in sheets:
            db= pd.read_excel(arq,s)
            db["Type"]=s
            dbs.append(db)
        dfinal = pd.concat(dbs)
        if not values['FY_v']:
            # Projection Database
            print("Starting Projection calculation")
            #window.refresh() if window else None
            dclientes= pd.read_excel(arq,'Icestar_Clients') # Clientes
            dproj= pd.read_excel(arq,'Volumes_Projection')
            dprojservs= pd.read_excel(arq,'Services_List')
            dproj.rename(columns={'Month': "Mes"}, inplace=True)
            dproj.rename(columns={'Week': "Semana"}, inplace=True)
            dproj.rename(columns={'BLAST FREEZE': "FREEZE"}, inplace=True)
            dproj= dproj.fillna(0)
            # Promedio Mensal Saldos ( Monthly Average Storage Balance )
            dproj["PromedioM_Frozen"]=0.0
            dproj["PromedioM_Refrigerated"]=0.0
            dproj["PromedioM_Dry"]=0.0
            # Ingrssos Mensal
            dproj["Ingressos($)_Frozen"]=0.0
            dproj["Ingressos($)_Refrigerated"]=0.0
            dproj["Ingressos($)_Dry"]=0.0
            sems_meses= dproj.groupby(['Mes']).apply(lambda x: len(x["Semana"].unique())).to_dict()
            m = currentMonth-1
            clientes= dproj.Client_Code.unique()
            meses= dproj[dproj.Mes>m].Mes.unique()
            sems= dproj[dproj.Mes>0].Semana.unique()
            Storage_Types=["Frozen","Refrigerated","Dry"]
            Storage_Types_Dict={"Frozen":"Storage Frozen","Refrigerated":"Storage Refrigerated","Dry":"Storage Dry"}
            for f in dproj.Facility.unique():
                print(f)
                #window.refresh() if window else None
                for stt in Storage_Types:
                    PromedioMT= "PromedioM_"+ stt
                    SaldoS= "Balance(" + stt +")"
                    print('Gathering', stt, 'data for months: ', end = " ")
                    for m in meses:
                        print(m, end = " ")
                        #window.refresh() if window else None
                        for c in clientes:
                            dpm= dproj[(dproj.Facility == f)&(dproj.Mes == m) & (dproj.Client_Code ==c)]
                            dproj.loc[(dproj.Facility == f)&(dproj.Mes == m) & (dproj.Client_Code ==c), PromedioMT] = dpm[(dpm.Mes == m )][SaldoS].mean()
                    print('')
            
            # Ingressos en pesos (Storage projection calculation in revenue)
            df= pd.read_excel(arq, "Rates")
            ufs= pd.read_csv("Databases/External/UF_"+str(currentYear)+".csv",sep=';')
            ufs = ufs.melt(id_vars=["Día"], var_name="Mes", value_name="Total")
            meses= ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep','Oct', 'Nov', 'Dic']
            i=1
            for m in meses:
                ufs.loc[(ufs['Mes']== m),"Mes"]=i
                i+=1

            dftp = pd.read_excel(arq, "Contracts") # contratos
            dftp = dftp.fillna(0)
            df = df.fillna(0)
            dproj= dproj.fillna(0)
            dftp.rename(columns={'Fixed Positions Rate': "Rate"}, inplace=True)
            dftp["Variable Positions Rate"]= dftp["Variable Positions Rate"].fillna(0)
            # Sem minímo
            ufs['Total'] = ufs['Total'].str.replace(".", "")
            ufs['Total'] = ufs['Total'].str.replace(",", ".").astype(float)
            ufs= ufs.groupby(['Mes']).apply(lambda x: x["Total"].mean()).reset_index(name ='Total')
            ufs.loc[(ufs.Mes > currentMonth), 'Total'] = ufs[ufs.Mes == currentMonth]["Total"].sum()
            dproj= dproj[dproj.Client_Code!=0]
            # Ingressos (Storage Calculation)
            m=1
            for f in dftp.Facility.unique():
                print('Running storage calculation for:',f)
                #window.refresh() if window else None
                for stt in Storage_Types:
                    PromedioMT = "PromedioM_"+ stt
                    IngressosT = "Ingressos($)_"+ stt
                    clientes_tarifas_base= df[(df.Facility==f)&(df["IceStar Contract Type"]==Storage_Types_Dict[stt])].groupby(['Client_Code',"Moneda_Geral"]).apply(lambda x: x["Valor"].mean()).to_dict()
                    try:
                        if not clientes_tarifas_base['Facility']:
                            break
                    except:
                        pass
                    clientes_TO= dftp[(dftp.Facility==f)&(dftp.ContractType =="TakeorPay")].Client_Code.unique()
                    for c in clientes_TO:
                        checkvalue=1
                        for cli in clientes_tarifas_base:
                            if cli[0]== c:
                                clientes_tarifas_base[cli] = dftp[dftp.Client_Code ==c].Rate.mean()
                                checkvalue=0
                        if checkvalue:
                            clientes_tarifas_base[(c,"PESOS")] = dftp[dftp.Client_Code ==c].Rate.mean()
                    m=1
                    countm = {cm[0]:0 for cm in clientes_tarifas_base}
                    actualmonth = {cm[0]:0 for cm in clientes_tarifas_base}
                    rate = {}
                    revenue = {}
                    while m <=12:
                        clientes_tarifas={}
                        for cm in clientes_tarifas_base:
                            c=cm[0]
                            if c!=0:
                                if cm[1]=="UF":
                                    if clientes_tarifas_base[cm] <1:
                                        if m >= currentMonth:
                                            clientes_tarifas[cm]=clientes_tarifas_base[cm]*ufs[ufs.Mes==(currentMonth-1)].Total.mean()
                                        else:
                                            clientes_tarifas[cm]=clientes_tarifas_base[cm]*ufs[ufs.Mes==m].Total.mean()
                                    else:
                                        clientes_tarifas[cm]=clientes_tarifas_base[cm]
                                else:
                                    clientes_tarifas[cm]=clientes_tarifas_base[cm]
                                if not countm[c]:   
                                    rate[c]= clientes_tarifas[cm]
                                    revenue[c]= dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Revenue.sum()
                                if (cm[1]=="PESOS")|(cm[1]=="UF"):
                                    initialm=dftp[(dftp.Facility==f)&(dftp.Client_Code== c)]['Rate Adjustment Initial Month'].mean()
                                    frequency= dftp[(dftp.Facility==f)&(dftp.Client_Code== c)]['Rate Adjustment Frequency'].mean()
                                    inflation = dproj[(dproj.Facility==f)&(dproj.Client_Code == c)].Inflation.mean()
                                    if frequency == 12:
                                        rate[c]= clientes_tarifas[cm] 
                                        revenue[c]= dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Revenue.sum()
                                    elif m == initialm:
                                        rate[c] = (clientes_tarifas[cm]*(1+inflation)**(m))
                                        revenue[c]= (dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Revenue.sum()*(1+inflation)**m)
                                        actualmonth[c] = initialm + frequency
                                        countm[c]= 1
                                    elif frequency == 1:
                                        rate[c] = (clientes_tarifas[cm]*(1+inflation)**(m))
                                        revenue[c]= (dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Revenue.sum()*(1+inflation)**(m))
                                    elif m == actualmonth[c]:
                                        countm[c]+=1
                                        rate[c] = (clientes_tarifas[cm]*(1+inflation)**(m*countm[c]))
                                        revenue[c]= (dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Revenue.sum()*(1+inflation)**(m*countm[c]))
                                        actualmonth[c] = actualmonth[c] + frequency
                                    if c in clientes_TO:
                                        dif_positions= (dproj[(dproj.Facility==f)&(dproj.Client_Code == c)&(dproj.Mes == m)][PromedioMT].mean()-dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Positions.sum())
                                        if (dftp[(dftp.Client_Code== c)].Type.unique()[0]=="Positions"):
                                            if dif_positions >=0:
                                                var_rate= dftp[dftp.Client_Code ==c]["Variable Positions Rate"].mean()
                                                dproj.loc[(dproj.Facility==f)&(dproj.Client_Code == c)& (dproj.Mes == m), IngressosT] = rate[c] *dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Positions.sum()*monthrange(currentYear,m)[1]+var_rate * dif_positions * monthrange(currentYear,m)[1]
                                            else:
                                                dproj.loc[(dproj.Facility==f)&(dproj.Client_Code == c)& (dproj.Mes == m), IngressosT] = rate[c] *dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Positions.sum()*monthrange(currentYear,m)[1]
                                        elif (dftp[(dftp.Client_Code== c)].Type.unique()[0]=="Revenue"):
                                            if dif_positions >=0:
                                                dproj.loc[(dproj.Facility==f)&(dproj.Client_Code == c)& (dproj.Mes == m), IngressosT] = revenue[c] + rate[c] * dif_positions * monthrange(currentYear,m)[1]
                                            else:
                                                dproj.loc[(dproj.Facility==f)&(dproj.Client_Code == c)& (dproj.Mes == m), IngressosT] = revenue[c]
                                    else:
                                        dproj.loc[(dproj.Facility==f)&(dproj.Client_Code == c)& (dproj.Mes == m), IngressosT] = rate[c] * dproj[(dproj.Facility==f)&(dproj.Client_Code == c)&(dproj.Mes == m)][PromedioMT].mean() * monthrange(currentYear,m)[1]
                        m+=1
            dproj = dproj.drop_duplicates()
            clientes_TO= dftp[(dftp.ContractType =="TakeorPay")].Client_Code.unique()
            # Ingessos BLAST FREEZING (Blast Freezing Calculation)
            for f in dftp.Facility.unique():
                clientes_tarifas_base= df[(df.Facility==f)&(df["IceStar Contract Type"]=='Blast Freezing')].groupby(['Client_Code',"Moneda_Geral"]).apply(lambda x: x["Valor"].mean()).to_dict()
                m=1
                countm = {cm[0]:0 for cm in clientes_tarifas_base}
                actualmonth = {cm[0]:0 for cm in clientes_tarifas_base}
                rate = {}
                revenue = {}
                while m <=12:
                    clientes_tarifas={}
                    for cm in clientes_tarifas_base:
                        c=cm[0]
                        if cm[1]=="UF":
                            if clientes_tarifas_base[cm] <1:
                                if m >= currentMonth:
                                    clientes_tarifas[cm]=clientes_tarifas_base[cm]*ufs[ufs.Mes==(currentMonth-1)].Total.mean()
                                else:
                                    clientes_tarifas[cm]=clientes_tarifas_base[cm]*ufs[ufs.Mes==m].Total.mean()
                            else:
                                clientes_tarifas[cm]=clientes_tarifas_base[cm]   
                        else:
                            clientes_tarifas[cm]=clientes_tarifas_base[cm]   
                        if not countm[c]:   
                            rate[c]= clientes_tarifas[cm]
                            revenue[c]= dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Revenue.sum()
                        if (cm[1]=="PESOS")|(cm[1]=="UF"):
                            initialm=dftp[(dftp.Facility==f)&(dftp.Client_Code== c)]['Rate Adjustment Initial Month'].mean()
                            frequency= dftp[(dftp.Facility==f)&(dftp.Client_Code== c)]['Rate Adjustment Frequency'].mean()
                            inflation = dproj[(dproj.Facility==f)&(dproj.Client_Code == c)].Inflation.mean()
                            if frequency == 12:
                                rate[c]= clientes_tarifas[cm] 
                                revenue[c]= dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Revenue.sum()
                            elif m == initialm:
                                rate[c] = (clientes_tarifas[cm]*(1+inflation)**(m))
                                revenue[c]= (dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Revenue.sum()*(1+inflation)**m)
                                actualmonth[c] = initialm + frequency
                                countm[c]= 1
                            elif m == actualmonth[c]:
                                countm[c]+=1
                                rate[c] = (clientes_tarifas[cm]*(1+inflation)**(m*countm[c]))
                                revenue[c]= (dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Revenue.sum()*(1+inflation)**(m*countm[c]))
                                actualmonth[c] = actualmonth[c] + frequency
                            elif frequency == 1:
                                rate[c] = (clientes_tarifas[cm]*(1+inflation)**(m))
                                revenue[c]= (dftp[(dftp.Facility==f)&(dftp.Client_Code== c)].Revenue.sum()*(1+inflation)**(m))
                            dproj.loc[(dproj.Facility==f)&(dproj.Client_Code == c)& (dproj.Mes == m), "BLAST FREEZING"] = rate[c] * dproj[(dproj.Facility==f)&(dproj.Client_Code == c)&(dproj.Mes == m)]['Blast Freezing Volume'].sum() 
                    m+=1                    
            # Por servicio
            dproj=dproj.fillna(0)
            dfat= dfinal[dfinal.Type== "ACT_" + str(currentYear)]
            dfat=dfat.rename(columns={'Month': "Mes"})
            dfat=dfat[dfat.Mes!=0]
            meses= list(range(1,13))
            clientes= dfat.Client_Code.unique()
            m = currentMonth - 1
            servs= dprojservs.dropna().Specified_Services.unique()
            sems_meses= dproj.groupby(['Mes']).apply(lambda x: len(x["Semana"].unique())).to_dict()
            dfat=dfat.fillna(0)
            count=0 
            for f in dfat.Facility.unique():  
                clientes= dfat[dfat.Facility==f].Client_Code.unique()
                for c in clientes:
                    mesflag={x:0 for x in servs}
                    mesvalor={x:0 for x in servs}
                    for me in meses:
                        Total= dfat[(dfat.Facility == f)&(dfat.Client_Code == c) &(dfat.Mes == me) & (dfat.Service == "STORAGE")].Total.sum()
                        if Total!= 0:
                            for s in servs:
                                strexp= 'Expected_'+ s+' (0 or 1)'
                                if dproj[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me)][strexp].max()==1:
                                    mesflag[s]=1
                                    mesvalor[s]=dproj[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me)&(dproj[strexp]!=0)][s].mean()
                                    dproj.loc[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me),s]= mesvalor[s]
                                else:
                                    if mesflag[s]:
                                        dproj.loc[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me),s]= mesvalor[s]
                                    else:
                                        dproj.loc[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me),s] = dfat[(dfat.Facility == f)&(dfat.Client_Code == c) &(dfat.Mes == me)&(dfat.Service == s)].Total.sum()/Total
                            dproj.loc[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me),"STORAGE"] = 1
                        else:
                            if me>= currentMonth:
                                m= currentMonth-1
                                Total =dfat[(dfat.Facility == f)&(dfat.Client_Code == c) &(dfat.Mes <=m) & (dfat.Service == "STORAGE")].Total.sum()
                                if not Total !=0:
                                    Total= dfat[(dfat.Facility == f)&(dfat.Mes <= m)].Total.sum()
                            else:
                                Total= dfat[(dfat.Facility == f)&(dfat.Mes == me)].Total.sum()
                                m=me
                            for s in servs:
                                strexp= 'Expected_'+s+' (0 or 1)'
                                if dproj[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me)][strexp].max()==1:
                                    mesflag[s]=1
                                    mesvalor[s]=dproj[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me)&(dproj[strexp]!=0)][s].mean()
                                else:
                                    if mesflag[s]:
                                        dproj.loc[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me),s]= mesvalor[s]
                                    else:
                                        dproj.loc[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me),s] = dfat[(dfat.Facility == f)&(dfat.Client_Code == c)&(dfat.Mes <= m)&(dfat.Service == s)].Total.sum()/Total
                                dproj.loc[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me),"STORAGE"] = 1 

            for f in dfat.Facility.unique():
                clientes= dfat[dfat.Facility==f].Client_Code.unique()
                clientes_budg=dproj[(~dproj.Client_Code.isin(clientes)) & (dproj.Facility==f)].Client_Code.unique()
                for c in clientes_budg:
                    mesflag={x:0 for x in servs}
                    mesvalor={x:0 for x in servs}
                    for me in meses:
                        if me>= currentMonth:
                            Total= dfat[(dfat.Facility == f)&(dfat.Mes <= currentMonth-1)].Total.sum()
                            m= currentMonth-1
                        else:
                            Total= dfat[(dfat.Facility == f)&(dfat.Mes == me)].Total.sum()
                            m=me
                        for s in servs:
                            strexp= 'Expected_'+ s+' (0 or 1)'
                            if dproj[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me)][strexp].max()==1:
                                mesflag[s]=1
                                mesvalor[s]=dproj[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me)][s].mean()
                            else:
                                if mesflag[s]:
                                        dproj.loc[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me),s]= mesvalor[s]
                                else:
                                    dproj.loc[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me),s] = dfat[(dfat.Facility == f)&(dfat.Client_Code == c) &(dfat.Mes <= m)&(dfat.Service == s)].Total.sum()/Total
                        dproj.loc[(dproj.Facility == f)&(dproj.Client_Code ==c) & (dproj.Mes ==me),"STORAGE"] = 1 
            dproj['STORAGE']=1
            dproj['OTHERS'] =dproj.OTHERS*(dproj['Ingressos($)_Frozen'])
            dproj['HANDLING'] =dproj.HANDLING*dproj['Ingressos($)_Frozen']
            dproj['STORAGE'] =dproj.STORAGE*(dproj['Ingressos($)_Frozen']+dproj['Ingressos($)_Refrigerated']+dproj['Ingressos($)_Dry'])
            dprojteste=dproj[dproj.Mes>(currentMonth - 1)]
            # Columns Organization
            dtiposconts= dftp[['Client_Code','ContractType']].drop_duplicates()
            dprojteste= pd.merge(dprojteste,dtiposconts,on ='Client_Code',how ='outer')
            dprojteste=dprojteste.drop_duplicates()
            dprojteste= dprojteste.groupby(['Mes','Client_Code','ContractType','Ingressos($)_Frozen',"Facility","BLAST FREEZING",'OTHERS','HANDLING']).apply(lambda x: x[['STORAGE']].mean()).reset_index()
            dprojteste.loc[(dprojteste.Client_Code.isin(clientes_TO)),'ContractType']="TakeorPay"
            dprojteste=dprojteste.drop_duplicates()
            dprojteste["Type"]="YTG"
            dproj = dprojteste.copy()
            print("Projection calculation Finished")
            #window.refresh() if window else None
            # Pipeline Database
            print("Reading Pipeline Database")
            window.refresh() if window else None
            arq_p = pd.ExcelFile('Databases/Pipeline_Database_CRM.xlsx')
            dpip= pd.read_excel(arq_p,"Pipeline Database")
            dpip['Operation Beginning Month'] = pd.DatetimeIndex(dpip['Operation Beginning Date']).month
            dpip['Operation Beginning Day'] = pd.DatetimeIndex(dpip['Operation Beginning Date']).day
            dpip['Operation Beginning Year'] = pd.DatetimeIndex(dpip['Operation Beginning Date']).year
            dpip.rename(columns={'Account Name (Client)':'Client'},inplace=True)
            colunas = dpip.columns
            dpip= dpip[(dpip['Sales Pipeline Stages']!=  '6. Closed - Won') & (dpip['Sales Pipeline Stages']!= '6. Closed - Lost') & (dpip['Probability (edit)']>=0.9)]
            for c in dpip.Client.unique():
                try:
                    mes=dpip[(dpip.Client==c)]['Operation Beginning Month'].unique()[0]
                    dia=dpip[(dpip.Client==c)]['Operation Beginning Day'].unique()[0]
                    ano=dpip[(dpip.Client==c)]['Operation Beginning Year'].unique()[0]
                    if ano == currentYear:
                        valor=dpip[(dpip.Client==c)]['Monthly Estimated Revenue (CLP$)'].mean()* dpip[(dpip.Client==c)]['Probability (edit)'].mean()
                        dur=dpip[(dpip.Client==c)]['Contract Duration (Months)'].unique()[0]
                        if not math.isnan(mes):
                            mes=int(mes)
                            dia=int(dia)
                            i=mes
                            count=0
                            mrange= monthrange(currentYear,mes)
                            valor_inicial = valor
                            if dia>1:
                                valor_inicial = valor/mrange[1]*(mrange[1]-dia)
                            while i <=12:
                                if count > dur:
                                    break
                                if i == mes:
                                    dpip.loc[(dpip.Client==c),i]= valor_inicial
                                else:
                                    dpip.loc[(dpip.Client==c),i]= valor
                                i+=1
                                count+=1
                except:
                    print(c)
                    window.refresh() if window else None        
            dpip=dpip.melt(id_vars=colunas, var_name="Mes", value_name="Total")
            dpip['Total'] = dpip['Total'].astype(float)
            dpipservs= pd.read_excel(arq_p,sheet_name="Data Validation")
            dpipservs=dpipservs[["Services"]]
            dpipservs= dpipservs.dropna()
            servispipe= dpipservs.Services.unique()
            for s in servispipe:
                sp="%"+s
                dpip[s]= dpip[sp] * dpip["Total"]
            dpip.rename(columns={'Client Segment': "Segment"}, inplace=True)
            dpip.rename(columns={'Contract Type': "ContractType"}, inplace=True)
            dpip=dpip[[ 'Client','Client_Code','GroupCode','Facility',"ContractType",'Segment','Mes','Sales Pipeline Stages','Probability (edit)','Operation Beginning Date','Opportunity Opening Date','STORAGE','HANDLING','BLASTFREEZING','OTHERS']]
            dpip["Type"]="Pipeline"
            m= currentMonth-1
            dpipfinal= dpip[dpip.Mes>m]
            dpip=dpipfinal
            dpip = dpip.melt(id_vars=['Facility',"Client",'Client_Code','GroupCode', "Mes","Type","ContractType", 'Segment','Sales Pipeline Stages','Probability (edit)','Operation Beginning Date','Opportunity Opening Date'], var_name="Service", value_name="Total")
            dpip.rename(columns={'Client': 'Client_Name'}, inplace=True)
            dpip.rename(columns={'ContractType': 'Client_Type'}, inplace=True)
            dpip = dpip[['Facility',"Client_Name","Client_Code","GroupCode", "Mes","Type","Client_Type", 'Segment','Service','Total']]
            #dpip.to_csv("Databases/BIinputs/Pipeline_Database.csv",index=False)
            # Final Database
            dprojfinal=dproj[['Client_Code','STORAGE','HANDLING','OTHERS','BLAST FREEZING','Mes','Type',"Facility","ContractType"]]
            dprojfinal = dprojfinal.fillna(0)
            dprojfinal=dprojfinal.drop_duplicates()
            dprojfinal = dprojfinal[['Client_Code', 'STORAGE', 'HANDLING', 'OTHERS', 'Mes', 'Type', 'Facility','BLAST FREEZING']]
            dprojfinal = dprojfinal.melt(id_vars=["Client_Code", "Mes","Type",'Facility'], var_name="Service", value_name="Total")
            dprojfinal.Client_Code= dprojfinal.Client_Code.replace(np.nan, 0)
            dprojfinal.Client_Code = dprojfinal.Client_Code.astype(int)
            dclientes_tot= pd.read_excel(arq,'Icestar_Clients')[['Client_Code', 'Client_Name']]
            dprojfinal= pd.merge(dprojfinal,dclientes_tot,on ='Client_Code',how ='left')
            # Organizacao base servicos (Services Database Organization)
            dfinal.rename(columns={'Month': "Mes"}, inplace=True)
            dfinal.Client_Code.unique()
            dfinal.Client_Code= dfinal.Client_Code.replace(np.nan, 0)
            dfinal.Client_Code = dfinal.Client_Code.astype(int)
            dfinal=pd.concat([dfinal,dprojfinal])
            dfinal.drop('ContractType',axis=1,inplace=True)
            dclientes= pd.read_excel(arq,'Icestar_Clients')[['Client_Code','Segment','Group','GroupCode','Client_Type','Operation_Type']]
            dclientes.Client_Code = dclientes.Client_Code.astype(int)
            dfinal.Client_Code.unique()
            dfinal.Client_Code= dfinal.Client_Code.replace(np.nan, 0)
            dfinal.Client_Code = dfinal.Client_Code.astype(int)
            dfinal= pd.merge(dfinal,dclientes,on ='Client_Code',how ='left')
            dfinal=pd.concat([dfinal,dpip])
            dfinal['Segment'].fillna("Others",inplace=True)
            # Base de dados conversao dolar
            arq_rates= pd.ExcelFile('Databases/External/Dollar_rates.xlsx')
            rates = pd.read_excel(arq_rates)
            rates_dict = dict(zip(rates['Facility'],rates['Rate']))
            rates_dict[np.nan] = 0
            dfinal['Total Local'] = dfinal['Total'].copy()
            dfinal['Total'] = dfinal.apply(lambda x: x['Total']*rates_dict[x['Facility']],axis = 1)
            dfinal.to_excel("Databases/BIinputs/IceStar_Commercial_Database.xlsx",index=False)
            today = date.today()
            dfinal.to_excel("Databases/BIinputs/old/IceStar_Commercial_Database_"+str(currentYear)+"_"+str(currentMonth)+"_"+str(today)+".xlsx",index=False)
            print("Database Ready --> Refresh Power BI")
            #window.refresh() if window else None
            print("You can exit now")
            #window.refresh() if window else None
        else:
            if currentMonth == 12:
                dfinal.rename(columns={'Month': "Mes"}, inplace=True)
                dfinal.Client_Code.unique()
                dfinal.Client_Code= dfinal.Client_Code.replace(np.nan, 0)
                dfinal.Client_Code = dfinal.Client_Code.astype(int)
                dfinal.drop('ContractType',axis=1,inplace=True)
                dclientes= pd.read_excel(arq,'Icestar_Clients')[['Client_Code','Segment','Group','GroupCode','Client_Type','Operation_Type']]
                dclientes.Client_Code = dclientes.Client_Code.astype(int)
                dfinal.Client_Code.unique()
                dfinal.Client_Code= dfinal.Client_Code.replace(np.nan, 0)
                dfinal.Client_Code = dfinal.Client_Code.astype(int)
                dfinal= pd.merge(dfinal,dclientes,on ='Client_Code',how ='left')
                dfinal['Segment'].fillna("Others",inplace=True)
                # Base de dados conversao dolar
                arq_rates= pd.ExcelFile('Databases/External/Dollar_rates.xlsx')
                rates = pd.read_excel(arq_rates)
                rates_dict = dict(zip(rates['Facility'],rates['Rate']))
                rates_dict[np.nan] = 0
                dfinal['Total Local'] = dfinal['Total'].copy()
                dfinal['Total'] = dfinal.apply(lambda x: x['Total']*rates_dict[x['Facility']],axis = 1)
                dfinal.to_excel("Databases/BIinputs/IceStar_Commercial_Database.xlsx",index=False)
                today = date.today()
                dfinal.to_excel("Databases/BIinputs/old/IceStar_Commercial_Database_"+str(currentYear)+"_"+str(currentMonth)+"_"+str(today)+".xlsx",index=False)
                print("Database Ready --> Refresh Power BI")
                #window.refresh() if window else None
                print("You can exit now")
                #window.refresh() if window else None
            else:
                print("You can only run for month 12")
                #window.refresh() if window else None
        #-------------------------------------------------------------------- Test Code -------------------------------------------------------

window.close()
