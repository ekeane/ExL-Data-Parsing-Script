#!/usr/bin/env python
# coding: utf-8

import pandas as pd
def exlRegCount(csvInStringformat):
    data = pd.read_csv(csvInStringformat, encoding = "ISO-8859-1")
    
    #Converting proper field names:
    data['PROG AREA CODE'] = data["PROG  AREA  CODE"]
    data['Session ID'] = data["Session ID"]
    data['SECTION TITLE'] = data["SECTION  TITLE"]
    data['START DATE'] = data["START  DATE"]
    data['END DATE'] = data["END  DATE"]
    data['STATUS'] = data["STATUS"]
    data['MAX QTY'] = data["MAX QTY"]
    data['SECTION_COURSE ACCOUNT CODE'] = data["SECTION_COURSE ACCOUNT CODE"]
    data['REGISTRATION DATE'] = data["REGISTRATION  DATE"]
    data['Current Enrollment'] = data["Current Enrollment"]
    data['Total Fees'] = data["Total Fees"]
    data['PROGRAM'] = data["PROGRAM"]
    data['COURSE CODE'] = data["COURSE  CODE"]
    data['PROGRAM TITLE'] = data["PROGRAM  TITLE"]
    data['TERM'] = data["TERM"]
    data['SECTION'] = data["SECTION"]
    data['SOURCE'] = data["SOURCE"]
    data['CITY'] = data["CITY"]
    data['COUNTRY'] = data["COUNTRY"]
    data['PROVINCE_STATE'] = data["PROVINCE_STATE"]
    
    #Drop old columns:
    del data['PROG  AREA  CODE']
    del data['SECTION  TITLE']
    del data['START  DATE']
    del data['END  DATE']
    del data['REGISTRATION  DATE']
    del data['COURSE  CODE']
    del data['PROGRAM  TITLE']
    
    #Reorder columns:
    data = data[['PROG AREA CODE', 'Session ID', 'SECTION TITLE', 'START DATE', 'END DATE', 'STATUS', 'MAX QTY', 'SECTION_COURSE ACCOUNT CODE', 
            'REGISTRATION DATE', 'Current Enrollment', 'Total Fees', 'PROGRAM', 'COURSE CODE', 'PROGRAM TITLE', 'TERM', 
            'SECTION', 'SOURCE', 'CITY', 'COUNTRY', 'PROVINCE_STATE']]

    # Converting dates:
    data['START DATE'] = pd.to_datetime(data['START DATE'])
    data['END DATE'] = pd.to_datetime(data['END DATE'])
    data['REGISTRATION DATE'] = pd.to_datetime(data['REGISTRATION DATE'])
    
    #Convert dates to proper format:
    data['START DATE'] = data['START DATE'].dt.strftime('%d-%m-%Y')
    data['END DATE'] = data['END DATE'].dt.strftime('%d-%m-%Y')
    data['REGISTRATION DATE'] = data['REGISTRATION DATE'].dt.strftime('%d-%m-%Y')
    
    return data.to_excel("ExL Reg Count by Program.xlsx")
    
    


# In[55]:


exlRegCount('Sec Count by ProgArea enrollment Doina.csv')


# In[ ]:




