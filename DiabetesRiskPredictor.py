#!/usr/bin/env python
# coding: utf-8

# In[37]:


import pandas as pd


# In[38]:


import xlwings as xw


# In[39]:


import datetime


# In[40]:


def evaluate_diabetes_risk(bmi, hba1c):
    """Custom risk classification based on BMI and HbA1c"""
    if (18.5 <= bmi <= 24.9) or (hba1c < 5.7):
        return "LOW RISK"
    elif (25.0 <= bmi <= 29.9) or (5.7 <= hba1c <= 6.4):
        return "MODERATE RISK"
    elif (bmi >= 30.0) or (6.3 <= hba1c <= 7.0):
        return "HIGH RISK"
    else:
        return "VERY HIGH RISK"


# In[41]:


def main():
    try:
        wb = xw.books.active
        sheet = wb.sheets['PREDICTIONS']
        user_bmi = sheet.range('F9').value  
        user_hba1c = sheet.range('F10').value
        risk_level = evaluate_diabetes_risk(user_bmi, user_hba1c)
        sheet.range('F12').value = risk_level
        wb.save()

        excel = xw.apps.active.api
        excel.MsgBox(f"Diabetes Risk Level: {risk_level}", 0, "Risk Assessment Result")
        
        with open(r"C:\Users\owner\OneDrive\Desktop\risk_log.txt", "a") as f:
         f.write(f"{datetime.datetime.now()}: Success - BMI: {user_bmi}, HbA1c: {user_hba1c}, Risk: {risk_level}\n")
    
    except Exception as e:
        with open(r"C:\Users\owner\OneDrive\Desktop\risk_log.txt", "a") as f:
            f.write(f"{datetime.datetime.now()}: Error - {str(e)}\n")


# In[42]:


main()

