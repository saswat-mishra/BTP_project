# -*- coding: utf-8 -*-
"""
Created on Mon Jan 11 10:42:59 2021

@author: ASUS
"""


import numpy as np
import xlrd

shirish = xlrd.open_workbook("Shirish.xlsx").sheet_by_index(0)
shreyas = xlrd.open_workbook("Shreyas.xlsx").sheet_by_index(0)
ashish = xlrd.open_workbook("Ashish.xlsx").sheet_by_index(0)
ankit = xlrd.open_workbook("Ankit.xlsx").sheet_by_index(0)
nirmal = xlrd.open_workbook("Nirmal.xlsx").sheet_by_index(0)
public = xlrd.open_workbook("Public.xlsx").sheet_by_index(0)


print(type(shirish.cell_value(12,0)))

wb1 = []
for row in range (shirish.nrows):
    temp = []
    for col in range (1, shirish.ncols):
        if isinstance(shirish.cell_value(row,col), float) or isinstance(shirish.cell_value(row,col), int):
            temp.append(shirish.cell_value(row,col))
    wb1.append(temp)
wb2 = []
for row in range (shreyas.nrows):
    temp = []
    for col in range (1, shreyas.ncols):
        if isinstance(shreyas.cell_value(row,col), float) or isinstance(shreyas.cell_value(row,col), int):
            temp.append(shreyas.cell_value(row,col))
    wb2.append(temp)
wb3 = []
for row in range (ashish.nrows):
    temp = []
    for col in range (1, ashish.ncols):
        if isinstance(ashish.cell_value(row,col), float) or isinstance(ashish.cell_value(row,col), int):
            temp.append(ashish.cell_value(row,col))
    wb3.append(temp)
wb4 = []
for row in range (ankit.nrows):
    temp = []
    for col in range (1, ankit.ncols):
        if isinstance(ankit.cell_value(row,col), float) or isinstance(ankit.cell_value(row,col), int):
            temp.append(ankit.cell_value(row,col))
    wb4.append(temp)
wb5 = []
for row in range (nirmal.nrows):
    temp = []
    for col in range (1, nirmal.ncols):
        if isinstance(nirmal.cell_value(row,col), float) or isinstance(nirmal.cell_value(row,col), int):
            temp.append(nirmal.cell_value(row,col))
    wb5.append(temp)

wb6 = []
for row in range (1,public.nrows):
    temp = []
    for col in range (6, public.ncols):
        if isinstance(public.cell_value(row,col), float) or isinstance(public.cell_value(row,col), int):
            temp.append(public.cell_value(row,col))
    wb6.append(temp)

#function for generating G.m. of 2 arrays
def gm_arr(arr1, arr2):
    temp = []
    for i in range(len(arr1)):
        temp.append((arr1[i]*arr2[i])**0.5)
    return temp
#function to generate G.m. of 2 matrices
def gm_mat(mat1, mat2):
    temp = []
    for i in range(len(mat1)):
        temp.append(gm_arr(mat1[i], mat2[i]))
    return temp

data = gm_mat(gm_mat(gm_mat(wb1, wb2),wb3), wb4)

tax_benefits = []
non_tax_financial_exemptions = []
refueling_benifits = []
battery_performance = []
vehicle_related_measures = []
section_comparision = []

#data = np.asarray(data)

for i in range(3,6):
    temp = []
    for j in range(3):
        temp.append(data[i][j])
    tax_benefits.append(temp)
    
for i in range(11,15):
    temp = []
    for j in range(4):
        temp.append(data[i][j])
    non_tax_financial_exemptions.append(temp)
    
for i in range(22,25):
    temp = []
    for j in range(3):
        temp.append(data[i][j])
    refueling_benifits.append(temp)
    
for i in range(30,33):
    temp = []
    for j in range(3):
        temp.append(data[i][j])
    battery_performance.append(temp)
    
for i in range(38,42):
    temp = []
    for j in range(4):
        temp.append(data[i][j])
    vehicle_related_measures.append(temp)
    
for i in range(6,11):
    temp = []
    for j in range(5):
        temp.append(data[i][j])
    section_comparision.append(temp)
    
def norm_mat(mat):
    for i in range(len(mat[0])):
        temp = 0
        for j in range(len(mat)):
            temp = temp + mat[j][i]
        for k in range(len(mat)):
            mat[k][i] = mat[k][i]/temp
    return mat

norm_mat(tax_benefits)
norm_mat(non_tax_financial_exemptions)
norm_mat(refueling_benifits)
norm_mat(battery_performance)
norm_mat(vehicle_related_measures)
norm_mat(section_comparision)

priority = []
multiplier = []
for i in range(len(tax_benefits)):
    priority.append(sum(tax_benefits[i]))
    
for i in range(len(non_tax_financial_exemptions)):
    priority.append(sum(non_tax_financial_exemptions[i]))
    
for i in range(len(refueling_benifits)):
    priority.append(sum(refueling_benifits[i]))
    
for i in range(len(battery_performance)):
    priority.append(sum(battery_performance[i]))
    
for i in range(len(vehicle_related_measures)):
    priority.append(sum(vehicle_related_measures[i]))
    
for i in range(len(section_comparision)):
    multiplier.append(sum(section_comparision[i]))
    
priority_factor = np.zeros(17)
for i in range(3):
    priority_factor[i] = priority[i]*multiplier[0]
    
for i in range(3,7):
    priority_factor[i] = priority[i]*multiplier[1]
    
for i in range(7,10):
    priority_factor[i] = priority[i]*multiplier[4]
    
for i in range(10,13):
    priority_factor[i] = priority[i]*multiplier[2]
    
for i in range(13,17):
    priority_factor[i] = priority[i]*multiplier[3]
    
arr = []
for i in range(17):
    temp = 0
    for j in range(74):
        temp = temp + wb6[j][i]
    arr.append(temp/74)
summ = sum(arr)
for i in range(17):
    arr[i] = arr[i]/summ


for i in range(17):
    priority_factor[i] = priority_factor[i]*arr[i]

#Creating ranked list
params = ["Exemption from sales tax / GST",	"Exemption from import tax on selected EV companies",	"No road tax",	"Subsidies when purchasing EVs",	"Exemption from transportation fees",	"Exemption from municipality parking charges",	"Free access to all of toll roads",	"Exemption from fees at designated charging stations",	 "Availability of Public charging stations",	"Tax incentives for In-home charging equipment",	"Maximum distance covered by vehicle in a single charge",	"Compatibility with different charging ports",	"Speed of charging",	"Maximum vehicle speed",	"Durability",	"Access to servicing",	"Proximity to point of purchase"]

sorted_index = np.argsort(priority_factor)
sorted_index[:] = sorted_index[::-1]
result = []

for i in range(17):
    result.append(params[sorted_index[i]])
    
print(result)