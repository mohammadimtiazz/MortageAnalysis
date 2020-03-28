# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

from openpyxl import Workbook
from enum import Enum

class FixedLoanType(Enum):
    fifteen_year = 1
    thirty_year = 2

def MonthlyMortagePayment(loan_amount, interest_rate, FixedLoanType):
    # formula of monthly rate 
    # L = interest_rate/12
    # %c = total number of year loan * total_num_pay_per_year
    # %L[c(1 + c)^n]/[(1 + c)^n - 1]    
    if (FixedLoanType == FixedLoanType.thirty_year):
        total_num_payment = 30*12
        c = ((interest_rate/100)/12)
        monthly_payment = loan_amount * c * (((1+c)**total_num_payment)/(((1+c)**total_num_payment) - 1))
        
    if (FixedLoanType == FixedLoanType.fifteen_year):
        total_num_payment = 15*12
        c = ((interest_rate/100)/12)
        monthly_payment = loan_amount * c * (((1+c)**total_num_payment)/(((1+c)**total_num_payment) - 1)) 
    
    return monthly_payment
        
        
    
#The price of the home
#home_price = input("Insert Home Price: ")
home_price = 300000.0
home_price = float(home_price)

#projected down payment
down_payment_percentage = range(5,21,5)
down_payment_dollar = []

#The interest rate
#interest_rate = input("Insert interest rate as percentage (x%): ")
interest_rate = 3.25

#property tax input and related calculation
property_tax = 3.0
#property tax in dollar = (home_price * (property_tax/100)) / 12 reflects per year dollar value
property_tax_dollar = (home_price * (property_tax/100.0)) / 12.0
property_tax_dollar_array = []

#Home Insurance input and related calculation
home_insurance_year = 1260.0
home_insurance_month = home_insurance_year / 12.0
home_insurance_month_array = []

#HOA -> Home owners association fee per month and calculation
new_home = True
if new_home:    
    if ((home_price >= 200000) and (home_price <= 300000)):
        hoa_per_month = float(700.0/12.0)
    else:
        hoa_per_month = input("Insert HOA per month: ")
else:
    if ((home_price >= 200000) and (home_price <= 300000)):
        hoa_per_month = float(1000.0/12.0)
    else:
        hoa_per_month = input("Insert HOA per month: ")
        
hoa_per_month = 40.0
hoa_per_month_array = []

#other lists
loan_amount = []
monthly_payment_thirty_year = []
total_pay_per_month = []
count = 0

#creating excel file and editing initially
wb = Workbook()

# grab the active worksheet
ws = wb.active

#initial 
ws.append(["Home Price", home_price])
ws.append(["Interest rate in percentage", str(interest_rate) + "%"])
ws.append(["Property tax in percentage", str(property_tax) + "%"])
ws.append(["Down Payment in percentage", "5%", "10%", "15%", "20%"])



for i in down_payment_percentage:
    #print(i)
    #calculating downpayment
    down_payment_dollar.append(home_price * (float(i)/100.0))

    #calculating total loan
    loan_amount.append(float(home_price - (home_price * (float(i)/100.0))))
    #print('loan amount: ', loan_amount[count])
    
    #calculating monthly paymnet toward total loan
    monthly_payment_thirty_year.append(MonthlyMortagePayment(loan_amount[count], interest_rate, FixedLoanType.thirty_year))
    #print("per month payment: ", monthly_payment_thirty_year[count])   
    
    #appendng propetry tax, home insurance and HOA
    property_tax_dollar_array.append(property_tax_dollar)
    home_insurance_month_array.append(home_insurance_month)
    hoa_per_month_array.append(hoa_per_month)
    
    #calculating total monthly payment based on loan, property tax, home onsurance and HOA
    total_pay = monthly_payment_thirty_year[count] + property_tax_dollar_array[count] + home_insurance_month_array[count] + hoa_per_month_array[count]
    total_pay_per_month.append(total_pay)
    count = count + 1

ws.append(["Down payment"] + down_payment_dollar)
ws.append(["Loan Amount"] + loan_amount)
ws.append(["Monthly payment 30 Years Toward just home loan"] + monthly_payment_thirty_year)
ws.append(["Monthly payment for Property tax"] + property_tax_dollar_array)
ws.append(["Monthly payment for Home Insurance"] + home_insurance_month_array)
ws.append(["Monthly payment for HOA"] + hoa_per_month_array)
ws.append(["Monthly payment for 30 year fixed loan"] + total_pay_per_month)


wb.save("result_of_analysis.xlsx")