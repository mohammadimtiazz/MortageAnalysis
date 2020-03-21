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
        monthly_payment = loan_amount * ( (c * (1+c)**total_num_payment) / (((1+c)**total_num_payment) - 1) )
        
    if (FixedLoanType == FixedLoanType.fifteen_year):
        total_num_payment = 15*12
        c = ((interest_rate/100)/12)
        monthly_payment = loan_amount * ( (c * (1+c)**total_num_payment) / (((1+c)**total_num_payment) - 1) )        
    
    return monthly_payment
        
        
    

home_price = input("Insert Home Price: ")
home_price = int(home_price)
down_payment = input("Insert down payment as percentage (x%): ")
down_payment = float(down_payment)
interest_rate = input("Insert interest rate as percentage (x%): ")
interest_rate = float(interest_rate)

loan_amount = home_price - (home_price * (down_payment/100))

print('loan amount: ',loan_amount)

monthly_payment = MonthlyMortagePayment(loan_amount, interest_rate, FixedLoanType.thirty_year)
print("per month payment: ", monthly_payment)


wb = Workbook()

# grab the active worksheet
ws = wb.active

ws.append(["Home Price", home_price])
ws.append(["Down Payment in percentage", down_payment])
ws.append(["Interest rate as percentage", interest_rate])


# # Data can be assigned directly to cells
# ws['A1'] = 42

# # Rows can also be appended
# ws.append([1, 2, 3])

# # Python types will automatically be converted
# import datetime
# ws['A2'] = datetime.datetime.now()

# Save the file
wb.save("result_of_analysis.xlsx")