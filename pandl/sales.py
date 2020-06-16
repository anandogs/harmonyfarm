import gspread
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from render import render_mpl_table
from matplotlib import pyplot as plt

def get_month_range():
    today = datetime.today()
    two_months_ago = (datetime.now() - relativedelta(months=2)).replace(day=1)
    year_ago_end = (datetime.now() - relativedelta(months=12))
    year_ago_start = (year_ago_end - relativedelta(months = 2)).replace(day=1)
    return two_months_ago, today, year_ago_start, year_ago_end

def get_kgs(a, b):
    if a == 'Egg(s)':
        return b/10
    else: 
        return b

def finyr(date):
    if date < pd.to_datetime('1-Apr-2020'):
        return 20
    return 21

def main():
        
    cystart, cyend, pystart, pyend = get_month_range()

    gc = gspread.service_account('/Users/anandoghose/Documents/harmony-new.json')

    wks = gc.open('Harmony produce')

    #data = wks.get_all_records()

    sales = wks.worksheet("Sales (Linked)")

    sales_data = sales.get_all_values()

    sales_headers = sales_data.pop(0)

    sales_df = pd.DataFrame(sales_data, columns=sales_headers)

    sales_df.loc[:,['Quantity Sold (in KGs)', 'Price Per KG / Unit']] = sales_df.loc[:,['Quantity Sold (in KGs)', 'Price Per KG / Unit']].astype(float)

    sales_df['Total Sales'] = sales_df['Quantity Sold (in KGs)']* sales_df['Price Per KG / Unit']

    sales_df['Amount in KGs'] = sales_df.apply(lambda row: get_kgs(row['Produce Name'], row['Quantity Sold (in KGs)']), axis=1)


    #=IF(C2="Egg(s)",D2/10,D2)


    sales_df['Timestamp'] = pd.to_datetime(sales_df['Timestamp'])

    sales_df['MonthA'] = pd.DatetimeIndex(sales_df['Timestamp']).month

    sales_df['finyr'] = sales_df.apply(lambda row: finyr(row['Timestamp']), axis=1)

    sales_cy = sales_df[(sales_df['Timestamp'] > cystart) & (sales_df['Timestamp'] < cyend)]
    sales_py = sales_df[(sales_df['Timestamp'] > pystart) & (sales_df['Timestamp'] < pyend)]

    filtered_sales = sales_py.append(sales_cy)

    sales_by_month = filtered_sales.pivot_table(values='Total Sales', index='Produce Name', columns=['finyr', 'MonthA'], aggfunc='sum', margins=True, margins_name='Total',  fill_value=0 ).round(0).astype(int)
    sales_by_month.reset_index(inplace = True)
    sales_by_month.replace(0,'', inplace=True)

    #fig, ax = render_mpl_table(sales_by_month, header_columns=1, col_width=3.5)

    #fig.savefig('foo.png')

    expenses = wks.worksheet("Expenses")

    expenses.update_acell('D2',0)

    expenses_data = expenses.get_all_values()

    expenses_headers = expenses_data.pop(0)

    expenses_df = pd.DataFrame(expenses_data, columns=expenses_headers)

    expenses_df['Timestamp'] = pd.to_datetime(expenses_df['Timestamp'])

    expenses_df['Expense Amount'] = expenses_df['Expense Amount'].astype(float)

    start_of_year =  pd.to_datetime('1-Apr-2020')
    ytd_expense = expenses_df[(expenses_df['Timestamp'] > start_of_year) & (expenses_df['Timestamp']<cyend) & (expenses_df['Type of Expense'] != 'CAPEX')].sum()['Expense Amount']
    ytd_sales = sales_df[(sales_df['Timestamp'] > start_of_year) & (sales_df['Timestamp']<cyend)].sum()['Total Sales']
    ytd_profit = ytd_sales - ytd_expense

    expenses.update_acell('D2',ytd_profit)
    expenses_df.loc[0,'Expense Amount'] = -ytd_profit

    expenses_cy = expenses_df[(expenses_df['Type of Expense'] == 'COGS') | (expenses_df['Type of Expense'] == 'OPEX Farm Improvements')]
    expenses_cy = expenses_cy[(expenses_cy['Timestamp'] > cystart) & (expenses_cy['Timestamp'] < cyend)]

    current_month_name = cyend.strftime("%B")
    start_month_name = cystart.strftime("%B")
    sales_des = 'Sales from '+start_month_name+' till '+current_month_name
    sales = sales_cy['Total Sales'].sum()
    date = cyend.strftime("%m/%d/%Y")

    summary_df = {'Date': date, 'Description': sales_des, 'Amount': sales}
    summary_df = pd.DataFrame(summary_df, index=[0])
    
    exp_summary = -expenses_cy.groupby(by='Expense Description').sum()
    exp_summary.reset_index(inplace = True)
    exp_summary['Date'] = date
    exp_summary.rename(columns={'Expense Description': 'Description', 'Expense Amount': 'Amount'}, inplace=True)

    summary_df = summary_df.append(exp_summary)

    summary_df = summary_df.append(summary_df.sum(numeric_only=True), ignore_index=True)
    pandl_des = 'Profit / Loss from '+start_month_name+' till '+current_month_name
    summary_df.iloc[-1,1] = pandl_des
    summary_df['Amount'] = summary_df['Amount'].astype(int)

    capex = expenses_df[expenses_df['Type of Expense'] == 'CAPEX']
    peanut_balance = int(-capex[capex['CAPEX Type'] == 'Peanuts'].sum()['Expense Amount'])
    trees_balance = int(-capex[capex['CAPEX Type'] == 'Trees'].sum()['Expense Amount'])

    with open("harmony.html", 'w') as _file:
        _file.write('<style> table{ text-align: center; width:100%}</style><link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous"><div class="container"><h1>Harmony P&L</h1><br><h2>'+ sales_des + ' compared with same period last year</h2>' + sales_by_month.to_html() + '<br><br><h2>'+ pandl_des +'</h2>' + summary_df.to_html()+'<br><br><h3>Remaining Peanut CAPEX to recoup: '+str(int(peanut_balance))+'<br>Remaining Tree CAPEX to recoup: '+str(int(trees_balance))+'</h3></div>')

if __name__ == '__main__':
    main()
