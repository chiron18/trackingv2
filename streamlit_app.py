import pandas as pd
import streamlit as st
import numpy as np
import openpyxl
import datetime  # Import the datetime module


st.set_page_config(layout="wide")

def appendDictToDF(df,dictToAppend):
  df = pd.concat([df, pd.DataFrame.from_records([dictToAppend])])
  return df

def expenses():
    file = '2023_expenses.xlsx'
    return file

def payments():
    file = '2023_payments.xlsx'
    return file

EXPENSE_FILE=expenses()
PAYMENT_FILE=payments()

input_df = pd.read_excel(EXPENSE_FILE)  
payments_df=pd.read_excel(PAYMENT_FILE)

input_df['Paid']=input_df['Paid'].astype(float)
input_df['Paid']=input_df['Paid'].round(2)
payments_df['Amount']=payments_df['Amount'].astype(float)

expenses_df = input_df.copy()
expenses_df['Split_count'] = expenses_df[['Greg', 'Ian', 'Jerry','Peter','Jason','Brent','Kellen']].sum(axis=1)
expenses_df['Amount_per_split'] = expenses_df['Paid']/expenses_df['Split_count'] 
expenses_df['Amount_per_split']=expenses_df['Amount_per_split'].astype(float)

owes_df = pd.DataFrame(columns=['Situation','Amount', 'Item']) # create empty dataframe

for row in range(0,expenses_df.shape[0]):
    for people in range(3,10):
        item = expenses_df.iat[row,0]
        paid = expenses_df.iat[row,1]
        paid_by = expenses_df.iat[row,2]
        count = expenses_df.iat[row,11]
        amount = paid/count
        debtor = expenses_df.columns[people]
        situation = debtor + " owes " + paid_by

        if (expenses_df.iat[row,people] == True):
            owes_df= appendDictToDF(owes_df,{'Situation':situation,'Amount': amount, 'Item': item })

for row in range(0,payments_df.shape[0]):
    payer = payments_df.iat[row,0]
    print(payer)
    payee = payments_df.iat[row,1]
    payment = payments_df.iat[row,2]
    item = "payment"
    situation = payee + " owes " + payer
    owes_df= appendDictToDF(owes_df,{'Situation':situation,'Amount': payment, 'Item': item })

group_owe = owes_df.groupby(['Situation'], as_index = False)['Amount'].sum()
lookup = group_owe.copy()
group_owe['Inverse'] = group_owe.Situation.str.split(' owes ').str[1] +" owes " + group_owe.Situation.str.split(' owes ').str[0]
mapLookup = dict(lookup[['Situation', 'Amount']].values)
group_owe['Inverse Amount'] = group_owe['Inverse'].map(mapLookup)
group_owe['Inverse Amount'] = group_owe['Inverse Amount'].replace(np.nan, 0)
group_owe['Final Amount']=group_owe['Amount'] -group_owe['Inverse Amount']
group_owe = group_owe[group_owe['Final Amount'] > 0]
group_owe['Amount']=group_owe['Amount'].round(2)
group_owe['Inverse Amount']=group_owe['Inverse Amount'].round(2)
group_owe['Final Amount']=group_owe['Final Amount'].round(2)

final_owe = group_owe.copy()
final_owe = final_owe.drop(['Amount', 'Inverse', 'Inverse Amount'], axis=1)

page_bg_img = """
<style>
[data-testid="block-container"]{
background-image: url(https://tigerarms.ca/wp-content/uploads/2022/09/egwuapfebhyoiogxiudc.jpg);
background-size: cover;
background-position: right left;
background-attachment: fixed;
}
</style>
"""

st.markdown(page_bg_img, unsafe_allow_html=True)

st.header('2023 Expense Tracking')

with open("hunt.css") as source_des:
    st.markdown(f"<style>{source_des.read()}</style>", unsafe_allow_html=True)

st.write("Expenses Submitted")
st.caption('Enter expense details, the checkmarks under names indicate the persons that will share the expense')
edited_df = st.data_editor(input_df, num_rows = "dynamic")

st.write("Payments Made")
st.caption('Enter payment details')
payment_edited_df = st.data_editor(payments_df, num_rows = "dynamic")

if st.button('Save Changes (manually refresh page to update final tally)'):
    edited_df.to_excel(EXPENSE_FILE, index=False)
    payment_edited_df.to_excel(PAYMENT_FILE, index=False)

st.write("Final Tally")
st.dataframe(final_owe, hide_index=True)

# Create a download button
if st.button('Generate Excel File (in memory) - this is a two button process'):
    # Create an Excel writer object
    excel_writer = pd.ExcelWriter('tempFile.xlsx', engine='openpyxl')
    
    # Write each DataFrame to a separate sheet
    expenses_df.to_excel(excel_writer, sheet_name='Expenses', index=False)
    payments_df.to_excel(excel_writer, sheet_name='Payments', index=False)
    final_owe.to_excel(excel_writer, sheet_name='Final_Tally', index=False)
    owes_df.to_excel(excel_writer, sheet_name='Situations', index=False)
    group_owe.to_excel(excel_writer, sheet_name='Situation_Summary', index=False)
    
    # Save the Excel file
    excel_writer.close()
    
    # Provide a link to download the Excel file
    st.download_button(
        label='Click here to download the generated Excel file',
        data=open('tempFile.xlsx', 'rb'),
        file_name=f'2023_Anzac_Expenses_{datetime.datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


