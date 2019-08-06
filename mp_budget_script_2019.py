#!/usr/bin/env python
# coding: utf-8

# In[10]:


# import modules
import pandas as pd


# In[11]:


#Read data from Expenses sheet into data frame
expenses = pd.read_excel('FY_19_reports/2019_06.xlsx',"EX", usecols=[2,4,13,14,16,18,19],skiprows=11)
#insert Accounting Type Column
expenses.insert(0,'Accounting Type',"Expenses")
#replace NaNs with blanks
expenses.fillna('',inplace=True)
#drop empty rows
expenses.dropna(how='all')
#rename columns
expenses.columns = ['Accounting Type','Accounting Date','Account','Employee','Vendor','Amount','Comments','Transaction Date']
#re-order columns
expenses = expenses[['Accounting Type','Transaction Date','Accounting Date','Account','Employee','Vendor','Amount','Comments']]
print(expenses.head())


# In[12]:


journals = pd.read_excel('FY_19_reports/2019_06.xlsx',"Journals", usecols=[3,7,9,16,17],skiprows=11)
journals.insert(0,'Accounting Type','Journal Entry')
journals.fillna('',inplace=True)
journals.dropna(how='all')
journals.insert(1, 'Transaction Date',"")
journals.insert(3, 'Employee',"")
journals.columns = ['Accounting Type','Transaction Date','Accounting Date','Employee','Vendor','Account','Amount','Comments']
journals = journals[['Accounting Type','Transaction Date','Accounting Date','Account','Employee','Vendor','Amount','Comments']]
#### strip out salaries
journals[~journals.Account.str.contains("SALARIES")]
print(journals.head())


# In[13]:


accounts_payable = pd.read_excel('FY_19_reports/2019_06.xlsx',"AP", usecols=[3,11,13,14,15],skiprows=11)
accounts_payable.insert(0,'Accounting Type','Accounts Payable')
accounts_payable.fillna('',inplace=True)
accounts_payable.dropna(how='all')
accounts_payable.insert(1, 'Transaction Date',"")
accounts_payable.insert(3, 'Employee',"")
accounts_payable.insert(8, 'Comments',"")
accounts_payable.columns = ['Accounting Type','Transaction Date','Account','Employee','Vendor','Description','Amount','Accounting Date','Comments']
accounts_payable = accounts_payable[['Accounting Type','Transaction Date','Accounting Date','Account','Employee','Vendor','Amount','Comments']]
print(accounts_payable.head())


# In[15]:


full_data = pd.concat([expenses,accounts_payable,journals])
full_data.head()


# In[21]:


#full_data.to_csv('out.csv')

with open('FY18-19_Master_Data.csv', 'a') as f:
    full_data.to_csv(f, mode='a', header=f.tell()==0)


# In[ ]:
