# -*- coding: utf-8 -*-
"""
Created on Wed Dec 14 10:19:25 2022

@author: akashdwivedi
"""

import copy
# import smtplib
import win32com.client as win32
from datetime import datetime
import os
import re 
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)



import random
import pandas as pd 
import copy
df = pd.read_csv('secretsanta.csv') #Enter your .csv form file name here 
mylist = df['Name'].tolist()
names=mylist
emails=(df['Name']+" "+df['Email']).tolist()
num=df['Phone_Number'].tolist()
email2=df['Email'].tolist()

def secret_santa(names):
    my_list = names
    choose = copy.copy(my_list)
    result = []
    for i in my_list:
        names = copy.copy(my_list)
        names.pop(names.index(i))
        chosen = random.choice(list(set(choose)&set(names)))
        result.append((i,chosen))
        choose.pop(choose.index(chosen))
    return result



ss_result = secret_santa(names)

ss_result = [list(ele) for ele in ss_result]

final = zip(ss_result,emails,num,email2)


df2 = pd.DataFrame(final, columns = ['Name', 'Name2','Phone_Number','email2'])

df2[['Reciver','Sender']] = pd.DataFrame(df2.Name.tolist(), index= df2.index)


df2['Name']=df2['Name'].astype(str)






df3=pd.merge(df2,df,on="Phone_Number")



df3=df3.drop(['Name_x', 'Name2','Name_y'], axis=1)

df3.rename(columns = {'Sender':'Name'}, inplace = True)



df4=pd.merge(df3,df,on="Name")


#df.dtypes
#df2.dtypes
#df3.dtypes
df4.dtypes

df4=df4.drop(['Email_x','Address_y','Phone_Number_y',
              'Wishlist_Gift_1_y','Wishlist_Gift_2_y','Wishlist_Gift_3_y'],axis=1)

df4.rename(columns = {'Phone_Number_x':'Reciver_Phone_Number' ,
                      'email2':'Reciver_email',
                      'Name':'Sender_Name','Email_y':'Sender_email'
                      ,'Address_x':'Reciver_Address'}, inplace = True)

df4.to_excel('Gifting_Final.xlsx',index=0)

for sendername,sender_email,reciver_name,recadd,recnum,item1,item2,item3 in zip(df4['Sender_Name'], df4['Sender_email'], df4['Reciver'],df4['Reciver_Address'],df4['Reciver_Phone_Number'],df4['Wishlist_Gift_1_x'],df4['Wishlist_Gift_2_x'],df4['Wishlist_Gift_3_x']):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = sender_email
    mail.Subject = 'Week of Gifting Buddy Announcement'
    mail.HTMLBody = """\
    <html>
        <head></head>
        <body>
            <p>Hi – """+ sendername +" 👋!<br> This is an automated email from Week Of Gifting Fun SPOCS .\n<br>You drew<br>✨✨✨✨<b>"+ reciver_name +"""</b>✨✨✨✨<br> \n 
            <ul>
            <li>Rule Number 1: Please do not tell anyone! 🤫\n  <br>
    <li> Rule Number 2: The budget is min 500 and max 1000 INR ! 👛 \n <br></ul> <br><i>What are you waiting for? Go ahead and get something nice for  """ + reciver_name +""" <br> </i>\n 
    <b>Following is the gifting details of  """ +reciver_name+"""</b ><rr>
     "<ul>
     <li> <b>Address</b> 🏠:- """ + recadd + """  <br>
     <li> <b> Phone_Number ☎️ </b>:- """+ str(recnum)+"""  <br>
      
      </ul>
      <b> Wish list Items 🌠 </b><br> 
      <ul>
    <li> <b>"""  + item1 + """\ </b> <br>
    <li> <b>""" + item2 + """</b><br>
    <li> <b> """ + item3 + """"</b><br>
       </ul>
            </p>
        </body>
    </html>
    """
    mail.Send()
# -*- coding: utf-8 -*-
"""
Created on Wed Dec 14 10:19:25 2022

@author: akashdwivedi

"""
