#!/usr/bin/env python
# coding: utf-8

# In[5]:


import os

def extract_url_table(input_url,folder_path=os.getcwd()):

    import pandas as pd
    import datetime

    url = input_url

    # Assign the table data to a Pandas dataframe
    table = pd.read_html(url)[0]

    # Print the dataframe
    time_stamp = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    
    new_excel_file=os.path.join(folder_path,"Excel_Table_Output_"+time_stamp+".xlsx")

    writer = pd.ExcelWriter(new_excel_file, engine='openpyxl')

    table.to_excel(writer,sheet_name="Output")
    
    writer.save()


    print("Table in Url Converted to Excel File and stored in.." ,new_excel_file)
    
    
    


# In[6]:



#urls to try:

# https://www.icai.org/category/bos-important-announcements
# https://www.icai.org/post.html?post_id=17843
#https://www.icai.org/post.html?post_id=17825
# https://cbic-gst.gov.in/central-tax-notifications.html
# https://trends.builtwith.com/websitelist/Responsive-Tables


# In[7]:


extract_url_table("https://cbic-gst.gov.in/central-tax-notifications.html")


# In[8]:


extract_url_table("https://trends.builtwith.com/websitelist/Responsive-Tables")


# In[ ]:




