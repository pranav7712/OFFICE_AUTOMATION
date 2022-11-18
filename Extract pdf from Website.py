#!/usr/bin/env python
# coding: utf-8

# In[4]:


import os

def extract_url_pdf(input_url,folder_path=os.getcwd()):
    
    import os
    import requests
    from urllib.parse import urljoin
    from bs4 import BeautifulSoup
    import pandas as pd
    import datetime
    
    url = input_url

    #If there is no such folder, the script will create one automatically
    folder_location = folder_path
    if not os.path.exists(folder_location):os.mkdir(folder_location)

    response = requests.get(url)
    soup= BeautifulSoup(response.text, "html.parser") 

    link_text=list()
    link_href=list()
    link_file=list()
    
    counter=0

    for link in soup.select("a[href$='.pdf']"):
        #Name the pdf files using the last portion of each link which are unique in this case
        
        filename = os.path.join(folder_location,link['href'].split('/')[-1])
        with open(filename, 'wb') as f:
            f.write(requests.get(urljoin(url,link['href'])).content)
            
        link_text.append(str(link.text))
        
        link_href.append(link['href'])

        link_file.append(link['href'].split('/')[-1])
        
        counter+=1

        print(counter, "-Files Extracted from URL named ",link['href'].split('/')[-1])
        
    table_dict={"Text":link_text,"Url_Link":link_href,"File Name":link_file}

    df=pd.DataFrame(table_dict)
    
    time_stamp = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    
    print("Creating an Excel file with Name of FIle, Url Link and Link Text...")
    

    new_excel_file=os.path.join(folder_location,"Excel_Output_"+time_stamp+".xlsx")

    writer = pd.ExcelWriter(new_excel_file, engine='openpyxl')

    df.to_excel(writer,sheet_name="Output")

    
    writer.save()


    print("All Pdf files downloaded and Excel File Created")


# In[2]:



#urls to try:

# https://www.icai.org/category/bos-important-announcements
# https://www.icai.org/post.html?post_id=17843
#https://www.icai.org/post.html?post_id=17825
# https://cbic-gst.gov.in/central-tax-notifications.html
# https://trends.builtwith.com/websitelist/Responsive-Tables


# In[5]:


extract_url_pdf(input_url="https://www.icai.org/post.html?post_id=17843")


# In[6]:


extract_url_pdf(input_url="https://cbic-gst.gov.in/central-tax-notifications.html")

# https://cbic-gst.gov.in/central-tax-notifications.html


# In[ ]:

extract_url_pdf(input_url="https://www.icai.org/post.html?post_id=17827")





