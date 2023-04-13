#!/usr/bin/env python
# coding: utf-8

# In[18]:


#IMPORT DEPENDENCIES 

import pandas as pd

import streamlit as st
import io 

from datetime import date
from st_aggrid import AgGrid


buffer =io.BytesIO()

st.markdown(
        f"""
<style>
    .reportview-container .main .block-container{{
        max-width: 90%;
        padding-top: 5rem;
        padding-right: 5rem;
        padding-left: 5rem;
        padding-bottom: 5rem;
    }}
    img{{
    	max-width:40%;
    	margin-bottom:40px;
    }}
</style>
""",
        unsafe_allow_html=True,
    )

header_container = st.container()
stats_container = st.container()	
#######################################



# You can place things (titles, images, text, plots, dataframes, columns etc.) inside a container
with header_container:

	# for example a logo or a image that looks like a website header
	st.image('images/logo.jpg')

	# different levels of text you can include in your app
	st.title("Consolidated Leads")
   
uploaded_file = st.file_uploader('Upload Excel file for processing', type=["xlsx"])
if uploaded_file is not None:
  df= pd.read_excel(uploaded_file)
  AgGrid(df.head(10))
else:
 st.warning("You need to upload an Excel file")

with stats_container:
  df['E-mail']= df['E-mail'].astype(str).str.lower()	
  df.drop_duplicates(subset=['E-mail','Program Version Name'],inplace = True)
  df.drop_duplicates(subset=['E-mail'], inplace= True) 
  df_noDups = df.copy()


# In[5]:




    
def pivot_table_w_subtotals(df, values, indices, columns, aggfunc, fill_value):
    '''
    Adds tabulated subtotals to pandas pivot tables with multiple hierarchical indices.
    
    Args:
    - df - dataframe used in pivot table
    - values - values used to aggregrate
    - indices - ordered list of indices to aggregrate by
    - columns - columns to aggregrate by
    - aggfunc - function used to aggregrate (np.max, np.mean, np.sum, etc)
    - fill_value - value used to in place of empty cells
    
    Returns:
    -flat table with data aggregrated and tabulated
    
    '''
    listOfTable = []
    for indexNumber in range(len(indices)):
        n = indexNumber+1
        table = pd.pivot_table(df,values=values,index=indices[:n],columns=columns,aggfunc=aggfunc,fill_value=fill_value).reset_index()
        for column in indices[n:]:
            table[column] = ''
        listOfTable.append(table)
    concatTable = pd.concat(listOfTable).sort_index()
    concatTable = concatTable.set_index(keys=indices)
    return concatTable.sort_index(axis=0,ascending=True)

# In[8]:


# delete all rows with column ' [Lead Name]'' contains 'TEST/SHAKEEL/ANTHOLOGY/YASTEEL'
#pattern= 'test|shakeel|anthology|yasteel'

#df.drop(df[df['Lead Name'].str.contains(pattern, case= False,na=False)].index,inplace=True)
#df.drop(df[df['E-mail'].str.contains(pattern, case= False,na=False)].index,inplace=True)

mask= 'puleng|thokozani'
df.drop(df[df['Owner'].str.contains(mask, case= False,na=False)].index,inplace=True)
today = date.today().strftime('%Y/%m/%d')
df['Created On']= pd.to_datetime(df['Created On'],format='%Y/%m/%d')
df.drop(df[df['Created On'] >= today].index,inplace=True)
df_noJunk= df.copy()



   





# In[48]:


#Process 4a
negleads= 'Already dealing with REGENT|Does not Qualify|Existing RBS Student|General Enquiry|Invalid Lead|Invalid Number|No Finance|Not Interested|NSFAS|Programme not Offered|Registered Elsewhere|Looking for Job|Nursing Degree|Teaching Degree|Do not call student'
NegativeLeads= df[df['Program Version Name'].isnull() & df['Stage Step'].str.contains(negleads, case= False,na=False)]
df.drop(df[df['Program Version Name'].isnull() & df['Stage Step'].str.contains(negleads, case= False,na=False)].index,inplace=True)
#Process 4b
pivotnegleads= pd.pivot_table(NegativeLeads, values='Lead Name',index='Stage Step', columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
 



# In[59]:


#Process 5a
access= 'Access for Success in Accounting'
AccessLeads=df[df['Program Version Name'].str.contains(access, case= False,na=False)]
  
   
df.drop(df[df['Program Version Name'].str.contains(access, case= False,na=False)].index,inplace = True)
#Process 5b
pivotaccessleads= pd.pivot_table(AccessLeads,values='Lead Name',index='Stage Step', columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
 
   


# In[ ]:


#Process 6 
school= 'schools'
schoolLeads = df[df['UTM Campaign'].str.contains(school, case= False,na=False) | df['UTM Medium'].str.contains(school, case= False,na=False)  | df['UTM Source'].str.contains(school, case= False,na=False)]
df.drop(schoolLeads.index,inplace=True)


# In[84]:


#process 7
df['Campus'].replace(to_replace=['Durban','Pietermaritzburg'],value='KZN',inplace=True)
df['Campus'].replace(to_replace=['Johannesburg'],value='JHB',inplace=True)
df['Campus'].replace(to_replace=['Cape Town'],value='CT',inplace=True)
df['Campus'].replace(to_replace=['East London'],value='EL',inplace=True)
df['Campus'].replace(to_replace=['Pretoria'],value='PTA',inplace=True)
df['Campus'].replace(to_replace=['Manzini','Mbabane'],value='SWAZILAND',inplace=True)
df['Campus'].replace(to_replace=['Ongwediva','Windhoek'],value='NAMIBIA',inplace=True)
df['Campus'].replace(to_replace=['Kimberly','Bloemfontein','Nelspruit','Polokwane'],value='OTHER',inplace=True)
df_cleanedRegion= df.copy()

#process 7b
pivotRegions= pd.pivot_table(df,values='Lead Name',index='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
  
  


# In[64]:





# In[70]:


#pieRegions=df['Campus'].value_counts().plot(kind='pie', autopct='%1.1f%%')

 
# In[85]:


# Helper function to insert 'Headings' into Excel cells



# In[11]:

dfl=df.copy()
dfl.fillna('(blank)',inplace=True)
pivotLeadsA= pd.pivot_table(dfl,values='Lead Name',index='Stage Step',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedPivotLeadsA=pivotLeadsA.sort_values(by=['Lead Name'],ascending= False)
dfl['Stage Step'].replace(to_replace=['Interested information sent','Online App Pending','Request a callback','Waiting for Results','Jan 2023','Waiting for Matric Results','Firm Offer','Appointment Set','DBA Enquiry'],value='POSITIVE',inplace=True)
dfl['Stage Step'].replace(to_replace=['Not Interested','Already dealing with REGENT','No Finance','General Enquiry','Invalid Number','Registered Elsewhere','NSFAS','Programme not Offered','Does not Qualify','Invalid Lead','Do Not Call Student','Looking for Job','Existing RBS Student','Teaching Degree','Nursing Degree'],value='NEGATIVE',inplace=True)
dfl['Stage Step'].replace(to_replace=['Still in Matric','Jul-23','Jan-24','July 2024','(blank)'],value='NEUTRAL',inplace=True)
dfl['Stage Step'].replace(to_replace=['No Answer','Voicemail'],value='NOT CONTACTABLE',inplace=True)
pivotLeadsDisposition= pd.pivot_table(dfl,values='Lead Name',index='Stage Step',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
 



#Process8b
df.fillna('(blank)',inplace=True)
pivotProgramme= pd.pivot_table(df,values='Lead Name',index='Program Version Name',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedPivotProgramme=pivotProgramme.sort_values(by=['Lead Name'],ascending= False)

#Process8c
pivotProgramme2= pd.pivot_table(df,values='Lead Name',index='Program Version Name', columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedPivotProgramme2=pivotProgramme2.sort_values(by=['Grand Total'],ascending= False)



#Process10a
todayL = '2022/09/12'
df['Created On']= pd.to_datetime(df['Created On'],format='%Y/%m/%d')
priorLeads = df[df['Created On'] < todayL]
df.drop(df[df['Created On'] < todayL].index,inplace=True)
newLeads= df.copy()

#Process10b

pivotpriorLeads= pd.pivot_table(priorLeads,values='Lead Name',index='Stage Step',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedPivotpriorLeads=pivotpriorLeads.sort_values(by=['Lead Name'],ascending= False)
priorLeads['Stage Step'].replace(to_replace=['Interested information sent','Online App Pending','Request a callback','Waiting for Results','Jan 2023','Waiting for Matric Results','Firm Offer','Appointment Set','DBA Enquiry'],value='POSITIVE',inplace=True)
priorLeads['Stage Step'].replace(to_replace=['Not Interested','Already dealing with REGENT','No Finance','General Enquiry','Invalid Number','Registered Elsewhere','NSFAS','Programme not Offered','Does not Qualify','Invalid Lead','Do Not Call Student','Looking for Job','Existing RBS Student','Teaching Degree','Nursing Degree'],value='NEGATIVE',inplace=True)
priorLeads['Stage Step'].replace(to_replace=['Still in Matric','Jul-23','Jan-24','July 2024','(blank)'],value='NEUTRAL',inplace=True)
priorLeads['Stage Step'].replace(to_replace=['No Answer','Voicemail'],value='NOT CONTACTABLE',inplace=True)
pivotPriorLeadsDisp= pd.pivot_table(priorLeads,values='Lead Name',index='Stage Step',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivotPriorLeadsDisp=pivotPriorLeadsDisp.sort_values(by=['Lead Name'],ascending= False)


#Process10c

pivot10c= pd.pivot_table(priorLeads,values='Lead Name',index='Program Version Name',columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivot10c = pivot10c.sort_values(by=['Grand Total'],ascending= False)
pivot10cA= pd.pivot_table(priorLeads,values='Lead Name',index='Program Version Name',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivot10cA = pivot10cA.sort_values(by=['Lead Name'],ascending= False)

#Process9new
pivotnewLeads= pd.pivot_table(newLeads,values='Lead Name',index='Stage Step',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedPivotnewLeads=pivotnewLeads.sort_values(by=['Lead Name'],ascending= False)
pivot9c= pd.pivot_table(newLeads,values='Lead Name',index='Program Version Name',columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedPivot9c=pivot9c.sort_values(by=['Grand Total'],ascending= False)
outlierdf= newLeads[~ newLeads['Stage Step'].isin(['Interested information sent','Online App Pending','Request a callback','Waiting for Results','Jan 2023','Waiting for Matric Results','Firm Offer','Appointment Set','DBA Enquiry','Requested a callback','Not Interested','Already dealing with REGENT','No Finance','General Enquiry','Invalid Number','Registered Elsewhere','NSFAS','Programme not Offered','Does not Qualify','Invalid Lead','Do Not Call Student','Looking for Job','Existing RBS Student','Teaching Degree','Nursing Degree','Still in Matric','Jul-23','Jan-24','July 2024','(blank)','B2B Contact','No Answer','Voicemail'])]
newLeads['Stage Step'].replace(to_replace=['Interested information sent','Online App Pending','Request a callback','Waiting for Results','Jan 2023','Waiting for Matric Results','Firm Offer','Appointment Set','DBA Enquiry','Requested a callback'],value='POSITIVE',inplace=True)
newLeads['Stage Step'].replace(to_replace=['Not Interested','Already dealing with REGENT','No Finance','General Enquiry','Invalid Number','Registered Elsewhere','NSFAS','Programme not Offered','Does not Qualify','Invalid Lead','Do Not Call Student','Looking for Job','Existing RBS Student','Teaching Degree','Nursing Degree'],value='NEGATIVE',inplace=True)
newLeads['Stage Step'].replace(to_replace=['Still in Matric','Jul-23','Jan-24','July 2024','(blank)','B2B Contact'],value='NEUTRAL',inplace=True)
newLeads['Stage Step'].replace(to_replace=['No Answer','Voicemail'],value='NOT CONTACTABLE',inplace=True)
pivotnewLeadsDisp= pd.pivot_table(newLeads,values='Lead Name',index='Stage Step',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivotnewLeadsDisp=pivotnewLeadsDisp.sort_values(by=['Lead Name'],ascending= False)

#Process9a
df['Program Version Name'].replace(to_replace=['Bachelor of Commerce in Law','Bachelor of Commerce','Bachelor of Business Administration','Bachelor of Commerce in Supply Chain Management','Bachelor of Public Administration','Bachelor of Commerce in Human Resource Management','Bachelor of Commerce in Accounting','Bachelor of Commerce in Retail Management'],value='DEGREE',inplace=True)
df['Program Version Name'].replace(to_replace=['Higher Certificate in Healthcare Services Management','Higher Certificate in Business Management','Higher Certificate in Supply Chain Management','Higher Certificate in Human Resource Management','Higher Certificate in Accounting','Higher Certificate in Marketing Management','Higher Certificate in Entrepreneurship','Higher Certificate in Retail Management','Higher Certificate in Islamic Finance, Banking and Law','Higher Certificate in Management for Estate Agents'], value='HC',inplace=True)
df['Program Version Name'].replace(to_replace=['Advanced Diploma in Management','Advanced Diploma in Financial Management','Diploma in Public Relations Management','Diploma in Financial Management'],value='DIPLOMA',inplace=True)
df['Program Version Name'].replace(to_replace=['Postgraduate Diploma in Supply Chain Management','Postgraduate Diploma in Project Management','Postgraduate Diploma in Management','Postgraduate Diploma in Educational Management and Leadership','Postgraduate Diploma in Digital Marketing','Postgraduate Diploma in Islamic Finance and Banking','Postgraduate Diploma in Accounting','Bachelor of Commerce Honours','Bachelor of Commerce Honours in Human Resource Management'],value='PG/H',inplace=True)
df['Program Version Name'].replace(to_replace=['Master of Business Administration','Master of Business Administration in Healthcare Management'],value='MBA',inplace=True)
df['Program Version Name'].replace(to_replace=['Doctor of Business Administration'],value= 'DBA',inplace=True)
df['Program Version Name'].replace(to_replace=['(blank)'],value='BLANKS',inplace=True)
df_cleanedProg= df.copy()
pivotProgress= pd.pivot_table(df,values='Lead Name',index='Program Version Name',columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
name_order = ['DEGREE','HC','DIPLOMA','PG/H','MBA','DBA','BLANKS']
sortedPivotProgress=pivotProgress.loc[name_order]
pivotProgress1= pd.pivot_table(df,values='Lead Name',index='Program Version Name',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedPivotProgress1=pivotProgress1.loc[name_order]
df['MM-DD'] = df['Created On'].dt.strftime('%Y/%m/%d')
df['month'] = df['Created On'].dt.month_name()
#pivotDay= pd.pivot_table(df,values='Lead Name',index=['month','MM-DD'],aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
pivotDay= pivot_table_w_subtotals(df=df,values='Lead Name',indices=['month','MM-DD'],columns=[],aggfunc='count',fill_value='')
#sortedPivotDay=pivotDay.sort_values(by=['Lead Name'],ascending= False)

#Process11c
pivot11c= pd.pivot_table(newLeads,values='Lead Name',index='Program Version Name',columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivot11c = pivot11c.sort_values(by=['Grand Total'],ascending= False)
pivot11cA= pd.pivot_table(newLeads,values='Lead Name',index='Program Version Name',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivot11cA = pivot11cA.sort_values(by=['Lead Name'],ascending= False)


#Process12
validation='(blank)'
OrganicLeads= df[df['UTM Campaign'].str.contains(validation, case= False,na=False) & df['UTM Medium'].str.contains(validation, case= False,na=False) & df['UTM Source'].str.contains(validation, case= False,na=False)]
unfiltereddf=  df[~ df['UTM Campaign'].str.contains(validation, case= False,na=False) & df['UTM Medium'].str.contains(validation, case= False,na=False) & df['UTM Source'].str.contains(validation, case= False,na=False)]
df.drop(df[df['UTM Campaign'].str.contains(validation, case= False,na=False) & df['UTM Medium'].str.contains(validation, case= False,na=False) & df['UTM Source'].str.contains(validation, case= False,na=False)].index,inplace=True)
validation1= 'campaign=gmb-jhb|source=google|medium=organic|content=website-link',
'campaign=(organic)|source=Google|medium=search',
'campaign=(direct)|source=(direct)',
'campaign=Canned_incomplete|source=CRM|medium=email',
'campaign=gmb-jhb|source=google|medium=organic|content=',
'campaign=Canned_enquiry-new|source=CRM|medium=email',
'campaign=(organic)|source=Bing|medium=search',
'campaign=(referral)|source=www.icancpd.net|medium=referral|content=/',
'campaign=gmb-ongwe|source=google|medium=organic|content=website-link',
'campaign=(organic)|source=Google|medium=searchcampaign=(organic)|source=Google|medium=search',
'campaign=SCM_articulation|source=Facebook|medium=post',
'campaign=HC_Jan23_Offer|source=Email|medium=email',
'campaign=Enquiries-canned|source=CRM|medium=email',
'campaign=gmb-cpt|source=google|medium=organic|content=website-link',
'campaign=(referral)|source=search-dra.dt.dbankcloud.com|medium=referral|content=/',
'campaign=(organic)|source=Facebook|medium=social',
'campaign=(referral)|source=southafricaportal.com|medium=referral|content=/',
'campaign=ITT_launch|source=Insta|medium=post',
'campaign=(referral)|source=en.m.wikipedia.org|medium=referral|content=/',
'campaign=Jivo_Remarketing|source=Email|medium=email',
'campaign=(referral)|source=m.trendads.co|medium=referral|content=/search/?search_term=regent business school&brand=uncategorized',
'campaign=(referral)|source=find-mba.com|medium=referral|content=/',
'campaign=gmb-pta|source=google|medium=organic|content=website-link',
'campaign=Post|source=Facebook|medium=Social',
'Facebook',
'https://regent.ac.za/programme/postgraduate-diploma-in-accounting',
'https://regent.ac.za/',
'https://regent.ac.za/programme/bachelor-of-commerce-in-supply-chain-management-regent',
'https://regent.ac.za/programme/postgraduate-diploma-in-project-management',
'https://regent.ac.za/my-account',
'https://regent.ac.za/programmes/postgraduate',
'https://regent.ac.za/programme/bachelor-of-business-administration-regent',
'https://regent.ac.za/programme/postgraduate-diploma-in-digital-marketing',
'https://regent.ac.za/logins',
'https://regent.ac.za/programme/advanced-diploma-in-management',
'https://regent.ac.za/programme/diploma-in-public-relations-management',
'https://regent.ac.za/programme/advanced-diploma-in-financial-management',
'https://regent.ac.za/programmes',
'https://regent.ac.za/apply-online?utm_source=CRM&utm_medium=email&utm_campaign=Canned_incomplete',
'https://regent.ac.za/programme/postgraduate-diploma-in-supply-chain-management',
'https://regent.ac.za/contact-us/learning-centres',
'https://regent.ac.za/contact-us/thank-you-qualifying',
'https://regent.ac.za/?utm_source=google&utm_medium=organic&utm_campaign=gmb-jhb&utm_content=website-link',
'https://regent.ac.za/apply-online',
'https://regent.ac.za/apply',
'https://regent.ac.za/programme/higher-certificate-in-health-care-services-management',
'https://regent.ac.za/programme/postgraduate-diploma-in-islamic-finance-and-banking',
'https://regent.ac.za/campus-news/higher-certificate-versus-diploma-versus-degree-help',
'Insta',
'https://regent.ac.za/programme/bachelor-of-commerce-in-human-resource-management',
'https://regent.ac.za/programme/master-of-business-administration-in-healthcare-management',
'https://regent.ac.za/programme/higher-certificate-in-human-resource-management',
'https://regent.ac.za/?utm_source=google&utm_medium=organic&utm_campaign=gmb-cpt&utm_content=website-link',
'https://regent.ac.za/degrees/12-things-you-can-do-with-your-bcom-degree',
'https://regent.ac.za/programmes/undergraduate/diplomas',
'https://regent.ac.za/programmes/short-learning-programmes',
'https://regent.ac.za/about-us',
'https://regent.ac.za/programmes/undergraduate/degrees',
'https://regent.ac.za/programmes?utm_source=CRM&utm_medium=email&utm_campaign=Enquiries-canned',
'https://regent.ac.za/programme/postgraduate-diploma-in-accounting-access-programme',
'https://regent.ac.za/programme/bachelor-of-commerce-in-law',
'https://regent.ac.za/programme/bachelor-of-commerce-honours',
'https://regent.ac.za/programmes/postgraduate/postgraduate-diplomas',
'https://regent.ac.za/programme/higher-certificate-in-business-management',
'https://regent.ac.za/vacancies',
'https://regent.ac.za/retail-management-development-programme-application',
'https://regent.ac.za/programme/mba-master-of-business-administration',
'https://regent.ac.za/programme/bachelor-of-public-administration',
'https://regent.ac.za/programme/bachelor-of-commerce',
'https://regent.ac.za/contact-us',
'https://regent.ac.za/programmes/undergraduate',
'https://regent.ac.za/programme/bachelor-of-commerce-in-accounting',
'https://regent.ac.za/programme/postgraduate-diploma-in-management',
'https://regent.ac.za/programme/higher-certificate-in-supply-chain-management',
'https://regent.ac.za/programme/higher-certificate-in-accounting',
'https://regent.ac.za/programme/higher-certificate-in-entrepreneurship',
'https://regent.ac.za/events',
'https://regent.ac.za/programmes/undergraduate/higher-certificates',
'https://regent.ac.za/about-us/why-choose-regent',
'https://regent.ac.za/programmes/postgraduate/honours',
'https://regent.ac.za/blog/a-21st-century-approach-to-human-resource-management',
'https://regent.ac.za/programme/diploma-in-financial-management',
'https://regent.ac.za/about-us/student-experience',
'https://regent.ac.za/category/distance-learning',
'https://regent.ac.za/locations/regent-eswatini-manzini',
'https://regent.ac.za/programme/master-of-business-administration',
'https://regent.ac.za/programme/bachelor-of-commerce-in-retail-management',
'https://regent.ac.za/programme/bachelor-of-commerce-in-human-resource-management-honours',
'https://regent.ac.za/category/higher-certificates',
'https://regent.ac.za/programme/e-commerce-starting-your-online-business',
'https://regent.ac.za/?s=Honors+project+management',
'https://regent.ac.za/programmes/undergraduate/advanced-diplomas',
'https://regent.ac.za/students/accessing-regentonline',
'https://regent.ac.za/programme/doctor-of-business-administration',
'https://regent.ac.za/?s=fees',
'https://regent.ac.za/programme/higher-certificate-in-marketing-management',
'https://regent.ac.za/#top',
'https://regent.ac.za/join-regent-connect-today',
'https://regent.ac.za/programme/postgraduate-diploma-in-educational-management-and-leadership',
'Jan 2023',
'https://regent.ac.za/programme/higher-certificate-in-business-management#request-form',
'https://regent.ac.za/programme/higher-certificate-in-islamic-finance-banking-and-law',
'https://regent.ac.za/about-us/accreditation',
'https://regent.ac.za/programme/higher-certificate-in-management-for-estate-agents/',
'https://regent.ac.za/programmes/undergraduate/',
'https://regent.ac.za/programme/bachelor-of-business-administration-regent/',
'https://regent.ac.za/programme/bachelor-of-commerce-honours/',
'https://regent.ac.za/programme/higher-certificate-in-accounting/',
'https://regent.ac.za/programme/bachelor-of-commerce/',
'https://regent.ac.za/programme/higher-certificate-in-retail-management/',
'https://regent.ac.za/programme/bachelor-of-commerce-in-accounting/',
'https://regent.ac.za/programme/postgraduate-diploma-in-project-management/',
'https://regent.ac.za/apply-online/',
'https://regent.ac.za/programme/bachelor-of-commerce-in-supply-chain-management-regent/',
'https://regent.ac.za/programme/master-of-business-administration/',
'https://regent.ac.za/programme/higher-certificate-in-business-management/',
'https://regent.ac.za/programme/postgraduate-diploma-in-accounting/',
'https://regent.ac.za/category/degrees/',
'https://regent.ac.za/programme/higher-certificate-in-supply-chain-management/',
'https://regent.ac.za/programmes/postgraduate/',
'Ongwediva Fair'

OrganicLeads1= df[df['UTM Campaign'].str.contains(validation, case= False,na=False) & df['UTM Medium'].str.contains(validation, case= False,na=False) &  df.loc[df['UTM Source'].isin(validation1)]]
unfiltereddf1 = df[~ df.loc[df['UTM Source'].isin(validation1)]]
df.drop(df[df['UTM Campaign'].str.contains(validation, case= False,na=False) & df['UTM Medium'].str.contains(validation, case= False,na=False) & df.loc[df['UTM Source'].isin(validation1)]].index,inplace=True)
OrganicLeads = OrganicLeads.append(OrganicLeads1, ignore_index = True)
unflitereddf= unfiltereddf.append(unfiltereddf1, ignore_index = True)		    
post='post|Banner|Email|Organic|Social'
OrganicLeads2=df[ df['UTM Medium'].str.contains(post, case= False,na=False)]
unfiltereddf2= df[~ df['UTM Medium'].str.contains(post, case= False,na=False)]
unfiltereddf= unfiltereddf.append(unfiltereddf2,ignore_index = True)
df.drop(df[ df['UTM Medium'].str.contains(post, case= False,na=False)].index, inplace=True)
OrganicLeads = OrganicLeads.append(OrganicLeads2, ignore_index = True)
OverallPaid = df.copy()

#Process13
Walkin='Walk In'
OrganicSeg= pd.DataFrame(OrganicLeads)
dfWalk=OrganicSeg[OrganicSeg['Method of contact'].str.contains(Walkin,case=False,na=False)]
OrganicSeg.drop(OrganicSeg[OrganicSeg['Method of contact'].str.contains(Walkin,case=False,na=False)].index,inplace=True)
pivotwalk= pd.pivot_table(dfWalk,values='Lead Name',index='Program Version Name',columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivotwalk = pivotwalk.loc[name_order]

#Process14

Call='Inbound Call'
dfCall=OrganicSeg[OrganicSeg['Method of contact'].str.contains(Call,case=False,na=False)]
OrganicSeg.drop(OrganicSeg[OrganicSeg['Method of contact'].str.contains(Call,case=False,na=False)].index,inplace=True)
pivotcall= pd.pivot_table(dfCall,values='Lead Name',index='Program Version Name',columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivotcall = pivotcall.loc[name_order]

#Process15
Live= 'Live Chat'

UTMSource= 'campaign=gmb-jhb|source=google|medium=organic|content=website-link',
'campaign=(organic)|source=Google|medium=search',
'campaign=(direct)|source=(direct)',
'campaign=Canned_incomplete|source=CRM|medium=email',
'campaign=gmb-jhb|source=google|medium=organic|content=',
'campaign=Canned_enquiry-new|source=CRM|medium=email',
'campaign=(organic)|source=Bing|medium=search',
'campaign=(referral)|source=www.icancpd.net|medium=referral|content=/',
'campaign=gmb-ongwe|source=google|medium=organic|content=website-link',
'campaign=(organic)|source=Google|medium=searchcampaign=(organic)|source=Google|medium=search',
'campaign=SCM_articulation|source=Facebook|medium=post',
'campaign=HC_Jan23_Offer|source=Email|medium=email',
'campaign=Enquiries-canned|source=CRM|medium=email',
'campaign=gmb-cpt|source=google|medium=organic|content=website-link',
'campaign=(referral)|source=search-dra.dt.dbankcloud.com|medium=referral|content=/',
'campaign=(organic)|source=Facebook|medium=social',
'campaign=(referral)|source=southafricaportal.com|medium=referral|content=/',
'campaign=ITT_launch|source=Insta|medium=post',
'campaign=(referral)|source=en.m.wikipedia.org|medium=referral|content=/',
'campaign=Jivo_Remarketing|source=Email|medium=email',
'campaign=(referral)|source=m.trendads.co|medium=referral|content=/search/?search_term=regent business school&brand=uncategorized',
'campaign=(referral)|source=find-mba.com|medium=referral|content=/',
'campaign=gmb-pta|source=google|medium=organic|content=website-link',
'campaign=Post|source=Facebook|medium=Social',

dflive= OrganicSeg[OrganicSeg['UTM Campaign'].str.contains(validation, case= False,na=False) & OrganicSeg['UTM Medium'].str.contains(validation, case= False,na=False) &  OrganicSeg.loc[OrganicSeg['UTM Source'].isin(UTMSource)]]
OrganicSeg.drop(OrganicSeg[OrganicSeg['UTM Campaign'].str.contains(validation, case= False,na=False) & OrganicSeg['UTM Medium'].str.contains(validation, case= False,na=False) & OrganicSeg.loc[OrganicSeg['UTM Source'].isin(UTMSource)]].index,inplace=True)
dflive1=OrganicSeg[OrganicSeg['Method of contact'].str.contains(Live,case=False,na=False) & OrganicSeg['UTM Campaign'].str.contains(validation, case= False,na=False) & OrganicSeg['UTM Medium'].str.contains(validation, case= False,na=False) & OrganicSeg['UTM Source'].str.contains(validation, case= False,na=False)]
OrganicSeg.drop(OrganicSeg[OrganicSeg['Method of contact'].str.contains(Live,case=False,na=False) & OrganicSeg['UTM Campaign'].str.contains(validation, case= False,na=False) & OrganicSeg['UTM Medium'].str.contains(validation, case= False,na=False) & OrganicSeg['UTM Source'].str.contains(validation, case= False,na=False)].index,inplace=True)
dfliveappend= dflive.append(dflive1, ignore_index = True)
source= 'Facebook',
'https://regent.ac.za/programme/postgraduate-diploma-in-accounting',
'https://regent.ac.za/',
'https://regent.ac.za/programme/bachelor-of-commerce-in-supply-chain-management-regent',
'https://regent.ac.za/programme/postgraduate-diploma-in-project-management',
'https://regent.ac.za/my-account',
'https://regent.ac.za/programmes/postgraduate',
'https://regent.ac.za/programme/bachelor-of-business-administration-regent',
'https://regent.ac.za/programme/postgraduate-diploma-in-digital-marketing',
'https://regent.ac.za/logins',
'https://regent.ac.za/programme/advanced-diploma-in-management',
'https://regent.ac.za/programme/diploma-in-public-relations-management',
'https://regent.ac.za/programme/advanced-diploma-in-financial-management',
'https://regent.ac.za/programmes',
'https://regent.ac.za/apply-online?utm_source=CRM&utm_medium=email&utm_campaign=Canned_incomplete',
'https://regent.ac.za/programme/postgraduate-diploma-in-supply-chain-management',
'https://regent.ac.za/contact-us/learning-centres',
'https://regent.ac.za/contact-us/thank-you-qualifying',
'https://regent.ac.za/?utm_source=google&utm_medium=organic&utm_campaign=gmb-jhb&utm_content=website-link',
'https://regent.ac.za/apply-online',
'https://regent.ac.za/apply',
'https://regent.ac.za/programme/higher-certificate-in-health-care-services-management',
'https://regent.ac.za/programme/postgraduate-diploma-in-islamic-finance-and-banking',
'https://regent.ac.za/campus-news/higher-certificate-versus-diploma-versus-degree-help',
'https://regent.ac.za/programme/bachelor-of-commerce-in-human-resource-management',
'https://regent.ac.za/programme/master-of-business-administration-in-healthcare-management',
'https://regent.ac.za/programme/higher-certificate-in-human-resource-management',
'https://regent.ac.za/?utm_source=google&utm_medium=organic&utm_campaign=gmb-cpt&utm_content=website-link',
'https://regent.ac.za/degrees/12-things-you-can-do-with-your-bcom-degree',
'https://regent.ac.za/programmes/undergraduate/diplomas',
'https://regent.ac.za/programmes/short-learning-programmes',
'https://regent.ac.za/about-us',
'https://regent.ac.za/programmes/undergraduate/degrees',
'https://regent.ac.za/programmes?utm_source=CRM&utm_medium=email&utm_campaign=Enquiries-canned',
'https://regent.ac.za/programme/postgraduate-diploma-in-accounting-access-programme',
'https://regent.ac.za/programme/bachelor-of-commerce-in-law',
'https://regent.ac.za/programme/bachelor-of-commerce-honours',
'https://regent.ac.za/programmes/postgraduate/postgraduate-diplomas',
'https://regent.ac.za/programme/higher-certificate-in-business-management',
'https://regent.ac.za/vacancies',
'https://regent.ac.za/retail-management-development-programme-application',
'https://regent.ac.za/programme/mba-master-of-business-administration',
'https://regent.ac.za/programme/bachelor-of-public-administration',
'https://regent.ac.za/programme/bachelor-of-commerce',
'https://regent.ac.za/contact-us',
'https://regent.ac.za/programmes/undergraduate',
'https://regent.ac.za/programme/bachelor-of-commerce-in-accounting',
'https://regent.ac.za/programme/postgraduate-diploma-in-management',
'https://regent.ac.za/programme/higher-certificate-in-supply-chain-management',
'https://regent.ac.za/programme/higher-certificate-in-accounting',
'https://regent.ac.za/programme/higher-certificate-in-entrepreneurship',
'https://regent.ac.za/events',
'https://regent.ac.za/programmes/undergraduate/higher-certificates',
'https://regent.ac.za/about-us/why-choose-regent',
'https://regent.ac.za/programmes/postgraduate/honours',
'https://regent.ac.za/blog/a-21st-century-approach-to-human-resource-management',
'https://regent.ac.za/programme/diploma-in-financial-management',
'https://regent.ac.za/about-us/student-experience',
'https://regent.ac.za/category/distance-learning',
'https://regent.ac.za/programme/master-of-business-administration',
'https://regent.ac.za/programme/bachelor-of-commerce-in-retail-management',
'https://regent.ac.za/programme/bachelor-of-commerce-in-human-resource-management-honours',
'https://regent.ac.za/category/higher-certificates',
'https://regent.ac.za/programme/e-commerce-starting-your-online-business',
'https://regent.ac.za/?s=Honors+project+management',
'https://regent.ac.za/programmes/undergraduate/advanced-diplomas',
'https://regent.ac.za/students/accessing-regentonline',
'https://regent.ac.za/programme/doctor-of-business-administration',
'https://regent.ac.za/?s=fees',
'https://regent.ac.za/programme/higher-certificate-in-marketing-management',
'https://regent.ac.za/#top',
'https://regent.ac.za/join-regent-connect-today',
'https://regent.ac.za/programme/postgraduate-diploma-in-educational-management-and-leadership',
'https://regent.ac.za/programme/higher-certificate-in-business-management#request-form',
'https://regent.ac.za/programme/higher-certificate-in-islamic-finance-banking-and-law',
'https://regent.ac.za/about-us/accreditation',
'https://regent.ac.za/programme/higher-certificate-in-management-for-estate-agents/',
'https://regent.ac.za/programmes/undergraduate/',
'https://regent.ac.za/programme/bachelor-of-business-administration-regent/',
'https://regent.ac.za/programme/bachelor-of-commerce-honours/',
'https://regent.ac.za/programme/higher-certificate-in-accounting/',
'https://regent.ac.za/programme/bachelor-of-commerce/',
'https://regent.ac.za/programme/higher-certificate-in-retail-management/',
'https://regent.ac.za/programme/bachelor-of-commerce-in-accounting/',
'https://regent.ac.za/programme/postgraduate-diploma-in-project-management/',
'https://regent.ac.za/apply-online/',
'https://regent.ac.za/programme/bachelor-of-commerce-in-supply-chain-management-regent/',
'https://regent.ac.za/programme/master-of-business-administration/',
'https://regent.ac.za/programme/higher-certificate-in-business-management/',
'https://regent.ac.za/programme/postgraduate-diploma-in-accounting/',
'https://regent.ac.za/category/degrees/',
'https://regent.ac.za/programme/higher-certificate-in-supply-chain-management/',
'https://regent.ac.za/programmes/postgraduate/'
dflive2=OrganicSeg[OrganicSeg['Method of contact'].str.contains(Live,case=False,na=False) & OrganicSeg['UTM Campaign'].str.contains(validation, case= False,na=False) & OrganicSeg['UTM Medium'].str.contains(validation, case= False,na=False) & OrganicSeg.loc[OrganicSeg['UTM Source'].isin(source)]]
OrganicSeg.drop(OrganicSeg[OrganicSeg['Method of contact'].str.contains(Live,case=False,na=False) & OrganicSeg['UTM Campaign'].str.contains(validation, case= False,na=False) & OrganicSeg['UTM Medium'].str.contains(validation, case= False,na=False) & OrganicSeg.loc[OrganicSeg['UTM Source'].isin(source)]].index,inplace=True)
dflivefinal= dfliveappend.append(dflive2, ignore_index = True)
CRMOrg = OrganicSeg.copy() 
pivotlive= pd.pivot_table(dflivefinal,values='Lead Name',index='Program Version Name',columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivotlive = pivotlive.loc[name_order]
pivotCRM= pd.pivot_table(CRMOrg,values='Lead Name',index='Program Version Name',columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivotCRM = pivotCRM.loc[name_order]


#PROCESS18
PaidSeg= pd.DataFrame(OverallPaid)

Jivo = 'campaign=CAO_web-banner_2022|source=CAO_website|medium=banner',
'campaign=Lookalike_HCs|source=Facebook|medium=ad_post',
'campaign=Lookalike_Apps_HCs|source=Facebook|medium=ad_post',
'campaign=USSD_RDCampaign|source=mobile|medium=ussd',
utmpaidsource = 'campaign=Canned_enquiry-new|source=CRM|medium=email'
utmCampaign= 'Flume_B2B_Search_Reskill_and_Upskill'
utmMedium = 'paid'
dfJivo= PaidSeg[PaidSeg.loc[PaidSeg['UTM Source'].isin(Jivo)] & PaidSeg['Method of contact'].str.contains(Live,case=False,na=False) & PaidSeg['UTM Campaign'].str.contains(validation, case= False,na=False) & PaidSeg['UTM Medium'].str.contains(validation, case= False,na=False) ]
PaidSeg.drop(PaidSeg[PaidSeg.loc[PaidSeg['UTM Source'].isin(Jivo)] & PaidSeg['Method of contact'].str.contains(Live,case=False,na=False) & PaidSeg['UTM Campaign'].str.contains(validation, case= False,na=False) & PaidSeg['UTM Medium'].str.contains(validation, case= False,na=False) ].index,inplace=True)
dfJivo2=PaidSeg[PaidSeg['Method of contact'].str.contains(Live,case=False,na=False) & PaidSeg['UTM Campaign'].str.contains(utmCampaign, case= False,na=False) & PaidSeg['UTM Medium'].str.contains(utmMedium, case= False,na=False) & PaidSeg['UTM Source'].str.contains(utmpaidsource, case= False,na=False)]
PaidSeg.drop(PaidSeg[PaidSeg['Method of contact'].str.contains(Live,case=False,na=False) & PaidSeg['UTM Campaign'].str.contains(utmCampaign, case= False,na=False) & PaidSeg['UTM Medium'].str.contains(utmMedium, case= False,na=False) & PaidSeg['UTM Source'].str.contains(utmpaidsource, case= False,na=False)].index,inplace=True)
dfJivofinal= dfJivo.append(dfJivo2, ignore_index = True)
CRMPaid=PaidSeg.copy()
pivotJivo= pd.pivot_table(dfJivo,values='Lead Name',index='Program Version Name',columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivotJivo = pivotJivo.loc[name_order]


#Process19
pivotCRMPaid= pd.pivot_table(CRMPaid,values='Lead Name',index='Program Version Name',columns='Campus',aggfunc = 'count',margins=True,margins_name='Grand Total',fill_value=' ')
sortedpivotCRMPaid = pivotCRMPaid.loc[name_order]






with pd.ExcelWriter(buffer,engine='openpyxl') as writer:  
   df_noDups.to_excel(writer, sheet_name='NoDups',index=False)
   df_noJunk.to_excel(writer, sheet_name='No Junk',index =False)
   NegativeLeads.to_excel(writer, sheet_name='Negative Leads_No Progress',index =False)
   pivotnegleads.to_excel(writer, sheet_name='Pivot of Neg Leads_No Prog',index = True,startrow=1,startcol=1)
   AccessLeads.to_excel(writer, sheet_name='Access Leads',index =False)
   pivotaccessleads.to_excel(writer, sheet_name='Pivot Access Leads',index = True,startrow=1,startcol=1)
   schoolLeads.to_excel(writer, sheet_name='School Leads',index =False)
   df_cleanedRegion.to_excel(writer, sheet_name='Cleaned Region',index =False)
   pivotRegions.to_excel(writer, sheet_name='Pivot Regions',index = True,startrow=1,startcol=1)
   #sortedPivotLeadsA.to_excel(writer, sheet_name='Leads Analysis',index = True,startrow=1,startcol=1)
   #pivotLeadsDisposition.to_excel(writer, sheet_name='Leads Analysis',index = True,startrow=1,startcol=10)
   #sortedPivotProgramme.to_excel(writer, sheet_name='Programme Analysis 1',index = True,startrow=1,startcol=1)
  # sortedPivotProgramme2.to_excel(writer, sheet_name='Programme Analysis 2',index = True,startrow=1,startcol=1)
   priorLeads.to_excel(writer, sheet_name='Leads Carried Over @11 Sep',index =False)
   sortedPivotpriorLeads.to_excel(writer, sheet_name='Pivot Carried Over-Reg',index = True,startrow=1,startcol=1)
   sortedpivotPriorLeadsDisp.to_excel(writer, sheet_name='Pivot Carried Over-Reg',index = True,startrow=1,startcol=10)
   #sortedpivot10c.to_excel(writer, sheet_name='Pivot Carried Over-Reg & Prog',index = True,startrow=1,startcol=1)
   #sortedpivot10cA.to_excel(writer, sheet_name='Pivot Carried Over-Reg & Prog',index = True,startrow=1,startcol=10)
   newLeads.to_excel(writer, sheet_name='New Leads 12 Sep-Present',index =False)
   outlierdf.to_excel(writer, sheet_name='Unfiltered Data New Leads',index =False)
   sortedPivotnewLeads.to_excel(writer, sheet_name='Pivot New Leads',index = True,startrow=1,startcol=1)
   sortedpivotnewLeadsDisp.to_excel(writer, sheet_name='Pivot New Leads',index = True,startrow=1,startcol=12)
   #sortedpivot11c.to_excel(writer, sheet_name='Pivot New Carried Over-Reg & Prog',index = True,startrow=1,startcol=1)
   #sortedpivot11cA.to_excel(writer, sheet_name='Pivot New Carried Over-Reg & Prog',index = True,startrow=1,startcol=10)
   sortedPivot9c.to_excel(writer, sheet_name='Pivot of New Leads Per Programme',index = True,startrow=1,startcol=1)
   df_cleanedProg.to_excel(writer, sheet_name='Cleaned Progress',index =False)
   sortedPivotProgress.to_excel(writer, sheet_name='Pivot Cleaned Progress',index = True,startrow=1,startcol=1)
   sortedPivotProgress1.to_excel(writer, sheet_name='Pivot Cleaned Progress',index = True,startrow=1,startcol=12)
   pivotDay.to_excel(writer, sheet_name='Pivot New Leads Day on Day',index = True,startrow=1,startcol=1)
   OrganicLeads.to_excel(writer, sheet_name='Overall Organic',index =False)
   unfiltereddf.to_excel(writer, sheet_name='Unfiltered Organic',index =False)
   dfWalk.to_excel(writer, sheet_name='Walk-in',index =False)
   sortedpivotwalk.to_excel(writer, sheet_name='Pivot Walk-In',index = True,startrow=1,startcol=1)
   dfCall.to_excel(writer, sheet_name='Calls',index =False)
   sortedpivotcall.to_excel(writer, sheet_name='Pivot Calls',index = True,startrow=1,startcol=1)
   dflivefinal.to_excel(writer, sheet_name='Jivo Org',index =False)
   sortedpivotlive.to_excel(writer, sheet_name='Pivot Jivo Org',index = True,startrow=1,startcol=1)
   CRMOrg.to_excel(writer, sheet_name='CRM Org',index =False)
   sortedpivotCRM.to_excel(writer, sheet_name='Pivot CRM Org',index = True,startrow=1,startcol=1)
   OverallPaid.to_excel(writer, sheet_name='Overall Paid',index =False)
   dfJivofinal.to_excel(writer, sheet_name=' Paid Jivo',index =False)
   sortedpivotJivo.to_excel(writer, sheet_name='Pivot Paid Jivo',index = True,startrow=1,startcol=1)
   CRMPaid.to_excel(writer, sheet_name=' CRM Paid',index =False)
   sortedpivotCRMPaid.to_excel(writer, sheet_name='Pivot CRM Paid ',index = True,startrow=1,startcol=1)
   



   







writer.close()

st.download_button(
    label = "Download Processed Excel File",
    data=buffer,
    file_name="Consolidated Leads.xlsx",
    mime= "application/vnd.ms-excel"
)
# In[ ]:




