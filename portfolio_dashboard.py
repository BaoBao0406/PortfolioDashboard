#! python3
# portfolio_dashboard.py - 

import pyodbc
import plotly
from datetime import date
import pandas as pd
import plotly.offline as pyo
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import plotly.figure_factory as ff
import numpy as np

##################################################################################################

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=VOPPSCLDBN01\VOPPSCLDBI01;'
                      'Database=SalesForce;'
                      'Trusted_Connection=yes;')

##################################################################################################

# FDC User ID and Name list
user = pd.read_sql('SELECT DISTINCT(Id), Name \
                    FROM dbo.[User]', conn)
user = user.set_index('Id')['Name'].to_dict()


BK_tmp = pd.read_sql("SELECT BK.Id, BK.Booking_ID_Number__c, FORMAT(BK.nihrm__ArrivalDate__c, 'yyyy-MM-dd') AS ArrivalDate, ac.Id, ac.Name AS ACName, ag.Id, ag.Name AS AGName, BK.End_User_Region__c, BK.End_User_SIC__c, \
                             BK.nihrm__BookingTypeName__c, ac.nihrm__RegionName__c, ac.Industry, BK.nihrm__CurrentBlendedRoomnightsTotal__c, BK.nihrm__CurrentBlendedRevenueTotal__c, \
                             FORMAT(BK.nihrm__BookedDate__c, 'yyyy-MM-dd') AS BookedDate, BK.nihrm__Property__c, ac.Type \
                      FROM dbo.nihrm__Booking__c AS BK \
                             LEFT JOIN dbo.Account AS ac \
                                 ON BK.nihrm__Account__c = ac.Id \
                             LEFT JOIN dbo.Account AS ag \
                                 ON BK.nihrm__Agency__c = ag.Id \
                      WHERE (BK.nihrm__BookingTypeName__c NOT IN ('ALT Alternative', 'CN Concert', 'IN Internal')) AND \
                             (BK.nihrm__Property__c NOT IN ('Sands Macao Hotel')) AND (BK.nihrm__BookingStatus__c IN ('Definite'))", conn)

BK_tmp.columns = ['Id', 'Booking ID#', 'ArrivalDate', 'Accound Id', 'Account', 'Agency Id', 'Agency', 'End User Region', 'End User SIC', 'Booking Type', 'Account: Region', 'Account: Industry', 'Blended Roomnights',
                  'Blended Total Revenue', 'BookedDate', 'Property', 'Account Type']

##################################################################################################

BK_cv_tmp = BK_tmp[['Booking ID#', 'ArrivalDate', 'Account', 'Account: Region', 'Account: Industry', 'Agency', 'Booking Type', 'Blended Roomnights', 'Blended Total Revenue', 'Account Type']]

                 
# Noted that the current date is set to 2021-01-01
NOW = datetime.datetime(2021,1,1)
BK_cv_tmp['ArrivalDate'] = pd.to_datetime(BK_cv_tmp['ArrivalDate'])
BK_cv_tmp = BK_cv_tmp[BK_tmp['ArrivalDate'] < NOW]

BK_cv_tmp['Arrival Year'] = pd.DatetimeIndex(BK_cv_tmp['ArrivalDate']).year
BK_cv_tmp['Arrival Month'] = pd.DatetimeIndex(BK_cv_tmp['ArrivalDate']).month
BK_cv_tmp['Arrival Day'] = pd.DatetimeIndex(BK_cv_tmp['ArrivalDate']).day

# Noted that due to COVID-19, for arrival year before 2020, the year have been add for one year.
# (Orignal arrival: 2019, Now: 2020)
BK_cv_tmp['Arrival Year'] = BK_cv_tmp['Arrival Year'].apply(lambda x: x + 1 if x < 2020 else x)
BK_cv_tmp['Arrival_new'] = pd.to_datetime(BK_cv_tmp[['Arrival Year', 'Arrival Month', 'Arrival Day']])

# Take out Account 'ROTARY CLUB OF MACAU'
acct_name = ['ROTARY CLUB OF MACAU']
BK_cv_tmp = BK_cv_tmp[BK_cv_tmp['Account'] != 'ROTARY CLUB OF MACAU']


RFGM = BK_cv_tmp.groupby(['Accound Id', 'Account']).agg({'Arrival_new': lambda x: (NOW - x.max()).days,
                                                         'Accound Id': lambda x: len(x), 
                                                         'Blended Roomnights': lambda x: x.sum(),
                                                         'Blended Total Revenue': lambda x: x.sum()})

RFGM.rename(columns={'Arrival_new': 'Recency', 
                     'Accound Id': 'Frequency',
                     'Blended Roomnights': 'Guestroom_value',
                     'Blended Total Revenue': 'Monetary_value'}, inplace=True)

# Recency score
r_labels = range(5, 0, -1)
r_quartiles = pd.cut(RFGM['Recency'], 5, labels=r_labels)
RFGM = RFGM.assign(R = r_quartiles.values)

# Frequency score
f_labels = range(1, 6)
f_quartiles = pd.cut(RFGM['Frequency'], 5, labels=f_labels)
RFGM = RFGM.assign(F = f_quartiles.values)

# Guestroom value score
g_labels = range(1, 6)
g_quartiles = pd.cut(RFGM['Guestroom_value'], 5, labels=g_labels)
RFGM = RFGM.assign(G = g_quartiles.values)

# Monetary value score
m_labels = range(1, 6)
m_quartiles = pd.cut(RFGM['Monetary_value'], 5, labels=m_labels)
RFGM = RFGM.assign(M = m_quartiles.values)


def join_rfgm(x):
    return str(x['R']) + '-' + str(x['F']) + '-' + str(x['G']) + '-' + str(x['M'])

RFGM['RFGM_Segment'] = RFGM.apply(join_rfgm, axis=1)
RFGM['RFGM_Score'] = RFGM[['R', 'F', 'G', 'M']].sum(axis=1)

# Take out the outlier data for late process (combine it later)
RFGM = pd.DataFrame(RFGM).reset_index()
acct_name = ['JEUNESSE GLOBAL HOLDINGS, LLC TAIWAN BR']
outliers = RFGM.loc[RFGM['Account'].isin(acct_name)]
outliers

# Take out outliers
acct_name = ['JEUNESSE GLOBAL HOLDINGS, LLC TAIWAN BR']
outliers_list = BK_cv_tmp[BK_cv_tmp['Account'].isin(acct_name)].index
data1 = BK_cv_tmp.drop(outliers_list)

RFGM1 = data1.groupby(['Account Id', 'Account']).agg({'Arrival_new': lambda x: (NOW - x.max()).days,
                                                      'Account Id': lambda x: len(x), 
                                                      'Blended Roomnights': lambda x: x.sum(),
                                                      'Blended Total Revenue': lambda x: x.sum()})

RFGM1.rename(columns={'Arrival_new': 'Recency', 
                      'Account Id': 'Frequency',
                      'Blended Roomnights': 'Guestroom_value',
                      'Blended Total Revenue': 'Monetary_value'}, inplace=True)

# Recency score
r_labels = range(5, 0, -1)
r_quartiles = pd.cut(RFGM1['Recency'], 5, labels=r_labels)
RFGM1 = RFGM1.assign(R = r_quartiles.values)

# Frequency score
f_labels = range(1, 6)
f_quartiles = pd.cut(RFGM1['Frequency'], 5, labels=f_labels)
RFGM1 = RFGM1.assign(F = f_quartiles.values)

# Guestroom value score
g_labels = range(1, 6)
g_quartiles = pd.cut(RFGM1['Guestroom_value'], 5, labels=g_labels)
RFGM1 = RFGM1.assign(G = g_quartiles.values)

# Monetary value score
m_labels = range(1, 6)
m_quartiles = pd.cut(RFGM1['Monetary_value'], 5, labels=m_labels)
RFGM1 = RFGM1.assign(M = m_quartiles.values)


RFGM1['RFGM_Segment'] = RFGM1.apply(join_rfgm, axis=1)
RFGM1['RFGM_Score'] = RFGM1[['R', 'F', 'G', 'M']].sum(axis=1)

##################################################################################################





def figures_to_html(figs, filename):
    dashboard = open(filename, 'w')
    dashboard.write("<html><head></head><body>" + "\n")
    for fig in figs:
        inner_html = fig.to_html().split('<body>')[1].split('</body>')[0]
        dashboard.write(inner_html)
    dashboard.write("</body></html>" + "\n")

figures_to_html([fig1, fig2, fig3, fig4], filename='performance_dashboard.html')

##################################################################################################
