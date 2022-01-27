#! python3
# portfolio_dashboard.py - 

import pyodbc
import plotly
import datetime
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

BK_cv_tmp = BK_tmp[['Booking ID#', 'ArrivalDate', 'Accound Id', 'Account', 'Account: Region', 'Account: Industry', 'Agency Id', 'Agency', 'Booking Type', 'Blended Roomnights', 'Blended Total Revenue', 'Account Type']]

                 
# Noted that the current date is set to 2021-01-01
NOW = datetime.datetime(2021,1,1)
BK_cv_tmp['ArrivalDate'] = pd.to_datetime(BK_cv_tmp['ArrivalDate'])
BK_cv_tmp = BK_cv_tmp[BK_cv_tmp['ArrivalDate'] < NOW]

BK_cv_tmp['year'] = pd.DatetimeIndex(BK_cv_tmp['ArrivalDate']).year
BK_cv_tmp['month'] = pd.DatetimeIndex(BK_cv_tmp['ArrivalDate']).month
BK_cv_tmp['day'] = pd.DatetimeIndex(BK_cv_tmp['ArrivalDate']).day

# Noted that due to COVID-19, for arrival year before 2020, the year have been add for one year.
# (Orignal arrival: 2019, Now: 2020)
BK_cv_tmp['year'] = BK_cv_tmp['year'].apply(lambda x: x + 1 if x < 2020 else x)
BK_cv_tmp['Arrival_new'] = pd.to_datetime(BK_cv_tmp[['year', 'month', 'day']], errors='coerce')

# Take out Account 'ROTARY CLUB OF MACAU'
acct_name = ['ROTARY CLUB OF MACAU']
BK_cv_tmp = BK_cv_tmp[BK_cv_tmp['Account'] != 'ROTARY CLUB OF MACAU']


def join_rfgm(x):
    return str(x['R']) + '-' + str(x['F']) + '-' + str(x['G']) + '-' + str(x['M'])

# create RFGM table
def rfgm_df(BK_tmp):
    RFGM_df = BK_tmp.groupby(['Accound Id', 'Account']).agg({'Arrival_new': lambda x: (NOW - x.max()).days,
                                                             'Accound Id': lambda x: len(x), 
                                                             'Blended Roomnights': lambda x: x.sum(),
                                                             'Blended Total Revenue': lambda x: x.sum()})
    RFGM_df.rename(columns={'Arrival_new': 'Recency', 
                            'Accound Id': 'Frequency',
                            'Blended Roomnights': 'Guestroom_value',
                            'Blended Total Revenue': 'Monetary_value'}, inplace=True)
    return RFGM_df

# calculate RFGM quatiles scores
def rfgm_score_quat(RFGM_tmp, rfgm_dict):
    for key, values in rfgm_dict.items():
        tmp_labels = values[0]
        tmp_quartiles = pd.cut(RFGM_tmp[key], 5, labels=tmp_labels)
        RFGM_tmp[values[1]] = tmp_quartiles.values
    RFGM_tmp['RFGM_Segment'] = RFGM_tmp.apply(join_rfgm, axis=1)
    
    return RFGM_tmp

rfgm_dict = {'Recency': [range(5, 0, -1), 'R'], 'Frequency': [range(1, 6), 'F'], 
             'Guestroom_value': [range(1, 6), 'G'], 'Monetary_value': [range(1, 6), 'M']}

RFGM_df = rfgm_df(BK_cv_tmp)
RFGM_score = rfgm_score_quat(RFGM_df, rfgm_dict)

# Take out the outlier data for late process (combine it later)
RFGM_df = pd.DataFrame(RFGM_df).reset_index()
acct_name = ['JEUNESSE GLOBAL HOLDINGS, LLC TAIWAN BR']
outliers = RFGM_df.loc[RFGM_df['Account'].isin(acct_name)]
outliers_row = BK_cv_tmp[BK_cv_tmp['Account'].isin(acct_name)].index
BK_cv_wo_outliner_tmp = BK_cv_tmp.drop(outliers_row)

RFGM_df_no_outlier = rfgm_df(BK_cv_wo_outliner_tmp)

RFGM_score_no_outlier = rfgm_score_quat(RFGM_df_no_outlier, rfgm_dict)

RFGM_score_no_outlier['RFGM_Score'] = RFGM_score_no_outlier[['R', 'F', 'G', 'M']].sum(axis=1)

print(RFGM_score_no_outlier)
###################################################################################################

# Plot 1
fig1 = make_subplots(rows=3, cols=1, subplot_titles=('Tentative Bookings', 'Prospect Bookings', 'Inquiries'), 
                    column_widths=[0.05], row_heights=[0.3, 0.3, 0.3], vertical_spacing=0.1, horizontal_spacing=0.0, 
                    specs=[[{"type": "table"}], [{"type": "table"}], [{"type": "table"}]])

table1_obj = go.Table(header = dict(values=bk_display_col),
                      cells = dict(values=[sm_current_business_t[k].tolist() for k in sm_current_business_t.columns[0:]]))

table2_obj = go.Table(header = dict(values=bk_display_col),
                      cells = dict(values=[sm_current_business_p[k].tolist() for k in sm_current_business_p.columns[0:]]))

table3_obj = go.Table(header = dict(values=inq_display_col),
                      cells = dict(values=[sm_inquiry[k].tolist() for k in sm_inquiry.columns[0:]]))

fig1.add_trace(table1_obj, row=1, col=1)
fig1.add_trace(table2_obj, row=2, col=1)
fig1.add_trace(table3_obj, row=3, col=1)

fig1.update_layout(title='Current Business', autosize=False, width=1800, height=800)



def figures_to_html(figs, filename):
    dashboard = open(filename, 'w')
    dashboard.write("<html><head></head><body>" + "\n")
    for fig in figs:
        inner_html = fig.to_html().split('<body>')[1].split('</body>')[0]
        dashboard.write(inner_html)
    dashboard.write("</body></html>" + "\n")

figures_to_html([fig1], filename='performance_dashboard.html')

##################################################################################################
