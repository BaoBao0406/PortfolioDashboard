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
from dateutil.relativedelta import relativedelta

##################################################################################################

# Set date range
now = datetime.datetime.now()
start_year_def = now - relativedelta(years=3)
start_date_def = str(start_year_def.year) + '-01-01'
end_date_def = str(now.year) + '-12-31'


conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=VOPPSCLDBN01\VOPPSCLDBI01;'
                      'Database=SalesForce;'
                      'Trusted_Connection=yes;')

# FDC User ID and Name list
user = pd.read_sql('SELECT DISTINCT(Id), Name \
                    FROM dbo.[User]', conn)
user = user.set_index('Id')['Name'].to_dict()

##################################################################################################

# SQL queries

# Booked date data
start_year_booked = now - relativedelta(years=3)
start_date_booked = str(start_year_booked.year) + '-01-01'
end_date_booked = str(now.year) + '-12-31'
booked_date_df = pd.read_sql("SELECT BK.OwnerId, BK.nihrm__Property__c, ac.Name, ac.BillingCountry, ag.Name, BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'MM/dd/yyyy') AS ArrivalDate, BK.VCL_Arrival_Year__c, BK.VCL_Arrival_Month__c, FORMAT(BK.nihrm__DepartureDate__c, 'MM/dd/yyyy') AS DepartureDate, \
                                 BK.nihrm__CurrentBlendedRoomnightsTotal__c, BK.nihrm__BlendedGuestroomRevenueTotal__c, \
                                 BK.VCL_Blended_F_B_Revenue__c, BK.nihrm__CurrentBlendedEventRevenue7__c, BK.nihrm__BookingStatus__c, FORMAT(BK.nihrm__LastStatusDate__c, 'MM/dd/yyyy') AS LastStatusDate, \
                                 FORMAT(BK.nihrm__BookedDate__c, 'MM/dd/yyyy') AS BookedDate, BK.Booked_Year__c, BK.Booked_Month__c, BK.End_User_Region__c, BK.End_User_SIC__c, BK.nihrm__BookingTypeName__c, BK.Booking_ID_Number__c, FORMAT(BK.nihrm__DateDefinite__c, 'MM/dd/yyyy') AS DateDefinite, BK.Date_Definite_Month__c, BK.Date_Definite_Year__c \
                          FROM dbo.nihrm__Booking__c AS BK \
                          LEFT JOIN dbo.Account AS ac \
                              ON BK.nihrm__Account__c = ac.Id \
                          LEFT JOIN dbo.Account AS ag \
                              ON BK.nihrm__Agency__c = ag.Id \
                          WHERE (BK.nihrm__BookingTypeName__c NOT IN ('ALT Alternative', 'CN Concert', 'IN Internal')) AND \
                              (BK.nihrm__Property__c NOT IN ('Sands Macao Hotel')) AND \
                              (BK.nihrm__BookedDate__c BETWEEN CONVERT(datetime, '" + start_date_booked + "') AND CONVERT(datetime, '" + end_date_booked + "'))", conn)
booked_date_df.columns = ['Owner Name', 'Property', 'Account', 'Company Country', 'Agency', 'Booking: Booking Post As', 'Arrival', 'Arrival Year', 'Arrival Month', 'Departure', 
                      'Blended Roomnights', 'Blended Guestroom Revenue Total', 'Blended F&B Revenue', 'Blended Rental Revenue', 'Status',
                      'Last Status Date', 'Booked', 'Booked Year', 'Booked Month', 'End User Region', 'End User SIC', 'Booking Type', 'Booking ID#', 'DateDefinite', 'Date Definite Month', 'Date Definite Year']
booked_date_df['Owner Name'].replace(user, inplace=True)


# Arrival date data
start_year_arrival = now - relativedelta(years=3)
end_year_arrival = now + relativedelta(years=5)
start_date_arrival = str(start_year_arrival.year) + '-01-01'
end_date_arrival = str(end_year_arrival.year) + '-12-31'
arrival_date_df = pd.read_sql("SELECT BK.OwnerId, BK.nihrm__Property__c, ac.Name, ac.BillingCountry, ag.Name, BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'MM/dd/yyyy') AS ArrivalDate, BK.VCL_Arrival_Year__c, BK.VCL_Arrival_Month__c, FORMAT(BK.nihrm__DepartureDate__c, 'MM/dd/yyyy') AS DepartureDate, \
                                    BK.nihrm__CurrentBlendedRoomnightsTotal__c, BK.nihrm__BlendedGuestroomRevenueTotal__c, \
                                    BK.VCL_Blended_F_B_Revenue__c, BK.nihrm__CurrentBlendedEventRevenue7__c, BK.nihrm__BookingStatus__c, FORMAT(BK.nihrm__LastStatusDate__c, 'MM/dd/yyyy') AS LastStatusDate, \
                                    FORMAT(BK.nihrm__BookedDate__c, 'MM/dd/yyyy') AS BookedDate, BK.Booked_Year__c, BK.Booked_Month__c, BK.End_User_Region__c, BK.End_User_SIC__c, BK.nihrm__BookingTypeName__c, BK.Booking_ID_Number__c, FORMAT(BK.nihrm__DateDefinite__c, 'MM/dd/yyyy') AS DateDefinite, BK.Date_Definite_Month__c, BK.Date_Definite_Year__c \
                                FROM dbo.nihrm__Booking__c AS BK \
                                LEFT JOIN dbo.Account AS ac \
                                    ON BK.nihrm__Account__c = ac.Id \
                                LEFT JOIN dbo.Account AS ag \
                                    ON BK.nihrm__Agency__c = ag.Id \
                                WHERE (BK.nihrm__BookingTypeName__c NOT IN ('ALT Alternative', 'CN Concert', 'IN Internal')) AND \
                                    (BK.nihrm__Property__c NOT IN ('Sands Macao Hotel')) AND \
                                    (BK.nihrm__ArrivalDate__c BETWEEN CONVERT(datetime, '" + start_date_arrival + "') AND CONVERT(datetime, '" + end_date_arrival + "'))", conn)
arrival_date_df.columns = ['Owner Name', 'Property', 'Account', 'Company Country', 'Agency', 'Booking: Booking Post As', 'Arrival', 'Arrival Year', 'Arrival Month', 'Departure', 
                            'Blended Roomnights', 'Blended Guestroom Revenue Total', 'Blended F&B Revenue', 'Blended Rental Revenue', 'Status',
                            'Last Status Date', 'Booked', 'Booked Year', 'Booked Month', 'End User Region', 'End User SIC', 'Booking Type', 'Booking ID#', 'DateDefinite', 'Date Definite Month', 'Date Definite Year']
arrival_date_df['Owner Name'].replace(user, inplace=True)
arrival_date_df['Days before Arrive'] = ((pd.to_datetime(arrival_date_df['Arrival']) - pd.Timestamp.now().normalize()).dt.days).astype(int)
arrival_date_df.sort_values(by=['Days before Arrive'], inplace=True)


# All booking data
BK_tmp = pd.read_sql("SELECT BK.Id, BK.Booking_ID_Number__c, FORMAT(BK.nihrm__ArrivalDate__c, 'yyyy-MM-dd') AS ArrivalDate, ac.Id, ac.Name AS ACName, ag.Id, ag.Name AS AGName, BK.End_User_Region__c, BK.End_User_SIC__c, \
                             BK.nihrm__BookingTypeName__c, ac.nihrm__RegionName__c, ac.Industry, ag.nihrm__RegionName__c, BK.nihrm__CurrentBlendedRoomnightsTotal__c, BK.nihrm__CurrentBlendedRevenueTotal__c, \
                             FORMAT(BK.nihrm__BookedDate__c, 'yyyy-MM-dd') AS BookedDate, BK.nihrm__Property__c, ac.Type, ag.Type \
                      FROM dbo.nihrm__Booking__c AS BK \
                             LEFT JOIN dbo.Account AS ac \
                                 ON BK.nihrm__Account__c = ac.Id \
                             LEFT JOIN dbo.Account AS ag \
                                 ON BK.nihrm__Agency__c = ag.Id \
                      WHERE (BK.nihrm__BookingTypeName__c NOT IN ('ALT Alternative', 'CN Concert', 'IN Internal', 'CS Catering - Social')) AND \
                             (BK.nihrm__Property__c NOT IN ('Sands Macao Hotel')) AND (BK.nihrm__BookingStatus__c IN ('Definite'))", conn)

BK_tmp.columns = ['Id', 'Booking ID#', 'ArrivalDate', 'Accound Id', 'Account', 'Agency Id', 'Agency', 'End User Region', 'End User SIC', 'Booking Type', 'Account: Region', 'Account: Industry', 'Agency: Region',
                  'Blended Roomnights', 'Blended Total Revenue', 'BookedDate', 'Property', 'Account Type', 'Agency Type']

##################################################################################################

# Filter by industry or region

# TODO: filter by MICE and Tradeshow

# 

##################################################################################################

# RFM data preprocess

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

RFGM_score_no_outlier = pd.DataFrame(rfgm_score_quat(RFGM_df_no_outlier, rfgm_dict)).reset_index()
# 
region_sic = BK_tmp[['Accound Id', 'Account: Region', 'Account: Industry']].drop_duplicates(subset=['Accound Id'])
RFGM_score_no_outlier = pd.merge(RFGM_score_no_outlier, region_sic, on='Accound Id', how='left')


###################################################################################################

# Plot 1

months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

history_count = booked_date_df.groupby(['Booked Month', 'Booked Year']).size().reset_index(name='NumberOfBK')
history_count['to_sort']=history_count['Booked Month'].apply(lambda x: months.index(x))
history_count = history_count.sort_values('to_sort')

arrival_rn_revenue = arrival_date_df.groupby(['Arrival Month', 'Arrival Year']).sum().reset_index()
arrival_rn_revenue['to_sort']=arrival_rn_revenue['Arrival Month'].apply(lambda x: months.index(x))
arrival_rn_revenue = arrival_rn_revenue.sort_values('to_sort')



fig1 = make_subplots(rows=3, cols=1, subplot_titles=('3 years comparsion on RNs by Created Year and Month', '3 years comparsion on RNs by Arrival Year and Month', '3 years comparsion on Total Revenue by Arrival Year and Month'),
                     column_widths=[1.0], row_heights=[1.0, 1.0, 1.0], shared_xaxes=True)

# previous and next 3 years
years = sorted([now.year - i for i in range(-3, 3)])

# color index
cols = plotly.colors.DEFAULT_PLOTLY_COLORS

for i, year in enumerate(years):
    # 3 years comparsion on RN by Created Year and Month
    line1 = go.Scatter(x=history_count[history_count['Booked Year'] == int(year)]['Booked Month'], 
                       y=history_count[history_count['Booked Year'] == int(year)]['NumberOfBK'], 
                       mode='lines+markers', line={'color': cols[i]}, name=str(int(year)), legendgroup=str(year), showlegend=False)
    
    # 3 years comparsion on RNs by Arrival Year and Month
    line2 = go.Scatter(x=arrival_rn_revenue[arrival_rn_revenue['Arrival Year'] == int(year)]['Arrival Month'], 
                       y=arrival_rn_revenue[arrival_rn_revenue['Arrival Year'] == int(year)]['Blended Roomnights'], 
                       mode='lines+markers', line={'color': cols[i]}, name=str(int(year)), legendgroup=str(year))
    
    # 3 years comparsion on Total Revenue by Arrival Year and Month
    line3 = go.Scatter(x=arrival_rn_revenue[arrival_rn_revenue['Arrival Year'] == int(year)]['Arrival Month'], 
                       y=arrival_rn_revenue[arrival_rn_revenue['Arrival Year'] == int(year)]['Blended Guestroom Revenue Total'], 
                       mode='lines+markers', line={'color': cols[i]}, name=str(int(year)), legendgroup=str(year), showlegend=False)
    
    fig1.add_trace(line1, row=1, col=1)
    fig1.add_trace(line2, row=2, col=1)
    fig1.add_trace(line3, row=3, col=1)

fig1.update_layout(title='3 years Demand History Comparsion', xaxis_title='Month', xaxis_showticklabels=True, height=1000)


# Plot 2 & 3

top_account = BK_tmp[BK_tmp['Account Type'] == 'Account']
top_account = top_account.groupby(['Accound Id', 'Account', 'Account: Industry']).agg({'Accound Id': lambda x: len(x), 
                                                                                       'Blended Roomnights': lambda x: x.sum(),
                                                                                       'Blended Total Revenue': lambda x: x.sum()})
top_account.rename(columns={'Accound Id': 'No. of Booking', 'Blended Roomnights': 'RNs', 'Blended Total Revenue': 'Total Revenue'}, inplace=True)
top_account = pd.DataFrame(top_account).reset_index()
top_account.rename(columns={'Account: Industry': 'Industry'}, inplace=True)

top_agency = BK_tmp[BK_tmp['Agency Type'] == 'Agency']
top_agency = top_agency.groupby(['Agency Id', 'Agency']).agg({'Agency Id': lambda x: len(x), 
                                                              'Blended Roomnights': lambda x: x.sum(),
                                                              'Blended Total Revenue': lambda x: x.sum()})
top_agency.rename(columns={'Agency Id': 'No. of Booking', 'Blended Roomnights': 'RNs', 'Blended Total Revenue': 'Total Revenue'}, inplace=True)
top_agency = pd.DataFrame(top_agency).reset_index()


# function to create top 10 table
def top_10_table_info(tmp_fig, tmp_data, tmp_cols, sort_index):
    # create sorted table for top 10 account/agency
    for i, col in tmp_cols.items():
        tmp_top_data = tmp_data[col]
        tmp_top_data = tmp_top_data.sort_values(by=tmp_top_data.columns[sort_index], ascending=False).head(10)
        
        tmp_table = go.Table(header = dict(values=col),
                              cells = dict(values=[tmp_top_data[k].tolist() for k in tmp_top_data.columns[0:]]))
    
        tmp_fig.add_trace(tmp_table, row=1, col=i)



fig2 = make_subplots(rows=1, cols=3, subplot_titles=('Top 10 Most Bookings', 'Top 10 Most RNs', 'Top 10 Most Total Revenue'), 
                    column_widths=[0.05, 0.05, 0.05], row_heights=[0.3], vertical_spacing=0.0, horizontal_spacing=0.05, 
                    specs=[[{"type": "table"}, {"type": "table"}, {"type": "table"}]])

top_10_ac_cols = {1 : ['Account', 'Industry', 'No. of Booking'], 2 : ['Account', 'Industry', 'RNs'], 3 : ['Account', 'Industry', 'Total Revenue']}
top_10_table_info(fig2, top_account, top_10_ac_cols, 2)

fig2.update_layout(title='Top 10 Account information', autosize=False, width=1800, height=800)


fig3 = make_subplots(rows=1, cols=3, subplot_titles=('Top 10 Most Bookings', 'Top 10 Most RNs', 'Top 10 Most Total Revenue'), 
                    column_widths=[0.05, 0.05, 0.05], row_heights=[0.3], vertical_spacing=0.0, horizontal_spacing=0.05, 
                    specs=[[{"type": "table"}, {"type": "table"}, {"type": "table"}]])

top_10_ag_cols = {1 : ['Agency', 'No. of Booking'], 2 : ['Agency', 'RNs'], 3 : ['Agency', 'Total Revenue']}
top_10_table_info(fig3, top_agency, top_10_ag_cols, 1)

fig3.update_layout(title='Top 10 Agency information', autosize=False, width=1800, height=800)


# Plot 4

# TODO: Horizontal Percentage Bar Charts 
# https://plotly.com/python/horizontal-bar-charts/

# TODO: Sunburst Charts
# https://plotly.com/python/sunburst-charts/



# Plot 5 & 6

fig5 = make_subplots(rows=3, cols=1, subplot_titles=('Loyal Customers', 'Big Spenders on Guestroom', 'Big Spenders on Total Revenue'), 
                    column_widths=[0.05], row_heights=[0.3, 0.3, 0.3], vertical_spacing=0.1, horizontal_spacing=0.0, 
                    specs=[[{"type": "table"}], [{"type": "table"}], [{"type": "table"}]])

fig6 = make_subplots(rows=3, cols=1, subplot_titles=('Promising Customers', 'New Potential Customers', 'Gone Big Spenders'), 
                    column_widths=[0.05], row_heights=[0.3, 0.3, 0.3], vertical_spacing=0.1, horizontal_spacing=0.0, 
                    specs=[[{"type": "table"}], [{"type": "table"}], [{"type": "table"}]])

rfgm_display_col = ['Account', 'Account: Region', 'Account: Industry', 'Recency', 'Frequency', 'Guestroom_value', 'Monetary_value', 'R', 'F', 'G', 'M']
RFGM_score_no_outlier = RFGM_score_no_outlier[rfgm_display_col]


# function to create rfm analysis table
def plotly_table_figure(fig_tmp, dataframe, columns):
    global plot_number
    table_tmp = go.Table(header = dict(values=columns),
                         cells = dict(values=[dataframe[k].tolist() for k in dataframe.columns[0:]]))
    fig_tmp.add_trace(table_tmp, row=plot_number, col=1)
    plot_number += 1


# Reset plot_number
plot_number = 1
loyal_customer = RFGM_score_no_outlier[RFGM_score_no_outlier['F'] >= 3].sort_values('Frequency', ascending=False)
plotly_table_figure(fig5, loyal_customer, rfgm_display_col)

big_spender_RN = RFGM_score_no_outlier[RFGM_score_no_outlier['G'] >= 3].sort_values('Guestroom_value', ascending=False)
plotly_table_figure(fig5, big_spender_RN, rfgm_display_col)

big_spender_Rev = RFGM_score_no_outlier[(RFGM_score_no_outlier['M'] >= 3) & (RFGM_score_no_outlier['R'] <= 3)].sort_values('Monetary_value', ascending=False)
plotly_table_figure(fig5, big_spender_Rev, rfgm_display_col)

# Reset plot_number
plot_number = 1
promising_customer = RFGM_score_no_outlier[(RFGM_score_no_outlier['F'] == 2) & (RFGM_score_no_outlier['M'] <= 3) & (RFGM_score_no_outlier['R'] <= 3)].sort_values('Frequency', ascending=False)
plotly_table_figure(fig6, promising_customer, rfgm_display_col)

new_potential = RFGM_score_no_outlier[(RFGM_score_no_outlier['R'] == 5) & (RFGM_score_no_outlier['Frequency'] <= 2)].sort_values('Monetary_value', ascending=False)
plotly_table_figure(fig6, new_potential, rfgm_display_col)

gone_big_spender = RFGM_score_no_outlier[(RFGM_score_no_outlier['R'] >= 2) & (RFGM_score_no_outlier['M'] >= 2)].sort_values('Monetary_value', ascending=False)
plotly_table_figure(fig6, gone_big_spender, rfgm_display_col)

fig5.update_layout(title='Customer RFGM scores', autosize=False, width=1800, height=800)

fig6.update_layout(autosize=False, width=1800, height=800)

def figures_to_html(figs, filename):
    dashboard = open(filename, 'w')
    dashboard.write("<html><head></head><body>" + "\n")
    for fig in figs:
        inner_html = fig.to_html().split('<body>')[1].split('</body>')[0]
        dashboard.write(inner_html)
    dashboard.write("</body></html>" + "\n")

figures_to_html([fig1, fig2, fig3, fig5, fig6], filename='portfolio_dashboard.html')

##################################################################################################
