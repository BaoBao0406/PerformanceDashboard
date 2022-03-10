#! python3
# performance_dashboard.py - 

import pyodbc, pickle, re
from datetime import date, datetime
import plotly
import pandas as pd
import plotly.offline as pyo
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import plotly.figure_factory as ff
import numpy as np
from dateutil.relativedelta import relativedelta

##################################################################################################

# Set date range
now = datetime.now()
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

# Date Definite data
date_def_df = pd.read_sql("SELECT BK.OwnerId, BK.nihrm__Property__c, ac.Name, ac.BillingCountry, ag.Name, BK.Name, FORMAT(BK.nihrm__ArrivalDate__c, 'MM/dd/yyyy') AS ArrivalDate, FORMAT(BK.nihrm__DepartureDate__c, 'MM/dd/yyyy') AS DepartureDate, \
                                 BK.nihrm__CurrentBlendedRoomnightsTotal__c, BK.nihrm__BlendedGuestroomRevenueTotal__c, \
                                 BK.VCL_Blended_F_B_Revenue__c, BK.nihrm__CurrentBlendedEventRevenue7__c, BK.nihrm__BookingStatus__c, FORMAT(BK.nihrm__LastStatusDate__c, 'MM/dd/yyyy') AS LastStatusDate, \
                                 FORMAT(BK.nihrm__BookedDate__c, 'MM/dd/yyyy') AS BookedDate, BK.End_User_Region__c, BK.End_User_SIC__c, BK.nihrm__BookingTypeName__c, BK.Booking_ID_Number__c, FORMAT(BK.nihrm__DateDefinite__c, 'MM/dd/yyyy') AS DateDefinite, BK.Date_Definite_Month__c, BK.Date_Definite_Year__c \
                          FROM dbo.nihrm__Booking__c AS BK \
                          LEFT JOIN dbo.Account AS ac \
                              ON BK.nihrm__Account__c = ac.Id \
                          LEFT JOIN dbo.Account AS ag \
                              ON BK.nihrm__Agency__c = ag.Id \
                          WHERE (BK.nihrm__BookingTypeName__c NOT IN ('ALT Alternative', 'CN Concert', 'IN Internal')) AND \
                              (BK.nihrm__Property__c NOT IN ('Sands Macao Hotel')) AND (BK.nihrm__BookingStatus__c = 'Definite') AND \
                              (BK.nihrm__DateDefinite__c BETWEEN CONVERT(datetime, '" + start_date_def + "') AND CONVERT(datetime, '" + end_date_def + "'))", conn)
date_def_df.columns = ['Owner Name', 'Property', 'Account', 'Company Country', 'Agency', 'Booking: Booking Post As', 'Arrival', 'Departure', 
                      'Blended Roomnights', 'Blended Guestroom Revenue Total', 'Blended F&B Revenue', 'Blended Rental Revenue', 'Status',
                      'Last Status Date', 'Booked', 'End User Region', 'End User SIC', 'Booking Type', 'Booking ID#', 'DateDefinite', 'Date Definite Month', 'Date Definite Year']
date_def_df['Owner Name'].replace(user, inplace=True)


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
                                    BK.nihrm__CurrentBlendedRoomnightsTotal__c, BK.nihrm__BlendedGuestroomRevenueTotal__c, BK.VCL_Blended_F_B_Revenue__c, BK.nihrm__CurrentBlendedEventRevenue7__c, BK.nihrm__BookingStatus__c, FORMAT(BK.nihrm__LastStatusDate__c, 'MM/dd/yyyy') AS LastStatusDate, \
                                    FORMAT(BK.nihrm__BookedDate__c, 'MM/dd/yyyy') AS BookedDate, BK.Booked_Year__c, BK.Booked_Month__c, BK.End_User_Region__c, BK.End_User_SIC__c, BK.nihrm__BookingTypeName__c, BK.Booking_ID_Number__c, FORMAT(BK.nihrm__DateDefinite__c, 'MM/dd/yyyy') AS DateDefinite, \
                                    BK.Date_Definite_Month__c, BK.Date_Definite_Year__c,  FORMAT(BK.LastActivityDate, 'MM/dd/yyyy') AS LastActivityDate \
                                FROM dbo.nihrm__Booking__c AS BK \
                                LEFT JOIN dbo.Account AS ac \
                                    ON BK.nihrm__Account__c = ac.Id \
                                LEFT JOIN dbo.Account AS ag \
                                    ON BK.nihrm__Agency__c = ag.Id \
                                WHERE (BK.nihrm__BookingTypeName__c NOT IN ('ALT Alternative', 'CN Concert', 'IN Internal')) AND \
                                    (BK.nihrm__Property__c NOT IN ('Sands Macao Hotel')) AND \
                                    (BK.nihrm__ArrivalDate__c BETWEEN CONVERT(datetime, '" + start_date_arrival + "') AND CONVERT(datetime, '" + end_date_arrival + "'))", conn)
arrival_date_df.columns = ['Owner Name', 'Property', 'Account', 'Company Country', 'Agency', 'Booking: Booking Post As', 'Arrival', 'Arrival Year', 'Arrival Month', 'Departure', 'Blended Roomnights', 'Blended Guestroom Revenue Total', 
                           'Blended F&B Revenue', 'Blended Rental Revenue', 'Status', 'Last Status Date', 'Booked', 'Booked Year', 'Booked Month', 'End User Region', 'End User SIC', 'Booking Type', 'Booking ID#', 'DateDefinite', 
                           'Date Definite Month', 'Date Definite Year', 'LastActivityDate']
arrival_date_df['Owner Name'].replace(user, inplace=True)
arrival_date_df['Days before Arrive'] = ((pd.to_datetime(arrival_date_df['Arrival']) - pd.Timestamp.now().normalize()).dt.days).astype(int)
arrival_date_df.sort_values(by=['Days before Arrive'], inplace=True)


# SM Inquiry
inquiry = pd.read_sql("SELECT inq.OwnerId, inq.nihrm__Property__c, inq.nihrm__Account__c, inq.nihrm__Company__c, inq.Name, inq.nihrm__ArrivalDate__c, inq.nihrm__Guests__c, inq.nihrm__TotalRooms__c, inq.nihrm__Status__c \
                       FROM dbo.nihrm__Inquiry__c AS inq", conn)
inquiry.columns = ['Owner Name', 'Property', 'Account', 'Company', 'Name', 'Arrival', 'Guests', 'Total Rooms', 'Status']
inquiry['Owner Name'].replace(user, inplace=True)


# ML prediction
BK_ml_tmp = pd.read_sql("SELECT BK.Id, BK.Booking_ID_Number__c, FORMAT(BK.nihrm__ArrivalDate__c, 'yyyy-MM-dd') AS ArrivalDate, FORMAT(BK.nihrm__DepartureDate__c, 'yyyy-MM-dd') AS DepartureDate, BK.nihrm__CommissionPercentage__c, BK.Percentage_of_Attrition__c, BK.nihrm__Property__c, BK.nihrm__FoodBeverageMinimum__c, ac.Name AS ACName, ag.Name AS AGName, BK.End_User_Region__c, \
                                BK.End_User_SIC__c, BK.nihrm__BookingTypeName__c, BK.RSO_Manager__c, BK.Non_Compete_Clause__c, ac.nihrm__RegionName__c, ac.Industry, BK.nihrm__CurrentBlendedRoomnightsTotal__c, BK.nihrm__BlendedGuestroomRevenueTotal__c, BK.VCL_Blended_F_B_Revenue__c, BK.nihrm__CurrentBlendedEventRevenue7__c, BK.nihrm__CurrentBlendedEventRevenue4__c, \
                                BK.nihrm__BookingMarketSegmentName__c, BK.Promotion__c, BK.nihrm__CurrentBlendedADR__c, BK.nihrm__PeakRoomnightsBlocked__c, FORMAT(BK.nihrm__BookedDate__c, 'yyyy-MM-dd') AS BookedDate, FORMAT(BK.nihrm__LastStatusDate__c, 'yyyy-MM-dd') AS LastStatusDate \
                         FROM dbo.nihrm__Booking__c AS BK \
                             LEFT JOIN dbo.Account AS ac \
                                 ON BK.nihrm__Account__c = ac.Id \
                             LEFT JOIN dbo.Account AS ag \
                                 ON BK.nihrm__Agency__c = ag.Id \
                         WHERE (BK.nihrm__BookingTypeName__c NOT IN ('ALT Alternative', 'CN Concert', 'IN Internal')) AND \
                             (BK.nihrm__Property__c NOT IN ('Sands Macao Hotel')) AND (BK.nihrm__BookingStatus__c IN ('Tentative', 'Prospect'))", conn)
BK_ml_tmp['RSO_Manager__c'].replace(user, inplace=True)
BK_ml_tmp.columns = ['Id', 'BK_no', 'ArrivalDate', 'DepartureDate', 'Commission', 'Attrition', 'Property', 'F&B Minimum', 'Account', 'Agency', 'End User Region',
                  'End User SIC', 'Booking Type', 'RSO Manager', 'Non-compete', 'Account: Region', 'Account: Industry', 'Blended Roomnights', 'Blended Guestroom Revenue Total',
                  'Blended F&B Revenue', 'Blended Rental Revenue', 'Blended AV Revenue', 'Market Segment', 'Promotion', 'Blended ADR', 'Peak Roomnights Blocked', 
                  'BookedDate', 'LastStatusDate']

Event_ml_tmp = pd.read_sql("SELECT ET.nihrm__Booking__c, ET.nihrm__Property__c, MAX(ET.nihrm__AgreedAttendance__c) \
                         FROM dbo.nihrm__BookingEvent__c AS ET \
                         WHERE (ET.nihrm__Property__c NOT IN ('Sands Macao Hotel')) AND (ET.nihrm__StartDate__c BETWEEN CONVERT(datetime, '" + start_date_arrival + "') AND CONVERT(datetime, '" + end_date_arrival + "')) \
                         GROUP BY ET.nihrm__Booking__c, ET.nihrm__Property__c", conn)
Event_ml_tmp.columns = ['Id', 'property', 'Attendance']
Event_ml_tmp = Event_ml_tmp[['Id', 'Attendance']]
BK_ml_tmp = BK_ml_tmp.join(Event_ml_tmp.set_index('Id'), on='Id')



##################################################################################################

# ML prediction

BK_ml_percent = BK_ml_tmp[['Property', 'Account: Region', 'Account: Industry', 'Agency', 'End User Region', 'End User SIC', 'Booking Type', 'Blended Roomnights',
                        'Blended Guestroom Revenue Total', 'Blended F&B Revenue', 'Blended Rental Revenue', 'Blended AV Revenue', 'Attendance',
                        'RSO Manager', 'Market Segment', 'Promotion', 'Blended ADR', 'Peak Roomnights Blocked', 'ArrivalDate', 'DepartureDate', 
                        'BookedDate', 'LastStatusDate']]

# calculate Inhouse day (Departure - Arrival)    
BK_ml_percent['Inhouse day'] = (pd.to_datetime(BK_ml_percent['DepartureDate']).dt.date - pd.to_datetime(BK_ml_percent['ArrivalDate']).dt.date).dt.days
# calculate Lead day (Arrival - Booked) 
BK_ml_percent['Lead day'] = (pd.to_datetime(BK_ml_percent['ArrivalDate']).dt.date - pd.to_datetime(BK_ml_percent['BookedDate']).dt.date).dt.days
# calculate Decision day (Last Status date - Booked) 
BK_ml_percent['Decision day'] = (pd.to_datetime(BK_ml_percent['LastStatusDate']).dt.date - pd.to_datetime(BK_ml_percent['BookedDate']).dt.date).dt.days
# booking info Arrival Month
BK_ml_percent['Arrival Month'] = pd.DatetimeIndex(BK_ml_percent['ArrivalDate']).month
BK_ml_percent['Arrival Month_sin'] = np.sin(2 * np.pi * BK_ml_percent['Arrival Month']/12)
# booking info Booked Month
BK_ml_percent['Booked Month'] = pd.DatetimeIndex(BK_ml_percent['BookedDate']).month
BK_ml_percent['Booked Month_sin'] = np.sin(2 * np.pi * BK_ml_percent['Booked Month']/12)
# booking info Last Status Month
BK_ml_percent['Last Status Month'] = pd.DatetimeIndex(BK_ml_percent['LastStatusDate']).month
BK_ml_percent['Last Status Month_sin'] = np.sin(2 * np.pi * BK_ml_percent['Last Status Month']/12)

BK_ml_percent['Agency'] = BK_ml_percent['Agency'].apply(lambda x: 0 if x is np.nan else 1)
BK_ml_percent['RSO Manager'] = BK_ml_percent['RSO Manager'].apply(lambda x: 0 if x is np.nan else 1)
BK_ml_percent['Promotion'] = BK_ml_percent['Promotion'].apply(lambda x: 0 if x is np.nan else 1)

BK_ml_percent = BK_ml_percent[['Property', 'Account: Region', 'Account: Industry', 'Agency', 'End User Region', 'End User SIC', 'Booking Type', 'Blended Roomnights', 'Blended Guestroom Revenue Total', 
                                 'Blended F&B Revenue', 'Blended Rental Revenue', 'Blended AV Revenue', 'Attendance', 'RSO Manager', 'Market Segment', 'Promotion', 'Blended ADR', 'Peak Roomnights Blocked', 
                                 'Inhouse day', 'Lead day', 'Decision day', 'Arrival Month_sin', 'Booked Month_sin', 'Last Status Month_sin']]

# Load Transformer and prediction model
path = 'I:\\10-Sales\\+Dept Admin (3Y, Internal)\\2021\\Personal Folders\\Patrick Leong\\Python Code\\Sales Dashboard\\'
Transformer = open(path + 'Materization_percent_tf.pkl', 'rb')
Transformer = pickle.load(Transformer)

Model = open(path + 'Materization_percent_ml.pkl', 'rb')
Model = pickle.load(Model)

columns_to_standarize = ['Blended Roomnights', 'Blended Guestroom Revenue Total', 'Blended F&B Revenue', 'Blended Rental Revenue',
                         'Blended AV Revenue', 'Attendance', 'Blended ADR', 'Peak Roomnights Blocked']
columns_to_te = ['Property', 'Account: Region', 'Account: Industry', 'End User Region', 'End User SIC', 'Booking Type', 'Market Segment']
column_normal = ['Agency', 'RSO Manager', 'Promotion', 'Inhouse day', 'Lead day', 'Decision day', 'Arrival Month_sin',
                 'Booked Month_sin', 'Last Status Month_sin']

column_name = columns_to_standarize + columns_to_te + column_normal

BK_ml_percent_tf = Transformer.transform(BK_ml_percent)
BK_ml_percent_tf = pd.DataFrame(BK_ml_percent_tf, columns=column_name)
BK_ml_percent_tf = BK_ml_percent_tf.rename(columns = lambda x:re.sub('[^A-Za-z0-9_]+', '', x))
BK_ml_percent_tf = BK_ml_percent_tf.astype(float)
y_pred_tf = Model.predict_proba(BK_ml_percent_tf)

# Join predication to BK ID no
y_pred_tf = pd.DataFrame(y_pred_tf, columns = ['TD %', 'D %'])
y_pred_tf.reset_index(drop=True, inplace=True)
BK_ml_pred = BK_ml_tmp[['BK_no']]
BK_ml_pred.reset_index(drop=True, inplace=True)
BK_ml_pred = pd.concat([BK_ml_pred, y_pred_tf], axis=1)


##################################################################################################

# Filter data for each team

# Definite
sm_production = date_def_df[date_def_df['Owner Name'] == 'Luis Wan']

sm_current_production = sm_production[sm_production['Date Definite Year'] == now.year]
sm_current_production.fillna("-", inplace = True)

# Tentaive and Prospect
sm_current_business = arrival_date_df[arrival_date_df['Owner Name'] == 'Luis Wan']
# join ml prediction table to T & P table
sm_current_business = pd.merge(sm_current_business, BK_ml_pred, how='left', left_on=['Booking ID#'], right_on= ['BK_no'])
sm_current_business['D %'] = sm_current_business['D %'].map(lambda x: "{0:.2f}%".format(x*100))
sm_current_business.fillna("-", inplace = True)

sm_current_business['LastActivityDate'] = np.where(sm_current_business['LastActivityDate'] == '-', sm_current_business['Booked'], sm_current_business['LastActivityDate'])
sm_current_business['color_display'] = np.where((date.today() - pd.to_datetime(sm_current_business['LastActivityDate']).dt.date).dt.days < 14, 'rgb(189, 215, 231)', 'rgb(230, 230, 230)')
sm_current_business['color_display'] = np.where((date.today() - pd.to_datetime(sm_current_business['Booked']).dt.date).dt.days < 14, 'rgb(255, 255, 102)', sm_current_business['color_display'])

##################################################################################################

# Pre process data

# date range for bar chart
end_year = now + relativedelta(years=2)
#plot_start_date = str(current_year.year) + '-' + str(current_year.month) + '-01'
plot_start_date = str(now.year) + '-' + '01' + '-01'
plot_end_date = str(end_year.year) + '-' + str(end_year.month) + '-01'
arrival_date_range = pd.date_range(start=plot_start_date, end=plot_end_date, freq='MS')
arrival_date_label = arrival_date_range.strftime('%Y-%m').to_frame()
#current_month = current_year.strftime("%b")
current_month = 'Feb'


# Plot 1 - Definite Month bar chart figure
months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
sm_current_production['to_sort'] = sm_current_production['Date Definite Month'].apply(lambda x: months.index(x))
sm_current_production = sm_current_production.sort_values('to_sort')


# Plot 2 - 
sm_production = date_def_df
plt_2_current_production = sm_production[sm_production['Date Definite Year'] == now.year]
plt_2_current_production.fillna("-", inplace = True)
plt_2_current_production['Arrival_year_month'] = pd.to_datetime(plt_2_current_production['Arrival']).dt.strftime('%Y-%m')
# History production for bar chart
plt_2_current_production_hist = plt_2_current_production[plt_2_current_production['Date Definite Month'] != current_month]
plt_2_current_production_hist = plt_2_current_production_hist.groupby(['Arrival_year_month'])['Blended Roomnights'].sum().reset_index().set_index('Arrival_year_month')
plt_2_current_production_hist = pd.merge(arrival_date_label, plt_2_current_production_hist, how='left', left_index=True, right_index=True).reset_index(drop=True).fillna(0)
# Current production for bar chart
plt_2_current_production_curr = plt_2_current_production[plt_2_current_production['Date Definite Month'] == current_month]
plt_2_current_production_curr = plt_2_current_production_curr.groupby(['Arrival_year_month'])['Blended Roomnights'].sum().reset_index().set_index('Arrival_year_month')
plt_2_current_production_curr = pd.merge(arrival_date_label, plt_2_current_production_curr, how='left', left_index=True, right_index=True).reset_index(drop=True).fillna(0)
# Total team production for line chart
plt_2_total_production = arrival_date_df[arrival_date_df['Status'] == 'Definite']
plt_2_total_production['Arrival_year_month'] = pd.to_datetime(plt_2_total_production['Arrival']).dt.strftime('%Y-%m')
plt_2_total_production =  plt_2_total_production.groupby(['Arrival_year_month'])['Blended Roomnights'].sum().reset_index().set_index('Arrival_year_month')
plt_2_total_production = pd.merge(arrival_date_label, plt_2_total_production, how='left', left_index=True, right_index=True).reset_index(drop=True).fillna(0)


# Plot 3 - Production vs Budget
plt_3_total_production = sm_current_production
plt_3_total_production['to_sort'] = plt_3_total_production['Date Definite Month'].apply(lambda x: months.index(x))
plt_3_total_production = plt_3_total_production.sort_values('to_sort')


# Plot 4 - Tentative & Prospect
bk_display_col = ['Property', 'Account', 'Agency', 'Booking: Booking Post As', 'Arrival', 'Departure', 'Blended Roomnights', 'D %', 'Days before Arrive', 'color_display']
plt_4_current_business_t = sm_current_business[(sm_current_business['Status'] == 'Tentative') & (sm_current_business['Owner Name'] == 'Luis Wan')][bk_display_col]
plt_4_current_business_p = sm_current_business[(sm_current_business['Status'] == 'Prospect') & (sm_current_business['Owner Name'] == 'Luis Wan')][bk_display_col]


# Plot 5 - Inquiry
inq_display_col = ['Property', 'Account', 'Company', 'Name', 'Arrival', 'Guests', 'Total Rooms']
plt_5_inquiry = inquiry[(inquiry['Owner Name'] == 'Luis Wan') & (inquiry['Status'] == 'Opened')][inq_display_col]
plt_5_inquiry.fillna("-", inplace = True)

##################################################################################################

# Plotly function

# Plot 1
fig1 = make_subplots(rows=1, cols=2, subplot_titles=('Monthly Definite RNs', 'Monthly Definite RN Revenue and Rental Revenue'), 
                    column_widths=[0.05, 0.05], row_heights=[0.3], shared_xaxes=True)


bar1 = go.Bar(x=sm_current_production['Date Definite Month'], y=sm_current_production['Blended Roomnights'], name='RNs')

bar2 = go.Bar(x=sm_current_production['Date Definite Month'], y=sm_current_production['Blended Guestroom Revenue Total'], name='RN Revenue')

bar3 = go.Bar(x=sm_current_production['Date Definite Month'], y=sm_current_production['Blended Rental Revenue'], name='Rental Revenue')

fig1.add_trace(bar1, row=1, col=1)
fig1.add_trace(bar2, row=1, col=2)
fig1.add_trace(bar3, row=1, col=2)

fig1.layout.xaxis.tickvals = months
fig1.layout.xaxis.tickformat = '%b'

fig1.update_layout(title='Definite Bookings', autosize=False, width=1800, height=500)


# Plot 2
fig2 = make_subplots(rows=2, cols=1, subplot_titles=('Monthly Definite RNs', 'Total RNs Definite in Sales Team'), 
                    column_widths=[0.05], row_heights=[0.3, 0.3], shared_xaxes=True)

bar1 = go.Bar(x=plt_2_current_production_hist[0], y=plt_2_current_production_hist['Blended Roomnights'])
bar2 = go.Bar(x=plt_2_current_production_curr[0], y=plt_2_current_production_curr['Blended Roomnights'])
line1 = go.Scatter(x=plt_2_total_production[0], y=plt_2_total_production['Blended Roomnights'], mode='lines+markers')

fig2.add_trace(bar1, row=1, col=1)
fig2.add_trace(bar2, row=1, col=1)
fig2.add_trace(line1, row=2, col=1)

fig2.update_layout(title='Definite Bookings', autosize=False, width=2000, height=1000, barmode='stack', 
                   xaxis = dict(title='Arrival Month', showticklabels=True, type='category', tickangle=45), showlegend=False)
fig2.update_xaxes(row=2, col=1, type='category', tickangle=45)


# Plot 3
fig3 = make_subplots(rows=12, cols=1, 
                     column_widths=[0.3], row_heights=[0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5], shared_xaxes=True,
                     specs=[[{"type": "domain"}], [{"type": "domain"}], [{"type": "domain"}], [{"type": "domain"}], [{"type": "domain"}], [{"type": "domain"}], 
                            [{"type": "domain"}], [{"type": "domain"}], [{"type": "domain"}], [{"type": "domain"}], [{"type": "domain"}], [{"type": "domain"}]])

max_val = 1000
threshold = 500

for i, month in enumerate(months):
    bullet = go.Indicator(mode='number+gauge+delta', title={'text': month}, delta={'reference': threshold},
                          value=plt_3_total_production[plt_3_total_production['Date Definite Month'] == month]['Blended Roomnights'].sum(),
                          gauge={'shape': 'bullet', 'bordercolor': 'white', 'borderwidth': 2,
                                 'axis': {'range': [0, max_val], 'ticklen': 0.000005, 'tickwidth': 0.0005},
                                 'steps': [{'range': [0, max_val], 'color': 'lightgray'},
                                           {'range': [0, threshold], 'color': 'gray'}],
                                 'threshold': {'line': {'color': 'red', 'width': 2}, 'thickness': 0.75,
                                               'value': threshold}})
    fig3.add_trace(bullet, row=i+1, col=1)

fig3.update_layout(title='Production Vs Budget', autosize=False, width=1500, height=750)


# Plot 4
fig4 = make_subplots(rows=2, cols=1, subplot_titles=('Tentative Bookings', 'Prospect Bookings'), 
                    column_widths=[0.05], row_heights=[0.3, 0.3], vertical_spacing=0.1, horizontal_spacing=0.0, 
                    specs=[[{"type": "table"}], [{"type": "table"}]])

table1_obj = go.Table(header = dict(values=bk_display_col),
                      cells = dict(values=[plt_4_current_business_t[k].tolist() for k in plt_4_current_business_t.columns[0:]]))

table2_obj = go.Table(header = dict(values=bk_display_col),
                      cells = dict(values=[plt_4_current_business_p[k].tolist() for k in plt_4_current_business_p.columns[0:]], line_color='white', fill_color=[plt_4_current_business_p['color_display']]))

fig4.add_trace(table1_obj, row=1, col=1)
fig4.add_trace(table2_obj, row=2, col=1)

fig4.update_layout(title='Current Business', autosize=False, width=2000, height=1200)


# Plot 5
fig5 = make_subplots(rows=1, cols=1, subplot_titles='Inquiries', 
                    column_widths=[0.05], row_heights=[0.3], vertical_spacing=0.1, horizontal_spacing=0.0, 
                    specs=[[{"type": "table"}]])

table3_obj = go.Table(header = dict(values=inq_display_col),
                      cells = dict(values=[plt_5_inquiry[k].tolist() for k in plt_5_inquiry.columns[0:]]))

fig5.add_trace(table3_obj, row=1, col=1)

fig5.update_layout(title='Current Business', autosize=False, width=2000, height=1200)



def figures_to_html(figs, filename):
    dashboard = open(filename, 'w')
    dashboard.write("<html><head></head><body>" + "\n")
    for fig in figs:
        inner_html = fig.to_html().split('<body>')[1].split('</body>')[0]
        dashboard.write(inner_html)
    dashboard.write("</body></html>" + "\n")

figures_to_html([fig1, fig2, fig3, fig4, fig5], filename='performance_dashboard.html')

##################################################################################################
