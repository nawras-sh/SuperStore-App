# Import Library
import pandas as pd
import seaborn as sns
sns.set()
import matplotlib.pyplot as plt
import streamlit as st
import plotly.express as px
import os
import datetime as dt
import openpyxl 
from  io import BytesIO
import plotly.figure_factory as ff
import plotly.graph_objects as go
import warnings
warnings.filterwarnings('ignore')

# Import Data
df= pd.read_csv('Sample-Superstore.csv')


# Set config page.
st.set_page_config(page_title='Dashboard',
                   page_icon=':bar-chart:',
                   layout='wide',initial_sidebar_state='auto')

# Set title Dashboard.
_, col_title = st.columns((0.5,1))
st.write("")

with col_title:
    st.title(":bar_chart: Sales Dashboard")
    st.markdown('<style>div.block-container{padding-top:4rem;}</style>', unsafe_allow_html=True)

#________________________________________________________
# Filtering Date

y_col, q_col, m_col = st.columns(3)

# Create Filters.
df['Order Date'] = pd.to_datetime(df['Order Date'],dayfirst= True , format='%d/%m/%Y')
df['year'] = df['Order Date'].dt.year
df['month'] = df['Order Date'].dt.month_name()
df['quarter'] = df['Order Date'].dt.quarter


with y_col:

# Year Filter
    df = df.sort_values(['year'])   
    f_year = st.multiselect('📆 Year', df['year'].unique())
    if not f_year:
        f_year_df = df.copy()
    else:
        f_year_df = df[df['year'].isin(f_year)]


with q_col:
    # Quarter Filter.
    f_year_df = f_year_df.sort_values(['quarter'])
    f_quarter = st.multiselect('📆 Quarter',f_year_df['quarter'].unique())
    if not f_quarter:
        f_quarter_df = f_year_df.copy()
    else:
        f_quarter_df = f_year_df[f_year_df['quarter'].isin(f_quarter)]


with m_col:
    # Month Filter
    months_full = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" ]
    f_quarter_df['month'] = pd.Categorical(f_quarter_df['month'], categories=months_full, ordered=True)
    f_quarter_df = f_quarter_df.sort_values(['month'])

    f_month = st.multiselect('📆 Month', f_quarter_df['month'].unique())
    if not f_month:
        f_month_df = f_quarter_df.copy()
    else:
        f_month_df  = f_quarter_df[f_quarter_df['month'].isin(f_month)]


#________________________________________________________
# Sidebar
#
#import image to sidebar
st.sidebar.image(f'Carrefour.png')
# set lastup date the data depend on the lase order
last_date = df['Order Date'].max().strftime('%d/%b/%Y')
st.sidebar.write(f'♻Last Update: {last_date}')

st.sidebar.header('Choose Your Filter')

# Create region filter
f_region = st.sidebar.multiselect('Region',df['Region'].unique())
if not f_region:
    f_region_df = f_month_df.copy()
else:
    f_region_df = f_month_df[f_month_df['Region'].isin(f_region)]

# Create state filter
f_state = st.sidebar.multiselect('State' ,f_region_df['State'].unique())
if not f_state:
    f_state_df = f_region_df.copy()
else:
    f_state_df= f_region_df[f_region_df['State'].isin(f_state)]

# Create city filter
f_city = st.sidebar.multiselect('City', f_state_df['City'].unique())
if not f_city:
    f_city_df = f_state_df.copy()
else:
    f_city_df = f_state_df[f_state_df['City'].isin(f_city)]

# Show Creater the Project and his contact
st.sidebar.markdown(f"Made By Eng. Nawras Sharfaldeen😊  \n  Email: Nawras.sharf@gmail.com")

# Create a copy of the last filter as filter_df for dealing with it.
filter_df = f_city_df.copy()

#______________________________________________________
# Body

# Custom CSS for bordered metric
st.markdown("""
    <style>
    .metric-container {
        border: 0px solid #ffffff; /* Border color */
        border-radius: 10px; /* Rounded corners */
        padding: 3px; /* Space inside the border */
        margin-bottom: 2px; /* Space below the metric */
        margin-right: 110px;
        margin-left: 110px;
        text-align: center; /* Center the text */
        background-color:#555;
    }
    .metric-container .value {
        font-size: 1.2em; /* Font size for the value */
        font-weight: bold; /* Bold font for the value */
        color: #ffffff; /* Color for the value */
    }
    .metric-container .label {
        font-size: 1em; /* Font size for the label */
        color: #ffffff; /* Color for the label */
    }
    </style>
""", unsafe_allow_html=True)

# Function to display a bordered metric
def bordered_metric(label, value):
    st.markdown(f"""
        <div class="metric-container">
            <div class="label">{label}</div>
            <div class="value">{value}</div>
        </div>
    """, unsafe_allow_html=True)

 # Creae columns for matric
_, metric_col1,metric_col2 =st.columns([0.1,4,4])
with metric_col1:
    # Total Of Sales
    bordered_metric('Total Of Sales',  '{:,.0f}'.format(filter_df['Sales'].sum()))

with metric_col2:
    # Total of Profit
    bordered_metric('Total Of Profit', '{:,.0f}'.format(filter_df['Profit'].sum())) #💰

st.divider()

#________________________________________________________
# Function to convert data fram to excle 

def to_excel(x):
    output = BytesIO() #to create an in-memory buffer
    with pd.ExcelWriter(output, engine='openpyxl') as writer: #to write the DataFrame to the buffer.
        x.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data
#________________________________________________________
# Time series analysis

# Bar and Scatter chart to show monthly Sales and Profit on the same chart
st.subheader('Monthly (Sales & Profots)')
pvt_sales_profit_by_month = filter_df.pivot_table(index='month',values=['Sales','Profit'] ,aggfunc='sum').reset_index().round(2)
pvt_sales_profit_by_month[['Sales', 'Profit']] = pvt_sales_profit_by_month[['Sales', 'Profit']].applymap(lambda x: '{:,.2f}'.format(x))


fig_sales_profit_by_month = go.Figure()
fig_sales_profit_by_month.add_trace(go.Bar(x=pvt_sales_profit_by_month['month'], y= pvt_sales_profit_by_month['Sales'], name='$ Sales',
                                           text =pvt_sales_profit_by_month['Sales']))  # ['{:,.0f}'.format(x)for x in pvt_sales_profit_by_month['Sales']]

fig_sales_profit_by_month.add_trace(go.Scatter(x=pvt_sales_profit_by_month['month'] , y= pvt_sales_profit_by_month['Profit'], mode = 'lines' , marker=dict(color='#ee4e04'),name='$ Profit'))
                                    

fig_sales_profit_by_month.update_layout(title = '',
                                        xaxis= dict(title='Month'),
                                        yaxis = dict(title = 'Sales', showgrid = False),
                                        yaxis2 = dict(title = 'Profit', overlaying ='y', side = 'right'),
                                        legend = dict (x=1 , y=1.2),
                                        template= 'gridon',
                                        )

st.plotly_chart(fig_sales_profit_by_month,use_container_width=True)

with st.expander('View Data'):
    st.write(ff.create_table(pvt_sales_profit_by_month))
    st.download_button('Download Data',data=to_excel(pvt_sales_profit_by_month),file_name='Sales profit Data.xlsx',help='click here to Download Excle File')

st.divider()
#_________________________________________
# Create Pie Chart For Sales By Region
pvt_sales_by_region = filter_df.pivot_table(index='Region', values='Sales', aggfunc='sum').reset_index()

col1, col2 = st.columns([2,3])

with col1:
    st.subheader('Most Selling Region')
    fig_sales_by_region = px.pie(pvt_sales_by_region,values='Sales', names='Region', hole=0.5, template='gridon')
    fig_sales_by_region.update_traces(text = pvt_sales_by_region['Region'])
    st.plotly_chart(fig_sales_by_region,use_container_width=True)

# Create bar Chart For Sales By Category
pvt_sales_by_cat = filter_df.pivot_table(index='Category', values='Sales',aggfunc='sum').reset_index()

with col2:
    st.subheader('Most Selling Category')
    fig_sales_by_cat = px.bar(pvt_sales_by_cat,x='Category', y='Sales', template='gridon',
                             text_auto='.2s',color='Category', pattern_shape_sequence=["-", "x", "+"]) #text=['${:,.2f}'.format(x) for x in pvt_sales_by_cat['Sales']]
    
    fig_sales_by_cat.update_traces(textfont_size=12, textangle=0, textposition="outside", cliponaxis=False)
    st.plotly_chart(fig_sales_by_cat,use_container_width=True)

# View and download data 

view1, view2 = st.columns([0.45,0.45])
# view sales by region data
with view1:
   with st.expander('Show as Table (Sales per Region)'):
       st.write(ff.create_table(pvt_sales_by_region))
       st.download_button('Downlad Data', to_excel(pvt_sales_by_region), file_name='Sales per Region.xlsx')

# view sales by category data
with view2:
    with st.expander('Show as Table (Sales Per Category)'):
        st.write(ff.create_table(pvt_sales_by_cat))
        st.download_button('Downlad Data', to_excel(pvt_sales_by_cat), file_name='Sales Per Category.xlsx')
st.write("")
st.divider()
#_________________________________________
# Time Series Comparison Sales Across Years 

gp_sales_across_years = filter_df.groupby(['year', 'month'], as_index=False)['Sales'].sum()

st.subheader('Comparison Sales Across Years')
fig_sales_across_years =  px.line(gp_sales_across_years , x='month', y= 'Sales', color='year', template='gridon',labels={'month':'Month'}) 
st.plotly_chart(fig_sales_across_years,use_container_width=True)

st.divider()
#________________________________________
# Creat Treemap based on Region and category nad Subcategory

st.subheader('Hierarchy of Sales (Category & Subcategory)')
fig_treemap = px.treemap(data_frame=filter_df,
                         path=['Region','Category','Sub-Category'],
                         values='Sales',
                         color='Category',
                         template='gridon')
fig_treemap.update_layout(width=500, height = 500)
st.plotly_chart(fig_treemap,use_container_width=True)

st.divider()
#_________________________________
# Sales By Ship Mode

col3,col4 = st.columns((5.2,5))

with col3:
    st.subheader('Sales Per Ship Mode')
    fig_sales_by_shipmode = px.pie(filter_df,names='Ship Mode', values='Sales', hole=0.5,template='presentation') 
    # fig_sales_by_shipmode.update_traces(text = filter_df['Ship Mode'])
    st.plotly_chart(fig_sales_by_shipmode,use_container_width=True)

with col4:
    st.subheader('Sales Per Segment')
    fig_sales_by_segment = px.pie(data_frame=filter_df,values='Sales', names='Segment',hole=0.5,template='ggplot2')
    st.plotly_chart(fig_sales_by_segment,use_container_width=True)


                                                    #Statistical Analysis
#________________________________________
# Relationship between Sales and Profit by Scatter Plot

_, col5 = st.columns([2.2, 5])
with col5:
    st.title('Statistical Analysis') 

st.divider()

st.subheader('Relationship between Sales & Profit')
fig_scatter = px.scatter(filter_df,x='Sales', y='Profit',
                         size='Sales',color='Profit',
                         hover_data=['Sales', 'Quantity', 'Region', 'Category'],
                         template='presentation')
st.plotly_chart(fig_scatter,use_container_width=True)

#________________________________________
# Statistic Values.

statistc_df = filter_df[['Sales', 'Profit', 'Quantity']].describe().round(2).reset_index()
statistc_df.rename(columns={'index': 'Statistic'},inplace=True)

with st.expander('Statistic Value'):
    st.write(ff.create_table(statistc_df))
    st.download_button('Downlad Data', to_excel(statistc_df),file_name='Statistic Data.xlsx')

st.divider()
#______________________________
#Correlation between Sales & Profit & Quantity

corr = filter_df[['Sales', 'Profit','Quantity']].corr()
st.subheader('Strength of the correlation between Sales & Profit & Quantity')
fig_corr = go.Figure(data=go.Heatmap(
    z= corr.values,
    x=corr.columns,
    y=corr.index,
    colorscale='Viridis'))

st.plotly_chart(fig_corr, use_container_width=True)
with st.expander('Correlation Matrix'):
    st.write(ff.create_table(corr.round(2).reset_index()))

st.divider()
#__________________________________
# Download Source Data.
st.subheader('Source Data')
st.download_button('Download', to_excel(filter_df), file_name='Carrefour.xlsx' ,help='Click here to download data source')
