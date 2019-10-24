import pandas as pd

pd.set_option('display.max_columns', None)

orders = pd.read_excel('Input/191010_Global wind turbine order database - September 2019.xlsx',
              sheet_name='Order Tracking ',
              skiprows=16,
              skipfooter=2,
              usecols="B:AA")
              
MW = orders.melt(id_vars=['Supplier', 'Region', 'Sub-region', 'Country', 'Project location',
       'Project name', 'Buyer', 'Order Date', 'Order Status', 'Project Type',
       'Order MW', 'Turbine Model', 'Turbine #', 'MW Rating',
       'Registered Date', 'Latest change date'], value_vars=['2018', '2019', '2020', '2021', '2022'],\
        value_name='MW', var_name='Year')


turbs = orders.melt(id_vars=['Supplier', 'Region', 'Sub-region', 'Country', 'Project location',
       'Project name', 'Buyer', 'Order Date', 'Order Status', 'Project Type',
       'Order MW', 'Turbine Model', 'Turbine #', 'MW Rating',
       'Registered Date', 'Latest change date'], value_vars=['´2018', '´2019', '´2020', '\'2021', '\'2022'],\
        value_name='# Turbines', var_name='Year')

turbs.replace({'´2018':'2018','´2019':'2019','´2020':'2020','\'2021':'2021','\'2022':'2022'}, inplace=True)

turbs = turbs[['# Turbines']].copy()

trans = pd.concat([MW, turbs], axis=1)

trans.to_excel('Output/Orders Transformed.xlsx')
