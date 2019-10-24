import pandas as pd
import numpy as np

pd.set_option('display.max_columns', None)

excl = [np.nan,'','North America','Total North America','South America','Total South America',
                'Caribbean','Total Caribbean','Central America','Total Central America','Total Latin America',
                'Total Latin America (Excl. Brazil)','Total Americas','Northern Europe','Total Northern Europe',
                'Western Europe','Total Western Europe','Southern Europe','Total Southern Europe','Eastern Europe',
                'Total Eastern Europe','Russia and the Caspian','Total Russia and the Caspian',
                'Total Eastern Europe plus Russia and the Caspian','Total Europe without Russian and the Caspian',
                'Total Europe', 'Africa','Total Africa','Middle East','Total Middle East','Total Middle East & Africa',
                'Total EMEA (not including RC)','Total EMEA','Asia Pacific','Total Asia Pacific',
                'Total Asia Pacific (ex-China)','Total Asia Pacific (ex-China and India)','Global Total',
                'Global Total (Excl. China)']


on_nb = pd.read_excel('Input/MOU Master Data.xlsm',
              sheet_name='Onshore–New Build',
              skiprows=6,
              skipfooter=10,
              usecols="B:AF")
on_nb.rename({'Unnamed: 0':'Country'}, axis='columns', inplace=True)
on_nb = on_nb[~on_nb.Country.isin(excl)]


on_re = pd.read_excel('Input/MOU Master Data.xlsm',
              sheet_name='Onshore–Repowering',
              skiprows=6,
              skipfooter=10,
              usecols="B:AF")
on_re.rename({'Unnamed: 0':'Country'}, axis='columns', inplace=True)
on_re = on_re[~on_re.Country.isin(excl)]


on_de = pd.read_excel('Input/MOU Master Data.xlsm',
              sheet_name='Onshore–Decommissioning',
              skiprows=6,
              skipfooter=10,
              usecols="B:AF")
on_de.rename({'Unnamed: 0':'Country'}, axis='columns', inplace=True)
on_de = on_de[~on_de.Country.isin(excl)]


off_nb = pd.read_excel('Input/MOU Master Data.xlsm',
              sheet_name='Offshore-New Build',
              skiprows=6,
              skipfooter=10,
              usecols="B:AF")
off_nb.rename({'Unnamed: 0':'Country'}, axis='columns', inplace=True)
off_nb = off_nb[~off_nb.Country.isin(excl)]


off_re = pd.read_excel('Input/MOU Master Data.xlsm',
              sheet_name='Offshore–Repowering',
              skiprows=6,
              skipfooter=10,
              usecols="B:AF")
off_re.rename({'Unnamed: 0':'Country'}, axis='columns', inplace=True)
off_re = off_re[~off_re.Country.isin(excl)]


off_de = pd.read_excel('Input/MOU Master Data.xlsm',
              sheet_name='Offshore–Decommissioning',
              skiprows=6,
              skipfooter=10,
              usecols="B:AF")
off_de.rename({'Unnamed: 0':'Country'}, axis='columns', inplace=True)
off_de = off_de[~off_de.Country.isin(excl)]


on_nb = on_nb.melt(id_vars='Country', var_name='Year', value_name='Capacity')
on_de = on_de.melt(id_vars='Country', var_name='Year', value_name='Capacity')
on_re = on_re.melt(id_vars='Country', var_name='Year', value_name='Capacity')
off_nb = off_nb.melt(id_vars='Country', var_name='Year', value_name='Capacity')
off_de = off_de.melt(id_vars='Country', var_name='Year', value_name='Capacity')
off_re = off_re.melt(id_vars='Country', var_name='Year', value_name='Capacity')


on_nb['Site'] = 'Onshore'
on_de['Site'] = 'Onshore'
on_re['Site'] = 'Onshore'
off_nb['Site'] = 'Offshore'
off_de['Site'] = 'Offshore'
off_re['Site'] = 'Offshore'

on_nb['Status'] = 'New Build'
on_de['Status'] = 'Decommissioning'
on_re['Status'] = 'Repowering'
off_nb['Status'] = 'New Build'
off_de['Status'] = 'Decommissioning'
off_re['Status'] = 'Repowering'


all = pd.concat([on_nb,on_de,on_re,off_nb,off_de,off_re], axis=0)

all.to_excel('Output/MOU Transformed.xlsx')
