import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats

file_name = 'TN20_data'
raw_data = pd.read_excel(file_name+'.xlsx')
data = raw_data.copy()

time = data['tempo']
temp_5 = data['temp_5']
temp_15 = data['temp_15']
temp_25 = data['temp_25']
temp_60 = data['temp_60']
emissions = data['emissoes_co']

#new column with the max values between all temps
data['temp_max'] = ''
max_temp = data['temp_max']

#fill the max_temp column with the maximum value per row
for row in range(0, len(data)):
    t5 = temp_5.iloc[row]
    t15 = temp_15.iloc[row]
    t25 = temp_25.iloc[row]
    t60 = temp_60.iloc[row]
    array = np.array([t5,t15,t25,t60])
    
    max_temp.iloc[row] = np.max(array)

#loops the data and calculates the slope if the difference between the row and the previous row is higher than min_temp_dif
#interval defines when to check the slope (s)
slope_data = []
interval = 2 #time interval between slope points
min_temp_dif = 5 #temperature difference between points
for row in range(0, len(data)): #starts at 1 to exclude row = 0
    
    if row % interval == 0 and row != 0:
        previous_row = row - interval
        temp_dif = max_temp.iloc[previous_row] - max_temp.iloc[row]
        
        if np.abs(temp_dif) > min_temp_dif:
            array = np.array([[time.iloc[previous_row], time.iloc[row]], [max_temp.iloc[previous_row], max_temp.iloc[row]]])
            array_stats = stats.linregress(array)
            slope = array_stats.slope
        
        #exemple of row 50 for data validation
        if row == 50:
            print(f'Ponto anterior: {previous_row}, \n'
                      f'Tempo: {time.iloc[previous_row]}, \n'
                      f'Temperatura: {max_temp.iloc[previous_row]}, \n'
                      f'\n'
                      f'Ponto atual: {row}, \n'
                      f'Tempo: {time.iloc[row]}, \n'
                      f'Temperatura: {max_temp.iloc[row]}, \n'
                      f'\n'
                      f'Declive: {array_stats.slope}')
    else:
        slope = 0
            
    slope_data.append(slope)

data['slope_data'] = [i for i in slope_data]
data.to_excel(f'{file_name}_slope_data.xlsx', index=False)
