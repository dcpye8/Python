This script reads an excel file with temperature and emission data, calculates the maximum temperature per row and the slope between points where the temperature difference is higher than a minimum threshold. The resulting data is saved to a new excel file.

The script starts by importing the required modules: pandas for data handling, numpy for numerical operations, matplotlib.pyplot for plotting and scipy.stats for linear regression.

Then, the excel file is read and stored in a pandas dataframe called 'raw_data'. A copy of the original data is made to the 'data' variable.

Next, the script extracts the relevant columns from the data (time, temperature at 5mm, 15mm, 25mm and 60mm, and CO emissions), and creates an empty column called 'temp_max' to store the maximum temperature per row.

Then, a loop is used to fill the 'temp_max' column with the maximum temperature value between all temperatures at different depths. The loop goes through each row, extracts the temperature values and creates an array with them. The 'np.max' function is then used to find the maximum value in the array, which is stored in the 'temp_max' column for that row.

Finally, another loop is used to calculate the slope between points where the temperature difference is higher than a minimum threshold. The loop goes through each row and checks if the time difference between the current and previous row is equal to the interval defined by the 'interval' variable. If it is, the temperature difference between the maximum temperatures of the current and previous row is checked against the minimum threshold defined by the 'min_temp_dif' variable. If the temperature difference is higher than the threshold, a linear regression is performed on the data points (time and maximum temperature) between the current and previous row using the 'stats.linregress' function. The resulting slope value is stored in a list called 'slope_data'. If the temperature difference is lower than the threshold, the slope value is set to zero.

At the end of the loop, the 'slope_data' list is added to the 'data' dataframe as a new column, and the resulting data is saved to a new excel file with a '_slope_data' suffix added to the original file name.
