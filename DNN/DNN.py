import pandas as pd
import os
import joblib
import numpy as np
import sklearn
import sklearn.neural_network

print('start')
# loading model
estimator = joblib.load('DNN.model')

# get the current working directory
current_directory = os.getcwd()
# build the full file path
file_path = os.path.join(current_directory, 'Input_data.xlsx')
# read the Input data.xlsx file in the current directory
Input_data = pd.read_excel(file_path)

# normalize input data
Input_data_max = np.array([[7.9, 1, 1, 1, 299.26, 5.70131,  7.60894, 1]])
Input_data_min = np.array([[4.27, 0, 0, 0, 0.07, -2.65926, 4.49223, 0]])
Input_data = np.array(Input_data)
data_min = np.min(Input_data_min,axis=0)
data_max = np.max(Input_data_max,axis=0)
Input_data_normalized = (Input_data - data_min) / (data_max-data_min)

# calculate predicted value
y_predict = estimator.predict(Input_data_normalized)
y_predict = pd.DataFrame(y_predict)
# set header
y_predict.columns = ['Horizontal VEIa (T=0.05s)','0.1s','0.15s','0.2s','0.25s','0.3s','0.35s','0.4s','0.45s',
                     '0.5s','0.55s','0.6s','0.65s','0.7s','0.75s','0.8s','0.85s','0.9s','0.95s','1s','1.05s',
                     '1.1s','1.15s','1.2s','1.25s','1.3s','1.35s','1.4s','1.45s','1.5s','1.55s','1.6s','1.65s',
                     '1.7s','1.75s','1.8s','1.85s','1.9s','1.95s','2s','2.1s','2.2s','2.3s','2.4s','2.5s','2.6s',
                     '2.7s','2.8s','2.9s','3s','3.1s','3.2s','3.3s','3.4s','3.5s','3.6s','3.7s','3.8s','3.9s','4s',
                     'Vertical VEIa (T=0.05s)','0.1s','0.15s','0.2s','0.25s','0.3s','0.35s','0.4s','0.45s',
                     '0.5s','0.55s','0.6s','0.65s','0.7s','0.75s','0.8s','0.85s','0.9s','0.95s','1s','1.05s',
                     '1.1s','1.15s','1.2s','1.25s','1.3s','1.35s','1.4s','1.45s','1.5s','1.55s','1.6s','1.65s',
                     '1.7s','1.75s','1.8s','1.85s','1.9s','1.95s','2s','2.1s','2.2s','2.3s','2.4s','2.5s','2.6s',
                     '2.7s','2.8s','2.9s','3s','3.1s','3.2s','3.3s','3.4s','3.5s','3.6s','3.7s','3.8s','3.9s','4s',
                      'Horizontal VEIr (T=0.05s)','0.1s','0.15s','0.2s','0.25s','0.3s','0.35s','0.4s','0.45s',
                     '0.5s','0.55s','0.6s','0.65s','0.7s','0.75s','0.8s','0.85s','0.9s','0.95s','1s','1.05s',
                     '1.1s','1.15s','1.2s','1.25s','1.3s','1.35s','1.4s','1.45s','1.5s','1.55s','1.6s','1.65s',
                     '1.7s','1.75s','1.8s','1.85s','1.9s','1.95s','2s','2.1s','2.2s','2.3s','2.4s','2.5s','2.6s',
                     '2.7s','2.8s','2.9s','3s','3.1s','3.2s','3.3s','3.4s','3.5s','3.6s','3.7s','3.8s','3.9s','4s',
                      'Vertical VEIr (T=0.05s)','0.1s','0.15s','0.2s','0.25s','0.3s','0.35s','0.4s','0.45s',
                     '0.5s','0.55s','0.6s','0.65s','0.7s','0.75s','0.8s','0.85s','0.9s','0.95s','1s','1.05s',
                     '1.1s','1.15s','1.2s','1.25s','1.3s','1.35s','1.4s','1.45s','1.5s','1.55s','1.6s','1.65s',
                     '1.7s','1.75s','1.8s','1.85s','1.9s','1.95s','2s','2.1s','2.2s','2.3s','2.4s','2.5s','2.6s',
                     '2.7s','2.8s','2.9s','3s','3.1s','3.2s','3.3s','3.4s','3.5s','3.6s','3.7s','3.8s','3.9s','4s']

writer = pd.ExcelWriter('Output_data.xlsx')
y_predict.iloc[:, :120].to_excel(writer, sheet_name='VEIa', index=False)
y_predict.iloc[:, 120:].to_excel(writer, sheet_name='VEIr', index=False)

# save the DataFrame as an Excel file
writer.save()
print('finish')
