# -*- coding: utf-8 -*-
"""Untitled2.ipynb

Automatically generated by Colaboratory.

Original file is located at
    https://colab.research.google.com/drive/1l7gFctFLTKh_XQ4mc4mwsZZXEnT8UNKI
"""

import pandas as pd
import numpy as np
from collections import defaultdict
from numpy import array
from keras.models import Sequential
from keras.layers import LSTM
from keras.layers import Dense

def split_sequence(sequence, n_steps):
	X, y = list(), list()
	for i in range(len(sequence)):
		# find the end of this pattern
		end_ix = i + n_steps
		# check if we are beyond the sequence
		if end_ix > len(sequence)-1:
			break
		# gather input and output parts of the pattern
		seq_x, seq_y = sequence[i:end_ix], sequence[end_ix]
		X.append(seq_x)
		y.append(seq_y)
	return np.array(X), np.array(y)

raw_data = pd.read_csv('/content/drive/MyDrive/Colab/true_1000_fulldata.csv')
train_data = raw_data[(raw_data['Year']>=1991) & (raw_data['Year']<=2002)]
test_data = raw_data[(raw_data['Year']>=2003) & (raw_data['Year']<=2011)]
sa3_num = len(raw_data['SA3 Code'].unique())
sa3_codes = raw_data['SA3 Code'].unique()
sa3_names = raw_data['SA3 Name'].unique()
age_groups = ['0-4','5-9', '10-14', '15-19', '20-24', '25-29', '30-34', '35-39','40-44', '45-49', '50-54', '55-59', '60-64', '65-69', '70-74','75-79', '80-84', '85+']
population_m_dict = defaultdict(dict)
population_f_dict = defaultdict(dict)
for sa3_code in sa3_codes:
    population_m_dict[sa3_code] = dict()
    population_f_dict[sa3_code] = dict()
    for year in range(1991,2012):
        if(raw_data[(raw_data['Year']==year) & (raw_data['SA3 Code']==sa3_code)]['Total'].size>0):
            population_m_dict[sa3_code][year] = raw_data[(raw_data['Year']==year) & (raw_data['SA3 Code']==sa3_code)][['m0-4', 'm5-9', 'm10-14', 'm15-19', 'm20-24', 'm25-29', 'm30-34', 'm35-39', 'm40-44', 'm45-49', 'm50-54', 'm55-59', 'm60-64', 'm65-69', 'm70-74', 'm75-79', 'm80-84', 'm85+']].values.tolist()[0]
            population_f_dict[sa3_code][year] = raw_data[(raw_data['Year']==year) & (raw_data['SA3 Code']==sa3_code)][['f0-4', 'f5-9', 'f10-14', 'f15-19', 'f20-24', 'f25-29', 'f30-34', 'f35-39', 'f40-44', 'f45-49', 'f50-54', 'f55-59', 'f60-64', 'f65-69', 'f70-74', 'f75-79', 'f80-84', 'f85+']].values.tolist()[0]

n_steps = 5
n_features = 18

output = pd.DataFrame(index = range(sa3_num*36), columns = ['Code','Area name','Sex','Age group',2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011])
for c in range(0,sa3_num):
  for s in range(18):
    output.loc[c*36+s] = {'Code':sa3_codes[c],'Area name':sa3_names[c],'Sex':'Females','Age group':age_groups[s]}
    output.loc[c*36+18+s] = {'Code':sa3_codes[c],'Area name':sa3_names[c],'Sex':'Males','Age group':age_groups[s]}

output

for code in sa3_codes:
  male_data = list(population_m_dict[code].values())
  X, y = split_sequence(male_data, n_steps)
  train_x = X[:6]
  train_y = y[:6]
  test_x = X[6:]
  test_y = y[6:]
  train_x = train_x.reshape((train_x.shape[0], train_x.shape[1], n_features))

  model = Sequential()
  model.add(LSTM(100, activation='relu', input_shape=(n_steps, n_features)))
  model.add(Dense(18))
  model.compile(optimizer='adam', loss='mse')
  model.fit(train_x, train_y, epochs=100, verbose=0)
  x_input = test_x
  for i in range(len(x_input)):
      temp_x = x_input[i].reshape((1, n_steps, n_features))
      yhat = model.predict(temp_x, verbose=0)
      output.loc[(output['Code']==code)&(output['Sex']=='Males'),2002+i]=yhat

for code in sa3_codes:
  female_data = list(population_f_dict[code].values())
  X, y = split_sequence(female_data, n_steps)
  train_x = X[:6]
  train_y = y[:6]
  test_x = X[6:]
  test_y = y[6:]
  train_x = train_x.reshape((train_x.shape[0], train_x.shape[1], n_features))

  model = Sequential()
  model.add(LSTM(100, activation='relu', input_shape=(n_steps, n_features)))
  model.add(Dense(18))
  model.compile(optimizer='adam', loss='mse')
  model.fit(train_x, train_y, epochs=100, verbose=0)
  x_input = test_x
  for i in range(len(x_input)):
      temp_x = x_input[i].reshape((1, n_steps, n_features))
      yhat = model.predict(temp_x, verbose=0)
      output.loc[(output['Code']==code)&(output['Sex']=='Females'),2002+i]=yhat

output.to_csv('/content/drive/MyDrive/Colab/1000_fulldata_output.csv',index=False,header=True)