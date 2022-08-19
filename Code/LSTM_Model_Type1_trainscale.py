import numpy as np
from keras.models import Sequential
from keras.layers import LSTM
from keras.layers import Dense
from collections import defaultdict

'''Function for Creating list for X (Time-Series) and y (Response / Predicted Value) to Match the LSTM Model Sliding Window'''
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

'''Function for Creating list for X and y (with Predicted Amount Controlled by the 'shift' variable) to Match the LSTM Model Sliding Window'''
def new_split_sequence(sequence, n_steps, shift):
    X, y = list(), list()
    for i in range(len(sequence)):
        # find the end of this pattern
        end_ix = i + n_steps
        # check if we are beyond the sequence
        if end_ix > len(sequence)-1:
            break
        # gather input and output parts of the pattern
        seq_x, seq_y = sequence[i:end_ix], sequence[end_ix:end_ix + shift]
        X.append(seq_x)
        y.append(seq_y)
    return np.array(X), np.array(y)

'''Function to generate splitting position of the dataset with given n_step size of the sliding window and the fixed amount of validation set(s)'''
def split_position(n_steps, train_start, train_end, fixed_num_val):
    len_full_train = train_end - train_start + 1
    window_num = len_full_train - n_steps
    train_ends = window_num - fixed_num_val
    train_bounds = slice(None, train_ends)
    val_bounds = slice(train_ends, window_num)
    test_bounds = window_num
    return train_bounds, val_bounds, test_bounds

def scale_df(df, minimum, maximum):
    data_range = maximum-minimum
    
    new_df = raw_data.copy()
    new_df.loc[:,['m0-4','m5-9', 'm10-14', 'm15-19', 'm20-24', 'm25-29', 'm30-34', 'm35-39','m40-44', 'm45-49', 'm50-54', 'm55-59', 'm60-64', 'm65-69', 'm70-74','m75-79', 'm80-84', 'm85+','f0-4', 'f5-9', 'f10-14', 'f15-19','f20-24', 'f25-29', 'f30-34', 'f35-39', 'f40-44', 'f45-49', 'f50-54','f55-59', 'f60-64', 'f65-69', 'f70-74', 'f75-79', 'f80-84', 'f85+']] = new_df.loc[:,['m0-4','m5-9', 'm10-14', 'm15-19', 'm20-24', 'm25-29', 'm30-34', 'm35-39','m40-44', 'm45-49', 'm50-54', 'm55-59', 'm60-64', 'm65-69', 'm70-74','m75-79', 'm80-84', 'm85+','f0-4', 'f5-9', 'f10-14', 'f15-19','f20-24', 'f25-29', 'f30-34', 'f35-39', 'f40-44', 'f45-49', 'f50-54','f55-59', 'f60-64', 'f65-69', 'f70-74', 'f75-79', 'f80-84', 'f85+']] - minimum
    new_df.loc[:,['m0-4','m5-9', 'm10-14', 'm15-19', 'm20-24', 'm25-29', 'm30-34', 'm35-39','m40-44', 'm45-49', 'm50-54', 'm55-59', 'm60-64', 'm65-69', 'm70-74','m75-79', 'm80-84', 'm85+','f0-4', 'f5-9', 'f10-14', 'f15-19','f20-24', 'f25-29', 'f30-34', 'f35-39', 'f40-44', 'f45-49', 'f50-54','f55-59', 'f60-64', 'f65-69', 'f70-74', 'f75-79', 'f80-84', 'f85+']] = new_df.loc[:,['m0-4','m5-9', 'm10-14', 'm15-19', 'm20-24', 'm25-29', 'm30-34', 'm35-39','m40-44', 'm45-49', 'm50-54', 'm55-59', 'm60-64', 'm65-69', 'm70-74','m75-79', 'm80-84', 'm85+','f0-4', 'f5-9', 'f10-14', 'f15-19','f20-24', 'f25-29', 'f30-34', 'f35-39', 'f40-44', 'f45-49', 'f50-54','f55-59', 'f60-64', 'f65-69', 'f70-74', 'f75-79', 'f80-84', 'f85+']] * (1/data_range)
    return new_df



def unscale_prediction(arr, minimum, maximum):
    return (arr * (maximum-minimum) + minimum)


'''Function for fitting the LSTM Model with each Sex data for all age-cohort in each area and return the prediction result (dataframe) for evaluation'''
def LSTM_FitPredict(sa3_codes, population_dict, n_steps, train_bounds, val_bounds, test_bounds, n_features, epochs_num, sex_label, pred_start, tuner, output, train_minimum, train_maximum):

    # Fit the Model with select sex's age-cohorts in the selected area
    for code in sa3_codes:
        male_sample_data = list(population_dict[code].values())
        X, y = split_sequence(male_sample_data, n_steps)
        train_x = X[train_bounds]
        train_y = y[train_bounds]
        val_x = X[val_bounds]
        val_y = y[val_bounds]
        test_x = X[test_bounds]
        test_y = y[test_bounds]
        train_x = train_x.reshape((train_x.shape[0], train_x.shape[1], n_features))

        # Search for the best LSTM Model of this Area's data
        tuner.search(train_x, train_y, epochs = epochs_num, validation_data=(val_x, val_y), verbose = 0)
        best_model = tuner.get_best_models(num_models=2)[0]
        
        # Fit and Record the LSTM Model based on the selected sex in the selected area
        history = best_model.fit(train_x, train_y, epochs = epochs_num, validation_data = (val_x, val_y), verbose = 0)

        # Create the first x_input window with data from 2002 (start year of the test set) 
        x_input = test_x
        x_input = x_input.reshape((1, n_steps, n_features))
        
        # Create the Prediction_list for recording each selected sex in the selected area to store each round's predicted result
        prediction_list = []

        # Rolling Update with Predicting 1 Year once a time
        for iter in range(10):

            # use the best LSTM Model to predict the next year's value with x_input
            prediction = best_model.predict(x_input, verbose=0)
            
            # If the prediction result is negative, replace it with 0
            prediction[prediction < 0] = 0
            unscaled_prediction = unscale_prediction(prediction, train_minimum, train_maximum)
            print("Area code: ", code)
            print("Year: ", 2002+iter)
            print("Scaled prediction: ", prediction)
            print("Unscaled prediction: ", unscaled_prediction)
            
            # Store the prediction result just in case for further checking        
            prediction_list.append(prediction)

            # Reconstruct the new x_input for next round's prediction
            prediction = prediction.reshape(1,1,18)   
            output.loc[(output['Code'] == code) & (output['Sex'] == sex_label), pred_start + iter] = unscaled_prediction
            x_input = np.hstack((x_input,prediction)) # Add the latest prediction
            x_input = x_input[0][1:].reshape(1,n_steps,n_features)  # Delete the first value

    # Return the prediction result for the selected sex
    return output