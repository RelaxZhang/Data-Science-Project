import numpy as np
from sklearn.model_selection import train_test_split

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
def split_position(n_steps, train_start, train_end):
    len_full_train = train_end - train_start + 1
    window_num = len_full_train - n_steps
    train_val_bounds = slice(None, window_num)
    test_bounds = window_num
    return train_val_bounds, test_bounds

'''Function for fitting the LSTM Model with each Sex data for all age-cohort in each area and return the prediction result (dataframe) for evaluation'''
def LSTM_FitPredict(sa3_codes, population_dict, n_steps, train_val_bounds, test_bounds, n_features, epochs_num, sex_label, pred_start, tuner, output, search_callbacks, fit_callbacks):

    # Fit the Model with select sex's age-cohorts in the selected area
    for code in sa3_codes:
        male_sample_data = list(population_dict[code].values())
        X, y = split_sequence(male_sample_data, n_steps)
        train_val_x = X[train_val_bounds]
        train_val_y = y[train_val_bounds]
        train_x, val_x, train_y, val_y = train_test_split(train_val_x, train_val_y, test_size=0.2, random_state=42)
        test_x = X[test_bounds]
        train_x = train_x.reshape((train_x.shape[0], train_x.shape[1], n_features))

        # Search for the best LSTM Model of this Area's data
        tuner.search(train_x, train_y, epochs = epochs_num, validation_data=(val_x, val_y), verbose = 0,
                    callbacks = search_callbacks)
        best_model = tuner.get_best_models(num_models = 1)[0]
        
        # Fit and Record the LSTM Model based on the selected sex in the selected area
        history = best_model.fit(train_x, train_y, epochs = epochs_num, validation_data = (val_x, val_y), verbose = 0,
                                callbacks = fit_callbacks)

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

            # Store the prediction result just in case for further checking        
            prediction_list.append(prediction)

            # Reconstruct the new x_input for next round's prediction
            prediction = prediction.reshape(1,1,18)   
            output.loc[(output['Code'] == code) & (output['Sex'] == sex_label), pred_start + iter] = prediction
            x_input = np.hstack((x_input,prediction)) # Add the latest prediction
            x_input = x_input[0][1:].reshape(1,n_steps,n_features)  # Delete the first value

    # Return the prediction result for the selected sex
    return output