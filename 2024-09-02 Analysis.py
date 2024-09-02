#!/usr/bin/env python
# coding: utf-8

# In[52]:


import pandas as pd
from sklearn.utils.class_weight import compute_class_weight
import numpy as np
from twelvedata import TDClient
import matplotlib.pyplot as plt
from imblearn.over_sampling import SMOTE
import os
from datetime import datetime
from typing import Dict
from sklearn.ensemble import RandomForestClassifier
import glob
import time
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Dense, Dropout, BatchNormalization
from tensorflow.keras.callbacks import EarlyStopping
from tensorflow.keras.regularizers import l2
import seaborn as sns
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error
from sklearn.ensemble import RandomForestRegressor
from sklearn.impute import SimpleImputer
from sklearn.inspection import partial_dependence
from sklearn.linear_model import LassoCV
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_squared_error, r2_score,mean_absolute_error
from sklearn.linear_model import Lasso
from sklearn.neural_network import MLPClassifier
from sklearn.model_selection import GridSearchCV
from sklearn.model_selection import cross_val_score
import joblib
from sklearn.neural_network import MLPRegressor
from sklearn.metrics import classification_report, confusion_matrix, accuracy_score


# In[10]:



file_path = r"C:\Users\thoma\Desktop\Projects\Project Oversight\Technology Ratios\Technology 1B ratios.xlsx"
excel_file = pd.ExcelFile(file_path)

# Create a writer to save the modified Excel file
with pd.ExcelWriter('path_to_modified_file.xlsx', engine='xlsxwriter') as writer:
    # Iterate through each sheet in the Excel file
    for sheet_name in excel_file.sheet_names:
        # Load the sheet into a DataFrame
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # List of columns to calculate percentage change
        columns = [
            'Current', 'Receivables Turnover',
            'Days Sales Outstanding',
            'Fixed Asset Turnover', 'Total Asset Turnover',
            'Gross Margin', 'Operating Margin', 'Pretax Margin', 'Net Profit Margin',
            'Operating Return on Assets', 'Return on Assets', 'Return on Equity',
            'Debt to Equity', 'ROIC', 'Tax Burden', 'Interest Burden', 'EBIT Margin',
            'Financial Leverage Ratio', 'Basic EPS'
        ]
        
        # Calculate percentage change and create new columns
        for column in columns:
            # Ensure the column exists in the DataFrame
            if column in df.columns:
                # Shift the DataFrame to align the current value with the value below
                df[f'Increase {column}'] = df[column] - df[column].shift(-4)
                df[f'Increase {column}'] = df[f'Increase {column}'].divide(df[column].shift(-4)).multiply(100)
        
        # Save the modified DataFrame to the new Excel file
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("Process completed. The modified file has been saved.")


# In[4]:



# Load the Excel file
file_path = r"C:\Users\thoma\Desktop\Projects\Project Oversight\Industrials Ratios\Industrial 1B Ratios.xlsx"
excel_file = pd.ExcelFile(file_path)

# Read all sheets into a dictionary of DataFrames
dfs = {sheet_name: excel_file.parse(sheet_name) for sheet_name in excel_file.sheet_names}

# Combine all sheets into a single DataFrame
combined_df = pd.concat(dfs.values(), ignore_index=True)
print(combined_df.head())
# Remove the 'Period End' and 'close' columns if they exist
columns_to_drop = ['Period End', 'volume', 'Ticker','datetime_utc', 'datetime_market']
combined_df = combined_df.drop(columns=[col for col in columns_to_drop if col in combined_df.columns])
print(combined_df.head())
# Replace Inf and -Inf with NaN
combined_df.replace([np.inf, -np.inf], np.nan, inplace=True)
print(combined_df.head())
# Drop rows with any missing values (including previously Inf values now replaced with NaN)
combined_df = combined_df.dropna()
print(combined_df.head())
# Optional: Save the result to a new Excel file or CSV file
combined_df.to_excel('combined_sheets.xlsx', index=False)  # Save to Excel


# In[22]:


directory_path = r"C:\Users\thoma\Desktop\Projects\Project Oversight\Combined Industrials"

# Step 2: List all Excel files in the specified directory
excel_files = glob.glob(os.path.join(directory_path, "*.xlsx"))

# Step 3: Read each Excel file and combine them into one DataFrame
combined_data = pd.DataFrame()  # Empty DataFrame to store the combined data

for file in excel_files:
    # Read the first sheet of each Excel file
    df = pd.read_excel(file, sheet_name=0)  # Adjust sheet_name if necessary
    combined_data = pd.concat([combined_data, df], ignore_index=True)  # Append to the combined DataFrame

# Step 4: Write the combined DataFrame to a new Excel file with a specific name
combined_data.to_excel(r"C:\Users\thoma\Desktop\Projects\Project Oversight\Combined Industrials\combined_data.xlsx", index=False)

print("All files have been combined into 'combined_data.xlsx'.")


# In[67]:


# Step 1: Load the Data
file_path = r"C:\Users\thoma\Desktop\Projects\Project Oversight\Combined Industrials\Combined Industrials.xlsx"
df = pd.read_excel(file_path)

# Step 2: Create the target variable
df['Target'] = (df['Next Period Close'] > df['close']).astype(int)  # 1 if 'Next Period Close' is higher, 0 otherwise

# Step 3: Feature Engineering
# Use all columns except 'Target', 'Next Period Close', and 'close' as features (customize as needed)
features = df.drop(['Target', 'Next Period Close', 'close'], axis=1)
target = df['Target']

# Fill missing values (if any) - can be adjusted to fit your specific case
features = features.fillna(features.mean())

# Step 4: Normalize the features
scaler = StandardScaler()
scaled_features = scaler.fit_transform(features)
smote = SMOTE(random_state=42)
# Step 5: Split the data into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(scaled_features, target, test_size=0.2, random_state=42)
X_train_resampled, y_train_resampled = smote.fit_resample(X_train, y_train)

# Step 6: Create the Neural Network Model
model = Sequential()
model.add(Dense(50, input_dim=X_train.shape[1], activation='tanh', kernel_regularizer=l2(0.01)))  # Input layer with L2 regularization
model.add(BatchNormalization())
model.add(Dense(200, activation='tanh', kernel_regularizer=l2(0.01)))  # Hidden layer with L2 regularization
model.add(Dropout(0.3))
model.add(Dense(1, activation='sigmoid'))  # Output layer

# Step 7: Compile the Model
model.compile(optimizer='RMSprop', loss='binary_crossentropy', metrics=['accuracy'])

# Step 8: Train the Model with Early Stopping
early_stopping = EarlyStopping(monitor='val_loss', patience=10, restore_best_weights=True)
history = model.fit(X_train_resampled, y_train_resampled, epochs=1000, batch_size=16, validation_data=(X_test, y_test), 
                    verbose=1, callbacks=[early_stopping])
# Step 9: Make Predictions on the Testing Set
y_pred = (model.predict(X_test) > 0.5).astype(int)

# Step 10: Evaluate the Model
print("Accuracy Score:", accuracy_score(y_test, y_pred))
print("Confusion Matrix:\n", confusion_matrix(y_test, y_pred))
print("Classification Report:\n", classification_report(y_test, y_pred))

# Optional: Plot Training Loss and Accuracy

# Plot loss
plt.plot(history.history['loss'], label='train_loss')
plt.plot(history.history['val_loss'], label='val_loss')
plt.title('Model Loss')
plt.ylabel('Loss')
plt.xlabel('Epoch')
plt.legend()
plt.show()

# Plot accuracy
plt.plot(history.history['accuracy'], label='train_accuracy')
plt.plot(history.history['val_accuracy'], label='val_accuracy')
plt.title('Model Accuracy')
plt.ylabel('Accuracy')
plt.xlabel('Epoch')
plt.legend()
plt.show()

print("Training set class distribution:\n", y_train.value_counts())
print("Testing set class distribution:\n", y_test.value_counts())


# In[90]:


file_path = r"C:\Users\thoma\Desktop\Projects\Project Oversight\Combined Industrials\Combined Industrials.xlsx"
df = pd.read_excel(file_path)

# Step 2: Create the target variable for regression
target = df['Next Period Close']  # Predicting the actual value of 'Next Period Close'

# Step 3: Feature Engineering
# Use all columns except 'Target', 'Next Period Close', and 'close' as features (customize as needed)
features = df.drop(['Next Period Close'], axis=1)

# Fill missing values (if any) - can be adjusted to fit your specific case
features = features.fillna(features.mean())

# Step 4: Normalize the features
scaler = StandardScaler()
scaled_features = scaler.fit_transform(features)

# Step 5: Split the data into training and testing sets
X_train, X_test, y_train, y_test, train_indices, test_indices = train_test_split(
    scaled_features, target, df.index, test_size=0.2, random_state=42)

# Step 6: Create the Neural Network Model for regression
model = Sequential()
model.add(Dense(50, input_dim=X_train.shape[1], activation='tanh', kernel_regularizer=l2(0.01)))  # Input layer with L2 regularization
model.add(BatchNormalization())
model.add(Dropout(0.2))  # Dropout layer for regularization
model.add(Dense(200, activation='leaky_relu', kernel_regularizer=l2(0.01)))  # Hidden layer with L2 regularization
model.add(Dropout(0.2))
model.add(Dense(200, activation='leaky_relu', kernel_regularizer=l2(0.01)))  # Hidden layer with L2 regularization
model.add(Dropout(0.2))
model.add(Dense(1, activation='linear'))  # Output layer for regression (linear activation by default)

# Step 7: Compile the Model for regression
model.compile(optimizer='RMSprop', loss='mean_squared_error', metrics=['mean_absolute_error'])

# Step 8: Train the Model with Early Stopping
early_stopping = EarlyStopping(monitor='val_loss', patience=10, restore_best_weights=True)
history = model.fit(X_train, y_train, epochs=1000, batch_size=16, validation_data=(X_test, y_test), verbose=1, callbacks=[early_stopping])

# Step 9: Make Predictions on the Testing Set
y_pred = model.predict(X_test)

# Step 10: Evaluate the Model for regression
print("Mean Squared Error (MSE):", mean_squared_error(y_test, y_pred))
print("Mean Absolute Error (MAE):", mean_absolute_error(y_test, y_pred))
r2 = r2_score(y_test, y_pred)

# Calculate adjusted R²
n = X_test.shape[0]  # Number of observations
p = X_test.shape[1]  # Number of predictors

adjusted_r2 = 1 - ((1 - r2) * (n - 1)) / (n - p - 1)

# Print adjusted R²
print(f"Adjusted R² Score: {adjusted_r2}")

y_pred_df = pd.DataFrame(y_pred, columns=['Predicted Next Period Close'])

# Add predictions to the original DataFrame for the testing set
df_test = df.loc[test_indices].copy()
df_test['Predicted Next Period Close'] = y_pred_df

# Save the modified DataFrame to a new Excel file
output_file_path = r"C:\Users\thoma\Desktop\Projects\Project Oversight\Combined Industrials\Predictions.xlsx"
#df_test.to_excel(output_file_path, index=False)

# Plot loss
plt.plot(history.history['loss'], label='train_loss')
plt.plot(history.history['val_loss'], label='val_loss')
plt.title('Model Loss')
plt.ylabel('Loss')
plt.xlabel('Epoch')
plt.legend()
plt.show()

# Plot Mean Absolute Error
plt.plot(history.history['mean_absolute_error'], label='train_mae')
plt.plot(history.history['val_mean_absolute_error'], label='val_mae')
plt.title('Model Mean Absolute Error')
plt.ylabel('MAE')
plt.xlabel('Epoch')
plt.legend()
plt.show()


# In[89]:


df_test.to_excel(output_file_path, index=False)


# In[ ]:




