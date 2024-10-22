import os
import pandas as pd
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import streamlit as st
import plotly.graph_objects as go

@st.cache_data
def load_all_excel_files(folder_path, sheet_name):
    dataframes = []
    for file in os.listdir(folder_path):
        if file.endswith('.xlsm'):
            file_path = os.path.join(folder_path, file)
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            dataframes.append(df)
    return pd.concat(dataframes, ignore_index=True)

@st.cache_data
def forecast_profit(data, seasonal_period=50, forecast_horizon=50):
    daily_profit = data[['TANGGAL', 'LABA']].copy()
    daily_profit['TANGGAL'] = pd.to_datetime(daily_profit['TANGGAL'])
    daily_profit = daily_profit.groupby('TANGGAL').sum()
    daily_profit = daily_profit[~daily_profit.index.duplicated(keep='first')]

    # Resample the data weekly and interpolate missing values
    daily_profit = daily_profit.resample('W').mean().interpolate()

    train_size = int(len(daily_profit) * 0.9)
    train, test = daily_profit[:train_size], daily_profit[train_size:]

    # Fit the model on the training data
    hw_model = ExponentialSmoothing(train, trend='add', seasonal='add', seasonal_periods=seasonal_period).fit()

    # Forecast the future values (including the test period)
    hw_forecast_future = hw_model.forecast(forecast_horizon)
    test_forecast = hw_model.forecast(len(test))  # Predict the test set
    
    # Store fitted values (train predictions)
    fitted_values = hw_model.fittedvalues

    return daily_profit, fitted_values, test, test_forecast, hw_forecast_future
