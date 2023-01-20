import os
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression

#build file pathing
dirname = os.path.dirname(__file__)
MCD_factors = os.path.join(dirname, 'MCD_Factors.xlsx')
output_file = os.path.join(dirname, 'MCD_output.xlsx')
factors = pd.read_excel(file_name, sheet_name = 'Factors', engine = 'openpyxl')
stocks = pd.read_excel(file_name, sheet_name = 'Stocks', engine = 'openpyxl')

#construct master dataframe
stocks['MCD_close_return'] = stocks['MCD_close'].pct_change()
stocks['CMG_close_return'] = stocks['CMG_close'].pct_change()
stocks['CAKE_close_return'] = stocks['CAKE_close'].pct_change()
stocks['SHAK_close_return'] = stocks['SHAK_close'].pct_change()
stocks['PLAY_close_return'] = stocks['PLAY_close'].pct_change()
stock_returns = stocks[['Date','MCD_close_return','CMG_close_return','CAKE_close_return','SHAK_close_return','PLAY_close_return']]
master_df = factors.merge(right = stock_returns, on = 'Date', how = 'left')

#fit model and run betas
X = master_df[['SP500','CPI_food_bev', 'real_GDP', 'AAA_yield']]
y = master_df['MCD_close_return']
reg = LinearRegression().fit(X, y)
betas = reg.coef_

#historical means
cmg_mean = master_df['CMG_close_return'].mean()
cake_mean = master_df['CAKE_close_return'].mean()
shak_mean = master_df['SHAK_close_return'].mean()
play_mean = master_df['PLAY_close_return'].mean()

#build matrixes
dep_arrays = [betas, betas, betas, betas]
dep_matrix = np.stack(dep_arrays)
ind_arrays = [cmg_mean, cake_mean, shak_mean, play_mean]

#solve for required rate of return and write output to excel
solutions = np.linalg.lstsq(dep_matrix,ind_arrays)[0]
RequiredRate = (0.0049 + sum(betas * solutions))
output = pd.DataFrame({'Required Rate of Return': [RequiredRate]})
output.to_excel(file_name2, sheet_name = 'output', engine = 'openpyxl')
