import pandas as pd
from scipy.stats import ttest_rel

data_file = 'MSC_DATA.xlsx'

# To analyze the effects of BS supplementation on physical performance in middle-aged and elderly adults with type 2 diabetes in the maximal incremental test. 
# Anthropometric characteristics were presented as mean and standard deviation. 
# The comparison of the supplementation effects was conducted through the Student's t-test.
# The significance level adopted for all data was p < 0.05.

# Load subjects data
subjects_attributes = pd.read_excel(
    data_file, sheet_name='SUBJECTS').loc[:, 'Age':'Water']

# Calculate mean and standard deviation for subjects data
subjects_mean, subjects_std = subjects_attributes.mean().round(
    2), subjects_attributes.std().round(2)

# Load intervention and placebo data
intervention_columns = pd.read_excel(
    data_file, sheet_name='SB_DATA').loc[:, 'Time':'HR_60']
placebo_columns = pd.read_excel(
    data_file, sheet_name='PLA_DATA').loc[:, 'Time':'HR_60']

# Perform paired t-test
_, p_values = ttest_rel(intervention_columns, placebo_columns)

# Print results
for column, p_value in zip(intervention_columns.columns, p_values):
    significance = "Significant" if p_value < 0.05 else "No significant difference"
    print(f'{column}, p_value: {p_value:.2f} ({significance})')