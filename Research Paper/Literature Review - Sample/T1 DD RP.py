import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats
import os

# Try to import statsmodels, but handle the case where it is missing
try:
    import statsmodels.api as sm
    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False

"""
PRE-RUN CHECK:
1. Ensure the Excel file path is correct for your system.
2. Install dependencies if missing: %pip install pandas numpy matplotlib scipy statsmodels openpyxl
"""

def run_analysis():
    # FIXED: Added 'r' prefix to make this a raw string, preventing Unicode escape errors
    file_path = r'C:\Users\shrey\Downloads\Master_Data_Yields_Weather_2005-2025.xlsx'
    
    # 1. Load Data with Error Handling
    if not os.path.exists(file_path):
        print(f"ERROR: File '{file_path}' not found.")
        print("Please check the file path and ensure the file exists.")
        return

    try:
        # Using openpyxl engine for .xlsx files
        df = pd.read_excel(file_path, engine='openpyxl')
    except Exception as e:
        print(f"ERROR reading Excel file: {e}")
        return

    # Clean Column Names (Remove leading/trailing spaces)
    df.columns = df.columns.str.strip()
    
    # Expected Columns (UPDATED to match your Excel file headers)
    col_yield = 'Foodgrains Yield kg per ha'
    col_rain = 'Kharif Monsoon mm JunSep'
    col_temp = 'Kharif Mean Temp C JunSep'
    col_year = 'Year'

    # Check if columns exist
    required_cols = [col_yield, col_rain, col_temp, col_year]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        print(f"ERROR: Missing columns in Excel: {missing}")
        print(f"Available columns: {list(df.columns)}")
        return

    # Drop rows with missing values to prevent errors in stats
    initial_count = len(df)
    df = df.dropna(subset=required_cols)
    if len(df) < initial_count:
        print(f"Note: Dropped {initial_count - len(df)} rows containing missing values.")

    print("\n--- DATA PREVIEW ---")
    print(df.head())

   

    # 3. Correlation Analysis
    r_rain, p_rain = stats.pearsonr(df[col_rain], df[col_yield])
    r_temp, p_temp = stats.pearsonr(df[col_temp], df[col_yield])

    correlation_results = pd.DataFrame({
        'Variables': ['Rainfall ↔ Yield', 'Temp ↔ Yield'],
        'Pearson r': [r_rain, r_temp],
        'P-value': [p_rain, p_temp],
        'Significant (p<0.05)': ['Yes' if p_rain < 0.05 else 'No', 
                                 'Yes' if p_temp < 0.05 else 'No']
    })
    print("\n=== CORRELATION ANALYSIS ===")
    print(correlation_results)
    correlation_results.to_csv('Correlation_Analysis.csv', index=False)

    # 4. Multiple Regression (Requires statsmodels)
    if not HAS_STATSMODELS:
        print("\n!!! WARNING: 'statsmodels' library is not installed.")
        print("Regression and Elasticity analysis cannot be performed.")
        print("Please run: %pip install statsmodels in the Spyder console.")
        return

    X = df[[col_rain, col_temp]].copy()
    X = sm.add_constant(X)  # Adds the intercept (const)
    y = df[col_yield]

    model = sm.OLS(y, X).fit()
    print("\n=== REGRESSION SUMMARY ===")
    print(model.summary())

    # Safely extract coefficients
    coeffs = model.params
    p_vals = model.pvalues
    stderr = model.bse

    regression_output = pd.DataFrame({
        'Variable': ['Intercept', 'Rainfall', 'Temperature'],
        'Coefficient': coeffs.values,
        'Std Error': stderr.values,
        'P-value': p_vals.values,
        'Significant': ['Yes' if p < 0.05 else 'No' for p in p_vals]
    })
    regression_output.to_csv('Regression_Coefficients.csv', index=False)

    # 5. Elasticity Calculation
    mean_rain = df[col_rain].mean()
    mean_temp = df[col_temp].mean()
    mean_yield = df[col_yield].mean()

    beta_rain = coeffs.iloc[1] 
    beta_temp = coeffs.iloc[2]

    rain_elasticity = (beta_rain * mean_rain) / mean_yield
    temp_elasticity = (beta_temp * mean_temp) / mean_yield

    elasticity_df = pd.DataFrame({
        'Variable': ['Rainfall', 'Temperature'],
        'Elasticity': [rain_elasticity, temp_elasticity],
        'Interpretation': [f'1% rain increase -> {rain_elasticity:.2f}% yield change',
                          f'1% temp increase -> {temp_elasticity:.2f}% yield change']
    })
    print("\n=== ELASTICITY ANALYSIS ===")
    print(elasticity_df)
    elasticity_df.to_csv('Elasticity_Analysis.csv', index=False)

    # 6. Time Series Visualization
    plt.rcParams['axes.grid'] = True
    fig, ax1 = plt.subplots(figsize=(12, 6))

    color_yield = '#2ca02c'
    ax1.set_xlabel('Year', fontsize=12)
    ax1.set_ylabel('Yield (kg/ha)', color=color_yield, fontsize=12, fontweight='bold')
    line1 = ax1.plot(df[col_year], df[col_yield], color=color_yield, marker='o', linewidth=2, label='Yield')
    ax1.tick_params(axis='y', labelcolor=color_yield)

    ax2 = ax1.twinx()
    color_rain = '#1f77b4'
    ax2.set_ylabel('Rainfall (mm)', color=color_rain, fontsize=12, fontweight='bold')
    line2 = ax2.plot(df[col_year], df[col_rain], color=color_rain, marker='s', linestyle='--', label='Rainfall')
    ax2.tick_params(axis='y', labelcolor=color_rain)

    plt.title('Crop Yield vs Rainfall Trends (2005-2025)', fontsize=14, fontweight='bold')
    
    lns = line1 + line2
    labs = [l.get_label() for l in lns]
    ax1.legend(lns, labs, loc='upper left')

    plt.tight_layout()
    plt.savefig('Yield_Weather_Analysis.png', dpi=300)
    plt.show()
    print("\n✓ Chart saved as 'Yield_Weather_Analysis.png'")

    # 7. Generate Final Summary Text
    summary_text = f"""
ANALYSIS SUMMARY: Impact of Weather on Yield
================================================================================
MODEL FIT:
R-squared: {model.rsquared:.4f}
Adj. R-squared: {model.rsquared_adj:.4f}
F-statistic p-value: {model.f_pvalue:.6f}

COEFFICIENTS & ELASTICITY:
Rainfall (beta): {beta_rain:.4f} | Elasticity: {rain_elasticity:.4f}
Temperature (beta): {beta_temp:.4f} | Elasticity: {temp_elasticity:.4f}

HYPOTHESIS TESTING (p < 0.05):
- Rainfall predicts Yield: {'SIGNIFICANT' if p_vals.iloc[1] < 0.05 else 'NOT SIGNIFICANT'}
- Temperature predicts Yield: {'SIGNIFICANT' if p_vals.iloc[2] < 0.05 else 'NOT SIGNIFICANT'}

KEY POLICY FINDING:
A 1% increase in rainfall is associated with a {rain_elasticity:.2f}% change in yield.
A 1 degree increase in temperature leads to a {beta_temp:.2f} kg/ha change in yield.
================================================================================
"""
    print(summary_text)
    with open('Analysis_Report.txt', 'w') as f:
        f.write(summary_text)
    print("✓ Report saved to 'Analysis_Report.txt'")

if __name__ == "__main__":
    run_analysis()