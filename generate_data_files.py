import pandas as pd
import numpy as np
from sklearn.impute import KNNImputer
import os

def create_raw_and_cleaned_datasets():
    excel_path = 'Marketing Campaign Research Database.xlsx'
    df_survey = pd.read_excel(excel_path, sheet_name='Survey Responses')
    df_exp = pd.read_excel(excel_path, sheet_name='Experiment Responses')
    
    # --------------------------------------------------------------------------
    # 1. Respondent-Level Wide Experimental Metrics (Pivoting 3 scenarios)
    # --------------------------------------------------------------------------
    exp_piv = df_exp.pivot(
        index='respondent_id',
        columns='scenario_condition',
        values=['purchase_intention', 'offer_attractiveness', 'value_for_money_confidence']
    )
    exp_piv.columns = [f"{col}_{cond.lower().replace(' ', '_')}" for col, cond in exp_piv.columns]
    exp_piv = exp_piv.reset_index()
    
    # Left join to preserve all 80 respondents (74 Completed + 6 In-progress)
    df_raw = pd.merge(df_survey, exp_piv, on='respondent_id', how='left')
    
    # --------------------------------------------------------------------------
    # 2. Introduce controlled 10% MCAR missingness into numeric Likert features
    # --------------------------------------------------------------------------
    np.random.seed(42)
    numeric_cols = [
        'ad_trust', 'discount_importance', 'choose_promoted_product_likelihood',
        'purchase_intention_discount', 'offer_attractiveness_discount', 'value_for_money_confidence_discount',
        'purchase_intention_cashback', 'offer_attractiveness_cashback', 'value_for_money_confidence_cashback',
        'purchase_intention_no_promotion', 'offer_attractiveness_no_promotion', 'value_for_money_confidence_no_promotion'
    ]
    
    df_raw_with_missing = df_raw.copy()
    for col in numeric_cols:
        if col in df_raw_with_missing.columns:
            # Inject 10% MCAR missingness only into observed values
            observed_idx = df_raw_with_missing[df_raw_with_missing[col].notnull()].index
            mask = np.random.rand(len(observed_idx)) < 0.10
            df_raw_with_missing.loc[observed_idx[mask], col] = np.nan

    # Save raw_data.csv (80 rows: 74 completed + 6 in-progress)
    raw_path_root = 'raw_data.csv'
    raw_path_assets = os.path.join('assets', 'raw_data.csv')
    df_raw_with_missing.to_csv(raw_path_root, index=False)
    df_raw_with_missing.to_csv(raw_path_assets, index=False)
    
    in_prog_raw = (df_raw_with_missing['status'] == 'In progress').sum()
    comp_raw = (df_raw_with_missing['status'] == 'Completed').sum()
    print("="*80)
    print(f"[OK] Saved {raw_path_root}: {df_raw_with_missing.shape[0]} Rows x {df_raw_with_missing.shape[1]} Columns")
    print(f"     Status Breakdown -> Completed: {comp_raw}, In progress: {in_prog_raw} (Total: {len(df_raw_with_missing)})")
    print(f"     Missing Likert Values: {df_raw_with_missing[numeric_cols].isna().sum().sum()}")
    print("="*80)
    
    # --------------------------------------------------------------------------
    # 3. Cleaned Dataset (Strictly 74 completed respondents, 0 in-progress)
    # --------------------------------------------------------------------------
    # Step A: Filter out all 6 'In progress' incomplete rows
    df_cleaned = df_raw_with_missing[df_raw_with_missing['status'] == 'Completed'].copy()
    
    # Step B: 5D Distance-Weighted KNN Imputation on numeric Likert constructs
    knn = KNNImputer(n_neighbors=5, weights='distance')
    imputed_vals = knn.fit_transform(df_cleaned[numeric_cols])
    df_cleaned[numeric_cols] = np.clip(np.round(imputed_vals), 1, 5).astype(int)
    
    # Step C: Mode Imputation for Nominal Categorical Attributes
    nominal_cols = ['discovery_platform', 'preferred_promotion', 'ad_ignore_reason']
    for c in nominal_cols:
        if c in df_cleaned.columns and df_cleaned[c].isna().sum() > 0:
            df_cleaned[c] = df_cleaned[c].fillna(df_cleaned[c].mode()[0])
            
    # Step D: Survey Skip Logic enforcement
    df_cleaned.loc[df_cleaned['purchased_after_ad_3m'] == 'No', 'purchase_ad_platform'] = 'Not applicable'
    df_cleaned['purchase_ad_platform'] = df_cleaned['purchase_ad_platform'].fillna('Not applicable')
    
    # Step E: Feature Engineering - Composite Ad Receptivity Score
    df_cleaned['ad_receptivity_score'] = (
        df_cleaned['ad_trust'] + df_cleaned['discount_importance'] + df_cleaned['choose_promoted_product_likelihood']
    ) / 3.0
    
    # Save cleaned_data.csv (strictly 74 rows)
    clean_path_root = 'cleaned_data.csv'
    clean_path_assets = os.path.join('assets', 'cleaned_data.csv')
    clean_path_step2 = os.path.join('assets', 'cleaned_data_step_2.csv')
    
    df_cleaned.to_csv(clean_path_root, index=False)
    df_cleaned.to_csv(clean_path_assets, index=False)
    df_cleaned.to_csv(clean_path_step2, index=False)
    
    in_prog_clean = (df_cleaned['status'] == 'In progress').sum()
    comp_clean = (df_cleaned['status'] == 'Completed').sum()
    print(f"[OK] Saved {clean_path_root}: {df_cleaned.shape[0]} Rows x {df_cleaned.shape[1]} Columns")
    print(f"     Status Breakdown -> Completed: {comp_clean}, In progress: {in_prog_clean} (Strictly 74 Rows)")
    print(f"     Total Missing Values Remaining: {df_cleaned.isna().sum().sum()}")
    print("="*80)
    
    # --------------------------------------------------------------------------
    # 4. Also Export Long-Form Datasets (For Repeated-Measures Trial Analyses)
    # --------------------------------------------------------------------------
    # Cleaned Long
    df_exp_clean = df_exp[df_exp['respondent_id'].isin(df_cleaned['respondent_id'])].copy()
    survey_meta = df_cleaned[['respondent_id', 'age_group', 'discovery_platform', 'ad_frequency', 
                              'ad_trust', 'discount_importance', 'choose_promoted_product_likelihood', 
                              'ad_receptivity_score', 'status']]
    df_long_cleaned = pd.merge(df_exp_clean, survey_meta, on='respondent_id', how='inner')
    df_long_cleaned.to_csv('cleaned_data_long.csv', index=False)
    df_long_cleaned.to_csv(os.path.join('assets', 'cleaned_data_long.csv'), index=False)
    
    # Raw Long (Outer join to include 6 in-progress rows)
    df_long_raw = pd.merge(df_exp, df_survey, on='respondent_id', how='outer')
    df_long_raw.to_csv('raw_data_long.csv', index=False)
    df_long_raw.to_csv(os.path.join('assets', 'raw_data_long.csv'), index=False)
    
    print(f"[OK] Saved cleaned_data_long.csv: {df_long_cleaned.shape[0]} Rows x {df_long_cleaned.shape[1]} Columns")
    print(f"[OK] Saved raw_data_long.csv: {df_long_raw.shape[0]} Rows x {df_long_raw.shape[1]} Columns")
    print("="*80)

if __name__ == '__main__':
    create_raw_and_cleaned_datasets()
