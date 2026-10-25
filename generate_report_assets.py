import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import pandas as pd

# Global styling configuration
plt.rcParams['figure.dpi'] = 300
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['axes.edgecolor'] = '#333333'
plt.rcParams['axes.linewidth'] = 0.8
plt.rcParams['grid.color'] = '#E0E0E0'
plt.rcParams['grid.linestyle'] = '--'
plt.rcParams['grid.alpha'] = 0.7
sns.set_theme(style="whitegrid", font_scale=1.05)

out_dir = 'assets/report_figures'
os.makedirs(out_dir, exist_ok=True)

# Ingest and clean data
excel_path = 'Marketing Campaign Research Database.xlsx'
df_survey = pd.read_excel(excel_path, sheet_name='Survey Responses')
df_exp = pd.read_excel(excel_path, sheet_name='Experiment Responses')

# Remove in progress rows
df_survey_clean = df_survey[df_survey['status'] == 'Completed'].copy()
df_survey_clean['age_group'] = df_survey_clean['age_group'].astype(str).str.replace('\u2013', '-').str.replace('\u2014', '-').str.replace('\ufffd', '-')
df_exp_clean = df_exp[df_exp['respondent_id'].isin(df_survey_clean['respondent_id'])].copy()

df_long = pd.merge(df_exp_clean, df_survey_clean, on='respondent_id', suffixes=('_exp', '_survey'))
resp_df = df_long.drop_duplicates(subset=['respondent_id']).copy()

# FIG 1: AGE DISTRIBUTION
fig, ax = plt.subplots(figsize=(8, 4.2))
age_order = ['Under 18', '18-22', '23-30', '31-45', 'Above 45']
age_cts = resp_df['age_group'].value_counts().reindex(age_order)
bars = ax.barh(age_order, age_cts.values, color='#3B6FB6', edgecolor='#1F4478', height=0.6)
ax.set_title('Figure 1. Respondent Age-Group Distribution (N = 74 Completed)', fontsize=12, fontweight='bold', pad=12)
ax.set_xlabel('Number of Respondents (n)', fontsize=11, labelpad=8)
ax.set_xlim(0, 30)
for bar in bars:
    w = bar.get_width()
    ax.text(w + 0.6, bar.get_y() + bar.get_height()/2, f"{int(w)} ({w/len(resp_df)*100:.1f}%)",
            va='center', ha='left', fontsize=10, fontweight='bold', color='#1F4478')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'fig1_age_distribution.png'), bbox_inches='tight')
plt.close(fig)

# FIG 2: DISCOVERY PLATFORMS
fig, ax = plt.subplots(figsize=(8.5, 4.5))
plat_cts = resp_df['discovery_platform'].value_counts()
bars = ax.barh(plat_cts.index[::-1], plat_cts.values[::-1], color='#4682B4', edgecolor='#254B70', height=0.6)
ax.set_title('Figure 2. Primary Digital Product Discovery Channels', fontsize=12, fontweight='bold', pad=12)
ax.set_xlabel('Number of Respondents (n)', fontsize=11, labelpad=8)
ax.set_xlim(0, 32)
for bar in bars:
    w = bar.get_width()
    ax.text(w + 0.5, bar.get_y() + bar.get_height()/2, f"{int(w)} ({w/len(resp_df)*100:.1f}%)",
            va='center', ha='left', fontsize=10, color='#254B70')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'fig2_discovery_platforms.png'), bbox_inches='tight')
plt.close(fig)

# FIG 3: ADVERTISING CONVERSION FUNNEL
fig, ax = plt.subplots(figsize=(8, 4.0))
tot_aud = len(resp_df)
click_aud = (resp_df['clicked_ad_3m'] == 'Yes').sum()
purch_aud = (resp_df['purchased_after_ad_3m'] == 'Yes').sum()
stages = ['1. Total Audience', '2. Ad Clickers', '3. Ad Purchasers']
counts = [tot_aud, click_aud, purch_aud]
colors = ['#2B5C8F', '#4A90E2', '#50E3C2']
bars = ax.bar(stages, counts, color=colors, edgecolor='#1E3D5C', width=0.55)
ax.set_title('Figure 3. Advertising Funnel Progression and Conversion Efficiency', fontsize=12, fontweight='bold', pad=12)
ax.set_ylabel('Number of Participants (n)', fontsize=11)
ax.set_ylim(0, 90)
for bar, c in zip(bars, counts):
    h = bar.get_height()
    pct = (c / tot_aud) * 100
    ax.text(bar.get_x() + bar.get_width()/2, h + 2, f"n={c}\n({pct:.1f}%)",
            ha='center', va='bottom', fontsize=10, fontweight='bold')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'fig3_conversion_funnel.png'), bbox_inches='tight')
plt.close(fig)

# FIG 4: REASONS FOR IGNORING ADS
fig, ax = plt.subplots(figsize=(9, 4.5))
ign_cts = resp_df['ad_ignore_reason'].value_counts()
bars = ax.barh(ign_cts.index[::-1], ign_cts.values[::-1], color='#E67E22', edgecolor='#A04000', height=0.6)
ax.set_title('Figure 4. Primary Reasons Respondents Ignore Online Advertisements', fontsize=12, fontweight='bold', pad=12)
ax.set_xlabel('Number of Respondents (n)', fontsize=11, labelpad=8)
ax.set_xlim(0, 42)
for bar in bars:
    w = bar.get_width()
    ax.text(w + 0.6, bar.get_y() + bar.get_height()/2, f"{int(w)} ({w/len(resp_df)*100:.1f}%)",
            va='center', ha='left', fontsize=10, color='#6E2C00')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'fig4_ignore_reasons.png'), bbox_inches='tight')
plt.close(fig)

# FIG 5: EXPERIMENTAL SCENARIOS
fig, axes = plt.subplots(1, 3, figsize=(14, 4.2), sharey=True)
scen_order = ['Discount', 'Cashback', 'No Promotion']
metrics = [
    ('purchase_intention', 'Purchase Intention', '#27AE60'),
    ('offer_attractiveness', 'Offer Attractiveness', '#2980B9'),
    ('value_for_money_confidence', 'Value Confidence', '#8E44AD')
]
for ax, (metric, title, colr) in zip(axes, metrics):
    means = df_long.groupby('scenario_condition')[metric].mean().reindex(scen_order)
    sems = df_long.groupby('scenario_condition')[metric].sem().reindex(scen_order) * 1.96
    bars = ax.bar(scen_order, means.values, yerr=sems.values, capsize=5, color=colr, alpha=0.85, edgecolor='#1C2833', width=0.55)
    ax.set_title(title, fontsize=12, fontweight='bold', pad=10)
    ax.set_ylim(1, 5)
    ax.set_ylabel('Mean Rating (1-5 Scale)' if ax == axes[0] else '')
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    for bar, m in zip(bars, means):
        ax.text(bar.get_x() + bar.get_width()/2, 1.2, f"{m:.2f}",
                ha='center', va='bottom', fontsize=11, fontweight='bold', color='white')
fig.suptitle('Figure 5. Experimental Scenario Comparison Across Promotional Framing Conditions (Mean ± 95% CI)',
             fontsize=13, fontweight='bold', y=1.03)
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'fig5_experimental_scenarios.png'), bbox_inches='tight')
plt.close(fig)

# FIG 6: SPEARMAN HEATMAP (FULL MATRIX WITH 1.00 DIAGONAL)
fig, ax = plt.subplots(figsize=(6.5, 5.2))
corr_vars = ['ad_trust', 'discount_importance', 'choose_promoted_product_likelihood', 'purchase_intention']
corr_lbls = ['Ad Trust', 'Discount\nImportance', 'Promoted\nChoice', 'Purchase\nIntention']
spear_mat = df_long[corr_vars].corr(method='spearman')

sns.heatmap(spear_mat, annot=True, fmt=".2f", cmap='Blues',
            vmin=-0.20, vmax=1.00, linewidths=1.2, linecolor='white',
            cbar_kws={'label': 'Spearman Correlation (ρ)'},
            ax=ax, xticklabels=corr_lbls, yticklabels=corr_lbls,
            annot_kws={'fontsize': 11, 'fontweight': 'bold'})
ax.set_title('Figure 6. Spearman Rank Correlation Matrix', fontsize=12, fontweight='bold', pad=14)
plt.xticks(rotation=0, ha='center', fontsize=10)
plt.yticks(rotation=0, fontsize=10)
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'fig6_spearman_heatmap.png'), dpi=300, bbox_inches='tight')
plt.close(fig)

# FIG 7: CROSS-TABULATION HEATMAP
fig, ax = plt.subplots(figsize=(9, 4.8))
ct_norm = pd.crosstab(resp_df['age_group'], resp_df['preferred_promotion'], normalize='index') * 100
ct_raw = pd.crosstab(resp_df['age_group'], resp_df['preferred_promotion'])
annot_txt = ct_raw.astype(str) + "\n(" + ct_norm.round(1).astype(str) + "%)"
sns.heatmap(ct_norm, annot=annot_txt, fmt="", cmap='YlGnBu', cbar_kws={'label': 'Within Age Share (%)'}, ax=ax, linewidths=1)
ax.set_title('Figure 7. Promotional Preference Distribution Across Age Cohorts', fontsize=12, fontweight='bold', pad=14)
ax.set_xlabel('Preferred Promotional Offer', fontsize=11, labelpad=8)
ax.set_ylabel('Age Cohort', fontsize=11, labelpad=8)
plt.xticks(rotation=25, ha='right')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'fig7_age_promotion_crosstab.png'), bbox_inches='tight')
plt.close(fig)

# FIG 8: STACKED BEHAVIORAL BARS
fig, ax = plt.subplots(figsize=(8.5, 4.5))
freq_ord = ['Never', 'Rarely', 'Sometimes', 'Often', 'Very often']
ct_frq = pd.crosstab(resp_df['ad_frequency'], resp_df['purchased_after_ad_3m'], normalize='index') * 100
ct_frq = ct_frq.reindex(freq_ord).fillna(0)
colors = {'No': '#3B6FB6', 'Not sure': '#E67E22', 'Yes': '#27AE60'}
bottom = np.zeros(len(ct_frq))
for col in ['No', 'Not sure', 'Yes']:
    if col in ct_frq.columns:
        vals = ct_frq[col].values
        ax.bar(freq_ord, vals, bottom=bottom, label=col, color=colors[col], width=0.55, edgecolor='#2C3E50')
        for idx, (v, b) in enumerate(zip(vals, bottom)):
            if v > 8:
                ax.text(idx, b + v/2, f"{v:.1f}%", ha='center', va='center', color='white', fontweight='bold', fontsize=9)
        bottom += vals
ax.set_title('Figure 8. Reported Purchase-After-Ad Rates Across Ad Exposure Frequency', fontsize=12, fontweight='bold', pad=12)
ax.set_ylabel('Percentage Share within Group (%)', fontsize=11)
ax.set_xlabel('Encounter Frequency with Online Advertisements', fontsize=11, labelpad=8)
ax.set_ylim(0, 105)
ax.legend(title='Purchased After Ad?', bbox_to_anchor=(1.02, 1), loc='upper left')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'fig8_ad_frequency_purchase.png'), bbox_inches='tight')
plt.close(fig)

# FIG 9: REGRESSION COEFFICIENTS
fig, ax = plt.subplots(figsize=(8, 4.2))
reg_vars = ['Discount Framing', 'Cashback Framing', 'Discount Importance', 'Ad Trust']
betas = [0.905, 0.797, 0.165, 0.110]
errors = [0.195 * 1.96, 0.195 * 1.96, 0.072 * 1.96, 0.071 * 1.96]
y_pos = np.arange(len(reg_vars))
ax.errorbar(betas, y_pos, xerr=errors, fmt='o', color='#2C3E50', ecolor='#E74C3C', elinewidth=2, capsize=5, markersize=8)
ax.axvline(0, color='gray', linestyle='--', alpha=0.7)
ax.set_yticks(y_pos)
ax.set_yticklabels(reg_vars)
ax.set_xlabel('Regression Coefficient (β ± 95% CI)', fontsize=11)
ax.set_title('Figure 9. Multivariable OLS Regression Coefficients on Purchase Intention', fontsize=12, fontweight='bold', pad=12)
for i, (b, e) in enumerate(zip(betas, errors)):
    ax.text(b, i + 0.25, f"β = {b:.3f}", ha='center', va='bottom', fontsize=9, fontweight='bold', color='#2C3E50')
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'fig9_regression_coefficients.png'), bbox_inches='tight')
plt.close(fig)

# FIG 10: RATING BOXPLOTS
fig, ax = plt.subplots(figsize=(8.5, 4.5))
box_data = [df_long['ad_trust'], df_long['discount_importance'], df_long['choose_promoted_product_likelihood']]
box_labels = ['Ad Trust', 'Discount Importance', 'Promoted Preference']
bp = ax.boxplot(box_data, patch_artist=True, tick_labels=box_labels, widths=0.45,
                showmeans=True,
                meanprops=dict(marker='D', markeredgecolor='#196F3D', markerfacecolor='#2ECC71', markersize=6),
                medianprops=dict(color='#C0392B', linewidth=2.2),
                boxprops=dict(facecolor='#AED6F1', color='#1B4F72', linewidth=1.5),
                whiskerprops=dict(color='#1B4F72', linewidth=1.2, linestyle='--'),
                capprops=dict(color='#1B4F72', linewidth=1.5))
colors = ['#AED6F1', '#FAD7A0', '#A9DFBF']
for patch, color in zip(bp['boxes'], colors):
    patch.set_facecolor(color)
ax.set_title('Figure 10. Box Plots of Survey Rating Distributions (1–5 Likert Scale)', fontsize=12, fontweight='bold', pad=12)
ax.set_ylabel('Rating (1 to 5)', fontsize=11)
ax.set_ylim(0.5, 5.5)
ax.grid(axis='y', linestyle='--', alpha=0.6)
plt.tight_layout()
fig.savefig(os.path.join(out_dir, 'fig10_rating_boxplots.png'), dpi=300, bbox_inches='tight')
plt.close(fig)

print("All 10 figures generated successfully in assets/report_figures!")
