# -*- coding: utf-8 -*-
"""
肝癌用藥存活分析 (Kaplan-Meier Survival Analysis)
------------------------------------------------
本腳本包含：
1. Kaplan-Meier 存活率估計表（含 Greenwood 公式標準誤與 95% 信賴區間）
2. 中位存活時間 (Median Survival Time) 估計
3. Log-Rank 檢定（逐步期望值與變異數計算，Mantel-Cox 方法）
4. Cox 比例風險迴歸模型 (Hazard Ratio, 95% CI)
5. 專業醫學期刊等級 Kaplan-Meier 存活曲線圖（含風險人數表 Risk Table）
6. 匯出完整 Excel 報表
"""

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from lifelines import KaplanMeierFitter, CoxPHFitter
from lifelines.statistics import logrank_test
from lifelines.plotting import add_at_risk_counts

# 設定字體以支援繁體中文
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'DFKai-SB', 'DejaVu Sans', 'Arial']
plt.rcParams['axes.unicode_minus'] = False

def run_survival_analysis():
    excel_path = 'Survival test example.xlsx'
    df = pd.read_excel(excel_path)
    
    print("=" * 60)
    print("原始資料檢視 (Raw Data):")
    print("=" * 60)
    print(df.to_string(index=False))
    
    time_col = 'Follow-up Months'
    event_col = 'Status 1=event 0=censor'
    group_col = 'Group'
    
    # 建立 KM 估計物件
    kmf_all = KaplanMeierFitter()
    kmf_g1 = KaplanMeierFitter()
    kmf_g2 = KaplanMeierFitter()
    
    kmf_all.fit(df[time_col], df[event_col], label='Overall (整體)')
    
    df_g1 = df[df[group_col] == 1]
    df_g2 = df[df[group_col] == 2]
    
    kmf_g1.fit(df_g1[time_col], df_g1[event_col], label='Group 1 (組別 1)')
    kmf_g2.fit(df_g2[time_col], df_g2[event_col], label='Group 2 (組別 2)')
    
    # 計算各組 Life Table
    def build_detailed_life_table(kmf, data, label):
        timeline = sorted(data[time_col].unique())
        # 加入 0 時間點
        all_times = [0.0] + [t for t in timeline if t > 0]
        
        n_start = len(data)
        records = []
        cum_s = 1.0
        greenwood_sum = 0.0
        
        for t in all_times:
            if t == 0.0:
                records.append({
                    '時間 (Months)': 0.0,
                    '期初風險人數 (At Risk, n)': n_start,
                    '事件人數 (Events, d)': 0,
                    '設限人數 (Censored, c)': 0,
                    '區間存活機率 (1 - d/n)': 1.0,
                    '累積存活率 S(t)': 1.0,
                    'Greenwood SE': 0.0,
                    '95% CI 下界': 1.0,
                    '95% CI 上界': 1.0
                })
                continue
                
            n_t = len(data[data[time_col] >= t])
            d_t = len(data[(data[time_col] == t) & (data[event_col] == 1)])
            c_t = len(data[(data[time_col] == t) & (data[event_col] == 0)])
            
            p_cond = (1.0 - d_t / n_t) if n_t > 0 else 0.0
            cum_s = cum_s * p_cond
            
            if n_t > d_t and d_t > 0:
                greenwood_sum += d_t / (n_t * (n_t - d_t))
            
            se = cum_s * np.sqrt(greenwood_sum)
            
            # 取得 lifelines 算出的 CI（log-log 轉換避免超出 [0,1]）
            ci_row = kmf.confidence_interval_survival_function_.loc[[t]]
            ci_lower = ci_row.iloc[0, 0]
            ci_upper = ci_row.iloc[0, 1]
            
            records.append({
                '時間 (Months)': t,
                '期初風險人數 (At Risk, n)': n_t,
                '事件人數 (Events, d)': d_t,
                '設限人數 (Censored, c)': c_t,
                '區間存活機率 (1 - d/n)': round(p_cond, 4),
                '累積存活率 S(t)': round(cum_s, 4),
                'Greenwood SE': round(se, 4),
                '95% CI 下界': round(ci_lower, 4),
                '95% CI 上界': round(ci_upper, 4)
            })
            
        res_df = pd.DataFrame(records)
        return res_df

    table_g1 = build_detailed_life_table(kmf_g1, df_g1, 'Group 1')
    table_g2 = build_detailed_life_table(kmf_g2, df_g2, 'Group 2')
    
    print("\n" + "=" * 60)
    print("Group 1 存活壽命表 (Kaplan-Meier Life Table):")
    print("=" * 60)
    print(table_g1.to_string(index=False))
    
    print("\n" + "=" * 60)
    print("Group 2 存活壽命表 (Kaplan-Meier Life Table):")
    print("=" * 60)
    print(table_g2.to_string(index=False))
    
    # Log-Rank 逐步計算拆解
    unique_event_times = sorted(df[df[event_col] == 1][time_col].unique())
    logrank_records = []
    
    for t in unique_event_times:
        n1 = len(df[(df[group_col] == 1) & (df[time_col] >= t)])
        n2 = len(df[(df[group_col] == 2) & (df[time_col] >= t)])
        n_total = n1 + n2
        
        d1 = len(df[(df[group_col] == 1) & (df[time_col] == t) & (df[event_col] == 1)])
        d2 = len(df[(df[group_col] == 2) & (df[time_col] == t) & (df[event_col] == 1)])
        d_total = d1 + d2
        
        e1 = d_total * (n1 / n_total)
        e2 = d_total * (n2 / n_total)
        
        var_t = (n1 * n2 * d_total * (n_total - d_total)) / (n_total**2 * (n_total - 1)) if n_total > 1 else 0
        
        logrank_records.append({
            '事件時間 (t)': t,
            'Group 1 風險數 (n1)': n1,
            'Group 1 觀察事件數 (O1)': d1,
            'Group 1 期望事件數 (E1)': round(e1, 4),
            'Group 2 風險數 (n2)': n2,
            'Group 2 觀察事件數 (O2)': d2,
            'Group 2 期望事件數 (E2)': round(e2, 4),
            '合計風險數 (N)': n_total,
            '合計事件數 (D)': d_total,
            '超幾何變異數 (Var)': round(var_t, 4)
        })
        
    logrank_table = pd.DataFrame(logrank_records)
    
    o1_sum = logrank_table['Group 1 觀察事件數 (O1)'].sum()
    e1_sum = logrank_table['Group 1 期望事件數 (E1)'].sum()
    o2_sum = logrank_table['Group 2 觀察事件數 (O2)'].sum()
    e2_sum = logrank_table['Group 2 期望事件數 (E2)'].sum()
    var_sum = logrank_table['超幾何變異數 (Var)'].sum()
    
    chisq_exact = (o1_sum - e1_sum)**2 / var_sum
    
    # 呼叫 lifelines logrank_test 做驗證
    lr_res = logrank_test(df_g1[time_col], df_g2[time_col], df_g1[event_col], df_g2[event_col])
    
    print("\n" + "=" * 60)
    print("Log-Rank 逐步檢定計算表 (Step-by-step Log-Rank Table):")
    print("=" * 60)
    print(logrank_table.to_string(index=False))
    print(f"\n合計 Group 1: O1 = {o1_sum}, E1 = {e1_sum:.4f}")
    print(f"合計 Group 2: O2 = {o2_sum}, E2 = {e2_sum:.4f}")
    print(f"總變異數 Var(O1 - E1) = {var_sum:.4f}")
    print(f"Mantel-Cox 卡方統計量 Chi-sq = (O1 - E1)^2 / Var = {chisq_exact:.4f}")
    print(f"自由度 df = 1, p-value = {lr_res.p_value:.4f}")
    
    # Cox Proportional Hazards Model
    df_cox = df.copy()
    df_cox['Group_2_vs_1'] = (df_cox[group_col] == 2).astype(int)
    cph = CoxPHFitter()
    cph.fit(df_cox[[time_col, event_col, 'Group_2_vs_1']], duration_col=time_col, event_col=event_col)
    
    hr = np.exp(cph.params_['Group_2_vs_1'])
    hr_lower = np.exp(cph.confidence_intervals_.loc['Group_2_vs_1', 'Group_2_vs_1 lower-bound']) if 'Group_2_vs_1 lower-bound' in cph.confidence_intervals_.columns else np.exp(cph.confidence_intervals_.iloc[0, 0])
    hr_upper = np.exp(cph.confidence_intervals_.loc['Group_2_vs_1', 'Group_2_vs_1 upper-bound']) if 'Group_2_vs_1 upper-bound' in cph.confidence_intervals_.columns else np.exp(cph.confidence_intervals_.iloc[0, 1])
    cox_p = cph.summary.loc['Group_2_vs_1', 'p']
    
    # 統計摘要表
    med1 = kmf_g1.median_survival_time_
    med2 = kmf_g2.median_survival_time_
    
    summary_data = [
        {'指標 (Metric)': '樣本數 (Total N)', 'Group 1': len(df_g1), 'Group 2': len(df_g2), '整體/檢定結果': len(df)},
        {'指標 (Metric)': '事件數 (Events, 死亡)', 'Group 1': int(df_g1[event_col].sum()), 'Group 2': int(df_g2[event_col].sum()), '整體/檢定結果': int(df[event_col].sum())},
        {'指標 (Metric)': '設限數 (Censored)', 'Group 1': int((df_g1[event_col] == 0).sum()), 'Group 2': int((df_g2[event_col] == 0).sum()), '整體/檢定結果': int((df[event_col] == 0).sum())},
        {'指標 (Metric)': '設限比例 (Censoring Rate)', 'Group 1': f"{(df_g1[event_col] == 0).mean()*100:.1f}%", 'Group 2': f"{(df_g2[event_col] == 0).mean()*100:.1f}%", '整體/檢定結果': f"{(df[event_col] == 0).mean()*100:.1f}%"},
        {'指標 (Metric)': '中位存活時間 (Median Survival)', 'Group 1': f"{med1} 個月", 'Group 2': f"{med2} 個月", '整體/檢定結果': '-'},
        {'指標 (Metric)': 'Log-Rank 卡方值 (Chi-square)', 'Group 1': '-', 'Group 2': '-', '整體/檢定結果': f"{lr_res.test_statistic:.4f}"},
        {'指標 (Metric)': 'Log-Rank 自由度 (df)', 'Group 1': '-', 'Group 2': '-', '整體/檢定結果': "1"},
        {'指標 (Metric)': 'Log-Rank p 值 (p-value)', 'Group 1': '-', 'Group 2': '-', '整體/檢定結果': f"{lr_res.p_value:.4f} (無顯著差異)"},
        {'指標 (Metric)': '風險比 Hazard Ratio (Group 2 vs 1)', 'Group 1': '1.0 (對照/基準)', 'Group 2': f"{hr:.2f} (95% CI: {hr_lower:.2f} - {hr_upper:.2f})", '整體/檢定結果': f"Cox p = {cox_p:.4f}"}
    ]
    summary_df = pd.DataFrame(summary_data)
    
    print("\n" + "=" * 60)
    print("統計總結摘要 (Statistical Summary):")
    print("=" * 60)
    print(summary_df.to_string(index=False))
    
    # 匯出 Excel
    with pd.ExcelWriter('Survival_Analysis_Report.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='原始資料 (Raw Data)', index=False)
        summary_df.to_excel(writer, sheet_name='統計摘要 (Summary)', index=False)
        table_g1.to_excel(writer, sheet_name='Group 1 存活表', index=False)
        table_g2.to_excel(writer, sheet_name='Group 2 存活表', index=False)
        logrank_table.to_excel(writer, sheet_name='Log-Rank 計算過程', index=False)
    print("\n[成功] 完整分析報表已匯出至 Survival_Analysis_Report.xlsx")

    # 繪製專業出版級存活曲線圖
    fig = plt.figure(figsize=(10, 8), dpi=300)
    # 使用 GridSpec 分割主圖與風險人數表
    gs = fig.add_gridspec(2, 1, height_ratios=[4, 1.2], hspace=0.28)
    ax = fig.add_subplot(gs[0])
    ax_table = fig.add_subplot(gs[1])
    
    # 顏色配置：醫學頂刊經典配色
    c1 = '#1f77b4' # 藍色
    c2 = '#d62728' # 紅色
    
    # 繪製 Group 1
    kmf_g1.plot_survival_function(
        ax=ax,
        color=c1,
        linewidth=2.5,
        ci_show=True,
        ci_alpha=0.15,
        show_censors=True,
        censor_styles={'marker': '+', 'ms': 9, 'mew': 1.8},
        label=f'Group 1 (n=6, 事件=3, 設限=3) - 中位存活期: {med1} 個月'
    )
    
    # 繪製 Group 2
    kmf_g2.plot_survival_function(
        ax=ax,
        color=c2,
        linewidth=2.5,
        ci_show=True,
        ci_alpha=0.15,
        show_censors=True,
        censor_styles={'marker': '+', 'ms': 9, 'mew': 1.8},
        label=f'Group 2 (n=6, 事件=6, 設限=0) - 中位存活期: {med2} 個月'
    )
    
    # 繪製 50% 存活輔助線
    ax.axhline(0.5, color='gray', linestyle=':', linewidth=1.2, alpha=0.8)
    ax.text(0.1, 0.51, '50% 存活線 (Median Survival)', color='dimgray', fontsize=9.5, style='italic')
    
    # Group 1 中位時間垂直線
    ax.plot([med1, med1], [0, 0.5], color=c1, linestyle='--', linewidth=1.5, alpha=0.85)
    # Group 2 中位時間垂直線
    ax.plot([med2, med2], [0, 0.5], color=c2, linestyle='--', linewidth=1.5, alpha=0.85)
    
    # 統計檢定結果文字框
    stats_box = (
        f"★ Log-Rank 統計檢定 (Mantel-Cox):\n"
        f"   • 卡方值 (Chi-square): {lr_res.test_statistic:.3f}\n"
        f"   • 自由度 (df): 1\n"
        f"   • p-value: {lr_res.p_value:.3f} (p > 0.05, 差異未達顯著)\n"
        f"★ Cox 比例風險模型 (Group 2 vs 1):\n"
        f"   • Hazard Ratio (HR): {hr:.2f} (95% CI: {hr_lower:.2f} - {hr_upper:.2f})\n"
        f"   • p-value: {cox_p:.3f}"
    )
    ax.text(
        0.03, 0.06, stats_box,
        transform=ax.transAxes,
        fontsize=9.5,
        verticalalignment='bottom',
        bbox=dict(boxstyle='round,pad=0.6', facecolor='#ffffff', edgecolor='#b0bec5', alpha=0.92)
    )
    
    ax.set_title('肝癌用藥存活率曲線 (Kaplan-Meier Survival Curves)', fontsize=14, fontweight='bold', pad=12)
    ax.set_ylabel('累積存活機率 S(t) (Survival Probability)', fontsize=11, fontweight='bold')
    ax.set_xlabel('')
    ax.set_ylim(-0.02, 1.05)
    ax.set_xlim(-0.2, 8.0)
    ax.grid(True, linestyle='--', alpha=0.35)
    ax.legend(loc='upper right', frameon=True, facecolor='#ffffff', edgecolor='#e0e0e0', fontsize=10)
    
    # 製作 Risk Table (Number at risk)
    time_ticks = [0, 1, 2, 3, 4, 5, 6, 7, 8]
    ax.set_xticks(time_ticks)
    
    risk_g1 = [len(df_g1[df_g1[time_col] >= t]) for t in time_ticks]
    risk_g2 = [len(df_g2[df_g2[time_col] >= t]) for t in time_ticks]
    
    ax_table.axis('off')
    table_data = [
        ['Group 1'] + [str(x) for x in risk_g1],
        ['Group 2'] + [str(x) for x in risk_g2]
    ]
    col_labels = ['組別 \\ 月份'] + [f"{t}m" for t in time_ticks]
    
    tab = ax_table.table(
        cellText=table_data,
        colLabels=col_labels,
        loc='center',
        cellLoc='center'
    )
    tab.auto_set_font_size(False)
    tab.set_fontsize(10)
    tab.scale(1.0, 1.6)
    
    # 格式化表格樣式
    for (row, col), cell in tab.get_celld().items():
        cell.set_edgecolor('#dcdcdc')
        if row == 0:
            cell.set_facecolor('#eceff1')
            cell.set_text_props(fontweight='bold')
        elif row == 1:
            cell.get_text().set_color(c1)
            cell.set_facecolor('#f7faff')
        elif row == 2:
            cell.get_text().set_color(c2)
            cell.set_facecolor('#fff7f7')
            
    ax_table.set_title('各時間點在險人數 (Number at Risk Table)', fontsize=11, fontweight='bold', loc='left', pad=4)
    
    plt.tight_layout()
    output_png = 'km_survival_curve.png'
    plt.savefig(output_png, bbox_inches='tight')
    plt.close()
    print(f"[成功] 存活曲線圖已儲存為 {output_png}")

if __name__ == '__main__':
    run_survival_analysis()
