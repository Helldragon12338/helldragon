import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import warnings
import matplotlib.font_manager as fm
import os
import sys  # 新增导入
from openpyxl import Workbook
from matplotlib.font_manager import FontProperties
warnings.filterwarnings('ignore')

# 创建中文字体对象
chinese_font = FontProperties(fname=r"C:\Windows\Fonts\simhei.ttf")  # Windows系统
# 如果上述路径不行，尝试：
#chinese_font = FontProperties(family='Microsoft YaHei')

# 实验常数
D = 0.102
F = np.pi * D**2 / 4
M_w = 18.015
M_O2 = 32.00
rho_w = 1000

# 氧平衡浓度表
temp_x_star = {
    0: 8.23e-6, 1: 8.01e-6, 2: 7.79e-6, 3: 7.58e-6, 4: 7.38e-6,
    5: 7.19e-6, 6: 7.01e-6, 7: 6.83e-6, 8: 6.66e-6, 9: 6.50e-6,
    10: 6.35e-6, 11: 6.20e-6, 12: 6.06e-6, 13: 5.92e-6, 14: 5.79e-6,
    15: 5.67e-6, 16: 5.55e-6, 17: 5.44e-6, 18: 5.33e-6, 19: 5.22e-6,
    20: 5.12e-6, 21: 5.02e-6, 22: 4.92e-6, 23: 4.83e-6, 24: 4.74e-6,
    25: 4.65e-6, 26: 4.57e-6, 27: 4.48e-6, 28: 4.40e-6, 29: 4.33e-6,
    30: 4.25e-6
}

# ========== 新增：氧饱和浓度表 ==========
C_sat_dict = {
    0: 14.62, 1: 14.22, 2: 13.83, 3: 13.46, 4: 13.11,
    5: 12.77, 6: 12.44, 7: 12.13, 8: 11.83, 9: 11.55,
    10: 11.27, 11: 11.01, 12: 10.76, 13: 10.52, 14: 10.29,
    15: 10.07, 16: 9.86, 17: 9.66, 18: 9.46, 19: 9.27,
    20: 9.09, 21: 8.92, 22: 8.74, 23: 8.58, 24: 8.42,
    25: 8.26, 26: 8.11, 27: 7.96, 28: 7.82, 29: 7.68,
    30: 7.54
}

def get_C_sat(T):
    """根据温度获取氧饱和浓度(mg/L)"""
    T_int = int(T)
    if T_int in C_sat_dict:
        if T == T_int:
            return C_sat_dict[T_int]
        else:
            T1, T2 = T_int, T_int + 1
            if T2 in C_sat_dict:
                C1, C2 = C_sat_dict[T1], C_sat_dict[T2]
                return C1 + (C2 - C1) * (T - T1)
    if T < 0:
        return C_sat_dict[0]
    elif T > 30:
        return C_sat_dict[30]
    return 8.26  # 25°C的默认值

def validate_data_input(T, C1, C2):
    """验证输入数据是否满足条件"""
    C_sat = get_C_sat(T)
    
    # 检查条件1: (C2 - C_sat) ≥ 0
    condition1 = (C2 - C_sat) >= 0
    
    # 检查条件2: C1在18-28 mg/L范围内
    condition2 = 18 <= C1 <= 28
    
    if condition1 and condition2:
        return True, ""
    else:
        error_msg = "请确保C2≥C_sat（设定温度下对应的值），C1在18-28 mg/L范围内\n"
        if not condition1:
            error_msg += f"当前：C2({C2:.2f}) - C_sat({C_sat:.2f}) = {C2-C_sat:.2f} < 0\n"
        if not condition2:
            error_msg += f"当前：C1 = {C1:.2f} mg/L，不在18-28 mg/L范围内"
        return False, error_msg

def get_x_star(T):
    """根据温度获取平衡摩尔分数"""
    T_int = int(T)
    if T_int in temp_x_star:
        if T == T_int:
            return temp_x_star[T_int]
        else:
            T1, T2 = T_int, T_int + 1
            if T2 in temp_x_star:
                x1, x2 = temp_x_star[T1], temp_x_star[T2]
                return x1 + (x2 - x1) * (T - T1)
    if T < 0:
        return temp_x_star[0]
    elif T > 30:
        return temp_x_star[30]
    return 4.65e-6

def concentration_to_mole_fraction(C):
    """将mg/L浓度转换为摩尔分数"""
    return C / (M_O2 * 1000) / (1000 / M_w)

def calculate_kxa_h(L_v, T, C1, C2, h):
    """计算Kxa和H_OL"""
    L = (L_v * rho_w) / (M_w * 1000)
    x1 = concentration_to_mole_fraction(C1)
    x2 = concentration_to_mole_fraction(C2)
    x_star = get_x_star(T)
    
    # 确保推动力为正
    if x2 <= x_star:
        x_star = x2 * 0.9
    
    if (x1 - x_star) > 0 and (x2 - x_star) > 0:
        ln_term = np.log((x1 - x_star) / (x2 - x_star))
    else:
        ratio = max(x1 / max(x2, 1e-10), 1.1)
        ln_term = np.log(ratio)
    
    Kxa = (L / (F * h)) * ln_term
    H_OL = h / ln_term if ln_term > 0 else h
    U_L = L_v / (F * 1000)
    
    return Kxa, H_OL, U_L, ln_term, x1, x2, x_star

def process_series_data(series_name, data, h):
    """处理一个系列的数据"""
    results = []
    
    for i, (L_v, V_g, T, C1, C2) in enumerate(data, 1):
        Kxa, H_OL, U_L, ln_term, x1, x2, x_star = calculate_kxa_h(L_v, T, C1, C2, h)
        
        u = (V_g / 3600) / F
        L_mol = (L_v * rho_w) / (M_w * 1000)
        
        result = {
            '组号': f'{series_name}-{i}',
            '液体流量_L_v_L_h': L_v,
            '气体流量_V_g_m3_h': V_g,
            '水温_T_C': T,
            '入口浓度_C1_mg_L': C1,
            '出口浓度_C2_mg_L': C2,
            '喷淋密度_U_L_m3_m2_h': U_L,
            '空塔气速_u_m_s': u,
            '液体摩尔流量_L_kmol_h': L_mol,
            '入口摩尔分数_x1': x1,
            '出口摩尔分数_x2': x2,
            '平衡摩尔分数_x_star': x_star,
            '对数项_ln': ln_term,
            '体积传质系数_Kxa_kmol_m3_h': Kxa,
            '传质单元高度_H_OL_m': H_OL
        }
        results.append(result)
    
    return pd.DataFrame(results)

def save_to_excel(df1, df2, filename, h):
    """保存数据到Excel文件"""
    try:
        print(f"\n正在保存数据到: {filename}")
        
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # 保存详细数据
            df1.to_excel(writer, sheet_name='系列I_详细数据', index=False)
            df2.to_excel(writer, sheet_name='系列II_详细数据', index=False)
            
            # 创建汇总表
            summary_df1 = pd.DataFrame({
                '组号': df1['组号'],
                '液体流量_L_v_L_h': df1['液体流量_L_v_L_h'],
                '喷淋密度_U_L_m3_m2_h': df1['喷淋密度_U_L_m3_m2_h'],
                '体积传质系数_Kxa_kmol_m3_h': df1['体积传质系数_Kxa_kmol_m3_h'],
                '传质单元高度_H_OL_m': df1['传质单元高度_H_OL_m']
            })
            
            summary_df2 = pd.DataFrame({
                '组号': df2['组号'],
                '气体流量_V_g_m3_h': df2['气体流量_V_g_m3_h'],
                '空塔气速_u_m_s': df2['空塔气速_u_m_s'],
                '体积传质系数_Kxa_kmol_m3_h': df2['体积传质系数_Kxa_kmol_m3_h'],
                '传质单元高度_H_OL_m': df2['传质单元高度_H_OL_m']
            })
            
            summary_df1.to_excel(writer, sheet_name='系列I_汇总', index=False)
            summary_df2.to_excel(writer, sheet_name='系列II_汇总', index=False)
            
            # 添加实验条件说明
            conditions_df = pd.DataFrame({
                '参数': ['塔内径_D_m', '塔截面积_F_m2', '填料层高度_h_m', 
                        '水的密度_rho_w_g_L', '水的摩尔质量_M_w_g_mol', '氧的摩尔质量_M_O2_g_mol'],
                '数值': [D, F, h, rho_w, M_w, M_O2],
                '单位': ['m', 'm2', 'm', 'g/L', 'g/mol', 'g/mol']
            })
            conditions_df.to_excel(writer, sheet_name='实验条件', index=False)
        
        print(f"✓ Excel文件已成功保存: {filename}")
        return True
            
    except Exception as e:
        print(f"✗ 保存Excel文件时出错: {e}")
        return False

def plot_figures(series1_df, series2_df, h):
    """绘制所有图表 - 修复中文显示、负号和标签重叠问题"""
    # ========== 第一步：优先配置全局字体（必须在创建figure之前） ==========
    # 1. 验证系统可用字体（排查字体是否存在）
    def check_font_available(font_name):
        """检查指定字体是否存在于系统中"""
        font_list = [f.name for f in fm.fontManager.ttflist]
        return font_name in font_list

    # 2. 定义优先级字体列表（优先中文字体，最后兜底西文字体）
    font_candidates = [
        'Microsoft YaHei',    # 微软雅黑（Windows）
        'SimHei',             # 黑体（Windows）
        'PingFang SC',        # 苹方（macOS）
        'Noto Sans SC',       # 思源黑体（Linux/macOS/Windows）
        'DejaVu Sans'         # 兜底西文字体（无中文）
    ]
    
    # 筛选系统实际存在的第一个字体
    available_font = 'DejaVu Sans'  # 默认
    for font in font_candidates:
        if check_font_available(font):
            available_font = font
            print(f"✓ 使用字体: {font}")
            break
    
    # 3. 核心配置（修复负号+指定可用中文字体）
    plt.rcParams['font.sans-serif'] = [available_font]  # 仅保留可用的中文字体
    plt.rcParams['axes.unicode_minus'] = False          # 关键：关闭unicode减号，正确显示负号
    plt.rcParams['font.family'] = 'sans-serif'          # 明确字体族
    
    # ========== 第二步：创建画布（配置后创建） ==========
    fig = plt.figure(figsize=(18, 8))
    
    # ========== 第三步：中文标签函数（优化字体大小/防重叠） ==========
    def set_chinese_label(ax, xlabel, ylabel, title):
        """设置中文标签，优化防重叠"""
        ax.set_xlabel(xlabel, fontsize=14, fontfamily='sans-serif')
        ax.set_ylabel(ylabel, fontsize=14, fontfamily='sans-serif')
        ax.set_title(title, fontsize=16, fontweight='bold', fontfamily='sans-serif')
        # 自动调整标签布局，防止重叠
        ax.tick_params(labelsize=12)  # 刻度字体大小
    
    # ========== 第四步：计算相关系数 ==========
    print("\n" + "="*70)
    print("相关系数计算")
    print("="*70)
    
    correlation_results = {}  # 存储相关系数结果
    
    # 系列I：Kxa与U_L的相关系数
    valid_mask1 = (series1_df['体积传质系数_Kxa_kmol_m3_h'] > 0) & (series1_df['传质单元高度_H_OL_m'] > 0)
    if valid_mask1.any() and sum(valid_mask1) >= 2:
        valid_data1 = series1_df[valid_mask1]
        U_L_values = valid_data1['喷淋密度_U_L_m3_m2_h'].values
        Kxa_values1 = valid_data1['体积传质系数_Kxa_kmol_m3_h'].values
        
        # 计算相关系数
        corr_Kxa_U_L = np.corrcoef(U_L_values, Kxa_values1)[0, 1]
        correlation_results['Kxa_U_L'] = corr_Kxa_U_L
        print(f"系列I - Kxa与喷淋密度U_L的相关系数: {corr_Kxa_U_L:.4f}")
        
        # 计算H_OL与U_L的相关系数
        H_OL_values1 = valid_data1['传质单元高度_H_OL_m'].values
        corr_H_OL_U_L = np.corrcoef(U_L_values, H_OL_values1)[0, 1]
        correlation_results['H_OL_U_L'] = corr_H_OL_U_L
        print(f"系列I - H_OL与喷淋密度U_L的相关系数: {corr_H_OL_U_L:.4f}")
    
    # 系列II：Kxa与u的相关系数
    valid_mask2 = (series2_df['体积传质系数_Kxa_kmol_m3_h'] > 0) & (series2_df['传质单元高度_H_OL_m'] > 0)
    if valid_mask2.any() and sum(valid_mask2) >= 2:
        valid_data2 = series2_df[valid_mask2]
        u_values = valid_data2['空塔气速_u_m_s'].values
        Kxa_values = valid_data2['体积传质系数_Kxa_kmol_m3_h'].values
        
        # 计算相关系数
        corr_Kxa_u = np.corrcoef(u_values, Kxa_values)[0, 1]
        correlation_results['Kxa_u'] = corr_Kxa_u
        print(f"系列II - Kxa与空塔气速u的相关系数: {corr_Kxa_u:.4f}")
        
        # 计算H_OL与u的相关系数
        H_OL_values = valid_data2['传质单元高度_H_OL_m'].values
        corr_H_OL_u = np.corrcoef(u_values, H_OL_values)[0, 1]
        correlation_results['H_OL_u'] = corr_H_OL_u
        print(f"系列II - H_OL与空塔气速u的相关系数: {corr_H_OL_u:.4f}")
    print("="*70)
    
    # ========== 第五步：智能文本位置管理器 ==========
    class TextPositionManager:
        """智能管理文本位置，防止重叠"""
        def __init__(self, ax):
            self.ax = ax
            self.positions = []
            self.min_distance = 0.15  # 最小距离阈值
            
        def add_text(self, text, x, y, transform='axes', **kwargs):
            """添加文本，自动调整位置避免重叠"""
            # 转换坐标为相对坐标
            if transform == 'axes':
                rel_x, rel_y = x, y
            else:
                # 如果是数据坐标，转换为相对坐标
                rel_x, rel_y = self.ax.transData.transform((x, y))
                rel_x = rel_x / self.ax.figure.bbox.width
                rel_y = rel_y / self.ax.figure.bbox.height
            
            # 检查是否与已有文本太近
            too_close = False
            for pos in self.positions:
                distance = np.sqrt((rel_x - pos[0])**2 + (rel_y - pos[1])**2)
                if distance < self.min_distance:
                    too_close = True
                    break
            
            if too_close:
                # 尝试几个备选位置
                candidates = [
                    (rel_x, rel_y + 0.1), (rel_x, rel_y - 0.1),
                    (rel_x + 0.1, rel_y), (rel_x - 0.1, rel_y),
                    (rel_x + 0.1, rel_y + 0.1), (rel_x - 0.1, rel_y - 0.1)
                ]
                for cand_x, cand_y in candidates:
                    cand_too_close = False
                    for pos in self.positions:
                        distance = np.sqrt((cand_x - pos[0])**2 + (cand_y - pos[1])**2)
                        if distance < self.min_distance:
                            cand_too_close = True
                            break
                    if not cand_too_close and 0 <= cand_x <= 1 and 0 <= cand_y <= 1:
                        rel_x, rel_y = cand_x, cand_y
                        break
            
            # 添加文本
            text_obj = self.ax.text(rel_x, rel_y, text, transform=self.ax.transAxes, 
                                   fontfamily='sans-serif', **kwargs)
            self.positions.append((rel_x, rel_y))
            return text_obj
    
    # ========== 第六步：图1: Kxa和H_OL与空塔气速u的关系 ==========
    ax1 = plt.subplot(1, 3, 1)
    text_manager1 = TextPositionManager(ax1)
    
    if valid_mask2.any():
        valid_data2 = series2_df[valid_mask2]
        u_values = valid_data2['空塔气速_u_m_s'].values
        Kxa_values = valid_data2['体积传质系数_Kxa_kmol_m3_h'].values
        
        # 绘制原始数据点
        ax1.loglog(u_values, Kxa_values, 'bo-', linewidth=2, 
                  markersize=10, label='Kxa', zorder=5)
        
        # 添加拟合线
        if len(u_values) >= 2:
            try:
                # 对数值进行线性拟合
                log_u = np.log10(u_values)
                log_Kxa = np.log10(Kxa_values)
                
                # 线性拟合
                coeffs = np.polyfit(log_u, log_Kxa, 1)
                a = 10**coeffs[1]  # 系数a
                b = coeffs[0]      # 指数b
                
                # 生成拟合曲线
                u_fit = np.logspace(np.log10(max(u_values.min()*0.9, 1e-3)), 
                                   np.log10(u_values.max()*1.1), 100)
                Kxa_fit = a * (u_fit**b)
                
                # 绘制拟合线
                ax1.loglog(u_fit, Kxa_fit, 'b--', linewidth=2, alpha=0.7, 
                          label='Kxa拟合', zorder=4)
                
                # 显示拟合公式（使用减号而不是负号）
                if b >= 0:
                    fit_text = f'Kxa = {a:.2f}·u^{b:.2f}'
                else:
                    fit_text = f'Kxa = {a:.2f}·u^(-{abs(b):.2f})'
                
                # 使用智能位置管理器添加文本
                text_manager1.add_text(fit_text, 0.05, 0.90,
                                      verticalalignment='top', fontsize=10,
                                      bbox=dict(boxstyle='round', facecolor='lightblue', alpha=0.8),
                                      zorder=6)
                
            except Exception as e:
                print(f"图1拟合计算时出错: {e}")
    
    set_chinese_label(ax1, '空塔气速 u (m/s)', '体积传质系数 Kxa (kmol/(m³·h))', 
                     '图1: 传质性能与空塔气速关系')
    ax1.tick_params(axis='y', labelcolor='b')
    ax1.grid(True, which="both", ls="--", alpha=0.3)
    
    # 添加相关系数
    if 'Kxa_u' in correlation_results:
        corr_text = f'相关系数 r = {correlation_results["Kxa_u"]:.4f}'
        text_manager1.add_text(corr_text, 0.05, 0.85,
                              verticalalignment='top', fontsize=10,
                              bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.7),
                              zorder=6)
    
    # 添加右侧坐标轴 (H_OL)
    ax1b = ax1.twinx()
    if valid_mask2.any():
        H_OL_values = valid_data2['传质单元高度_H_OL_m'].values
        ax1b.loglog(u_values, H_OL_values, 'rs--', linewidth=2, 
                   markersize=8, label='H_OL', zorder=5)
        
        # 添加H_OL的拟合线
        if len(u_values) >= 2:
            try:
                # 对数值进行线性拟合
                log_u = np.log10(u_values)
                log_H_OL = np.log10(H_OL_values)
                
                # 线性拟合
                coeffs = np.polyfit(log_u, log_H_OL, 1)
                a_H = 10**coeffs[1]
                b_H = coeffs[0]
                
                # 生成拟合曲线
                H_OL_fit = a_H * (u_fit**b_H)
                
                # 绘制拟合线
                ax1b.loglog(u_fit, H_OL_fit, 'r:', linewidth=2, alpha=0.7, 
                           label='H_OL拟合', zorder=4)
                
                # 显示拟合公式
                if b_H >= 0:
                    fit_text_H = f'H_OL = {a_H:.3f}·u^{b_H:.2f}'
                else:
                    fit_text_H = f'H_OL = {a_H:.3f}·u^(-{abs(b_H):.2f})'
                
                # 使用智能位置管理器添加文本
                text_manager1.add_text(fit_text_H, 0.95, 0.90,
                                      verticalalignment='top', 
                                      horizontalalignment='right', fontsize=10,
                                      bbox=dict(boxstyle='round', facecolor='mistyrose', alpha=0.8),
                                      zorder=6)
            except Exception as e:
                print(f"图1 H_OL拟合计算时出错: {e}")
    
    ax1b.set_ylabel('传质单元高度 H_OL (m)', fontsize=14, color='r')
    ax1b.tick_params(axis='y', labelcolor='r')
    
    # 添加H_OL的相关系数
    if 'H_OL_u' in correlation_results:
        corr_text_H = f'H_OL r = {correlation_results["H_OL_u"]:.4f}'
        text_manager1.add_text(corr_text_H, 0.95, 0.85,
                              verticalalignment='top', 
                              horizontalalignment='right', fontsize=9,
                              bbox=dict(boxstyle='round', facecolor='peachpuff', alpha=0.7),
                              zorder=6)
    
    # 合并图例
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax1b.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=9, ncol=2)
    
    # ========== 第七步：图2: Kxa和H_OL与喷淋密度的关系 ==========
    ax2 = plt.subplot(1, 3, 2)
    text_manager2 = TextPositionManager(ax2)
    
    if valid_mask1.any():
        valid_data1 = series1_df[valid_mask1]
        U_L_values = valid_data1['喷淋密度_U_L_m3_m2_h'].values
        Kxa_values1 = valid_data1['体积传质系数_Kxa_kmol_m3_h'].values
        
        # 绘制原始数据点
        ax2.loglog(U_L_values, Kxa_values1, 'go-', linewidth=2, 
                  markersize=10, label='Kxa', zorder=5)
        
        # 添加拟合线
        if len(U_L_values) >= 2:
            try:
                # 对数值进行线性拟合
                log_U_L = np.log10(U_L_values)
                log_Kxa1 = np.log10(Kxa_values1)
                
                # 线性拟合
                coeffs = np.polyfit(log_U_L, log_Kxa1, 1)
                a = 10**coeffs[1]
                b = coeffs[0]
                
                # 生成拟合曲线
                U_L_fit = np.logspace(np.log10(max(U_L_values.min()*0.9, 1e-3)), 
                                     np.log10(U_L_values.max()*1.1), 100)
                Kxa_fit1 = a * (U_L_fit**b)
                
                # 绘制拟合线
                ax2.loglog(U_L_fit, Kxa_fit1, 'g--', linewidth=2, alpha=0.7, 
                          label='Kxa拟合', zorder=4)
                
                # 显示拟合公式
                if b >= 0:
                    fit_text = f'Kxa = {a:.2f}·U_L^{b:.2f}'
                else:
                    fit_text = f'Kxa = {a:.2f}·U_L^(-{abs(b):.2f})'
                
                # 使用智能位置管理器添加文本
                text_manager2.add_text(fit_text, 0.05, 0.90,
                                      verticalalignment='top', fontsize=10,
                                      bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.8),
                                      zorder=6)
            except Exception as e:
                print(f"图2拟合计算时出错: {e}")
    
    set_chinese_label(ax2, '喷淋密度 U_L (m³/(m²·h))', '体积传质系数 Kxa (kmol/(m³·h))', 
                     '图2: 传质性能与喷淋密度关系')
    ax2.tick_params(axis='y', labelcolor='g')
    ax2.grid(True, which="both", ls="--", alpha=0.3)
    
    # 添加相关系数
    if 'Kxa_U_L' in correlation_results:
        corr_text = f'相关系数 r = {correlation_results["Kxa_U_L"]:.4f}'
        text_manager2.add_text(corr_text, 0.05, 0.85,
                              verticalalignment='top', fontsize=10,
                              bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.7),
                              zorder=6)
    
    # 添加右侧坐标轴 (H_OL)
    ax2b = ax2.twinx()
    if valid_mask1.any():
        H_OL_values1 = valid_data1['传质单元高度_H_OL_m'].values
        ax2b.loglog(U_L_values, H_OL_values1, 'ms--', linewidth=2, 
                   markersize=8, label='H_OL', zorder=5)
        
        # 添加H_OL的拟合线
        if len(U_L_values) >= 2:
            try:
                # 对数值进行线性拟合
                log_U_L = np.log10(U_L_values)
                log_H_OL1 = np.log10(H_OL_values1)
                
                # 线性拟合
                coeffs = np.polyfit(log_U_L, log_H_OL1, 1)
                a_H = 10**coeffs[1]
                b_H = coeffs[0]
                
                # 生成拟合曲线
                H_OL_fit1 = a_H * (U_L_fit**b_H)
                
                # 绘制拟合线
                ax2b.loglog(U_L_fit, H_OL_fit1, 'm:', linewidth=2, alpha=0.7, 
                           label='H_OL拟合', zorder=4)
                
                # 显示拟合公式
                if b_H >= 0:
                    fit_text_H = f'H_OL = {a_H:.3f}·U_L^{b_H:.2f}'
                else:
                    fit_text_H = f'H_OL = {a_H:.3f}·U_L^(-{abs(b_H):.2f})'
                
                # 使用智能位置管理器添加文本
                text_manager2.add_text(fit_text_H, 0.95, 0.90,
                                      verticalalignment='top', 
                                      horizontalalignment='right', fontsize=10,
                                      bbox=dict(boxstyle='round', facecolor='lavender', alpha=0.8),
                                      zorder=6)
            except Exception as e:
                print(f"图2 H_OL拟合计算时出错: {e}")
    
    ax2b.set_ylabel('传质单元高度 H_OL (m)', fontsize=14, color='m')
    ax2b.tick_params(axis='y', labelcolor='m')
    
    # 添加H_OL的相关系数
    if 'H_OL_U_L' in correlation_results:
        corr_text_H = f'H_OL r = {correlation_results["H_OL_U_L"]:.4f}'
        text_manager2.add_text(corr_text_H, 0.95, 0.85,
                              verticalalignment='top', 
                              horizontalalignment='right', fontsize=9,
                              bbox=dict(boxstyle='round', facecolor='peachpuff', alpha=0.7),
                              zorder=6)
    
    # 合并图例
    lines1, labels1 = ax2.get_legend_handles_labels()
    lines2, labels2 = ax2b.get_legend_handles_labels()
    ax2.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=9, ncol=2)
    
    # ========== 第八步：图3: y-x图（简化版，避免过多元素） ==========
    ax3 = plt.subplot(1, 3, 3)
    
    # 生成平衡线数据
    if len(series1_df) > 0 or len(series2_df) > 0:
        all_x = pd.concat([series1_df['入口摩尔分数_x1'], series2_df['入口摩尔分数_x1'], 
                          series1_df['出口摩尔分数_x2'], series2_df['出口摩尔分数_x2']])
        x_max = all_x.max() * 1.2 if len(all_x) > 0 else 2e-5
    else:
        x_max = 2e-5
    
    x_eq = np.linspace(0, x_max, 100)
    y_eq = np.full_like(x_eq, 0.21)
    
    # 平衡线 (水平直线 y = 0.21)
    ax3.plot(x_eq * 1e6, y_eq, 'k-', linewidth=3, label='平衡线', zorder=1)
    
    # 操作线
    if len(series1_df) > 0:
        max_x = max(series1_df['入口摩尔分数_x1'].max(), series2_df['入口摩尔分数_x1'].max())
        min_x = min(series1_df['出口摩尔分数_x2'].min(), series2_df['出口摩尔分数_x2'].min())
        
        x_op = np.array([max_x, min_x])
        y_op = np.array([0.21, 0.21])
        ax3.plot(x_op * 1e6, y_op, 'b--', linewidth=2.5, label='操作线', alpha=0.7, zorder=2)
        
        # 标记数据点（简化显示，只显示第一个点）
        if len(series1_df) > 0:
            ax3.plot(series1_df.iloc[0]['入口摩尔分数_x1'] * 1e6, 0.21, 'ro', 
                    markersize=10, label='入口点', zorder=3)
            ax3.plot(series1_df.iloc[0]['出口摩尔分数_x2'] * 1e6, 0.21, 'go', 
                    markersize=10, label='出口点', zorder=3)
            
            # 添加推动力箭头（简化显示）
            x_star = series1_df.iloc[0]['平衡摩尔分数_x_star'] * 1e6
            x1_point = series1_df.iloc[0]['入口摩尔分数_x1'] * 1e6
            x2_point = series1_df.iloc[0]['出口摩尔分数_x2'] * 1e6
            
            # 推动力箭头1
            ax3.annotate('', xy=(x1_point, 0.209), xytext=(x_star, 0.209),
                       arrowprops=dict(arrowstyle='<->', color='red', lw=2),
                       zorder=4)
            ax3.text((x1_point + x_star)/2, 0.2105, '推动力1', 
                    ha='center', va='bottom', fontsize=10, color='red', 
                    fontfamily='sans-serif',
                    bbox=dict(boxstyle='round', facecolor='white', alpha=0.7),
                    zorder=5)
            
            # 推动力箭头2
            ax3.annotate('', xy=(x2_point, 0.209), xytext=(x_star, 0.209),
                       arrowprops=dict(arrowstyle='<->', color='orange', lw=2),
                       zorder=4)
            ax3.text((x2_point + x_star)/2, 0.2075, '推动力2', 
                    ha='center', va='top', fontsize=10, color='orange',
                    fontfamily='sans-serif',
                    bbox=dict(boxstyle='round', facecolor='white', alpha=0.7),
                    zorder=5)
    
    set_chinese_label(ax3, '液相氧摩尔分数 x (×10^6)', '气相氧摩尔分数 y', '图3: 氧解吸过程 y-x 图')
    ax3.grid(True, alpha=0.3)
    
    # 图例放在不遮挡的位置
    ax3.legend(loc='upper right', fontsize=10)
    
    # ========== 第九步：全局优化和保存 ==========
    # 主标题
    plt.suptitle('氧解吸实验数据分析结果', fontsize=18, fontweight='bold', 
                fontfamily='sans-serif', y=1.02)
    
    # 优化布局
    plt.tight_layout(rect=[0, 0, 1, 0.96])  # 为主标题留出空间
    
    # 保存图表
    try:
        plt.savefig('氧解吸实验分析图表.png', dpi=300, bbox_inches='tight', facecolor='white')
        print("✓ 图表已保存为PNG文件: 氧解吸实验分析图表.png" 
        " (建议打印彩色版本)")
        print("2405 zjw")
    except Exception as e:
        print(f"✗ 保存PNG图表时出错: {e}")
        # 尝试使用英文文件名保存
        try:
            plt.savefig('oxygen_desorption_analysis.png', dpi=300, bbox_inches='tight', facecolor='white')
            print("✓ 图表已保存为英文名PNG文件: oxygen_desorption_analysis.png")
        except Exception as e2:
            print(f"✗ 英文名保存也失败: {e2}")
    
    plt.show()
    return fig
   
def print_processed_tables(df1, df2, h):
    """打印处理后的数据表 - h已作为参数传入"""
    print("=" * 120)
    print("（一）系列 I 数据处理表")
    print("=" * 120)
    
    for idx, row in df1.iterrows():
        print(f"{row['组号']:>6} | "
              f"L_v: {row['液体流量_L_v_L_h']:6.1f} L/h | "
              f"U_L: {row['喷淋密度_U_L_m3_m2_h']:6.2f} m3/(m2·h) | "
              f"Kxa: {row['体积传质系数_Kxa_kmol_m3_h']:7.2f} kmol/(m3·h) | "
              f"H_OL: {row['传质单元高度_H_OL_m']:6.3f} m")
    
    print("\n" + "=" * 120)
    print("（二）系列 II 数据处理表")
    print("=" * 120)
    
    for idx, row in df2.iterrows():
        print(f"{row['组号']:>6} | "
              f"V_g: {row['气体流量_V_g_m3_h']:6.1f} m3/h | "
              f"u: {row['空塔气速_u_m_s']:6.4f} m/s | "
              f"Kxa: {row['体积传质系数_Kxa_kmol_m3_h']:7.2f} kmol/(m3·h) | "
              f"H_OL: {row['传质单元高度_H_OL_m']:6.3f} m")
    
    print("\n" + "=" * 120)
    print("实验条件说明：")
    print(f"塔内径 D = {D*1000:.1f} mm")
    print(f"塔截面积 F = {F:.6f} m2")
    print(f"填料层高度 h = {h:.3f} m")
    print("=" * 120)

# ========== 新增：菜单系统 ==========

def clear_screen():
    """清屏"""
    os.system('cls' if sys.platform == 'win32' else 'clear')

def show_menu():
    """显示主菜单"""
    print("=" * 70)
    print("           氧解吸实验数据处理系统")
    print("=" * 70)
    print("1. 执行完整数据分析（输入新数据）")
    print("2. 使用测试数据分析")
    print("3. 查看历史结果文件")
    print("4. 重新绘制上次的图表")
    print("5. 系统设置与帮助")
    print("0. 退出程序")
    print("-" * 70)

def option1_full_analysis():
    """选项1：完整数据分析"""
    clear_screen()
    print("=" * 70)
    print("氧解吸实验数据处理系统（含数据验证）")
    print("=" * 70)
    
    # 获取填料层高度
    try:
        h = float(input("请输入填料层高度 h (m): "))
    except:
        print("输入错误，使用默认值 h = 0.8 m")
        h = 0.8
    
    print("\n" + "-" * 70)
    print("系列 I 数据输入")
    print("格式：液体流量(L/h), 气体流量(m3/h), 温度(°C), C1(mg/L), C2(mg/L)")
    print("示例：30.0, 20.0, 25.0, 25.5, 10.0")
    print("注意：C1应在18-28 mg/L范围内，C2 ≥ C_sat（温度对应饱和浓度）")
    print("-" * 70)
    
    series1_data = []
    for i in range(5):
        while True:
            try:
                input_str = input(f"第 {i+1} 组: ")
                values = [float(x.strip()) for x in input_str.split(',')]
                if len(values) == 5:
                    L_v, V_g, T, C1, C2 = values
                    
                    # ========== 数据验证 ==========
                    is_valid, error_msg = validate_data_input(T, C1, C2)
                    
                    if is_valid:
                        series1_data.append(values)
                        print(f"✓ 第 {i+1} 组数据验证通过")
                        break
                    else:
                        print(f"\n✗ 数据验证失败:")
                        print(error_msg)
                        print("请重新输入数据")
                else:
                    print("错误：需要5个数值")
            except ValueError:
                print("错误：请输入数字")
    
    print("\n" + "-" * 70)
    print("系列 II 数据输入")
    print("-" * 70)
    
    series2_data = []
    for i in range(5):
        while True:
            try:
                input_str = input(f"第 {i+1} 组: ")
                values = [float(x.strip()) for x in input_str.split(',')]
                if len(values) == 5:
                    L_v, V_g, T, C1, C2 = values
                    
                    # ========== 数据验证 ==========
                    is_valid, error_msg = validate_data_input(T, C1, C2)
                    
                    if is_valid:
                        series2_data.append(values)
                        print(f"✓ 第 {i+1} 组数据验证通过")
                        break
                    else:
                        print(f"\n✗ 数据验证失败:")
                        print(error_msg)
                        print("请重新输入数据")
                else:
                    print("错误：需要5个数值")
            except ValueError:
                print("错误：请输入数字")
    
    # 处理数据
    print("\n" + "=" * 70)
    print("正在处理数据...")
    print("=" * 70)
    
    series1_df = process_series_data('I', series1_data, h)
    series2_df = process_series_data('II', series2_data, h)
    
    # 打印结果
    print_processed_tables(series1_df, series2_df, h)
    
    # 保存到Excel
    import datetime
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f'氧解吸实验数据处理结果_{timestamp}.xlsx'
    success = save_to_excel(series1_df, series2_df, excel_filename, h)
    
    if success:
        print(f"\n✓ 数据已成功导出到Excel文件: {excel_filename}")
        print(f"文件位置: {os.path.abspath(excel_filename)}")
        
        # 保存当前数据到全局变量，以便后续使用
        global last_series1_df, last_series2_df, last_h
        last_series1_df = series1_df
        last_series2_df = series2_df
        last_h = h
    else:
        print("\n✗ Excel文件导出失败")
    
    # 绘制图表
    print("\n" + "=" * 70)
    print("正在生成图表...")
    print("=" * 70)
    
    plot_figures(series1_df, series2_df, h)
    
    input("\n数据分析完成！按回车键返回菜单...")

def option2_test_data():
    """选项2：使用测试数据分析"""
    clear_screen()
    print("使用测试数据运行程序...")
    
    # 定义h变量
    h = 0.8
    
    # 测试数据（已调整为满足验证条件）
    series1_test = [
        [15.0, 20.0, 25.0, 20.5, 9.0],   # C1=20.5 (18-28), C2=9.0 > C_sat=8.26
        [30.0, 20.0, 25.0, 22.0, 9.5],   # C1=22.0, C2=9.5 > C_sat
        [45.0, 20.0, 25.0, 24.0, 10.0],  # C1=24.0, C2=10.0 > C_sat
        [60.0, 20.0, 25.0, 26.0, 10.5],  # C1=26.0, C2=10.5 > C_sat
        [75.0, 20.0, 25.0, 28.0, 11.0]   # C1=28.0, C2=11.0 > C_sat
    ]
    
    series2_test = [
        [45.0, 10.0, 25.0, 22.0, 9.0],   # C1=22.0, C2=9.0 > C_sat
        [45.0, 15.0, 25.0, 22.0, 9.5],   # C1=22.0, C2=9.5 > C_sat
        [45.0, 20.0, 25.0, 22.0, 10.0],  # C1=22.0, C2=10.0 > C_sat
        [45.0, 25.0, 25.0, 22.0, 10.5],  # C1=22.0, C2=10.5 > C_sat
        [45.0, 30.0, 25.0, 22.0, 11.0]   # C1=22.0, C2=11.0 > C_sat
    ]
    
    # 验证测试数据
    print("\n验证测试数据...")
    for i, data in enumerate(series1_test + series2_test, 1):
        _, _, T, C1, C2 = data
        is_valid, error_msg = validate_data_input(T, C1, C2)
        if not is_valid:
            print(f"测试数据{i}验证失败: {error_msg}")
    
    series1_df = process_series_data('I', series1_test, h)
    series2_df = process_series_data('II', series2_test, h)
    
    # 打印结果
    print_processed_tables(series1_df, series2_df, h)
    
    # 保存测试数据
    import datetime
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    test_filename = f'氧解吸实验测试数据结果_{timestamp}.xlsx'
    success = save_to_excel(series1_df, series2_df, test_filename, h)
    
    if success:
        print(f"\n✓ 测试数据已成功导出到: {test_filename}")
        print(f"文件位置: {os.path.abspath(test_filename)}")
        
        # 保存到全局变量
        global last_series1_df, last_series2_df, last_h
        last_series1_df = series1_df
        last_series2_df = series2_df
        last_h = h
    
    # 绘制图表
    plot_figures(series1_df, series2_df, h)
    
    input("\n测试数据分析完成！按回车键返回菜单...")

def option3_view_history():
    """选项3：查看历史结果"""
    clear_screen()
    print("历史分析结果文件")
    print("=" * 70)
    
    try:
        import glob
        # 查找所有结果文件
        excel_files = glob.glob("氧解吸实验数据*.xlsx") + glob.glob("氧解吸实验测试*.xlsx")
        csv_files = glob.glob("*.csv")
        
        if not excel_files and not csv_files:
            print("暂无历史文件")
        else:
            if excel_files:
                print("Excel文件:")
                for i, f in enumerate(sorted(excel_files, reverse=True), 1):
                    size = os.path.getsize(f) / 1024
                    import datetime
                    mtime = datetime.datetime.fromtimestamp(os.path.getmtime(f))
                    print(f"{i}. {f} ({size:.1f}KB, {mtime.strftime('%Y-%m-%d %H:%M')})")
            
            if csv_files:
                print("\nCSV文件:")
                for i, f in enumerate(sorted(csv_files, reverse=True), 1):
                    if "结果" in f or "analysis" in f.lower():
                        size = os.path.getsize(f) / 1024
                        import datetime
                        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(f))
                        print(f"{i}. {f} ({size:.1f}KB, {mtime.strftime('%Y-%m-%d %H:%M')})")
            
            print(f"\nPNG图表文件:")
            png_files = glob.glob("*.png")
            if png_files:
                for f in png_files:
                    if os.path.exists(f):
                        print(f"  - {f}")
            else:
                print("  暂无PNG图表文件")
    
    except Exception as e:
        print(f"读取历史文件出错: {e}")
    
    input("\n按回车键返回菜单...")

def option4_replot_charts():
    """选项4：重新绘制上次的图表"""
    clear_screen()
    print("重新绘制图表")
    print("=" * 70)
    
    try:
        # 检查是否有上次的数据
        if 'last_series1_df' in globals() and 'last_series2_df' in globals():
            print("找到上次的数据，正在重新绘制图表...")
            plot_figures(last_series1_df, last_series2_df, last_h)
            print("\n图表重新绘制完成！")
        else:
            print("未找到上次的数据记录")
            print("请先执行选项1或2进行数据分析")
            
            # 尝试查找最近的数据文件
            import glob
            excel_files = glob.glob("氧解吸实验数据*.xlsx")
            if excel_files:
                latest_file = max(excel_files, key=os.path.getmtime)
                print(f"\n找到最近的数据文件: {latest_file}")
                choice = input("是否加载此文件并绘制图表？(y/n): ")
                if choice.lower() == 'y':
                    try:
                        import pandas as pd
                        # 读取Excel文件
                        excel_data = pd.read_excel(latest_file, sheet_name=None)
                        
                        if '系列I_详细数据' in excel_data and '系列II_详细数据' in excel_data:
                            series1_df = excel_data['系列I_详细数据']
                            series2_df = excel_data['系列II_详细数据']
                            
                            # 从实验条件sheet获取h值
                            if '实验条件' in excel_data:
                                conditions = excel_data['实验条件']
                                h_row = conditions[conditions['参数'] == '填料层高度_h_m']
                                if not h_row.empty:
                                    h = float(h_row.iloc[0]['数值'])
                                else:
                                    h = 0.8
                            else:
                                h = 0.8
                            
                            plot_figures(series1_df, series2_df, h)
                    except Exception as e:
                        print(f"加载文件失败: {e}")
    
    except Exception as e:
        print(f"重新绘制图表时出错: {e}")
    
    input("\n按回车键返回菜单...")

def option5_settings_help():
    """选项5：系统设置与帮助"""
    clear_screen()
    print("系统设置与帮助")
    print("=" * 70)
    
    print("\n📊 系统信息:")
    print(f"Python版本: {sys.version.split()[0]}")
    print(f"工作目录: {os.getcwd()}")
    print(f"Pandas版本: {pd.__version__}")
    print(f"Numpy版本: {np.__version__}")
    print(f"Matplotlib版本: {plt.matplotlib.__version__}")
    
    print("\n📋 使用说明:")
    print("1. 首次使用建议选择选项2测试数据分析")
    print("2. 实验数据输入格式: 液体流量,气体流量,温度,C1,C2")
    print("3. C1范围: 18-28 mg/L, C2需大于等于该温度下的饱和浓度")
    print("4. 结果会自动保存为Excel和PNG图表")
    print("5. 可使用选项3查看历史分析结果")
    
    print("\n⚠️ 注意事项:")
    print("• 确保已安装所有依赖库")
    print("• Windows系统请确保字体文件存在")
    print("• 图表保存为PNG格式，建议打印彩色版本")
    print("• 按Ctrl+C可强制退出程序")
    
    print("\n🛠️ 依赖库检查:")
    libraries = ['pandas', 'numpy', 'matplotlib', 'openpyxl']
    for lib in libraries:
        try:
            __import__(lib)
            print(f"✓ {lib}")
        except ImportError:
            print(f"✗ {lib} 未安装")
    
    input("\n按回车键返回菜单...")

def main_menu():
    """主菜单循环"""
    # 初始化全局变量
    global last_series1_df, last_series2_df, last_h
    last_series1_df = None
    last_series2_df = None
    last_h = 0.8
    
    # 检查必要的库
    try:
        import openpyxl
        print("✓ openpyxl 库已安装")
    except ImportError:
        print("✗ openpyxl 库未安装，正在安装...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
        print("✓ openpyxl 库安装完成")
        import openpyxl
    
    while True:
        clear_screen()
        show_menu()
        
        try:
            choice = input("\n请选择操作 (0-5): ").strip()
            
            if choice == '1':
                option1_full_analysis()
            elif choice == '2':
                option2_test_data()
            elif choice == '3':
                option3_view_history()
            elif choice == '4':
                option4_replot_charts()
            elif choice == '5':
                option5_settings_help()
            elif choice == '0':
                print("\n感谢使用氧解吸实验数据处理系统，再见！")
                import time
                time.sleep(1)
                break
            else:
                print("无效选择，请重新输入")
                import time
                time.sleep(1)
                
        except KeyboardInterrupt:
            print("\n\n程序被用户中断")
            break
        except Exception as e:
            print(f"\n发生错误: {e}")
            import traceback
            traceback.print_exc()
            input("按回车键继续...")

# ========== 程序入口 ==========

if __name__ == "__main__":
    # 直接进入菜单模式
    main_menu()