import pandas as pd
import numpy as np
from scipy.optimize import fsolve
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns

plt.rcParams['font.sans-serif'] = ['SimHei']  # Windows
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示为方块的问题

# 加载上传的 Excel 文件
file_path = "20永煤MTN01和03国债03的债券价格.xlsx"
excel_data = pd.ExcelFile(file_path)

# 从第一个工作表加载数据
data = excel_data.parse('Sheet1')

# 第一步：将 '日期' 列从 object 格式转换为字符串格式
data['日期'] = data['日期'].apply(str)

# 第二步：使用 strptime 将字符串日期转换为 datetime 格式
data['日期'] = data['日期'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d'))


# 定义一个函数，使用精确公式求解到期收益率 (YTM)
def ytm_function(r, price, coupon, face_value, years):
    # 计算债券价格为所有现金流折现值的总和
    cash_flows = sum([coupon / (1 + r) ** t for t in range(1, years + 1)])
    final_payment = face_value / (1 + r) ** years
    return cash_flows + final_payment - price


# 设置两只债券的相关信息
face_value = 100  # 债券面值
years_to_maturity_mtn001 = 3  # 假设 "20永煤MTN001" 的到期期限为 3 年
years_to_maturity_gz03 = 3  # 假设 "03国债03" 的到期期限为 3 年

# 初始化结果列表
ytm_results_mtn001 = []
ytm_results_gz03 = []

# 使用精确公式计算每只债券价格的 YTM
for index, row in data.iterrows():
    bond_price_mtn001 = row['20永煤MTN001']  # "20永煤MTN001" 的债券价格
    bond_price_gz03 = row['03国债03']  # "03国债03" 的债券价格

    coupon_payment_mtn001 = face_value * 0.0545  # "20永煤MTN001" 的票息率为 5.45%
    coupon_payment_gz03 = face_value * 0.034  # "03国债03" 的票息率为 3.4%

    # 使用 fsolve 求解给定价格的 "20永煤MTN001" 的 YTM
    ytm_solution_mtn001 = fsolve(ytm_function, 0.05, args=(bond_price_mtn001, coupon_payment_mtn001, face_value, years_to_maturity_mtn001))
    ytm_results_mtn001.append(ytm_solution_mtn001[0])

    # 使用 fsolve 求解给定价格的 "03国债03" 的 YTM
    ytm_solution_gz03 = fsolve(ytm_function, 0.05, args=(bond_price_gz03, coupon_payment_gz03, face_value, years_to_maturity_gz03))
    ytm_results_gz03.append(ytm_solution_gz03[0])

# 确保 ytm_results 的长度与数据框的行数匹配
if len(ytm_results_mtn001) == len(data) and len(ytm_results_gz03) == len(data):
    data['YTM_MTN001'] = ytm_results_mtn001
    data['YTM_GZ03'] = ytm_results_gz03
else:
    print(f"Error: ytm_results 的长度 ({len(ytm_results_mtn001)}) 与数据的长度 ({len(data)}) 不匹配")

# 显示包含 YTM 结果的更新后的数据框
print(data[['日期', '20永煤MTN001', 'YTM_MTN001', '03国债03', 'YTM_GZ03']].head())  # 显示前 5 行作为示例

# 可视化结果
sns.set_theme(style="whitegrid")
plt.figure(figsize=(14, 8))
sns.lineplot(x=data['日期'], y=data['YTM_MTN001'], label='20永煤MTN001 的精确 YTM', color='blue')
sns.lineplot(x=data['日期'], y=data['YTM_GZ03'], label='03国债03 的精确 YTM', color='green')
plt.title("不同时间的精确 YTM")
plt.xlabel("日期")
plt.ylabel("精确 YTM")
plt.xticks(rotation=45)
plt.legend()
plt.show()

# 计算 '20永煤MTN001' 和 '03国债03' 之间的债券利差
data['Bond_Yield_Spread'] = data['YTM_MTN001'] - data['YTM_GZ03']

# 存储相关列到数据框
bond_yield_data = data[['日期', 'YTM_MTN001', 'YTM_GZ03', 'Bond_Yield_Spread']]

# 显示前几行数据检查
print(bond_yield_data.head())

# 可视化债券利差
sns.set_theme(style="whitegrid")
plt.figure(figsize=(14, 8))
sns.lineplot(x=data['日期'], y=data['Bond_Yield_Spread'], label='债券利差 (20永煤MTN001 - 03国债03)', color='purple')
plt.title("20永煤MTN001 和 03国债03 之间的债券利差")
plt.xlabel("日期")
plt.ylabel("债券利差")
plt.xticks(rotation=45)
plt.legend()
plt.show()

# 常量
recovery_rate = 0.5  # 违约回收率 R = 50%
years_to_maturity = 3  # 假设债券的到期期限为 3 年

# 使用给定公式 (式 7-12) 计算违约概率 lambda
def calculate_default_probability(ytm_risky, ytm_risk_free, recovery_rate, years_to_maturity):
    # 使用违约概率的公式
    try:
        exp_term = np.exp(-ytm_risky * years_to_maturity) - recovery_rate * np.exp(-ytm_risk_free * years_to_maturity)
        ln_term = np.log(exp_term / (1 - recovery_rate))
        lambda_value = -ln_term / years_to_maturity - ytm_risk_free
        return lambda_value
    except ValueError:
        # 处理可能的数学域错误，例如 log 或 exp 得到无效值时
        return np.nan


# 计算每一行的违约概率
default_probabilities = []
for index, row in data.iterrows():
    ytm_risky = row['YTM_MTN001']  # "20永煤MTN001" 的到期收益率
    ytm_risk_free = row['YTM_GZ03']  # "03国债03" 的到期收益率

    # 使用函数计算违约概率
    default_probability = calculate_default_probability(ytm_risky, ytm_risk_free, recovery_rate, years_to_maturity)
    default_probabilities.append(default_probability)


# 将违约概率添加到数据框
data['Default_Probability'] = default_probabilities

# 显示包含违约概率的更新后的数据框前几行
print(data[['日期', 'YTM_MTN001', 'YTM_GZ03', 'Default_Probability']].head())

# 可视化违约概率随时间的变化
sns.set_theme(style="whitegrid")
plt.figure(figsize=(14, 8))
sns.lineplot(x=data['日期'], y=data['Default_Probability'], label='20永煤MTN001 的违约概率', color='red')
plt.title("20永煤MTN001 随时间变化的违约概率")
plt.xlabel("日期")
plt.ylabel("违约概率 (λ)")
plt.xticks(rotation=45)
plt.legend()
plt.show()
