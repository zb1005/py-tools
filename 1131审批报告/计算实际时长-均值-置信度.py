import pandas as pd
import numpy as np
from scipy import stats

def calculate_work_duration(start_time, end_time, holidays):
    # 计算工作时间（排除节假日和周末）
    current = start_time
    work_duration = timedelta()
    
    while current < end_time:
        next_day = (current + timedelta(days=1)).replace(hour=0, minute=0, second=0)
        if is_workday(current.date(), holidays):
            work_duration += min(next_day, end_time) - current
        current = next_day
    
    return work_duration.total_seconds() / (24 * 3600)  # 转换为天数


holiday_df = pd.read_excel(r"C:\Users\zhangbon\Desktop\临时活\1131统计\Q1审批时效数据分析打样-0513.xlsx",sheet_name="附2 方太春节假期", engine='openpyxl')
holidays = [datetime.strptime(str(date).strip(), '%Y-%m-%d %H:%M:%S').date() for date in holiday_df['方太假期']]


def analyze_approval_duration(df, group_field1, group_field2):
    """
    对审批时长进行分组统计和置信度分析
    :param df: 包含审批数据的DataFrame
    :param group_field1: 第一个分组字段名
    :param group_field2: 第二个分组字段名
    :return: 包含统计结果的DataFrame
    """
    # 分组计算计数和均值
    grouped = df.groupby([group_field1, group_field2])['实际时长'].agg(
        ['count', 'mean', 'std']
    ).reset_index()
    
    # 计算90%和95%的置信区间
    def calc_ci(row, confidence=0.95):
        if row['count'] <= 1 or pd.isna(row['std']):
            return (np.nan, np.nan)
        return stats.t.interval(
            confidence, 
            df=row['count']-1,
            loc=row['mean'],
            scale=row['std']/np.sqrt(row['count'])
        )
    
    grouped['90%_CI'] = grouped.apply(lambda x: calc_ci(x, 0.90), axis=1)
    grouped['95%_CI'] = grouped.apply(calc_ci, axis=1)
    
    # 拆分置信区间为上下限
    grouped['90%_CI_lower'] = grouped['90%_CI'].apply(lambda x: x[0])
    grouped['90%_CI_upper'] = grouped['90%_CI'].apply(lambda x: x[1])
    grouped['95%_CI_lower'] = grouped['95%_CI'].apply(lambda x: x[0])
    grouped['95%_CI_upper'] = grouped['95%_CI'].apply(lambda x: x[1])
    
    # 删除原始CI列
    grouped.drop(['90%_CI', '95%_CI'], axis=1, inplace=True)
    
    return grouped

# 使用示例
# df = pd.read_excel("您的数据文件.xlsx")
# result = analyze_approval_duration(df, "字段1", "字段2")
# result.to_excel("分析结果.xlsx", index=False)