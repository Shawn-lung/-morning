import os
import pandas as pd
import numpy as np
from sklearn.preprocessing import MinMaxScaler

def calculate_scores(input_file, output_file):
    # 读取Excel文件到DataFrame
    df = pd.read_excel(input_file, sheet_name="Sheet")
    
    # 确保DataFrame包含所有行
    print(f"Total rows read: {len(df)}")
    
    # 只保留需要的列并去除包含NaN值的行
    required_columns = ["EPS", "Beta", "PE Ratio", "ROE", "Gross margin", "P/B Ratio", "Revenue per share", "Operating margin"]
    df = df.dropna(subset=required_columns)

    # 将数值列转换为数值类型
    for col in required_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # 标准化这些列
    scaler = MinMaxScaler()
    df[['Normalized ROE', 'Normalized EPS', 'Normalized Gross Margin', 'Normalized Revenue per Share', 'Normalized PB Ratio', 'Normalized PE Ratio', 'Normalized Operating Margin']] = scaler.fit_transform(
        df[['ROE', 'EPS', 'Gross margin', 'Revenue per share', 'P/B Ratio', 'PE Ratio', 'Operating margin']]
    )

    # 定义名人选股策略
    def buffet_score(row):
        weights = {
            'Normalized ROE': 0.575,
            'Normalized EPS': 0.2462,
            'Normalized Gross Margin': 0.0695,
            'Normalized Revenue per Share': 0.0669,
            'Normalized PB Ratio': 0.0424
        }
        return (row['Normalized ROE'] * weights['Normalized ROE'] +
                row['Normalized EPS'] * weights['Normalized EPS'] +
                row['Normalized Gross Margin'] * weights['Normalized Gross Margin'] +
                row['Normalized Revenue per Share'] * weights['Normalized Revenue per Share'] -
                row['Normalized PB Ratio'] * weights['Normalized PB Ratio'])

    def graham_score(row):
        weights = {
            'Normalized PE Ratio': 0.4286,
            'Normalized PB Ratio': 0.4286,
            'Normalized EPS': 0.1429
        }
        return (row['Normalized PE Ratio'] * weights['Normalized PE Ratio'] +
                row['Normalized PB Ratio'] * weights['Normalized PB Ratio'] +
                row['Normalized EPS'] * weights['Normalized EPS'])

    def o_shaughnessy_score(row):
        weights = {
            'Normalized EPS': 0.637,
            'Normalized PE Ratio': 0.2583,
            'Normalized ROE': 0.1047
        }
        return (row['Normalized EPS'] * weights['Normalized EPS'] +
                row['Normalized PE Ratio'] * weights['Normalized PE Ratio'] +
                row['Normalized ROE'] * weights['Normalized ROE'])

    def lynch_score(row):
        weights = {
            'Normalized PE Ratio': 0.5825,
            'Normalized Revenue per Share': 0.2362,
            'Normalized Gross Margin': 0.0789,
            'Normalized PB Ratio': 0.1024
        }
        return (row['Normalized PE Ratio'] * weights['Normalized PE Ratio'] +
                row['Normalized Revenue per Share'] * weights['Normalized Revenue per Share'] +
                row['Normalized Gross Margin'] * weights['Normalized Gross Margin'] -
                row['Normalized PB Ratio'] * weights['Normalized PB Ratio'])

    def murphy_score(row):
        weights = {
            'Normalized ROE': 0.6132,
            'Normalized Operating Margin': 0.5868
        }
        return (row['Normalized ROE'] * weights['Normalized ROE'] +
                row['Normalized Operating Margin'] * weights['Normalized Operating Margin'])

    # 计算每只股票的评分
    df['Buffet Score'] = df.apply(buffet_score, axis=1)
    df['Graham Score'] = df.apply(graham_score, axis=1)
    df['O\'Shaughnessy Score'] = df.apply(o_shaughnessy_score, axis=1)
    df['Lynch Score'] = df.apply(lynch_score, axis=1)
    df['Murphy Score'] = df.apply(murphy_score, axis=1)

    # 将结果保存回Excel
    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name="Sheet", index=False)

if __name__ == "__main__":
    # 输入文件路径与输出文件路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    input_file = os.path.join(current_dir, "taiex_mid100_stock_data.xlsx")
    output_file = os.path.join(current_dir, "taiex_mid100_stock_data_with_scores.xlsx")

    # 计算并保存评分
    calculate_scores(input_file, output_file)
