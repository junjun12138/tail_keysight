import pandas as pd
import re

# 定义正则表达式模式来匹配邮箱
email_pattern = r'^[a-zA-Z0-9._%+-]+@(keysight\.com)$'

# 读取Excel文件
file_path = 'C:/Users/jun1yin5/vsCode/DES.xlsx'  
df = pd.read_excel(file_path, sheet_name="Sheet1")

df['Web Email'] = df['Web Email'].fillna('')
# 筛选出“Web Email”列中符合邮箱格式的行
filtered_df = df[df['Web Email'].apply(lambda x: re.match(email_pattern, x) is not None)]

# 将筛选结果保存到新的Excel文件中
output_file_path = 'filtered_emails.xlsx'  # 替换为你想要保存的文件路径
filtered_df.to_excel(output_file_path, index=False)

print(f"筛选结果已保存到 {output_file_path}")
