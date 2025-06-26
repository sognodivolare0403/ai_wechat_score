import pandas as pd
import json
import os

def convert_jsonl_to_excel():
    """
    读取 results.jsonl 文件并将其转换为 Excel 表格
    """
    # 检查文件是否存在
    if not os.path.exists('results.jsonl'):
        print("错误：results.jsonl 文件不存在")
        return
    
    # 读取 JSONL 文件
    data = []
    with open('results.jsonl', 'r', encoding='utf-8') as file:
        for line in file:
            if line.strip():  # 跳过空行
                data.append(json.loads(line))
    
    if not data:
        print("警告：results.jsonl 文件为空，先运行main.py获取数据")
        return

    # 将数据转换为 DataFrame
    df = pd.DataFrame(data)
    
    # 处理嵌套的 sub_scores 字段
    if 'sub_scores' in df.columns:
        # 展开 sub_scores 字典
        sub_scores_df = pd.json_normalize(df['sub_scores'])
        # 重命名列以避免冲突
        sub_scores_df.columns = [f'sub_{col}' for col in sub_scores_df.columns]
        # 删除原始的 sub_scores 列
        df = df.drop('sub_scores', axis=1)
        # 合并展开后的数据
        df = pd.concat([df, sub_scores_df], axis=1)

    # 根据scored_position分组数据
    guihua_df = df[df['scored_position'] == 'guihua'].copy()
    xingzheng_df = df[df['scored_position'] == 'xingzheng'].copy()
    
    def prepare_dataframe_for_position(position_df, position_name):
        """为特定职位准备DataFrame，只保留相关的sub_columns"""
        if position_df.empty:
            return position_df
            
        # 基础列
        base_columns = ['file', 'sheet_name', 'scored_position', 'scored_person', 'total_score','comment']
        
        # 找出该职位类型的sub_columns
        position_sub_columns = []
        for col in position_df.columns:
            if col.startswith('sub_') and not position_df[col].isna().all():
                position_sub_columns.append(col)
        
        # 组合所有需要的列
        required_columns = base_columns + position_sub_columns
        
        # 确保所有列都存在
        existing_columns = [col for col in required_columns if col in position_df.columns]
        
        # 只保留相关的列
        result_df = position_df[existing_columns]
        
        return result_df
    
    # 为每个职位类型准备数据
    guihua_processed = prepare_dataframe_for_position(guihua_df, 'guihua')
    xingzheng_processed = prepare_dataframe_for_position(xingzheng_df, 'xingzheng')
    
    # 创建ExcelWriter对象，用于写入多个工作表
    output_file = 'results.xlsx'
    with pd.ExcelWriter(output_file) as writer:
        if not guihua_processed.empty:
            guihua_processed.to_excel(writer, sheet_name='guihua', index=False)
            print(f"guihua工作表：{len(guihua_processed)}行数据，{len([col for col in guihua_processed.columns if col.startswith('sub_')])}个sub_columns")
        
        if not xingzheng_processed.empty:
            xingzheng_processed.to_excel(writer, sheet_name='xingzheng', index=False)
            print(f"xingzheng工作表：{len(xingzheng_processed)}行数据，{len([col for col in xingzheng_processed.columns if col.startswith('sub_')])}个sub_columns")
    
    print(f"数据已成功导出到 {output_file}，包含 guihua 和 xingzheng 两个工作表")

if __name__ == "__main__":
    convert_jsonl_to_excel()