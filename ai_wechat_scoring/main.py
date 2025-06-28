import requests
import json
from dotenv import load_dotenv
import os
import pandas as pd



def read_system_prompt(scored_position, scored_person):
    if scored_position == "xingzheng":
        system_prompt = open("system_prompt_xingzheng.txt", "r", encoding="utf-8").read()
        system_prompt = system_prompt.replace("{scored_person}", scored_person)
    elif scored_position == "guihua":
        system_prompt = open("system_prompt_guihua.txt", "r", encoding="utf-8").read()
        system_prompt = system_prompt.replace("{scored_person}", scored_person)
    
    return system_prompt

def read_user_message(excel_file):
    """
    从Excel文件中读取用户消息
    
    Args:
        excel_file (str): Excel文件路径
    
    Returns:
        str: 用户消息内容
    """
    
    try:
        # 读取Excel文件中的所有表格
        excel_data = pd.read_excel(excel_file, sheet_name=None)
        
        # 存储所有表格数据
        all_sheets_data = []
        
        # 需要读取的字段
        required_columns = ['Remark', 'StrContent', 'StrTime']
        
        # 遍历每个表格
        for sheet_name, df in excel_data.items():
            # 检查是否包含必要的列
            if 'Type' not in df.columns:
                print(f"表格 {sheet_name} 中没有找到 Type 列，跳过")
                continue

            # 筛选 type 为 1 的数据
            filtered_df = df[df['Type'] == 1]
            
            if filtered_df.empty:
                print(f"表格 {sheet_name} 中没有 Type 为 1 的数据，跳过")
                continue
            
            # 只选择需要的字段
            selected_df = filtered_df[required_columns]
            
            # 处理表格中的数据
            sheet_data = []
            for _, row in selected_df.iterrows():
                # 将行数据转换为字典
                row_dict = {}
                for column in required_columns:
                    # 处理NaN值
                    if pd.isna(row[column]):
                        row_dict[column] = ""
                    else:
                        # 处理日期时间类型
                        if isinstance(row[column], pd.Timestamp):
                            row_dict[column] = row[column].strftime("%Y-%m-%d %H:%M:%S")
                        else:
                            row_dict[column] = str(row[column])
                
                sheet_data.append(row_dict)
            
            # 将表格数据添加到总列表中
            all_sheets_data.append({
                "sheet_name": sheet_name,
                "data": sheet_data
            })
        return all_sheets_data
        
        # # 将数据转换为JSON字符串
        # user_message_list = json.dumps(all_sheets_data, ensure_ascii=False, indent=2)
        # return user_message_list
        
    except Exception as e:
        print(f"读取Excel文件出错: {e}")
        return f"读取Excel文件出错: {e}"

def get_api_config():
    """
    从环境变量获取API配置
    
    Returns:
        dict: API配置信息
    """
    load_dotenv() 
    
    # 构建配置
    config = {
        "url": f"{os.getenv('API_URL')}/v1/chat/completions",
        "key": os.getenv("API_KEY"),
        "model": os.getenv("MODEL"),
    }
    
    if not config["key"]:
        raise ValueError(f"请在.env文件中设置 {config['key_name']}")
    
    return config

def send_chat_request(system_prompt, user_message, model=None):
    """
    向AI API发送聊天请求，支持OpenAI、DeepSeek、Claude
    
    Args:
        system_prompt (str): 系统提示
        user_message (str or list): 用户消息，如果是list会转换为JSON字符串
        model (str): 使用的模型，如果不指定则使用环境变量配置
    
    Returns:
        dict: API响应
    """
    config = get_api_config()
    
    # 使用传入的模型或配置中的默认模型
    if model is None:
        model = config["model"]
    
    # 如果user_message是list，转换为JSON字符串
    if isinstance(user_message, list):
        user_message_content = json.dumps(user_message, ensure_ascii=False, indent=2)
    else:
        user_message_content = str(user_message)
    

    
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {config['key']}"
    }
    
    data = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message_content}
        ]
    }
    
    try:
        response = requests.post(config["url"], headers=headers, json=data, timeout=60)
        response.raise_for_status()
        result = response.json()
        return result
        
    except requests.exceptions.RequestException as e:
        print(f"❌ API请求出错: {e}")
        return None
    
def output_result(result, file, sheet_name):
    result_json = json.dumps(result, indent=2, ensure_ascii=False)
    with open(f"{file[:-5]}_{sheet_name}.json", "w", encoding="utf-8") as f:
        f.write(result_json)

# 使用示例
if __name__ == "__main__":
    scored_position = "xingzheng"
    scored_person = "董晗"
    system_prompt = read_system_prompt(scored_position, scored_person)

    # print(system_prompt)
    
    file = "xlsx/蒋博文_25申请.xlsx"
    all_sheets_data = read_user_message(file)
    
    rows = []
    for i, sheet in enumerate(all_sheets_data):
        if i == 1: break
        print(f"{file[5:-5]} - 第{i+1}个表格 - {sheet['sheet_name']}")
        user_message = sheet["data"]
        # print(user_message)
        result = send_chat_request(system_prompt, user_message)    
        if not result:
            print("请求失败")
        try:
            content = result["choices"][0]["message"]["content"]
        except TypeError as e:
            print(f"request failure: {e}")
            print(f"raw response: {result}")
            continue

        try:
            content_json = json.loads(content)
        except json.JSONDecodeError:
            print(f"JSON解析错误: {content}")
            continue

        content_json["file"] = file[5:-5]
        content_json["sheet_name"] = sheet["sheet_name"]
        content_json["scored_position"] = scored_position
        content_json["scored_person"] = scored_person
        
        with open("results.jsonl","a",encoding="utf-8") as f:
            f.write(json.dumps(content_json, ensure_ascii=False) + "\n") 
        
        print(f"{file[5:-5]} - 第{i+1}个表格 - {sheet['sheet_name']} - 完成")

