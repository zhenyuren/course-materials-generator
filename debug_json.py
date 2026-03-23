import json

# 读取JSON文件
json_path = "/Users/dudu/Desktop/智能金融学院/智能体设计/期初资料1/任渝_课程信息_2026年03月21日.json"

try:
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    print(f"JSON数据类型: {type(data)}")
    
    if isinstance(data, dict):
        print("JSON数据是字典")
        print(f"字典键: {list(data.keys())}")
        
        if 'courses' in data:
            courses = data['courses']
            print(f"courses类型: {type(courses)}")
            if isinstance(courses, list):
                print(f"courses列表长度: {len(courses)}")
                if courses:
                    print(f"第一个课程类型: {type(courses[0])}")
                    print(f"第一个课程内容: {courses[0]}")
                    
    elif isinstance(data, list):
        print(f"JSON数据是列表，长度: {len(data)}")
        if data:
            print(f"第一个元素类型: {type(data[0])}")
            print(f"第一个元素内容: {data[0]}")
            
except Exception as e:
    print(f"读取JSON文件失败: {e}")
