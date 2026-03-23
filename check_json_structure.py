import json

# 加载JSON文件
with open('json/任渝_课程信息_2026年03月21日.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print('JSON数据结构:')
print('=' * 60)
print(f'数据类型: {type(data)}')

if isinstance(data, dict):
    print(f'键列表: {list(data.keys())}')
    # 假设课程数据在某个键下
    for key, value in data.items():
        print(f'\n键 {key}:')
        if isinstance(value, list):
            print(f'  列表长度: {len(value)}')
            if value:
                print(f'  第一个元素类型: {type(value[0])}')
                if isinstance(value[0], dict):
                    print(f'  字段列表: {list(value[0].keys())}')
                    # 显示第一个课程的信息
                    course = value[0]
                    print(f'\n第一个课程信息:')
                    for field, val in course.items():
                        print(f'  {field}: {val}')
else:
    print('数据不是字典类型')