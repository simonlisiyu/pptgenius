#!/usr/bin/env python3
import markdown
from bs4 import BeautifulSoup

def check_markdown_elements(md_text):
    """
    将Markdown转换为HTML并检查是否包含特定元素
    返回HTML内容和检查结果
    """
    # 配置Markdown解析器，启用表格和代码块扩展
    md = markdown.Markdown(extensions=['tables', 'fenced_code'])
    
    # 转换为HTML
    html = md.convert(md_text)
    
    # 使用BeautifulSoup解析HTML
    soup = BeautifulSoup(html, 'html.parser')
    
    # 检查各种元素
    has_images = len(soup.find_all('img')) > 0
    has_tables = len(soup.find_all('table')) > 0
    has_code = len(soup.find_all(['code', 'pre'])) > 0
    
    # 统计元素数量
    image_count = len(soup.find_all('img'))
    table_count = len(soup.find_all('table'))
    code_count = len(soup.find_all(['code', 'pre']))
    
    # 返回结果
    result = {
        'has_images': has_images,
        'has_tables': has_tables,
        'has_code': has_code,
        'image_count': image_count,
        'table_count': table_count,
        'code_count': code_count,
        'html': html
    }
    
    return result

# 测试用的Markdown文本
test_md = """# 测试文档

## 图片测试
![示例图片](images/example.jpg)

## 表格测试
| 项目 | 数值 | 说明 |
|------|------|------|
| 销售额 | 100万 | 同比增长20% |
| 用户数 | 1000 | 活跃用户 |

## 代码测试
```python
def hello():
    print("Hello, World!")
```
"""

if __name__ == "__main__":
    # 运行测试
    result = check_markdown_elements(test_md)
    
    # 打印结果
    print("HTML输出:")
    print("-" * 50)
    print(result['html'])
    print("\n检查结果:")
    print("-" * 50)
    print(f"包含图片: {'是' if result['has_images'] else '否'} (数量: {result['image_count']})")
    print(f"包含表格: {'是' if result['has_tables'] else '否'} (数量: {result['table_count']})")
    print(f"包含代码: {'是' if result['has_code'] else '否'} (数量: {result['code_count']})") 