from DrissionPage import SessionPage

import pandas as pd
import os
import time

# 定义Excel文件路径
excel_path = os.path.join(r'c:\Users\MAO\Desktop\爬虫', 'bank_codes.xlsx')
# 定义结果保存路径
output_excel_path = os.path.join(r'c:\Users\MAO\Desktop\爬虫', 'bank_details.xlsx')

# 确保目录存在
os.makedirs(os.path.dirname(output_excel_path), exist_ok=True)

# 检查是否安装了必要的依赖
try:
    import openpyxl
except ImportError:
    print("警告: 未找到openpyxl库，将尝试使用其他引擎")


def crawl_bank_details(url):
    """爬取单个银行页面的详细信息"""
    print(f"正在访问: {url}")
    page = SessionPage()
    try:
        page.get(url)
        # 等待页面加载完成
        time.sleep(2)
        
        # 查找所有信息行
        rows = page.eles('@class=SwiftCard_row__lQUWR')
        
        results = []
        
        for row in rows:
            # 提取左侧标签文本
            label_span = row.ele('@class=SwiftCard_label__jE2Cj', timeout=1)
            label_text = label_span.text.strip() if label_span else "Unknown Label"
            
            # 提取右侧区域
            right_div = row.ele('@class=SwiftCard_right__CTbJK', timeout=1)
            
            # 根据不同标签类型处理
            if label_text == "SWIFT 代码":
                # 提取SWIFT代码值
                code_span = right_div.ele('tag:span', timeout=0.5)
                swift_code = code_span.text.strip() if code_span else "N/A"
                
                # 只保存标签和代码值
                results.append({
                    'label': label_text,
                    'value': swift_code
                })
            else:
                # 提取普通信息值
                value_text = right_div.text.strip() if right_div else "N/A"
                
                results.append({
                    'label': label_text,
                    'value': value_text
                })
        
        # 打印结果
        print(f"找到 {len(results)} 条信息:")
        for i, item in enumerate(results, 1):
            print(f"{i}. [{item['label']}] {item['value']}")
        
        # 提取特定值
        bank_info = {'网址': url}
        if results:
            bank_info['SWIFT代码'] = next((item['value'] for item in results if item['label'] == "SWIFT 代码"), "未找到")
            bank_info['银行名称'] = next((item['value'] for item in results if item['label'] == "银行名称"), "未找到")
            bank_info['分行信息'] = next((item['value'] for item in results if item['label'] == "分行信息"), "未找到")
            bank_info['城市'] = next((item['value'] for item in results if item['label'] == "城市"), "未找到")
            bank_info['国家'] = next((item['value'] for item in results if item['label'] == "国家"), "未找到")
            
            print("\n关键信息汇总:")
            print(f"SWIFT代码: {bank_info['SWIFT代码']}")
            print(f"银行名称: {bank_info['银行名称']}")
            print(f"分行信息: {bank_info['分行信息']}")
            print(f"城市: {bank_info['城市']}")
            print(f"国家: {bank_info['国家']}")
        
        return bank_info
    
    except Exception as e:
        print(f"访问 {url} 时出错: {str(e)}")
        return None
    finally:
        page.close()


def save_to_excel(data, final=False):
    """将银行详细信息保存到Excel文件"""
    try:
        # 创建DataFrame
        df = pd.DataFrame(data)
        
        # 保存到Excel
        df.to_excel(output_excel_path, index=False, engine='openpyxl')
        
        if final:
            print(f"\n数据已成功保存到: {output_excel_path}")
            print(f"共保存 {len(df)} 条银行详细信息")
        else:
            print(f"已保存 {len(df)} 条银行详细信息到Excel")
    except Exception as e:
        print(f"保存数据到Excel时出错: {str(e)}")


def main():
    try:
        # 读取Excel文件
        print(f"正在读取Excel文件: {excel_path}")
        # 尝试使用openpyxl引擎读取Excel文件
        try:
            df = pd.read_excel(excel_path, engine='openpyxl')
        except Exception as e:
            print(f"使用openpyxl引擎读取Excel失败: {str(e)}")
            # 尝试使用xlrd引擎
            try:
                df = pd.read_excel(excel_path, engine='xlrd')
                print("成功使用xlrd引擎读取Excel文件")
            except Exception as e2:
                print(f"使用xlrd引擎读取Excel也失败: {str(e2)}")
                raise
        
        # 检查是否包含'网址'列
        if '网址' not in df.columns:
            print("错误: Excel文件中未找到'网址'列")
            return
        
        # 存储爬取结果
        all_bank_details = []
        total_urls = len(df['网址'])
        
        # 遍历每个URL
        for idx, url in enumerate(df['网址'], 1):
            print(f"\n进度: {idx}/{total_urls}")
            bank_info = crawl_bank_details(url)
            
            if bank_info:
                all_bank_details.append(bank_info)
                
                # 每爬取10条数据保存一次
                if len(all_bank_details) % 10 == 0:
                    save_to_excel(all_bank_details)
            
            # 添加延迟，避免请求过快
            time.sleep(1)
        
        # 最终保存所有数据
        save_to_excel(all_bank_details, final=True)
        
    except Exception as e:
        print(f"程序执行出错: {str(e)}")


if __name__ == "__main__":
    from DrissionPage import SessionPage
    main()