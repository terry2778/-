from DrissionPage import SessionPage
import string
import time
import pandas as pd
import os

# 定义Excel文件保存路径
excel_path = os.path.join(r'c:\Users\MAO\Desktop\爬虫', 'bank_codes.xlsx')

# 确保目录存在
os.makedirs(os.path.dirname(excel_path), exist_ok=True)

# 检查是否安装了必要的依赖
try:
    import openpyxl
except ImportError:
    print("警告: 未找到openpyxl库，将尝试使用其他引擎")


def crawl_page(letter, page_num):
    """爬取指定字母和页码的页面"""
    url = f'https://www.myswiftcodes.com/check/{page_num}?i={letter}'
    print(f"正在爬取: {url}")
    
    page = SessionPage()
    try:
        page.get(url)
        # 等待页面加载完成
        time.sleep(2)
        
        # 获取所有目标列表项
        list_items = page.eles('@@tag()=li@@class=CheckResultList_card___c74g')
        
        results = []
        for item in list_items:
            # 在列表项内查找目标链接
            header_div = item.ele('@class=CheckResultList_header__BH8D4')
            if header_div:
                link_tag = header_div.ele('@tag()=a')
                if link_tag:
                    # 提取链接和文本内容
                    link = link_tag.attr('href')
                    full_text = link_tag.texts()
                    results.append({
                        'link': f"{link}",
                        'text': ''.join(full_text).strip()
                    })
        
        print(f"字母 {letter} 第 {page_num} 页: 找到 {len(results)} 个结果")
        return results
    
    except Exception as e:
        print(f"爬取字母 {letter} 第 {page_num} 页时出错: {str(e)}")
        return None
    finally:
        page.close()


def main():
    # 存储所有银行数据
    all_bank_data = []
    total_count = 0  # 总数据计数
    save_threshold = 10  # 每10条保存一次
    
    # 总进度计算变量
    total_letters = len(string.ascii_uppercase)
    max_pages = 1000
    
    # 从字母A到Z循环
    for letter_idx, letter in enumerate(string.ascii_uppercase, 1):
        print(f"\n开始爬取字母 {letter}")
        
        # 从第1页到第1000页循环
        for page_num in range(1, 1001):
            max_retries = 5  # 设置最大重试次数
            retries = 0
            results = None
            while retries < max_retries:
                results = crawl_page(letter, page_num)
                if results is not None:
                    break  # 成功获取结果（可能为空列表），退出重试循环
                retries += 1
                print(f"重试字母 {letter} 第 {page_num} 页，第 {retries} 次")
                time.sleep(3)  # 增加延迟后重试
            
            if retries >= max_retries:
                print(f"达到最大重试次数 {max_retries} 次，跳过字母 {letter} 第 {page_num} 页")
                continue  # 跳过当前页，继续下一页
            
            # 如果当前页没有结果（正常情况），停止该字母的爬取
            if not results:
                print(f"字母 {letter} 第 {page_num} 页没有找到结果，停止该字母的爬取")
                break
            
            # 处理结果并添加到总数据中
            for i, res in enumerate(results, 1):
                # 提取银行编号（假设文本内容的开头部分是银行编号）
                bank_code = res['text'].split()[0] if res['text'] else f"未知编号-{letter}-{page_num}-{i}"
                
                # 完整网址（假设链接是相对路径）
                full_url = f"https://www.myswiftcodes.com{res['link']}" if res['link'].startswith('/') else res['link']
                
                # 添加到总数据
                all_bank_data.append({
                    '字母': letter,
                    '页码': page_num,
                    '索引': i,
                    '银行编号': bank_code,
                    '网址': full_url,
                    '完整文本': res['text']
                })
                
                total_count += 1
                print(f"{letter}-{page_num}-{i}. {bank_code} -> {full_url}")
                print(f"当前进度: 字母 {letter} ({letter_idx}/{total_letters}), 第 {page_num} 页, 共 {total_count} 条数据")
                
                # 每爬取10条数据保存一次
                if total_count % save_threshold == 0:
                    save_to_excel(all_bank_data)
            
            # 添加延迟，避免请求过快
            time.sleep(1)
        
    # 最终保存一次
    save_to_excel(all_bank_data, final=True)


def save_to_excel(data, final=False):
    """将数据保存到Excel文件"""
    try:
        # 创建DataFrame
        df = pd.DataFrame(data)
        
        # 去重，避免重复数据
        df = df.drop_duplicates(subset=['银行编号', '网址'])
        
        # 删除'完整文本'列
        if '完整文本' in df.columns:
            df = df.drop(columns=['完整文本'])
        
        # 保存到Excel
        df.to_excel(excel_path, index=False, engine='openpyxl')
        
        if final:
            print(f"\n数据已成功保存到: {excel_path}")
            print(f"共保存 {len(df)} 条银行数据")
        else:
            print(f"已保存 {len(df)} 条银行数据到Excel")
    except Exception as e:
        print(f"保存数据到Excel时出错: {str(e)}")


if __name__ == "__main__":
    main()