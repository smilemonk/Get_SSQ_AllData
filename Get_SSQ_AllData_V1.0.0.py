#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
双色球历史数据获取脚本
功能：获取双色球开奖历史数据并保存到Excel文件
作者：山猫
版本：1.0.0
最后更新：2024-10-30
"""

import requests
import json
import time
import random
from datetime import datetime
import pandas as pd
from pathlib import Path
import re
import openpyxl


class Colors:
    """终端输出颜色定义"""
    RED = '\033[91m'      # 红色 - 用于红球
    BLUE = '\033[94m'     # 蓝色 - 用于蓝球
    YELLOW = '\033[93m'   # 黄色 - 用于期号
    GREEN = '\033[92m'    # 绿色 - 用于日期
    WHITE = '\033[97m'    # 白色 - 用于固定文字
    END = '\033[0m'       # 结束颜色


def fetch_ssq_history():
    """获取双色球历史数据并保存到Excel文件"""
    
    # 设置文件路径
    excel_path = Path(__file__).parent / 'ssq_history.xlsx'
    print(f"数据文件路径: {excel_path.absolute()}")
    
    # 检查是否存在历史文件
    latest_code = None
    if excel_path.exists():
        try:
            df_existing = pd.read_excel(excel_path)
            if not df_existing.empty:
                latest_code = str(df_existing['期号'].iloc[0])
                print(f"读取本地文件，最近一期为: {latest_code}")
        except Exception as e:
            print(f"读取历史文件出错: {e}")
    
    # 设置请求头
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Connection': 'keep-alive',
        'Host': 'www.cwl.gov.cn',
        'Referer': 'https://www.cwl.gov.cn/ygkj/wqkjgg/ssq/',
        'X-Requested-With': 'XMLHttpRequest'
    }
    
    url = 'https://www.cwl.gov.cn/cwl_admin/front/cwlkj/search/kjxx/findDrawNotice'
    
    all_data = []
    page_no = 1
    session = requests.Session()
    
    try:
        # 先访问主页获取cookies
        session.get('https://www.cwl.gov.cn/ygkj/wqkjgg/ssq/', headers=headers)
        
        # 创建示例行计算分隔线长度
        sample_line = f"{Colors.YELLOW}新增第2024124期{Colors.END} {Colors.GREEN}(2024年10月29日){Colors.END} {Colors.WHITE}开奖号码：{Colors.RED}02 14 15 17 25 30{Colors.END} {Colors.WHITE}+{Colors.END} {Colors.BLUE}11{Colors.END}"
        display_line = sample_line.replace(Colors.YELLOW, '').replace(Colors.GREEN, '').replace(Colors.WHITE, '').replace(Colors.RED, '').replace(Colors.BLUE, '').replace(Colors.END, '')
        separator = "-" * len(display_line)
        
        while True:
            params = {
                'name': 'ssq',
                'pageNo': page_no,
                'pageSize': 30,
                'systemType': 'PC'
            }
            
            # 添加随机延时防止请求过快
            time.sleep(random.uniform(1, 2))
            
            response = session.get(
                url, 
                headers=headers, 
                params=params,
                timeout=10
            )
            response.raise_for_status()
            
            data = response.json()
            
            if data['state'] == 0:
                lottery_list = data['result']
                if not lottery_list:
                    break
                    
                # 第一页第一条数据前添加分隔线
                if len(all_data) == 0:
                    print("\n" + "=" * len(display_line))
                    print(f"{Colors.YELLOW}新增数据如下：{Colors.END}")
                    print(separator)
                
                for lottery in lottery_list:
                    draw_num = str(lottery['code'])
                    
                    # 如果找到已存在的期号，停止获取
                    if latest_code and draw_num <= latest_code:
                        print(separator)
                        print("\n已获取到最新数据，停止抓取")
                        break
                    
                    red_balls = lottery['red'].split(',')
                    blue_ball = lottery['blue']
                    date_str = lottery['date']
                    date_str = re.sub(r'\([^)]*\)', '', date_str)
                    draw_date = datetime.strptime(date_str.strip(), '%Y-%m-%d').strftime('%Y年%m月%d日')
                    
                    # 带颜色输出开奖信息
                    print(f"{Colors.YELLOW}新增第{draw_num}期{Colors.END} {Colors.GREEN}({draw_date}){Colors.END} {Colors.WHITE}开奖号码：{Colors.RED}{' '.join(red_balls)}{Colors.END} {Colors.WHITE}+{Colors.END} {Colors.BLUE}{blue_ball}{Colors.END}")
                    
                    all_data.append({
                        '期号': draw_num,
                        '开奖日期': draw_date,
                        '红球1': red_balls[0],
                        '红球2': red_balls[1],
                        '红球3': red_balls[2],
                        '红球4': red_balls[3],
                        '红球5': red_balls[4],
                        '红球6': red_balls[5],
                        '蓝球': blue_ball
                    })
                
                if latest_code and draw_num <= latest_code:
                    break
                    
                page_no += 1
            else:
                break
        
        # 保存数据到Excel
        if all_data:
            new_df = pd.DataFrame(all_data)
            newest_code = new_df['期号'].iloc[0]
            
            if excel_path.exists():
                old_df = pd.read_excel(excel_path)
                old_df['期号'] = old_df['期号'].astype(str)
                new_df['期号'] = new_df['期号'].astype(str)
                final_df = pd.concat([new_df, old_df]).drop_duplicates(subset=['期号'])
                print(f"成功更新{len(new_df)}期数据（从{latest_code}更新到{newest_code}）")
            else:
                final_df = new_df
                print(f"创建新文件，获取了{len(new_df)}期数据（最新一期为{newest_code}）")
            
            # 按期号排序
            final_df['期号'] = final_df['期号'].astype(str)
            final_df = final_df.sort_values('期号', ascending=False, key=lambda x: x.astype(str))
            
            # 保存并格式化Excel
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='双色球历史数据')
                
                worksheet = writer.sheets['双色球历史数据']
                
                # 设置列宽
                worksheet.column_dimensions['A'].width = 12  # 期号
                worksheet.column_dimensions['B'].width = 15  # 开奖日期
                for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I']:  # 红球1-6和蓝球
                    worksheet.column_dimensions[col].width = 8
                
                # 居中对齐
                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
            
            print(f"数据已保存到Excel文件，当前最新一期为：{newest_code}")
            print(f"文件保存位置: {excel_path.absolute()}")
        else:
            print("没有新数据需要更新，本地文件已是最新")
            
    except requests.exceptions.RequestException as e:
        print(f"请求出错: {e}")
    except json.JSONDecodeError as e:
        print(f"JSON解析错误: {e}")
    except Exception as e:
        print(f"发生错误: {e}")


if __name__ == "__main__":
    fetch_ssq_history()