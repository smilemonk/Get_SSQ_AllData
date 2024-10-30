# 双色球历史数据获取工具

一个用Python编写的双色球历史开奖数据获取工具，可以自动下载并保存所有历史数据到Excel文件。

## 功能特点

- 自动获取双色球历史开奖数据
- 增量更新，只获取新数据
- 保存为Excel格式，支持自动格式化
- 终端彩色输出，直观展示数据
- 支持断点续传

## 使用方法

1. 确保已安装Python 3.x
2. 安装依赖包：
bash
pip install requests pandas openpyxl
3. 运行脚本：
bash
python Get_SSQ_AllData.py

## 数据格式

Excel文件包含以下列：
- 期号
- 开奖日期
- 红球1-6
- 蓝球

## 作者

山猫

## 许可证

MIT