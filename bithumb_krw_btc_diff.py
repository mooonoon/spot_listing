import requests
import pandas as pd
from datetime import datetime
import os


def fetch_bithumb_markets():
    """获取Bithumb的KRW和BTC交易对并进行比较"""
    try:
        # 获取KRW市场交易对
        krw_response = requests.get("https://api.bithumb.com/public/ticker/ALL_KRW")
        krw_data = krw_response.json()

        # 获取BTC市场交易对
        btc_response = requests.get("https://api.bithumb.com/public/ticker/ALL_BTC")
        btc_data = btc_response.json()

        # 提取KRW市场交易对列表
        krw_pairs = []
        for symbol in krw_data.get('data', {}).keys():
            if symbol != 'date':  # 排除时间戳字段
                krw_pairs.append(f'KRW-{symbol}')

        # 提取BTC市场交易对列表
        btc_pairs = []
        for symbol in btc_data.get('data', {}).keys():
            if symbol != 'date':  # 排除时间戳字段
                btc_pairs.append(f'BTC-{symbol}')

        # 提取基础货币（不包含计价货币）
        krw_base_currencies = {pair.split('-')[1] for pair in krw_pairs}
        btc_base_currencies = {pair.split('-')[1] for pair in btc_pairs}

        # 计算仅存在于KRW市场的交易对
        only_krw = [f'KRW-{coin}' for coin in krw_base_currencies - btc_base_currencies]

        # 计算仅存在于BTC市场的交易对
        only_btc = [f'BTC-{coin}' for coin in btc_base_currencies - krw_base_currencies]

        # 计算同时存在于两个市场的交易对
        both_markets = [f'KRW-{coin}, BTC-{coin}' for coin in krw_base_currencies & btc_base_currencies]

        return {
            'KRW_pairs': krw_pairs,
            'BTC_pairs': btc_pairs,
            'only_KRW': only_krw,
            'only_BTC': only_btc,
            'both_markets': both_markets
        }

    except Exception as e:
        print(f"获取数据时出错: {e}")
        return None


def save_to_excel(data):
    """将结果保存到Excel文件的不同工作表"""
    if not data:
        print("没有数据可保存")
        return

    # 创建输出目录
    output_dir = 'output'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 生成带时间戳的文件名
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = os.path.join(output_dir, f'bithumb_market_comparison_{timestamp}.xlsx')

    # 创建Excel写入器
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # 保存各个市场的交易对
        for sheet_name, pairs in data.items():
            df = pd.DataFrame(pairs, columns=[sheet_name])
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # 调整列宽
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                worksheet.column_dimensions[chr(65 + i)].width = max_len

    print(f"结果已保存到: {filename}")


def print_summary(data):
    """打印结果摘要"""
    if not data:
        return

    print("\n=== Bithumb市场交易对分析 ===")
    print(f"KRW市场交易对数量: {len(data['KRW_pairs'])}")
    print(f"BTC市场交易对数量: {len(data['BTC_pairs'])}")
    print(f"仅存在于KRW市场的交易对数量: {len(data['only_KRW'])}")
    print(f"仅存在于BTC市场的交易对数量: {len(data['only_BTC'])}")
    print(f"同时存在于两个市场的交易对数量: {len(data['both_markets'])}")

    print("\n同时存在于两个市场的交易对:")
    for pair in data['both_markets']:
        print(f"- {pair}")


if __name__ == "__main__":
    print("正在获取Bithumb交易所的交易对数据...")
    market_data = fetch_bithumb_markets()

    if market_data:
        print_summary(market_data)
        save_to_excel(market_data)
        print("\n分析完成！")
    else:
        print("获取数据失败，请检查API连接和响应格式。")