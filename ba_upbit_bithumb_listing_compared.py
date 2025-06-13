import requests
import pandas as pd
from datetime import datetime
import os

class ExchangeListings:
    def __init__(self):
        self.output_dir = "output"
        os.makedirs(self.output_dir, exist_ok=True)

    def get_binance_usdt_pairs(self):
        """获取币安USDT交易对"""
        url = "https://api.binance.com/api/v3/exchangeInfo"
        try:
            print("获取币安USDT交易对...")
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            data = response.json()

            usdt_pairs = []
            for symbol in data['symbols']:
                if (symbol['quoteAsset'] == 'USDT'
                        and symbol['status'] == 'TRADING'
                        and symbol['isSpotTradingAllowed']):
                    price_filter = next((f for f in symbol['filters'] if f['filterType'] == 'PRICE_FILTER'), {})
                    lot_size_filter = next((f for f in symbol['filters'] if f['filterType'] == 'LOT_SIZE'), {})

                    pair_info = {
                        'Symbol': symbol['symbol'],
                        'Base Asset': symbol['baseAsset'],
                        'Quote Asset': symbol['quoteAsset'],
                        'Price Precision': price_filter.get('tickSize', 'N/A'),
                        'Min Qty': lot_size_filter.get('minQty', 'N/A'),
                        'Qty Precision': lot_size_filter.get('stepSize', 'N/A'),
                        'Listing Date': datetime.fromtimestamp(symbol['onboardDate'] / 1000).strftime('%Y-%m-%d')
                        if 'onboardDate' in symbol else 'N/A'
                    }
                    usdt_pairs.append(pair_info)

            print(f"找到 {len(usdt_pairs)} 个币安USDT交易对")
            return pd.DataFrame(usdt_pairs)
        except Exception as e:
            print(f"获取币安数据失败: {e}")
            return pd.DataFrame()

    def get_bithumb_krw_pairs(self):
        """获取Bithumb KRW交易对"""
        try:
            print("获取Bithumb KRW交易对...")
            url = "https://api.bithumb.com/public/ticker/ALL_KRW"
            headers = {'User-Agent': 'Mozilla/5.0'}
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            data = response.json()

            krw_pairs = []
            for currency in data['data']:
                if currency == 'date':
                    continue
                krw_pairs.append({
                    'Market': f'KRW-{currency}',
                    'Currency': currency,
                    'Korean Name': '',
                    'English Name': ''
                })

            print(f"找到 {len(krw_pairs)} 个Bithumb KRW交易对")
            return pd.DataFrame(krw_pairs)
        except Exception as e:
            print(f"获取Bithumb数据失败: {e}")
            return pd.DataFrame()

    def get_upbit_krw_pairs(self):
        """获取Upbit KRW交易对"""
        try:
            print("获取Upbit KRW交易对...")
            url = "https://api.upbit.com/v1/market/all"
            headers = {'User-Agent': 'Mozilla/5.0'}
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            markets = response.json()

            krw_pairs = []
            for market in markets:
                if market['market'].startswith('KRW-'):
                    krw_pairs.append({
                        'Market': market['market'],
                        'Korean Name': market['korean_name'],
                        'English Name': market['english_name']
                    })

            print(f"找到 {len(krw_pairs)} 个Upbit KRW交易对")
            return pd.DataFrame(krw_pairs)
        except Exception as e:
            print(f"获取Upbit数据失败: {e}")
            return pd.DataFrame()

    def save_to_excel(self):
        """获取数据并进行比较，保存结果到Excel文件"""
        # 获取币安数据
        binance_df = self.get_binance_usdt_pairs()
        if not binance_df.empty:
            binance_df = binance_df[['Base Asset']]
            binance_df.columns = ['Asset']  # 统一列名

        # 获取Upbit数据
        upbit_df = self.get_upbit_krw_pairs()
        if not upbit_df.empty:
            upbit_df['Market'] = upbit_df['Market'].str.replace('KRW-', '')
            upbit_df = upbit_df[['Market']]
            upbit_df.columns = ['Asset']  # 统一列名

        # 获取Bithumb数据
        bithumb_df = self.get_bithumb_krw_pairs()
        if not bithumb_df.empty:
            bithumb_df = bithumb_df[['Currency']]
            bithumb_df.columns = ['Asset']  # 统一列名

        # 获取各个交易所的币种集合
        binance_set = set(binance_df['Asset']) if not binance_df.empty else set()
        upbit_set = set(upbit_df['Asset']) if not upbit_df.empty else set()
        bithumb_set = set(bithumb_df['Asset']) if not bithumb_df.empty else set()

        # 计算在三家交易所同时上线的币种
        all_three = binance_set & upbit_set & bithumb_set

        # 计算只在币安上线的币种
        only_binance = binance_set - upbit_set - bithumb_set

        # 计算只在 Upbit 上线的币种
        only_upbit = upbit_set - binance_set - bithumb_set

        # 计算只在 bithumb 上线的币种
        only_bithumb = bithumb_set - binance_set - upbit_set

        # 计算同时只在币安和 upbit 上线的币种
        binance_upbit = binance_set & upbit_set - bithumb_set

        # 计算同时只在币安和 bithumb 上线的币种
        binance_bithumb = binance_set & bithumb_set - upbit_set

        # 计算同时只在 bithumb 和 upbit 上线的币种
        bithumb_upbit = bithumb_set & upbit_set - binance_set

        # 创建新的 Excel 文件
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(self.output_dir, f"Exchange_Listings_Results_{timestamp}.xlsx")
        with pd.ExcelWriter(filename) as writer:
            # 将结果保存到不同的 sheet 中
            pd.DataFrame(all_three, columns=['Asset']).to_excel(writer, sheet_name='ALL', index=False)
            pd.DataFrame(only_binance, columns=['Asset']).to_excel(writer, sheet_name='only_ba', index=False)
            pd.DataFrame(only_upbit, columns=['Asset']).to_excel(writer, sheet_name='only_upbit', index=False)
            pd.DataFrame(only_bithumb, columns=['Asset']).to_excel(writer, sheet_name='only_bithumb', index=False)
            pd.DataFrame(binance_upbit, columns=['Asset']).to_excel(writer, sheet_name='ba_upbit', index=False)
            pd.DataFrame(binance_bithumb, columns=['Asset']).to_excel(writer, sheet_name='ba_bithumb', index=False)
            pd.DataFrame(bithumb_upbit, columns=['Asset']).to_excel(writer, sheet_name='upbit_bithumb', index=False)

        print(f"\n数据已保存到: {filename}")


if __name__ == "__main__":
    print("=== 交易所上币情况整合工具 ===")
    print("正在获取币安、Upbit和Bithumb的上币信息...")

    el = ExchangeListings()
    el.save_to_excel()

    print("程序执行完毕！")