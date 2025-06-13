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
        """保存到Excel文件，每个交易所一个sheet"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(self.output_dir, f"Exchange_Listings_{timestamp}.xlsx")

        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            # 币安数据
            binance_df = self.get_binance_usdt_pairs()
            if not binance_df.empty:
                binance_df.to_excel(writer, sheet_name='Binance_USDT', index=False)

            # Upbit数据
            upbit_df = self.get_upbit_krw_pairs()
            if not upbit_df.empty:
                upbit_df.to_excel(writer, sheet_name='Upbit_KRW', index=False)

            # Bithumb数据
            bithumb_df = self.get_bithumb_krw_pairs()
            if not bithumb_df.empty:
                bithumb_df.to_excel(writer, sheet_name='Bithumb_KRW', index=False)



            # 设置格式
            workbook = writer.book
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                worksheet.set_column('A:A', 15)
                worksheet.set_column('B:B', 15)
                worksheet.set_column('C:C', 15)
                if sheet_name == 'Binance_USDT':
                    worksheet.set_column('D:G', 12)
                else:
                    worksheet.set_column('D:D', 30)  # 韩文名称可能较长

        print(f"\n数据已保存到: {filename}")


if __name__ == "__main__":
    print("=== 交易所上币情况整合工具 ===")
    print("正在获取币安、Upbit和Bithumb的上币信息...")

    el = ExchangeListings()
    el.save_to_excel()

    print("程序执行完毕！")

#     帮我写个脚本处理这份文档
#     列出在三家交易所同时上线的
#     分别列出只在币安的，只在UPbit的只在bithumb上线的币种
#     并且分别列出同时只在币安和upbit上线的
#     同时只在币安和bithumb上线的
#     同时只在bithumb和upbit上线的