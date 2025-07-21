import pandas as pd

# 读取 upbit_pairs 文件
upbit_file = pd.ExcelFile(r'D:\ana\arbitrage\listing\output\upbit_pairs_20250721_164050.xlsx')

# 读取 bithumb_market_comparison 文件
bithumb_file = pd.ExcelFile(r'D:\ana\arbitrage\listing\output\bithumb_market_comparison_20250721_181858.xlsx')

# 提取 upbit 中 only_krw_pairs 和 only_btc_pairs 工作表的报价货币列
upbit_only_krw_pairs = upbit_file.parse('only_KRW_pairs')['报价货币'].tolist()
upbit_only_btc_pairs = upbit_file.parse('only_BTC_pairs')['报价货币'].tolist()

# 提取 bithumb 中 only_krw 和 only_btc 工作表的数据
bithumb_only_krw = bithumb_file.parse('only_KRW').iloc[:, 0].tolist()
bithumb_only_btc = bithumb_file.parse('only_BTC').iloc[:, 0].tolist()

# 对比并找出不同
# 对于 only_krw 部分
unique_upbit_only_krw = [pair for pair in upbit_only_krw_pairs if pair not in bithumb_only_krw]
unique_bithumb_only_krw = [pair for pair in bithumb_only_krw if pair not in upbit_only_krw_pairs]

# 对于 only_btc 部分
unique_upbit_only_btc = [pair for pair in upbit_only_btc_pairs if pair not in bithumb_only_btc]
unique_bithumb_only_btc = [pair for pair in bithumb_only_btc if pair not in upbit_only_btc_pairs]

# 创建新的 Excel 写入器
with pd.ExcelWriter('upbit_bithumb_result.xlsx', engine='openpyxl') as writer:
    # 保存仅在 upbit only_krw_pairs 中的报价货币到一个 sheet
    pd.DataFrame({'upbit_krw_only': unique_upbit_only_krw}).to_excel(writer,
                                                                                     sheet_name='upbit_krw_only',
                                                                                     index=False)
    # 保存仅在 bithumb only_krw 中的报价货币到一个 sheet
    pd.DataFrame({'bithumb_krw_only': unique_bithumb_only_krw}).to_excel(writer,
                                                                                     sheet_name='bithumb_krw_only',
                                                                                     index=False)
    # 保存仅在 upbit only_btc_pairs 中的报价货币到一个 sheet
    pd.DataFrame({'upbit_btc_only': unique_upbit_only_btc}).to_excel(writer,
                                                                                     sheet_name='upbit_btc_only',
                                                                                     index=False)
    # 保存仅在 bithumb only_btc 中的报价货币到一个 sheet
    pd.DataFrame({'bithumb_btc_only': unique_bithumb_only_btc}).to_excel(writer,
                                                                                     sheet_name='bithumb_btc_only',
                                                                                     index=False)

print("对比结果已保存到 comparison_result.xlsx")