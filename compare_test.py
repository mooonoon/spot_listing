import pandas as pd

# 读取文件
excel_file = pd.ExcelFile('/mnt/Exchange_Listings_20250613_170436.xlsx')

# 获取各个 sheet 的数据
binance_df = excel_file.parse('Binance_USDT')
upbit_df = excel_file.parse('Upbit_KRW')
bithumb_df = excel_file.parse('Bithumb_KRW')

# 获取各个交易所的币种集合
binance_set = set(binance_df['Asset'])
upbit_set = set(upbit_df['Asset'])
bithumb_set = set(bithumb_df['Asset'])

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
with pd.ExcelWriter('/mnt/Exchange_Listings_Results.xlsx') as writer:
    # 将结果保存到不同的 sheet 中
    pd.DataFrame(all_three, columns=['Asset']).to_excel(writer, sheet_name='ALL', index=False)
    pd.DataFrame(only_binance, columns=['Asset']).to_excel(writer, sheet_name='only_ba', index=False)
    pd.DataFrame(only_upbit, columns=['Asset']).to_excel(writer, sheet_name='only_upbit', index=False)
    pd.DataFrame(only_bithumb, columns=['Asset']).to_excel(writer, sheet_name='only_bithumb', index=False)
    pd.DataFrame(binance_upbit, columns=['Asset']).to_excel(writer, sheet_name='ba_upbit', index=False)
    pd.DataFrame(binance_bithumb, columns=['Asset']).to_excel(writer, sheet_name='ba_bithumb', index=False)
    pd.DataFrame(bithumb_upbit, columns=['Asset']).to_excel(writer, sheet_name='upbit_bithumb', index=False)