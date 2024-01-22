import glob
import pandas as pd

# 定义文件的搜索模式
pattern_order = './*楽々出荷分.csv'
pattern_jiaxu = './加須商品別在庫情報_*.csv'
pattern_xiqu = './福岡商品別在庫情報_*.csv'

# 使用 glob.glob 来找到匹配的文件
files_order = glob.glob(pattern_order)
files_jiaxu = glob.glob(pattern_jiaxu)
files_xiqu = glob.glob(pattern_xiqu)

# 检查是否找到文件，并选择第一个匹配项（如果有多个匹配）
file_path_order = files_order[0] if files_order else None
file_path_jiaxu = files_jiaxu[0] if files_jiaxu else None
file_path_xiqu = files_xiqu[0] if files_xiqu else None

# 检查所有文件是否都已找到
if not file_path_order or not file_path_jiaxu or not file_path_xiqu:
    print("未能找到所有必要的CSV文件，请检查它们是否在同一目录下。")
    # 这里可以加入退出程序的代码，如：exit()

# 尝试使用不同的编码格式读取文件
try:
    df_order = pd.read_csv(file_path_order, encoding='utf-8')
    df_jiaxu = pd.read_csv(file_path_jiaxu, encoding='utf-8')
    df_xiqu = pd.read_csv(file_path_xiqu, encoding='utf-8')
except UnicodeDecodeError:
    try:
        df_order = pd.read_csv(file_path_order, encoding='cp932')
        df_jiaxu = pd.read_csv(file_path_jiaxu, encoding='cp932')
        df_xiqu = pd.read_csv(file_path_xiqu, encoding='cp932')
    except UnicodeDecodeError:
        df_order = pd.read_csv(file_path_order, encoding='')
        df_jiaxu = pd.read_csv(file_path_jiaxu, encoding='shift_jis')
        df_xiqu = pd.read_csv(file_path_xiqu, encoding='shift_jis')


# 映射关系
column_mapping = {
    '請求ID':'店舗伝票番号',
    '請求日':'受注日' ,
    '顧客:郵便番号':'受注郵便番号',
    '顧客:住所':'受注住所１',
    '顧客:住所２':'受注住所２',
    '顧客:社名':'受注名',
    '顧客:社名カナ':'受注名カナ',
    '顧客:電話':'受注電話番号',
    '顧客:担当者アドレス':'受注メールアドレス',
    '納品先:郵便番号':'発送郵便番号',
    '納品先:住所':'発送先住所１',
    '納品先:住所２':'発送先住所２',
    '納品先:社名':'発送先名',
    '納品先:社名カナ':'発送先カナ',
    '納品先:電話':'発送電話番号',
    '支払方法':'支払方法',
    '税抜金額（合計）':'商品計',
    '消費税額（合計）':'税金',
    '発送料':'発送料',
    '手数料':'手数料',
    'ポイント':'ポイント',
    'その他費用':'その他費用',
    '税込金額（小計）':'合計金額',
    'ギフトフラグ':'ギフトフラグ',
    '時間帯指定':'時間帯指定',
    '日付指定':'日付指定', 
    '物流指示３':'備考',
    '品目':'商品名',
    '商品コード':'商品コード',
    '単価':'商品価格',
    '数量':'受注数量',
    '商品オプション':'商品オプション',
    '出荷済みフラグ':'出荷済フラグ',
    '顧客区分':'顧客区分',
    '顧客ID':'顧客コード',
    
}

# 定义最终 DataFrame 的列顺序
final_columns = list(column_mapping.values()) #+ ['作業者欄', '消費税率（%)','発送方法']  # 在这里添加其他所有需要的列

# 在 '備考' 之前插入 '作業者欄'
biko_index = final_columns.index('備考') if '備考' in final_columns else len(final_columns)
final_columns.insert(biko_index, '作業者欄')

# 在 '商品計' 之前插入 '発送方法'
shohin_kei_index = final_columns.index('商品計') if '商品計' in final_columns else len(final_columns)
final_columns.insert(shohin_kei_index, '発送方法')

# 添加剩余的列
final_columns += ['消費税率（%)']

# 初始化满足条件的DataFrame
df_jiaxu_full = pd.DataFrame(columns=final_columns )
df_xiqu_full = pd.DataFrame(columns=final_columns )
df_jiaxu_half = pd.DataFrame(columns=final_columns )
df_xiqu_half = pd.DataFrame(columns=final_columns )
df_stock_insufficient = pd.DataFrame(columns=['請求ID','商品総在庫数','商品コード','数量'])
error_df = pd.DataFrame(columns=['請求ID', '问题'])


# 找出所有品目为'送料'的行
shipping_rows = df_order[df_order['品目'] == '送料']

# 对于每个含有'送料'的'請求ID'，找到相同'請求ID'的第一个行并更新'発送料'
for index, row in shipping_rows.iterrows():
    request_id = row['請求ID']
    shipping_cost = row['税込金額（小計）']
    
    # 找到第一个匹配的'請求ID'的行
    matching_rows = df_order[df_order['請求ID'] == request_id].index
    if len(matching_rows) > 0:
        df_order.at[matching_rows[0], '発送料'] = shipping_cost

# 删除品目为送料的所有行
df_order = df_order[df_order['品目'] != '送料']

# 筛选出 '出荷ステータス' 为空或为 '出荷指示済' 的行
df_order_filtered = df_order[(df_order['出荷ステータス'].isna()) | (df_order['出荷ステータス'] == '出荷指示済')]

# 检查是否有 '出荷ステータス' 不为空且不为 '出荷指示済' 的行
non_compliant_rows = df_order[(df_order['出荷ステータス'].notna()) & (df_order['出荷ステータス'] != '出荷指示済')]

# 打印每个不符合条件的行的数量
print(f"Number of rows to be deleted: {len(non_compliant_rows)}")

# 如果有不符合条件的行，则为每行打印 'y'
for _ in range(len(non_compliant_rows)):
    print('y')


# 处理并填充数据
for index, row in df_order_filtered.iterrows():
    # 应用映射关系并创建新行
    new_row = {new_col: row[old_col] for old_col, new_col in column_mapping.items() if old_col in row}

     # 特殊处理：如果'出荷ステータス'为'出荷指示済'，则使用'明細キー'作为'店舗伝票番号'
    if row['出荷ステータス'] == '出荷指示済':
        new_row['店舗伝票番号'] = row['明細キー']

    # 特殊处理：如果'支払方法'为'FT代引'，则更改为'代引'
    if row['支払方法'] == 'FT代引':
        new_row['支払方法'] = '代引'

    # 继续其他处理
    new_row['作業者欄'] = ''
    new_row['消費税率（%)'] = 10
    new_row['発送方法'] = '佐川急便'

    # 清理空值
    new_row_cleaned = {k: v for k, v in new_row.items() if pd.notna(v)}

    # 检查 '出荷ステータス' 
    if pd.isna(row['出荷ステータス'])or (row['出荷ステータス'] == '出荷指示済'):
        print('x')
        
        # 检查是否有问题
        for col in ['納品先:郵便番号', '納品先:住所', '納品先:電話']:
            if pd.isna(row[col]) or '?' in str(row[col]):
                error_index = len(error_df)
                error_df.loc[error_index, '請求ID'] = row['請求ID']
                error_df.loc[error_index, '问题'] = f"{col}含有不明确的值或空值"
            # 对 '納品先:電話' 进行额外的格式检查
            elif col == '納品先:電話' and '-' not in str(row[col]):
                error_index = len(error_df)
                error_df.loc[error_index, '請求ID'] = row['請求ID']
                error_df.loc[error_index, '问题'] = '納品先:電話格式不正确（缺少"-")'

        
        # 查找对应的商品库存
        matching_row_jiaxu = df_jiaxu[df_jiaxu['品番'] == row['商品コード']]
        matching_row_xiqu = df_xiqu[df_xiqu['品番'] == row['商品コード']]

        # 计算总库存
        total_stock_jiaxu = matching_row_jiaxu['総在庫数'].sum() if not matching_row_jiaxu.empty else 0
        total_stock_xiqu = matching_row_xiqu['総在庫数'].sum() if not matching_row_xiqu.empty else 0
        total_stock = total_stock_jiaxu + total_stock_xiqu

        # 检查库存情况
        if not matching_row_jiaxu.empty and matching_row_jiaxu['総在庫数'].iloc[0] >= row['数量']:
            # df_jiaxu 可以完全满足订单
            if pd.isna(row['出荷ステータス']):
                df_jiaxu_full.loc[len(df_jiaxu_full)] = new_row
            else:
                df_jiaxu_half.loc[len(df_jiaxu_half)] = new_row
                # 更新 df_jiaxu 的库存
                df_jiaxu.loc[matching_row_jiaxu.index, '総在庫数'] -= row['数量']
        else:
            # 需要从 df_xiqu 发货的部分
            remaining_quantity = row['数量'] - total_stock_jiaxu if total_stock_jiaxu < row['数量'] else 0
            new_row['受注数量'] = remaining_quantity

            if not matching_row_xiqu.empty and total_stock >= row['数量']:
                # df_xiqu 能够满足剩余的需求
                if pd.isna(row['出荷ステータス']):
                    df_xiqu_full.loc[len(df_xiqu_full)] = new_row
                else:
                    df_xiqu_half.loc[len(df_xiqu_half)] = new_row
                    # 更新 df_xiqu 的库存
                    df_xiqu.loc[matching_row_xiqu.index, '総在庫数'] -= remaining_quantity
            elif total_stock < row['数量']:
                # 库存不足
                new_index = len(df_stock_insufficient)
                df_stock_insufficient.loc[new_index] = {
                    '請求ID': row['請求ID'],
                    '商品総在庫数': total_stock,
                    '商品コード': row['商品コード'],
                    '数量': remaining_quantity
                }
            # df_jiaxu 不能完全满足订单，但部分满足
            if total_stock_jiaxu > 0:
                new_row_partial = new_row.copy()
                new_row_partial['受注数量'] = total_stock_jiaxu
                if pd.isna(row['出荷ステータス']):
                    df_jiaxu_full.loc[len(df_jiaxu_full)] = new_row_partial
                else:
                    df_jiaxu_half.loc[len(df_jiaxu_half)] = new_row_partial


def calculate_total_amount(df):
    tax_rate = 0.1  # 税率为10%
    # 确保 NaN 值被替换为0
    df.fillna({'受注数量': 0, '商品価格': 0, '発送料': 0}, inplace=True)
    # 计算合计金额
    df['合計金額'] = ((df['受注数量'] * df['商品価格'] * (1 + tax_rate)) + df['発送料']).astype(int)
    return df          

# 应用这个函数到每个DataFrame
df_jiaxu_full = calculate_total_amount(df_jiaxu_full)
df_xiqu_full = calculate_total_amount(df_xiqu_full)
df_jiaxu_half = calculate_total_amount(df_jiaxu_half)
df_xiqu_half = calculate_total_amount(df_xiqu_half)

df_jiaxu_full.to_csv('加须请求（正常）.csv', index=False, encoding='utf-8-sig')
df_xiqu_full.to_csv('西区请求（正常）.csv', index=False, encoding='utf-8-sig')
df_jiaxu_half.to_csv('加须请求（分割）.csv', index=False, encoding='utf-8-sig')
df_xiqu_half.to_csv('西区请求（分割）.csv', index=False, encoding='utf-8-sig')
error_df.to_csv('error.csv', index=False, encoding='utf-8-sig')
df_stock_insufficient.to_csv('在库不足.csv', index=False, encoding='utf-8-sig')    