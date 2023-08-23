# 确保有以下的库 IF NOT 利用下面命令安装
# pip install pandas openpyxl

import pandas as pd


def process_vx_file(file_path):
    df = pd.read_excel(file_path, engine='openpyxl')

    # 1. 取消所有数据的合并
    # 2. 删除空行
    df.dropna(how='all', inplace=True)

    # 3. 在交易单号这列前插入新的列
    df.insert(df.columns.get_loc("交易单号"), "DH", "DH" + df["交易单号"].astype(str))
    df.drop(columns="交易单号", inplace=True)

    # 4. 在商户单号这列前插入新的列
    df.insert(df.columns.get_loc("商户单号"), "SH", "SH" + df["商户单号"].astype(str))
    df.drop(columns="商户单号", inplace=True)

    # 5. 在金额(元)这列前插入新的三列
    df.insert(df.columns.get_loc("金额(元)"), "in", df["金额(元)"].where(df["收/支/其他"] == "收入", 0))
    df.insert(df.columns.get_loc("金额(元)"), "out", df["金额(元)"].where(df["收/支/其他"] == "支出", 0))
    df.insert(df.columns.get_loc("金额(元)"), "other", df["金额(元)"].where(df["收/支/其他"] == "其他", 0))

    # 6. 删除金额(元)这列和收/支/其他这列
    df.drop(columns=["金额(元)", "收/支/其他"], inplace=True)

    # 7. 调整各列的顺序
    df = df[["DH", "交易时间", "交易类型", "in", "out", "other", "交易方式", "交易对方", "SH"]]

    return df


def process_zfb_file(file_path):
    df = pd.read_excel(file_path, header=None, engine='openpyxl')

    # 0. 创建一行作为首行
    df.columns = ["收/支/其他", "交易对方", "交易类型", "交易方式", "金额(元)", "交易单号", "商户单号", "交易时间"]

    # 1. 取消所有的单元格合并
    # 2. 删除空行
    df.dropna(how='all', inplace=True)

    # 3. 在交易单号这列前插入新的列
    df.insert(df.columns.get_loc("交易单号"), "DH", "DH" + df["交易单号"].astype(str))
    df.drop(columns="交易单号", inplace=True)

    # 4. 在商户单号这列前插入新的列
    df.insert(df.columns.get_loc("商户单号"), "SH", "SH" + df["商户单号"].astype(str))
    df.drop(columns="商户单号", inplace=True)

    # 5. 在金额(元)这列前插入新的三列
    df.insert(df.columns.get_loc("金额(元)"), "in", df["金额(元)"].where(df["收/支/其他"] == "收入", 0))
    df.insert(df.columns.get_loc("金额(元)"), "out", df["金额(元)"].where(df["收/支/其他"] == "支出", 0))
    df.insert(df.columns.get_loc("金额(元)"), "other", df["金额(元)"].where(df["收/支/其他"] == "其他", 0))

    # 6. 删除金额(元)这列和收/支/其他这列
    df.drop(columns=["金额(元)", "收/支/其他"], inplace=True)

    # 7. 调整各列的顺序
    df = df[["DH", "交易时间", "交易类型", "in", "out", "other", "交易方式", "交易对方", "SH"]]

    return df


def main_vx(files):
    dfs = []
    for file in files:
        dfs.append(process_vx_file(file))

    # 9. 把所有处理结果合并成一个表
    combined_df = pd.concat(dfs, ignore_index=True)

    return combined_df


def main_zfb(files):
    dfs = []
    for file in files:
        dfs.append(process_zfb_file(file))

    # 9. 把所有处理结果合并成一个表
    combined_df = pd.concat(dfs, ignore_index=True)

    return combined_df


# 10. 把上面#1 VX 和#2 ZFB的结果合并成同一个Excel文件
def save_to_excel(vx_files, zfb_files, output_file):
    vx_df = main_vx(vx_files)
    zfb_df = main_zfb(zfb_files)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        vx_df.to_excel(writer, sheet_name='VX', index=False)
        zfb_df.to_excel(writer, sheet_name='ZFB', index=False)

# 您可以使用save_to_excel函数并传入所有的VX和ZFB文件路径列表以及输出文件的路径，该函数将处理所有的文件并将其保存到一个Excel文件中。