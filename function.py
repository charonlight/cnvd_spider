import time
import pandas as pd
import os


# 获取当前时间
def get_current_time():
    ticks = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    # print(get_current_time(), '[*]>>',)
    return ticks


# 定义拼接函数，并对字段进行去重
def concat_func(x):
    return pd.Series({
        'CNVD_ID': '、'.join(x['CNVD_ID'].unique()),
        '漏洞名称': '、'.join(x['漏洞名称'].unique()),
        '危害级别': '、'.join(x['危害级别']),  # 不能去重
        '公开日期': '、'.join(x['公开日期'].unique()),  # 基本上公开日期都在同一天，所以去重
        'CVE_ID': '、'.join(x['CVE_ID'].unique()),
        '漏洞描述': "".join(x['漏洞描述'].unique()),  # 由于漏洞描述后都是句号结束，所以不拼接顿号
        '漏洞解决方案': '、'.join(x['漏洞解决方案'].unique()),
    })


# csv漏洞信息进行合并去重提取（读取./Spider/CNVD_Vul_Info.csv，根据影响产品进行汇总）
def CsvVulInfo_Extract(Spider_Csv, New_Csv):
    print(get_current_time(), '>>>>>', "根据影响产品字段进行去重汇总")

    df = pd.read_csv(f'./tmp/{Spider_Csv}')

    # 分组聚合+拼接
    # result = df.groupby(df['影响产品']).apply(concat_func).reset_index()  # 有索引值
    result = df.groupby(df['影响产品']).apply(concat_func)
    # print(result)

    df_data = pd.DataFrame(data=result)

    # 序号列为从1开始的自增列，默认加在dataframe最右侧
    df_data['序号'] = range(1, len(df_data) + 1)
    # 对原始列重新排序，使自增列位于最左侧
    df_data = df_data[['序号', '漏洞名称', '公开日期', 'CNVD_ID', 'CVE_ID', '危害级别', '漏洞描述', '漏洞解决方案']]

    if not os.path.exists("./tmp"):
        os.mkdir("./tmp")
    df_data.to_csv(f"./tmp/{New_Csv}")
    # df1.to_excel("影响产品字段去重之后Vul_Info.xlsx")

    print(get_current_time(), '[*]>>', "去重汇总完成")
