# 导入自定义函数
from spider import main_Spider
from function import CsvVulInfo_Extract
from banner import title, animate_banner
from report import docx_report, excel_report


# 逻辑
def main():
    # 打印banner信息
    title()
    # animate_banner()

    # # 启动爬虫，可自定义爬取页数，爬取的数据保存在./Spider/xxx.csv
    Spider_Csv = "CNVD_VulSpider_Info.csv"
    main_Spider(StartPage=1, EndPage=10, Spider_Csv=Spider_Csv)

    # 可以先对爬取的未去重的csv表格进行手工筛选，然后再进行代码去重合并

    # 读取爬下来的csv，根据影响产品字段进行漏洞信息汇总去重，去重之后的数据保存在./Spider/xxx.csv
    New_Csv = "影响产品字段去重之后Vul_Info.csv"
    CsvVulInfo_Extract(Spider_Csv=Spider_Csv, New_Csv=New_Csv)

    # 将csv转换成excel,生成的excel报告保存在./Report/xxx.xlsx
    excel_report(csv_filename=New_Csv, excel_filename="网络与信息安全情报收集（已根据影响产品进行了去重汇总）-2022.xlsx",
                 sheet_name="漏洞信息")  # 暂时不这样写

    # 读取去重之后的csv进行报告生成，生成的报告保存在./Report/xxx.docx
    docx_report(New_Csv=New_Csv, docx_filename="网络与信息安全情报收集-2022.docx")


if __name__ == '__main__':
    # 主程序，调用各个脚本文件
    main()
