import time
from selenium.webdriver.common.by import By
from time import sleep
from function import get_current_time
from webdriver import get_webdriver
import csv
import os


# 爬虫程序
def main_Spider(StartPage, EndPage, Spider_Csv):
    print(get_current_time(), '>>>>>', "爬虫程序启动")

    web = get_webdriver()
    web.get("https://www.cnvd.org.cn/flaw/list")

    # # 会出现验证码
    # WafCode(web)

    if not os.path.exists("./tmp"):
        os.mkdir("./tmp")

    # 初始化
    with open(f'./tmp/{Spider_Csv}', 'w', newline='', encoding='utf-8') as f:
        f_csv = csv.DictWriter(f,
                               [
                                   "序号",
                                   "CNVD_ID",
                                   "漏洞名称",
                                   "危害级别",
                                   "公开日期",
                                   "影响产品",
                                   "CVE_ID",
                                   "漏洞描述",
                                   "漏洞解决方案"
                               ])
        f_csv.writeheader()  # 写入文件头

        count = 0  # 用于计数
        for page in range(StartPage, EndPage + 1):  # 爬取页数,自定义
            if page == 1:
                print(get_current_time(), '>>>>>', "开始爬取第1页")
            else:
                web.find_element(by=By.XPATH,
                                 value='//div[@class="pages clearfix"]//a[@class="nextLink"]').click()  # 点击下一页
                sleep(0.5)
                print(get_current_time(), '>>>>>', f"开始爬取第{page}页")
            try:
                for i in range(1, 11):  # 每页有10个漏洞标题需要点击
                    try:
                        time.sleep(10)  # 害怕被封ip，哈哈哈

                        count += 1
                        rows = vul_info_Spider(web, i, count)
                        # print(rows)
                        f_csv.writerows(rows)  # 写入csv
                        print(get_current_time(), '[*]>>', f'{rows[0]["序号"]} {rows[0]["漏洞名称"]}  已成功爬取该漏洞信息到csv表格中')
                    except:
                        print("当前漏洞 出错 ,下一个")
                        WafCode(web)
                        # time.sleep(10)
                        # 加几行代码判断一下是否被waf或者需要验证码识别
                        # 报错的3种可能  1数据解析出错 2waf 3验证码
                print(get_current_time(), '[*]>>', f"第{page}页已经爬取完毕")
            except:
                print(get_current_time(), '[-]>>', "爬取目标页漏洞出现未知错误,开始爬取下一页漏洞信息")
        print(get_current_time(), '[*]>>', f"爬虫程序结束,共计爬取了{count}个漏洞信息")


# 单页数据解析
def vul_info_Spider(web, i, count):
    sleep(1)  # 3秒爬一次,防止请求频繁被封了
    rows = []
    # 点击a标签进入漏洞详情页
    web.find_element(by=By.XPATH, value=f"/html/body/div[4]/div[1]/div/div[1]/table/tbody/tr[{i}]/td[1]/a").click()

    # 获取漏洞详情信息,进行数据解析操作
    # print(get_current_time(), '>>>>>', "开始对目标漏洞进行数据解析")
    try:
        # 漏洞名称  xxx问题漏洞（CNVD-2022-12745） 需要去除漏洞后面的文字
        vul_name = web.find_element(by=By.XPATH, value="/html/body/div[4]/div[1]/div[1]/div[1]/h1").text
        try:
            vul_name = vul_name.split("（")[0]
        except:
            pass

        # CNVD_ID
        cnvd_id = web.find_element(by=By.XPATH,
                                   value="/html/body/div[4]/div[1]/div[1]/div[1]/div[2]/div[1]/table/tbody/tr[1]/td[2]").text

        # 危害级别 需要去除非法字符
        vul_level = web.find_element(by=By.XPATH,
                                     value="/html/body/div[4]/div[1]/div[1]/div[1]/div[2]/div[1]/table/tbody/tr[3]/td[2]").text.split(
            " ")[0]
        # 公开日期  需要将2022-02-20 转化为2022年02月20日
        vul_data = web.find_element(by=By.XPATH,
                                    value="/html/body/div[4]/div[1]/div[1]/div[1]/div[2]/div[1]/table/tbody/tr[2]/td[2]").text.split(
            "-")
        vul_data = vul_data[0] + "年" + vul_data[1] + "月" + vul_data[2] + "日"
        # 影响产品  空行换成顿号
        affect_product = web.find_element(by=By.XPATH,
                                          value="/html/body/div[4]/div[1]/div[1]/div[1]/div[2]/div[1]/table/tbody/tr[4]/td[2]").text.replace(
            "\n", "、")

        # 解析需要注意的地方
        # #### CVE_ID 需要判断是否存在CVE_ID 不存在的话表格中的tr就会有变化
        try:
            cve_id = web.find_element(by=By.XPATH,
                                      value="/html/body/div[4]/div[1]/div[1]/div[1]/div[2]/div[1]/table/tbody/tr[5]/td[2]/a").text
            tr = 0
        except:
            cve_id = "该漏洞无CVE编号"
            # print("该漏洞无CVE编号")
            tr = -1
        # 漏洞描述---> 对漏洞描述进行分割，分为产品描述和漏洞危害  # 暂时先只用漏洞危害
        vul_description = web.find_element(by=By.XPATH,
                                           value=f"/html/body/div[4]/div[1]/div[1]/div[1]/div[2]/div[1]/table/tbody/tr[{6 + tr}]/td[2]").text.replace(
            "\n", '|')  # 没有替换为空
        # print(vul_description)
        vul_description = vul_description.split("|")[-1]  # 这是漏洞危害
        # print(vul_description)

        # 漏洞解决方案
        vul_solution = web.find_element(by=By.XPATH,
                                        value=f"/html/body/div[4]/div[1]/div[1]/div[1]/div[2]/div[1]/table/tbody/tr[{9 + tr}]/td[2]").text

        # 写入csv中的数据
        data = {
            "序号": count,
            "CNVD_ID": cnvd_id,
            "漏洞名称": vul_name,
            "危害级别": vul_level,
            "公开日期": vul_data,
            "影响产品": affect_product,
            "CVE_ID": cve_id,
            "漏洞描述": vul_description,
            "漏洞解决方案": vul_solution,
        }
        rows.append(data)
        # print(rows)  # 打印

    except:
        print(get_current_time(), '[-]>>', "爬取失败,请检查当前页面是否符合匹配规则")

    web.back()  # 回退到首页
    sleep(0.5)  # 慢一点爬,别着急
    return rows


# 判断是否被拦截，以及解决办法
def WafCode(web):
    time.sleep(3)
    web_title = web.title
    if web_title == "本站开启了验证码保护":
        print("遇到验证码,请在30秒内输入验证码并且提交")
        time.sleep(30)
    if web_title == "waf":
        print("请求太频繁了,遇到waf,慢一点,尝试刷新,并且等待60秒")
        time.sleep(60)


if __name__ == '__main__':
    # # 本套程序可单独使用，爬取最后结果为生成一个csv表格
    main_Spider(StartPage=1, EndPage=1, Spider_Csv="CNVD_Vul_Info.csv")