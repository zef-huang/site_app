#coding=utf-8
import encodings.idna
import random
import threading
import time
import os
import traceback

from selenium import webdriver
import xlrd
import xlsxwriter
import yagmail
from selenium.webdriver import ActionChains


#输出测试报告
# def Ouput_test_result(excel_ouput, result):
#     xlsx = xlsxwriter.Workbook(excel_ouput)
#     table = xlsx.add_worksheet('testresult')
#     table.set_column('A:C', 50)
#     color = xlsx.add_format({'bg_color':'red'})
#     table.write_string(0,0,result,color)
#     xlsx.close()
from selenium.webdriver.support.wait import WebDriverWait


def record_log(filename, msg):
    with open("error_log.txt", 'a', encoding='utf8') as f:
        f.write('#' * 40 + '\n')
        f.write(filename + ' '+ get_date() + ' ' + get_time() + '出现异常')
        f.write(msg)

def reacord_new_game(filename, msg):
    with open('new_game.txt', 'a', encoding='utf8') as f:
        f.write('#' * 40 + '\n')
        f.write(filename + ' '+ get_date() + ' ' + get_time() + '新游戏更新')
        f.write(msg)

def reacord_no_change(filename, msg):
    with open('no_update.txt', 'a', encoding='utf8') as f:
        f.write('#' * 40 + '\n')
        f.write(filename + ' '+ get_date() + ' ' + get_time() + '网站没有更新')
        f.write(msg)

def record_first_time(filename, msg):
    with open("new_game.txt", 'a', encoding='utf8') as f:
        f.write(filename + ' '+ get_date() + ' ' + get_time() + msg)

#打开浏览器
def Open_browser():
    option = webdriver.ChromeOptions()
    option.headless = True
    return webdriver.Chrome(chrome_options=option)

#登陆网站
def Open_url(browser, url):
    browser.get(url)

def get_time():
    return time.strftime("%H:%M:%S", time.localtime(time.time()))

def get_date():
    return time.strftime("%Y-%m-%d", time.localtime(time.time()))

def get_xlsx_file():
    files = os.listdir()
    ret = [file for file in files if file.endswith('.xlsx')]
    return ret

# 读取excel文件， 获取需要监视的xpath地址
def get_single_web_data(path):
    xl = xlrd.open_workbook(path)
    table = xl.sheets()[0]
    url = table.row_values(0)[1]
    clickbroad = table.row_values(1)[1]
    monitor_place = {}  # {"xpath_place":[1, value]}

    i = 3
    while True:
        try:
            if len(table.row_values(i))<2:
                break
        except:
            break
        # [1, "xpath_place"]
        monitor_place[table.row_values(i)[1]] = [table.row_values(i)[0], None]
        i += 1

    return url, clickbroad, monitor_place

def check_clickbroad(browser, clickbroad, filename):
    # 1.需要点击操作的，进行点击操作，否则跳过
    if not clickbroad:
        pass
    else:
        browser.implicitly_wait(10)
        try:
            hang_place = browser.find_element_by_xpath(clickbroad)
            # 鼠标悬浮切换到角色扮演上面
            ActionChains(browser).move_to_element(hang_place).perform()
        except:
            record_log(filename,"原因：鼠标悬浮到角色扮演的xpath地址定位不到")

        # browser.find_element_by_xpath(clickbroad).click()
    # 2.如果点击后出现新页面，则做页面跳转关闭功能


# 寻找元素
def search_ele(browser, url, xpath_place, data, filename):
    try:
        browser.implicitly_wait(10)
        ret = browser.find_element_by_xpath(xpath_place)
        # 无头模式会导致一些值读取不到
        # if xpath_place == '/html/body/div[4]/div/div[3]/ul/li[5]/p':
        #     print(browser.find_element_by_xpath('/html/body/div[4]/div/div[3]/ul/li[5]/p').text)
        return ret
    except:
        # 定位符在一开始要有序号这样才能做到正确是识别哪个定位符失效
        msg = "{}文件中：{}网站中没有找到{}定位符".format(filename,url, data[0])
        print(msg)
        record_log(filename, msg)
        # 失败后重新打开网站和点击角色扮演进行重试
        # print("重试中....")
        # TODO
        # try:
        #     Open_url(browser, url)
        #     check_clickbroad(browser, clickbroad, filename)
        #     ret = browser.find_element_by_xpath(xpath_place)
        #     return ret
        # except Exception as e:
        #     # 记录错误
        #     print("重试失败")
        #     record_log(filename, msg)


# 获取游戏链接地址
def get_game_detail_url(browser, xpath_place, filename):
    browser.implicitly_wait(10)
    try:
        game_name = browser.find_element_by_xpath(xpath_place)
        ActionChains(browser).move_to_element(game_name).perform()
        WebDriverWait(browser, 10, 0.5)
        game_name.click()

        # 点击后打开新窗口时窗口数为2
        if len(browser.window_handles) == 2:
            browser.switch_to.window(browser.window_handles[-1])
            jump_url = browser.current_url
            browser.close()
            browser.switch_to.window(browser.window_handles[0])
        #  点击后在原来窗口打开
        else:
            jump_url = browser.current_url
            browser.back()
        return jump_url

    except Exception as e:
        print(e)
        record_log(filename, ' 点击游戏名字时出错，导致获取游戏链接地址出错\n')
        return None



def monitor(browser, url, old_content, clickbroad, monitor_place, filename):
    print("#"*12 + get_time() + ':'+ filename + ' 网站:' + url)
    # 2 查询点击操作
    check_clickbroad(browser, clickbroad, filename)
    if not monitor_place:
        print("网站{}没有需要监控的定位符".format(url))
    else:
        new_content = []
        tmp_place = monitor_place
        for xpath_place,data in tmp_place.items():
            check_clickbroad(browser, clickbroad, filename)
            try:
                new_value = search_ele(browser, url, xpath_place, data, filename).text
            except:
                # record_log(filename, new_value)
                continue

            if data[1] == None:
                msg = "{}:读取标识符{}的游戏名:{} ".format(get_time(),data[0], new_value)
                print(msg)
                record_first_time(filename, msg+'\n')
                monitor_place[xpath_place][1] = new_value
                old_content.add(new_value)
            elif data[1] == new_value:
                # pass
                msg = "{}:网站{}标识符{}没有内容更新".format(get_time(),url, data[0])
                print(msg)
                reacord_no_change(filename, msg)

            else:
                if new_value not in old_content:
                    # 出现新的游戏名， 记录下来
                    old_content.add(new_value)
                    if get_game_detail_url(browser, xpath_place, filename):
                        new_content.append("‘{}’更新为‘{}’，游戏详情页：{}\n".format(data[1], new_value, get_game_detail_url(browser, xpath_place, filename)))
                    else:
                        new_content.append("‘{}’更新为‘{}’，游戏详情页：{}\n".format(data[1], new_value, url))
                        # 网页可能因为点击游戏名字无法跳转导致出错， 重新加载网页
                        Open_url(browser, url)
                        check_clickbroad(browser, clickbroad, filename)
                    monitor_place[xpath_place][1] = new_value
                else:
                    monitor_place[xpath_place][1] = new_value

    return new_content



# clickbroad = {'http://www.yoyou.com/':'//*[@id="c_btn4"]'}
# monitor_place = {'http://www.golue.com/':{'//*[@id="tab-three"]/div[2]/div/ul/li[1]/label/a[2]/span':None, '//*[@id="tab-three"]/div[2]/div/ul/li[2]/label/a[2]/span':None}}
# content_select = {'http://www.golue.com/':'//*[@id="body"]/div[2]/div[1]/div[2]/div[1]/p[1]'}

def monitor_site(browser, old_content, url, clickbroad, monitor_place, filename):
        Open_url(browser, url)
        # 3 读取定位符，检查定位符是否发生改变，发生改变则发送邮件提醒
        content = monitor(browser, url, old_content, clickbroad, monitor_place, filename)
        # 内容整理，发送邮件
        if content:
            # 发生改变，将发生改变的内容通过邮件发送给用户，内容最后附加一个游戏详情的网址
            msg = "网站{} 更新内容：\n{}".format(url, ' '.join(content))
            print(msg)
            reacord_new_game(filename, msg)
            if send_mail("网站{} 更新内容：\n{}".format(url, ' '.join(content))):
                print("邮件发送成功")
            else:
                print("邮件发送失败")

def task(filename):
    try:
        old_content = set()  # 返回更新的游戏之前过滤掉老数据
        cwd = os.getcwd()
        url, clickbroad, monitor_place = get_single_web_data(cwd + '\\' + filename)
        # 1.打开浏览器

        browser = Open_browser()
    except Exception as e:
        print(e)
        record_log(filename, traceback.format_exc())

    while True:
        try:
            monitor_site(browser, old_content, url, clickbroad, monitor_place,filename)
        except Exception as e:
            print(e)
            record_log(filename, traceback.format_exc())
        time.sleep(random.randint(60,100))
        # time.sleep(3)

def get_mail_addr():
    with open('邮箱地址.txt', 'r', encoding='utf8') as f:
        mail_addr = f.readline()
    return mail_addr

#  邮箱发送
def send_mail(content):
    yag = yagmail.SMTP(user='zef_huang@163.com', password='qqdm2020', host='smtp.163.com')
    try:
        yag.send(to=[get_mail_addr()], subject='网站更新提醒', contents=content)
        return True
    except:
        return False

# 定义文件
def create_log_file():
    with open("error_log.txt", 'w') as f:
        pass

    with open("new_game.txt", 'w') as f:
        pass

    with open("no_update.txt", 'w') as f:
        pass

if __name__ == '__main__':
    create_log_file()
    for i in get_xlsx_file():
        if i == '11773.xlsx':
            continue
        threading.Thread(target=task, args=(i,)).start()
    print("启动完成，不用关闭窗口\n")
    # input('不用关闭窗口\n')




