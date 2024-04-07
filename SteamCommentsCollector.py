'''
Author: 阿朝
Date: 2023/8/25
'''

import tkinter as tk
from tkinter import messagebox
import requests
from bs4 import BeautifulSoup
import xlwt
import re


def steam_review_spider(store_link, file_name, comment_count):
    '''这个函数是核心爬虫逻辑'''
    headers = {    'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3'} #这一部分处理评论语言的问题

    game_link = store_link
    game_name = file_name
    game_id = re.search(r"https://store.steampowered.com/app/(\d+)/", game_link).group(1) #获取游戏ID
    comment_number = int(comment_count)

    review_content = [] #初始化游戏测评内容的二维数组

    if comment_number % 10 == 0:
        comment_page = comment_number // 10 #一页10条
    else:
        comment_page = comment_number // 10 + 1 #还多一页
    for i in range(1, comment_page + 1):
        #根据游戏ID生成游戏评论页面的位置
        url = 'http://steamcommunity.com/app/' + game_id +'/homecontent/?userreviewsoffset=' + str(10 * (i - 1)) + '&p=' + str(
            i) + '&workshopitemspage=' + str(i) + '&readytouseitemspage=' + str(i) + '&mtxitemspage=' + str(
            i) + '&itemspage=' + str(i) + '&screenshotspage=' + str(i) + '&videospage=' + str(i) + '&artpage=' + str(
            i) + '&allguidepage=' + str(i) + '&webguidepage=' + str(i) + '&integratedguidepage=' + str(
            i) + '&discussionspage=' + str(
            i) + '&numperpage=10&browsefilter=toprated&browsefilter=toprated&appid=433850&appHubSubSection=10&l=schinese&filterLanguage=default&searchText=&forceanon=1'
        html = requests.get(url, headers=headers).text #爬取网页
        soup = BeautifulSoup(html, 'html.parser')
        reviews = soup.find_all('div', {'class': 'apphub_Card'})    
        for review in reviews:
            cell = []

            #解析评论
            nick = review.find('div', {'class': 'apphub_CardContentAuthorName'}) #评论ID
            title = review.find('div', {'class': 'title'}).text #推荐/不推荐
            hour = review.find('div', {'class': 'hours'}).text.split(' ')[1] #游戏时长
            link = nick.find('a').attrs['href'] #评论链接
            comment = review.find('div', {'class': 'apphub_CardTextContent'}).text.split('\n')[2].strip('\t') #评论

            cell.append(nick.text)
            cell.append(title)
            cell.append(hour)
            cell.append(link)
            cell.append(comment) #一个人的评论信息

            review_content.append(cell)

    book_name_xls = game_name+'_评论.xls' #文件保存到当前目录
    workbook = xlwt.Workbook()  # 新建一个表格
    #建立两个工作页
    sheet_p = workbook.add_sheet("好评")
    sheet_n = workbook.add_sheet("差评")
    #初始化工作表
    sheet_p.write(0,0,"ID")
    sheet_p.write(0,1,"推荐/不推荐")
    sheet_p.write(0,2,"游戏时长")
    sheet_p.write(0,3,"链接")
    sheet_p.write(0,4,"评论")
    sheet_n.write(0,0,"ID")
    sheet_n.write(0,1,"推荐/不推荐")
    sheet_n.write(0,2,"游戏时长")
    sheet_n.write(0,3,"链接")
    sheet_n.write(0,4,"评论")
    index = len(review_content)  # 获取需要写入数据的行数
    line_p = 1
    line_n = 1 #好评与差评的行数
    for i in range(0, index):
        for j in range(0, len(review_content[i])):
            if review_content[i][1] == "推荐":
                sheet_p.write(line_p, j, review_content[i][j])  #写入好评
            else:
                sheet_n.write(line_n, j, review_content[i][j])  #写入差评
        if review_content[i][1] == "推荐":
            line_p += 1
        else:
            line_n += 1
    workbook.save(book_name_xls) #写入并保存表格
    messagebox.showinfo("结果", "完成咯！")


# 创建主窗口
root = tk.Tk()
root.title("阿朝的Steam评论爬取工具")
root.geometry("320x240")

# 创建标签和文本框
label1 = tk.Label(root, text="商店链接：")
entry1 = tk.Entry(root)

label2 = tk.Label(root, text="游戏名称：")
entry2 = tk.Entry(root)

label3 = tk.Label(root, text="评论数量：")
entry3 = tk.Entry(root)

note_label = tk.Label(root, text="注意：请勿短时间内频繁爬取！")

# 创建处理按钮
def handle_button():
    a = entry1.get()
    b = entry2.get()
    c = entry3.get()
    steam_review_spider(a, b, c)

button = tk.Button(root, text="开始爬取", command=handle_button)

# 布局界面元素
label1.pack()
entry1.pack()

label2.pack()
entry2.pack()

label3.pack()
entry3.pack()

note_label.pack()

button.pack()

# 启动主循环
root.mainloop()