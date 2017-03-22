from urllib import request
from urllib import error
import json

from xlwt import Workbook

def getInfoList(response_result):
    dict_result = json.loads(response_result)
    return dict_result['resultValue']['items']


def crawl(pageIndex, cookie):
    # 请求地址
    url = 'http://10.4.44.245:8088/WebContent/s6000/rest/esEvent/'
    # 构造post数据
    # data = json.dumps({'pageIndex': pageIndex, 'pageSize': 10,"contj":"211.160.252.2","Login_Result":"成功"}).encode('utf-8')
    data = json.dumps({'pageIndex': pageIndex, 'pageSize': 10, "contj": "日志审计", "Device_Name": "安恒日志审计系统"}).encode(
        'utf-8')
    try:
        # 构造Header
        req = request.Request(url=url, data=data)

        req.add_header('Accept', 'application/json, text/javascript, */*; q=0.01')
        req.add_header('Accept-Encoding', 'gzip, deflate')
        req.add_header('Accept-Language', 'zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3')
        req.add_header('User-Agent', 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:49.0) Gecko/20100101 Firefox/49.0')
        req.add_header('X-Requested-With', 'XMLHttpRequest')
        req.add_header('Cookie', cookie)
        req.add_header('Content-Type', 'application/json')
        req.add_header('Referer', 'http://10.4.44.245:8088/WebContent/s6000/main/index.jsp')
        req.add_header('Connection', 'keep-alive')

        response = request.urlopen(req)

        return response.read().decode('utf-8')

    except error.URLError as e:
        print('爬取第%s页失败，重新爬取' % (pageIndex))
        crawl(pageIndex)


if __name__ == '__main__':
    # 输入cookie
    cookie = input("输入cookie")
    if cookie == '':
        cookie = 'JSESSIONID=1E85D84FF2E3DEB1D193ADD17092D64D'

    # 设置爬取页数
    num = input('输入爬取页数')
    if num == '':
        num = '100'

    # 设置excel的title
    # title = ['Operation', 'Dest_Place', 'Source_Country', 'CE_ID', 'LE_ID', 'Dest_MAC', 'RevTimeHour', 'Log_Type',
    #          'Device_Type', 'Dest_Network', 'Source_City', 'Repeat_Count', 'Event_Type', 'SN', 'Original_Log',
    #          'RevTimeDay', 'Source_Port', 'mainid', 'Source_Province', 'Dest_ID', 'Method', 'Application_System',
    #          'Affiliated_Network', 'DT_ID', 'Device_Name', 'RevIP', 'Device_Subtype', 'RevTime', 'Dest_Province',
    #          'Event_Name', 'SendID', 'Dest_Name', 'Packet', 'Dest_Country', 'Unified_Class', 'URL', 'Dest_Port',
    #          'Log_Detail', 'Event_Times', 'Affiliated_Unit', 'Protocol', 'Dest_IP', 'Dest_System', 'esRevTime',
    #          'Source_IP', 'SendIP', 'Dest_City', 'Source_MAC', 'Log_Level', 'Event_Time', 'Dest_Unit']
    title = ['Source_IP', 'Source_Port', 'Dest_IP', 'Dest_Port', 'Event_Name', 'Event_Time', 'Event_Type', 'LE_ID',
             'Log_Level', 'URL', 'Packet', 'Method', 'Original_Log', 'domain', 'uri', 'stat_time', 'policy_id',
             'rule_id', 'action', 'block', 'HTTP', 'user-agent', 'cookie', 'application', 'orglist', 'orglistNext',
             'Content-Type', 'rcolumn_code', 'alertinfo', 'proxy_info', 'characters', 'Dest_Country', 'Dest_City',
             'Dest_Place', 'Dest_Network']

    # long变量，表示每一个sheet的长度，long为6000，表示6000行为一个sheet
    long=100

    # 遍历sheet
    for n in range(int(int(num)/long)):# 爬取的页数/每个excel的行数=要生成的excel表数
        # 遍历从0开始，excel名字以1开始
        book_page = n + 1
        # excel表
        book = Workbook('%s.xlsx'%(book_page))

        # 添加sheet
        sheet = book.add_sheet('sheet %s' %(book_page))

        # 第一行为title
        title_row = sheet.row(0)
        for i in range(len(title)):
            title_row.write(i, title[i])

        # 表示行索引的变量
        row_index = 1

        # 遍历每一个excel的数据量
        for num_index in range(long):
            # 要爬取的页数
            page=long*n+num_index + 1
            print("爬取第%s页数据" % (page))

            # try:
            #
            # except error as e:
            #     print("爬取第%s页数据时停止" % (page))
            # finally:
            # 爬取数据
            response_result = crawl(page, cookie)

            # 返回信息列表
            info_list = getInfoList(response_result)
            # print(info_list)

            # 遍历写入excel
            for info in info_list:  # [[],[]...]-->{}
                row = sheet.row(row_index)
                row_index = row_index + 1
                for i in range(len(title)):
                    if title[i] in info.keys():
                        row.write(i, str(info[title[i]]))
                    else:
                        row.write(i, '')

        # 保存excel
        book.save('%s.xlsx' % (book_page))
        print("保存第%s.xlsx文件" % (book_page))

