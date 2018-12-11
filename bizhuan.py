import requests
import json
import xlwt



def base_request(url):

    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.110 Safari/537.36'}
    response = requests.get(url, headers=headers)
    return response

def first_request():
    url1 = 'https://pages.kaola.com/pages/region/detail/8878/30001,1005,1034/159925,155647,217513.shtml?'
    response = base_request(url1)
    json_dict = json.loads(response.text)
    item_list = []

    data_list1 = json_dict['data'][1]['businessObj']['list']
    for data in data_list1:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)

    data_list2 = json_dict['data'][2]['businessObj']['list']
    for data in data_list2:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        # item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['groupBuyPrice']
        item['topTextTag'] = data['content']['activityTag']
        item_list.append(item)
    return item_list

def second_request():
    url2 = 'https://pages.kaola.com/pages/region/detail/8878/10001,1005,1005/220491,159922,162393.html?'

    response = base_request(url2)
    json_dict = json.loads(response.text)
    item_list = []

    data_list1 = json_dict['data'][1]['businessObj']['list']
    for data in data_list1:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)

    data_list2 = json_dict['data'][2]['businessObj']['list']
    for data in data_list2:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)
    return item_list

def third_request():
    url3 = 'https://pages.kaola.com/pages/region/detail/8878/1005,1005,1005/165676,162394,162395.html?'

    response = base_request(url3)
    json_dict = json.loads(response.text)
    item_list = []

    data_list0 = json_dict['data'][0]['businessObj']['list']
    for data in data_list0:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)

    data_list1 = json_dict['data'][1]['businessObj']['list']
    for data in data_list1:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)

    data_list2 = json_dict['data'][2]['businessObj']['list']
    for data in data_list2:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)
    return item_list

def forth_request():
    url4 = 'https://pages.kaola.com/pages/region/detail/8878/1005,1005,1005/252093,162392,188855.html?'

    response = base_request(url4)
    json_dict = json.loads(response.text)
    item_list = []

    data_list0 = json_dict['data'][0]['businessObj']['list']
    for data in data_list0:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)

    data_list1 = json_dict['data'][1]['businessObj']['list']
    for data in data_list1:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)

    data_list2 = json_dict['data'][2]['businessObj']['list']
    for data in data_list2:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)
    return item_list

def five_request():
    url5 = 'https://pages.kaola.com/pages/region/detail/8878/1005,1005,1005/232119,253968,203820.html?'

    response = base_request(url5)
    json_dict = json.loads(response.text)
    item_list = []

    data_list0 = json_dict['data'][0]['businessObj']['list']
    for data in data_list0:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)

    data_list1 = json_dict['data'][1]['businessObj']['list']
    for data in data_list1:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)

    data_list2 = json_dict['data'][2]['businessObj']['list']
    for data in data_list2:
        item = {}
        item['imageUrl'] = data['content']['imageUrl']
        item['introduce'] = data['content']['introduce']
        item['title'] = data['content']['title']
        item['actualCurrentPrice'] = data['content']['actualCurrentPrice']
        item['topTextTag'] = data['content']['goodsConfigMap']['topTextTag']
        item_list.append(item)
    return item_list

def to_xml(item):
    # 创建excel工作表
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('sheet1')

    # 设置表头
    worksheet.write(0, 0, label='imageUrl')
    worksheet.write(0, 1, label='introduce')
    worksheet.write(0, 2, label='title')
    worksheet.write(0, 3, label='actualCurrentPrice')
    worksheet.write(0, 4, label='topTextTag')

    # 将json字典写入excel
    # 变量用来循环时控制写入单元格，感觉有更好的表达方式
    val1 = 1
    val2 = 1
    val3 = 1
    val4 = 1
    val5 = 1
    for list_item in item:
        for key, value in list_item.items():
            if key == "imageUrl":
                worksheet.write(val1, 0, value)
                val1 += 1
            elif key == "introduce":
                worksheet.write(val2, 1, value)
                val2 += 1
            elif key == "title":
                worksheet.write(val3, 2, value)
                val3 += 1
            elif key == "actualCurrentPrice":
                worksheet.write(val4, 3, value)
                val4 += 1
            elif key == "topTextTag":
                worksheet.write(val5, 4, value)
                val5 += 1
            else:
                pass

    # 保存
    workbook.save('OK.xls')


if __name__ == '__main__':

    first = first_request()
    second = second_request()
    third = third_request()
    forth = forth_request()
    five = five_request()
    item = first+second+third+forth+five
    to_xml(item)




