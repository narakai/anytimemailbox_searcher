import time

import requests
import openpyxl
from bs4 import BeautifulSoup
import re

def resident_query():
    global city, street_address, zip_code, price, title
    # us_states = [
    #     'Alabama'
    # ]
    us_states = [
        'Alabama', 'Alaska', 'Arizona', 'Arkansas', 'California', 'Colorado', 'Connecticut', 'DC', 'Delaware',
        'Florida',
        'Georgia', 'Hawaii', 'Idaho', 'Illinois', 'Indiana', 'Iowa', 'Kansas', 'Kentucky', 'Louisiana', 'Maine',
        'Maryland',
        'Massachusetts', 'Michigan', 'Minnesota', 'Mississippi', 'Missouri', 'Montana', 'Nebraska', 'Nevada',
        'New-Hampshire',
        'New-Jersey', 'New-Mexico', 'New-York', 'North-Carolina', 'North-Dakota', 'Ohio', 'Oklahoma', 'Oregon',
        'Pennsylvania',
        'South-Carolina', 'South-Dakota', 'Tennessee', 'Texas', 'Utah', 'Vermont', 'Virginia', 'Washington',
        'West-Virginia', 'Wisconsin', 'Wyoming'
    ]

    # 创建一个Excel工作簿
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # 添加Excel表头
    headers = ["price", "title", "cmar", "addressLine1", "city", "state", "zip5", "zip4", "carrierRoute", "countyName"]
    sheet.append(headers)

    for state in us_states:
        url = 'https://www.anytimemailbox.com/l/usa/' + state
        # 发送HTTP GET请求并获取页面内容
        response = requests.get(url)
        # 检查请求是否成功
        if response.status_code == 200:
            # 使用Beautiful Soup解析HTML
            soup = BeautifulSoup(response.text, 'html.parser')

            # 选择所有的theme-location-item元素
            location_items = soup.find_all('div', class_='theme-location-item')

            # 创建一个列表来存储结果
            results = []

            # 遍历每个theme-location-item元素
            for item in location_items:
                # 获取t-title的值
                title = item.find('h3', class_='t-title').text.strip()

                # 获取t-price的值
                price = item.find('div', class_='t-price').b.text.strip()

                # 使用正则表达式匹配数字部分
                matches = re.findall(r'\d+\.\d+', price)
                # 提取的数字
                if matches:
                    extracted_number = float(matches[0])
                    price = extracted_number
                    print(extracted_number)
                else:
                    print('未找到数字')

                # 找到<div class="t-addr">元素
                t_addr_div = item.find('div', class_='t-addr')

                # 提取文本内容并分割
                address_text = t_addr_div.get_text(separator='|').strip()
                address_parts = address_text.split('|')

                if len(address_parts) >= 3:
                    street_address = address_parts[0].strip()
                    city_and_state = address_parts[1].strip()

                    # 进一步分割City和State
                    city, state_zip = city_and_state.split(',')

                    # 进一步分割State和Zip Code
                    state, zip_code = state_zip.strip().split()

                    print('Street Address:', street_address)
                    print('City:', city.strip())
                    print('State:', state)
                    print('Zip Code:', zip_code)
                else:
                    print('Address format is not as expected.')

                # 将结果存储为字典
                location_info = {
                    't-title': title,
                    't-price': price,
                    'Street Address': street_address,
                    'City': city,
                    'State': state,
                    'Zip': zip_code
                }

                # 将结果添加到列表中
                results.append(location_info)

            # 打印抓取的信息
            uspsUrl = "https://tools.usps.com/tools/app/ziplookup/zipByAddress"
            for result in results:
                print(result)
                payload = {
                    "companyName": "",
                    "address1": result['Street Address'],
                    "address2": "",
                    "city": result['City'],
                    "state": result['State'],
                    "urbanCode": "",
                    "zip": result['Zip'],
                }

                # 自定义头部
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36',
                    # 设置用户代理
                    'Referer': 'https://tools.usps.com/zip-code-lookup.htm?byaddress',
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Origin': 'https://tools.usps.com',
                }

                # 发送POST请求
                response = requests.post(uspsUrl, data=payload, headers=headers)

                # 获取响应内容
                response_json = response.json()

                # 获取data字段中的数据
                data = response_json["addressList"]

                print(data)

                for state_date in data:
                    # 提取data中的数据
                    row_data = [
                        result["t-price"],
                        result["t-title"],
                        state_date["cmar"],
                        state_date["addressLine1"],
                        state_date["city"],
                        state_date["state"],
                        state_date["zip5"],
                        state_date["zip4"],
                        state_date["carrierRoute"],
                        state_date["countyName"],
                    ]

                    # 将数据写入Excel
                    sheet.append(row_data)
                    time.sleep(0.1)

        else:
            print('Failed to retrieve the web page. Status code:', response.status_code)

    # 保存Excel文件
    workbook.save("address.xlsx")


if __name__ == '__main__':
    resident_query()