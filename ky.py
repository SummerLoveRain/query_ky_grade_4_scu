import os
import re

import requests
import xlsxwriter
from PIL import Image
from aip import AipOcr
from bs4 import BeautifulSoup

VALICODE_URL = "https://yz.scu.edu.cn/User/Valicdoe?t=5870"
QUERY_URL = "https://yz.scu.edu.cn/score/Query/--"
FIRST_URL = "https://yz.scu.edu.cn/score"
IMG_PATH = "./image/"
VCODE_IMG_NAME = "img.png"
TMP_IMG_NAME = "img.gif"
QUERY_FILE = "file.txt"
headers = {
    'Content-Type': 'application/x-www-form-urlencoded',
    'Referer': 'https://yz.scu.edu.cn/score',
    'Cookie': 'ASP.NET_SessionId=gp02wsi31lzs0knkqtanleks;Hm_lvt_7c1c028cd7cd078d7d3e3db9b7b913c5=1587867242,1588211931,1588299478,1588410791;Hm_lpvt_7c1c028cd7cd078d7d3e3db9b7b913c5 = 1588421130',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.122 Safari/537.36',
}

def first_page():
    print(req.get(FIRST_URL,headers=headers))


def query(ksbh, xm, zjhm, vcode):
    form_data = {
        'ksbh': ksbh,
        'xm': xm,
        'zjhm': zjhm,
        'vcode': vcode
    }
    print("查询成绩...")
    response = req.post(QUERY_URL, headers=headers, data=form_data)
    res = response.text
    print(res)
    print("查询完成")
    return res


def parse_data(res):
    print("解析成绩...")
    soup = BeautifulSoup(res, 'html.parser')
    (ksxm_div, ksbh_div, bkzy_div) = soup.findAll(name="div", attrs={"class": "display-field"})

    ksxm = ksxm_div.get_text().strip().replace(" ", "").replace("考生姓名：", "")
    ksbh = ksbh_div.get_text().strip().replace(" ", "").replace("考生编号：", "")
    bkzy = bkzy_div.get_text().strip().replace(" ", "").replace("报考专业：", "")

    (class1_tr, class2_tr, class3_tr, class4_tr, sum_tr) = soup.findAll(name="tr")

    (tmp_td, class1_td, class1_grade_td) = class1_tr.findAll(name="td")
    class1 = class1_td.get_text().strip().replace(" ", "")
    class1_grade = class1_grade_td.get_text().strip().replace(" ", "")

    (tmp_td, class2_td, class2_grade_td) = class2_tr.findAll(name="td")
    class2 = class2_td.get_text().strip().replace(" ", "")
    class2_grade = class2_grade_td.get_text().strip().replace(" ", "")

    (tmp_td, class3_td, class3_grade_td) = class3_tr.findAll(name="td")
    class3 = class3_td.get_text().strip().replace(" ", "")
    class3_grade = class3_grade_td.get_text().strip().replace(" ", "")

    (tmp_td, class4_td, class4_grade_td) = class4_tr.findAll(name="td")
    class4 = class4_td.get_text().strip().replace(" ", "")
    class4_grade = class4_grade_td.get_text().strip().replace(" ", "")

    (tmp_td, sum_grade_td) = sum_tr.findAll(name="td")
    sum_grade = sum_grade_td.get_text().strip().replace(" ", "")

    print("解析完成")
    return [ksxm, ksbh, bkzy, class1, class1_grade, class2, class2_grade, class3, class3_grade, class4, class4_grade,
            sum_grade]


def down_valicode():
    os.makedirs(IMG_PATH, exist_ok=True)
    try:
        print('下载验证码')
        r = requests.get(VALICODE_URL,headers=headers)
        with open(IMG_PATH + TMP_IMG_NAME, 'wb') as f:
            f.write(r.content)

    except:
        print('验证码下载失败')


def processImage():
    im = Image.open(IMG_PATH + TMP_IMG_NAME)
    mypalette = im.getpalette()
    i = 0
    try:
        while 1:
            im.putpalette(mypalette)
            new_im = Image.new("RGBA", im.size)
            new_im.paste(im)
            new_im.save(IMG_PATH + VCODE_IMG_NAME)

            i += 1
            im.seek(im.tell() + 1)

        print('下载完成')
    except EOFError:
        pass  # end of sequence


""" 你的 APPID AK SK """
APP_ID = ''
API_KEY = ''
SECRET_KEY = ''
client = AipOcr(APP_ID, API_KEY, SECRET_KEY)


def check_valicode():
    print("识别验证码...")
    # """ 调用通用文字识别, 图片参数为本地图片 """
    # """ 可选参数 """
    # options = {}
    # options["language_type"] = "CHN_ENG"  # 中英文混合
    # options["detect_direction"] = "true"  # 检测朝向
    # options["detect_language"] = "true"  # 是否检测语言
    # options["probability"] = "false"  # 是否返回识别结果中每一行的置信度

    image = get_file_content(IMG_PATH + VCODE_IMG_NAME)
    print(image)
    """ 如果有可选参数 """
    options = {}
    options["detect_direction"] = "true"
    options["probability"] = "true"
    options["language_type"] = "ENG"

    """ 带参数调用通用文字识别（高精度版） """
    result = client.basicAccurate(image, options)
    data = str(result)
    # 定义正则，提取验证码识别的内容
    pat = re.compile(r"\'words\': \'(.*?)\'")
    print(data)
    result = pat.findall(data)
    print(result)
    print("验证码识别完成")
    if (len(result) > 0):
        return result[0].replace(" ", "")
    else:
        return ""


def get_file_content(filePath):
    with open(filePath, 'rb') as fp:
        return fp.read()


if __name__ == "__main__":
    f = open(QUERY_FILE, 'r', encoding='UTF-8', closefd=True)
    req = requests.session()
    workbook = xlsxwriter.Workbook('grade.xlsx')  # 创建一个Excel文件
    worksheet = workbook.add_worksheet()  # 创建一个sheet
    title = [U'考试姓名', U'考试编号', U'报考专业', U'课程1', U'成绩', U'课程2', U'成绩', U'课程3', U'成绩', U'课程4', U'成绩', U'总分']  # 表格title
    worksheet.write_row('A1', title)  # title 写入Excel

    num = 1;
    while True:
        line = f.readline()
        if not line:
            break

        if line.isspace():
            continue

        [ksbh, xm] = line.split(' ', 1)
        ksbh = ksbh.strip().replace(" ", "")
        xm = xm.strip().replace(" ", "")
        print("第" + str(num) + "行数据")
        print("---" + ksbh + "---" + xm + "---")
        zjhm = ""
        first_page()
        while 1:
            down_valicode()
            processImage()
            vcode = check_valicode()
            print(vcode)

            res = query(ksbh, xm, zjhm, vcode)
            if res.__contains__("校验码错误或失效！") or (len(res) == 0):
                print("校验码错误或失效！")
                continue
            else:
                data = parse_data(res)
                num += 1
                row = 'A' + str(num)
                worksheet.write_row(row, data)
                break
    f.close()
    workbook.close()
