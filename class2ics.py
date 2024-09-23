from lxml import html
import pandas as pd
import time
import re

'''
textutil -convert html example.webarchive
'''


csv_path = "./files/convertcsv.csv"
ics_path = "./output.ics"

csv_UTC = 8
monday_of_first_week = "20240916"
class_to_time = {
    1: {"start": "080000", "end": "084500"},
    2: {"start": "085500", "end": "094000"},
    3: {"start": "100000", "end": "104500"},
    4: {"start": "105500", "end": "114000"},
    5: {"start": "120000", "end": "124500"},
    6: {"start": "125500", "end": "134000"},
    7: {"start": "140000", "end": "144500"},
    8: {"start": "145500", "end": "154000"},
    9: {"start": "160000", "end": "164500"},
    10: {"start": "165500", "end": "174000"},
    11: {"start": "180000", "end": "184500"},
    12: {"start": "185500", "end": "194000"},
    13: {"start": "194000", "end": "202000"},
}


def change_each_class_to_ics(index, class_info):
    # read class info
    summary = class_info["课程名称"] + class_info["课程代码"]
    # description: 任课教师，课程代码，上课时间地点，上课人数：xx人，备注
    description = f'{class_info["任课教师"]}，{class_info["课程代码"]}，{class_info["上课时间地点"]}，上课人数：{class_info["上课人数"]}人，{class_info["备注"]}'

    class_week_start, class_week_end, class_day, class_time_start, class_time_end, class_room = re.findall(r"(\d+)-(\d+)周 星期(.)\[(\d+)-(\d+)节\](.*)", class_info["上课时间地点"])[0]

    ##### fix 新时代中国特色社会主义理论与实践 1-16周
    if class_info["课程代码"] == "MARX6001":
        class_week_start = '1'
        class_week_end = '16'
    ##### end fix

    class_time_start = class_to_time[int(class_time_start)]["start"]
    class_time_end = class_to_time[int(class_time_end)]["end"]
    # UTC, keep 6 digits
    class_time_start = str(int(class_time_start) - csv_UTC * 10000).zfill(6)
    class_time_end = str(int(class_time_end) - csv_UTC * 10000).zfill(6)

    # 汉字转换
    class_day = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "日": 7}[class_day]

    # class_date_start: 从第一周的周一算起，在第class_week_start周的周class_day
    class_date_start = time.mktime(time.strptime(monday_of_first_week, "%Y%m%d")) + 7 * 24 * 3600 * (int(class_week_start) - 1) + 24 * 3600 * (int(class_day) - 1)
    class_date_start = time.strftime("%Y%m%d", time.localtime(class_date_start))

    # class_date_end: 从第一周的周一算起，在第class_week_end周的周日
    class_date_end = time.mktime(time.strptime(monday_of_first_week, "%Y%m%d")) + 7 * 24 * 3600 * (int(class_week_end) - 1) + 24 * 3600 * 6

    # change to ics format
    dtstamp = time.strftime("%Y%m%dT%H%M%SZ", time.localtime())
    uid = "sjtu-class" + str(index)
    location = class_room
    dtstart = class_date_start + "T" + class_time_start + "Z"
    dtend = class_date_start + "T" + class_time_end + "Z"
    until = time.strftime("%Y%m%dT%H%M%SZ", time.localtime(class_date_end))

    ics_str = f"BEGIN:VEVENT\nDTSTAMP:{dtstamp}\nUID:{uid}\nLOCATION:{location}\nDESCRIPTION:{description}\nSUMMARY:{summary}\nDTSTART:{dtstart}\nDTEND:{dtend}\nRRULE:FREQ=WEEKLY;UNTIL={until};INTERVAL=1;\nEND:VEVENT\n\n"
    return ics_str


if __name__ == "__main__":
    html_path = "./input.html"
    csv_path = "./files/convertcsv.csv"

    XPath = '/html/body/main/article/section/div[3]/table'

    # read html
    with open(html_path, "r") as f:
        html_text = f.read()

    # find class table by XPath
    root = html.fromstring(html_text)
    table = root.xpath(XPath)[0]

    # change table to csv, no index
    table_str = html.tostring(table).decode()
    class_list = pd.read_html(table_str)[0]
    # if df columns is number index, use the first row as columns
    if class_list.columns[0] == 0:
        class_list.columns = class_list.iloc[0]
        class_list = class_list.drop(0)  

    ## change to ics format
    # new ics file
    ics_file = open(ics_path, "w")
    # head
    ics_file.write("BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:sjtu课表\n\n")
    # body
    for index, row in class_list.iterrows():
        ics_file.write(change_each_class_to_ics(index, row))
    # tail
    ics_file.write("END:VCALENDAR\n")
    ics_file.close()
    print("CSV file has been converted to ics file.")

