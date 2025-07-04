from pandas import read_html
from requests import Session
from ddddocr import DdddOcr

session = Session()
login_url = "http://192.168.16.209/loginAction.do"

headers = {
    'User-Agent': 'Mozilla/5.0',
    'Referer': login_url
}

vchart_link = "http://192.168.16.209/validateCodeAction.do?"  # 示例路径，需根据实际页面调整
vchart_response = session.get(vchart_link, headers=headers)
ocr = DdddOcr(show_ad=False, use_gpu=False)
result = ocr.classification(vchart_response.content)

payload = {
    'zjh': '2215113116',
    'mm': 'light004',
    'v_yzm': f'{result}',
}

response = session.post(login_url, data=payload, headers=headers)
grades = session.get("http://192.168.16.209/gradeLnAllAction.do?type=ln&oper=qbinfo")
credits = session.get("http://192.168.16.209/gradeLnAllAction.do?oper=queryXfjd")

gradesTable = read_html(grades.text)   
creditsTable = read_html(credits.text)

print("-" * 80)
print("{:\u3000<22}\t{:\u3000<5}\t{:\u3000<5}".format("课程名","学分","成绩"))

for index in range(10, len(gradesTable), 6):
    print("-" * 80)
    tmp_frame = gradesTable[index].iloc[:, [2, 4, 5, 7]]
    tmp_frame.columns = ['课程名', '学分', '课程属性', '成绩']  # 给列命名（你可改成你自己喜欢的顺序）

    for i, row in tmp_frame.iterrows():
        print("{:\u3000<22}\t{:\u3000<5}\t\t{:\u3000<5}".format(row['课程名'], row['学分'] if len(str(row['学分'])) == 3 else str(row['学分']) + ".0", row['成绩']))
        
print("-" * 80)
print("{:\u3000<13}\t{:\u3000^5}\t{:\u3000^5}\t{:\u3000^5}".format("学年学期","学分绩点","学位绩点", "加权学分学位绩点"))
print("-" * 80)
for i, row in creditsTable[11].iterrows():
    print("{:\u3000<17}\t\t{:\u3000<5}\t\t{:\u3000<5}\t\t{:\u3000<5}".format(row['学年学期'], row['学分绩点'], row['学位绩点'], row['加权学分学位绩点']))
print("-" * 80)

input()
