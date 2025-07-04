from sys import argv
import requests, base64
from os.path import exists
from ddddocr import DdddOcr
from time import sleep, time
from PyQt5.QtGui import QIcon
from selenium.webdriver import Edge
from pandas import read_excel, read_html
from selenium.webdriver.common.by import By
from PyQt5.QtCore import Qt, pyqtSignal, QThread
from selenium.webdriver.edge.options import Options
from PyQt5.QtWidgets import QWidget, QMainWindow, QLabel, QAction, QApplication, QHBoxLayout, QLineEdit, QDialog, QVBoxLayout, QPushButton, QCheckBox, QMenuBar, QTextEdit, QComboBox

class urp_tools:
    
    def __init__(self, zh, mm, mode="--headless", link="http://192.168.16.209/"):
        
        self.zh = zh
        self.mm = mm
        self.link = link
        self.mode = mode
        options = Options()
        options.add_argument(self.mode) if self.mode == "--headless" else None
        options.add_argument('log-level=3')
        options.add_argument('--mute-audio')
        options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
        
        self.driver = Edge(options=options)
        self.driver.set_window_size(800, 800)
        self.driver.get(self.link)
        
    def recognize_yzm(self):
        
        img_bytes = self.driver.find_element(By.XPATH, '//*[@id="vchart"]').screenshot_as_png
        ocr = DdddOcr(show_ad=False, use_gpu=False)
        key = ocr.classification(img_bytes)
        if "o" in key:
            key = key.replace("o",'0')
        
        return key

    def login(self):
           
        yzm = self.recognize_yzm()
       
        self.driver.find_elements(By.CLASS_NAME,"input01")[0].send_keys(self.zh)
        self.driver.find_elements(By.CLASS_NAME,"input01")[1].send_keys(self.mm)
        self.driver.execute_script(
        f"document.loginForm.v_yzm.value = '{yzm}';"
        "document.getElementsByClassName('buttonImg')[0].click();"
        )
        while True:
            try:
                if self.driver.find_element(By.CLASS_NAME,"errorTop").text == "你输入的验证码错误，请您重新输入！":
                    print("验证码识别错误，请重试")
                    sleep(1)
                    yzm = self.recognize_yzm()
                    self.driver.find_elements(By.CLASS_NAME,"input01")[0].send_keys(self.zh)
                    self.driver.find_elements(By.CLASS_NAME,"input01")[1].send_keys(self.mm)
                    self.driver.execute_script(
                    f"document.loginForm.v_yzm.value = '{yzm}';"
                    "document.getElementsByClassName('buttonImg')[0].click();"
                    )
                else: break
            except: break

    def offline_preprocess(self):

        self.driver.find_element(By.ID,"details-button").click()
        self.driver.find_element(By.ID,"proceed-link").click()

        self.driver.find_element(By.NAME,"username").send_keys("2215113116")
        self.driver.find_element(By.NAME,"password").send_keys("Ldb20040716@")
        self.driver.find_element(By.CLASS_NAME,"submit").click()
        
        sleep(1)
        try:
            self.driver.switch_to.alert.dismiss()
            self.driver.switch_to.alert.accept()
        except: pass
        self.driver.get("https://223.112.21.198:6443/7b68f983/")
        sleep(0.5)
        self.driver.get("https://223.112.21.198:6443/7b68f983/")
    def evaluation(self):
        
        return_text = '--------------------------------------------\n'
        self.driver.get(f"{self.link}/jxpgXsAction.do?oper=listWj&yzxh={self.zh}")
        try:
            a = self.driver.find_elements(By.CLASS_NAME,"even")
            b = self.driver.find_elements(By.CLASS_NAME,"odd")
            return_text += f"📚 一共需要评估{len(a + b)}门课\n"
            n = 0
            for i in range(len(a) + len(b)):
                if len(b) > 0:
                    if b[i].text[-1] == "是":
                        n += 1    
                elif len(a) > 0:
                    if a[i].text[-1] == "是":
                        n += 1
            if n == len(a) + len(b) and n != 0 :
                # print("评估已完成。")
                return_text += "🎓 评估已完成，干得漂亮！\n--------------------------------------------"
            else:
                A = []
                for i in range(0, int(len(a) + len(b)+1)):
                    self.driver.execute_script(f"window.open('{self.link}jxpgXsAction.do?oper=listWj&yzxh={self.zh}', '_blank');")
                    A.append(self.driver.window_handles)
                sleep(3)
                for k in range(0, int(len(a) + len(b)+1)):
                    try:
                        self.driver.switch_to.window(self.driver.window_handles[k])
                        ll = self.driver.find_elements(By.TAG_NAME,"img")
                        ll[k].click()
                        sleep(1)
                        oo = self.driver.find_elements(By.ID,"2")
                        for j in oo:
                            j.click()
                        self.driver.find_elements(By.ID,"1")[-1].click()
                        self.driver.execute_script(
                            "function testbutt() \
                            {testbutt1();}"
                            "testbutt();"
                            )
                        self.driver.find_element(By.ID,"submit1").click()
                        sleep(0.8)
                        try:
                            self.driver.switch_to.alert.dismiss()
                            self.driver.switch_to.alert.dismiss()
                        except: pass
                        # self.driver.close()
                    except:
                        pass
        except Exception as e:
            print(e)
            print("评估失败，请重试。")
        
        self.driver.quit()
        return return_text

class SearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("查询学生信息")
        self.setGeometry(700, 250, 500, 50)
        self.setModal(False)  # 设置为模态窗口
        self.init_ui()

    def init_ui(self):
        layout = QHBoxLayout(self)

        # 学号输入
        self.student_id_input = QLineEdit()
        self.student_id_input.setPlaceholderText("请输入姓名")

        self.check_button = QCheckBox("自动填入")
        self.check_button.setChecked(True)

        # 查询按钮
        btn_layout = QHBoxLayout()
        self.ok_button = QPushButton("确认")
        self.cancel_button = QPushButton("取消")

        btn_layout.addStretch()
        btn_layout.addWidget(self.ok_button)
        btn_layout.addWidget(self.cancel_button)

        layout.addWidget(self.student_id_input)
        layout.addWidget(self.check_button)
        layout.addLayout(btn_layout)

        # 绑定按钮事件
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

    def get_student_id(self):
        return self.student_id_input.text()
    
    def get_check(self):
        return self.check_button.isChecked()
    
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("紫金URP查询助手")
        self.setGeometry(700, 350, 470, 300)
        self.setWindowIcon(QIcon(r"C:\Users\Administrator\Pictures\icon\high@3x.ico"))
        self.setStyleSheet("""
        QMainWindow {
         background-color: rgba(255, 255, 255, .9);   
        }
        """)
        # 主窗口
        content = QWidget()
        layout = QVBoxLayout()
        content.setLayout(layout)
        self.setCentralWidget(content)

        self.menu_bar = QMenuBar()
        find_people_menu = self.menu_bar.addMenu("👀更多功能")
        action1 = QAction("❤获取所有人姓名", self)
        action2 = QAction("🧑‍🤝‍🧑获取指定人所有信息", self)
        action3 = QAction("🚮清空", self)
        action4 = QAction("🚪 退出", self)

        find_people_menu.addAction(action1)
        action1.triggered.connect(self.get_all_name)
        find_people_menu.addAction(action2)
        action2.triggered.connect(self.get_specific_information)
        find_people_menu.addAction(action3)
        action3.triggered.connect(self.clear)
        find_people_menu.addAction(action4)
        action4.triggered.connect(self.close)

        custom_peole_menu = self.menu_bar.addMenu("🧑常查询人员")
        people1  = QAction("曾政", self)
        people2  = QAction("张嘉奇", self)
        people3  = QAction("温炳兴", self)
        people4  = QAction("宋益凡", self)
        custom_peole_menu.addActions([people1, people2, people3, people4])
        layout.addWidget(self.menu_bar)


        # 下半窗口
        self.sub_layout = QHBoxLayout()
        layout.addLayout(self.sub_layout)

        # 下半左侧
        self.left_layout = QVBoxLayout()
        self.sub_layout.addLayout(self.left_layout)

        # 下半右侧
        self.right_layout = QVBoxLayout()
        self.sub_layout.addLayout(self.right_layout)
        layout.addStretch() 

        self._init_left_content()
        from functools import partial

        for people in [people1, people2, people3, people4]:
            people.triggered.connect(partial(self.get_people_information, people.text()))

    def _init_left_content(self):
        
        username_layout = QHBoxLayout()
        username_layout.setAlignment(Qt.AlignLeft)
        self.username_label = QLabel("学号：")
        self.username_label.setStyleSheet("font-size: 20px; font-family: '华文中宋';")

        self.username_input = QLineEdit()
        self.username_input.setStyleSheet(
            """
            font-size: 20px; font-family: '宋体'; border:none; outline: none;
            border-bottom: 2px solid red; border-radius: 2px;
            """
        )
        self.username_input.setPlaceholderText("请输入学号")
        self.username_input.setClearButtonEnabled(True)

        username_layout.addWidget(self.username_label)
        username_layout.addWidget(self.username_input)

        password_layout = QHBoxLayout()
        self.password_label = QLabel("密码：")
        self.password_label.setStyleSheet("font-size: 20px; font-family: '华文中宋';")

        self.password_input = QLineEdit()
        self.password_input.setStyleSheet(
            """
            font-size: 12px; font-family: '华文中宋'; border:none; outline: none;
            border-bottom: 2px solid red; border-radius: 2px; letter-spacing: 3px;
            """
        )   
        self.password_input.setPlaceholderText("请输入密码")
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setClearButtonEnabled(True)

        password_layout.addWidget(self.password_label)
        password_layout.addWidget(self.password_input)

        headless_layout = QHBoxLayout()
        self.headless_combobox = QComboBox()
        self.headless_combobox.setMinimumWidth(390)
        self.headless_combobox.setStyleSheet("font-size: 15px;")
        self.headless_combobox.addItem("无头模式")
        self.headless_combobox.addItem("有头模式")
        headless_layout.addWidget(self.headless_combobox)

        self.innerNet_button = QCheckBox("内网登录")
        self.innerNet_button.setChecked(True)
        headless_layout.addWidget(self.innerNet_button)

        self.evaulate_button = QCheckBox("完成评估")
        self.evaulate_button.setChecked(False)
        headless_layout.addWidget(self.evaulate_button)

        self.login_button = QPushButton("一键查询")
        self.login_button.setStyleSheet("""
            QPushButton  {color: rgb(255, 255, 255); font-size: 20px; font-family: '华文中宋'; border-radius: 5px; border: 2px solid rgb(255, 255, 255); background-color: royalblue; height: 30px}
            QPushButton:hover {background-color: lightcoral; color: rgb(255, 255, 255); border: 2px solid rgb(255, 255, 255)}
            """)
        self.login_button.setCursor(Qt.PointingHandCursor)
        self.left_layout.addLayout(username_layout)
        self.left_layout.addSpacing(10)
        self.left_layout.addLayout(password_layout)
        self.left_layout.addSpacing(8)
        self.left_layout.addLayout(headless_layout)
        self.left_layout.addSpacing(2)
        self.left_layout.addWidget(self.login_button)

        
        self.show_tips = QLabel("📋处理结果显示在这里...")
        self.show_tips.setStyleSheet("font-size: 16px; font-family: '华文中宋';")
        self.left_layout.addSpacing(5)
        self.left_layout.addWidget(self.show_tips)

        self.show_result = QTextEdit("")
        self.show_result_layout = QVBoxLayout()
        self.show_result.setMinimumSize(470, 250)
        self.show_result.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.show_result_layout.addWidget(self.show_result)
        self.show_result.setStyleSheet("font-size: 15px; font-family: '华文中宋'; border: 2px solid rgba(0, 0, 0, .3); border-radius: 5px;")
        
        self.left_layout.addLayout(self.show_result_layout)

        self.login_button.clicked.connect(self.process)

        if exists("./user.txt"):
            with open("./user.txt", "r") as f:
                username = f.readline().strip()
                password = f.readline().strip()
                self.username_input.setText(username)
                self.password_input.setText(password)

            self.show_result.append("载入用户信息成功！")
    def process(self):
        
        if self.username_input.text() == "":
            self.show_result.clear()
            self.show_result.append("👤 请输入用户名")
            return
        elif self.password_input.text() == "":
            self.show_result.clear()
            self.show_result.append("🔒 请输入密码")
            return
        
        self.show_result.setText("🚀 开始执行任务... ")
        self.setWindowTitle("查询中...")
        self.login_button.setDisabled(True)

        with open('./user.txt', 'w') as f:
            f.write(self.username_input.text()+"\n")
            f.write(self.password_input.text())

        need_evaluate = True if self.evaulate_button.isChecked() else False

        innerMode = True if self.innerNet_button.isChecked() else False

        if self.headless_combobox.currentText() == "无头模式":
            self.urp_thread = urpThread(self.username_input.text(), self.password_input.text(),mode="--headless",link="https://223.112.21.198:6443/7b68f983/", innerMode=innerMode, need_evaluate=need_evaluate)
        else:
            self.urp_thread = urpThread(self.username_input.text(), self.password_input.text(),mode="",link="https://223.112.21.198:6443/7b68f983/", innerMode=innerMode, need_evaluate=need_evaluate)
            self.show_result.append("可视化查询过程...")
        self.urp_thread.start()

        self.urp_thread.process.connect(self.show_text)
        self.urp_thread.finished.connect(self.show_info)
    def show_text(self, text):
        self.show_result.append(text.strip())

    def show_info(self):
        self.setWindowTitle("查询完毕")
        self.login_button.setDisabled(False)

    def get_all_name(self):
        
        df = read_excel('./2022级学生信息.xlsx')
        self.show_result.clear()
        self.show_result.setText("\n".join(df['姓名']))
    
    def get_specific_information(self):
        
        self.search_thread = searchThread(self)
        self.search_thread.start()
        self.search_thread.process.connect(self.show_result.clear)
        self.search_thread.info.connect(self.replace)
        self.search_thread.finished.connect(self.show_text)

    def replace(self, info):

        self.username_input.setText(info[0])
        self.password_input.setText(info[1])

    def clear(self):
        self.username_input.clear()
        self.password_input.clear()
        self.show_result.clear()

    def get_people_information(self, name):

        if name == "曾政":
            self.username_input.setText("2215113116")
            self.password_input.setText("light004")
        elif name == "张嘉奇":
            self.username_input.setText("2215113155")
            self.password_input.setText("276672")
        elif name == "温炳兴":
            self.username_input.setText("2215113148")
            self.password_input.setText("182116")
        elif name == "宋益凡":
            self.username_input.setText("2215113140")
            self.password_input.setText("284310")


class searchThread(QThread):

    process = pyqtSignal()
    info = pyqtSignal(tuple)
    finished = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__()

        self.text = ""
        self.diag = SearchDialog(parent)
        self.diag.show()

    def run(self):

        df = read_excel('./2022级学生信息.xlsx')
        
        try:
            self.diag.exec_()
            name = self.diag.get_student_id()
            index = list(df['姓名']).index(name)
            self.text += '*\t{:>8}'.format(df['姓名'][index])+'\t' + str(df['学号'][index])+'\t'+str(df['身份证'][index][-6:])+'*'.center(10)
            self.process.emit()
            if self.diag.get_check():
                self.info.emit((str(df['学号'][index]), str(df['身份证'][index][-6:])))
        except Exception as e:
            pass
        
        self.finished.emit(self.text)


class urpThread(QThread):
    
    process = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, zh, mm, mode, link, innerMode=True, need_evaluate=True):
        super().__init__()

        self.zh = zh
        self.mm = mm
        self.mode = mode
        self.link = link
        self.innerMode = innerMode
        self.need_evaluate = need_evaluate
        self.text = ""

        if self.innerMode:
            self.link = "http://192.168.16.207"
        else:
            self.link = "https://223.112.21.198:6443"
    
    def get_credit(self):
        
        session = requests.Session()
        username = base64.b64encode("2215113116".encode()).decode()
        password = base64.b64encode("Ldb20040716@".encode()).decode()

        headers = {
            'User-Agent': 'Mozilla/5.0',
            'Referer': f'{self.link}',
        }

        if not self.innerMode:
            session.get(f"https://223.112.21.198:6443/vpn/user/auth/password?username={username}&password={password}&encode=1&rmbpwd_browser=0",
                                    headers=headers, verify=False)
            session.get(f"https://223.112.21.198:6443/7b68f983/", headers=headers, verify=False)

            login_url = "https://223.112.21.198:6443/7b68f983/loginAction.do"
            vchart_link = "https://223.112.21.198:6443/7b68f983/validateCodeAction.do?"  # 示例路径，需根据实际页面调整
            vchart_response = session.get(vchart_link, headers=headers, verify=False)
            ocr = DdddOcr(show_ad=False, use_gpu=False)
            result = ocr.classification(vchart_response.content)
            payload = {
                'zjh': f'{self.zh}',
                'mm': f'{self.mm}',
                'v_yzm': f'{result}',
            }

            session.post(login_url, data=payload, headers=headers, verify=False)
            grades = session.get("http://192.168.16.207/gradeLnAllAction.do?type=ln&oper=qbinfo", verify=False)
            credits = session.get("http://192.168.16.207/gradeLnAllAction.do?oper=queryXfjd", verify=False)
        else:
            login_url = "http://192.168.16.207/loginAction.do"
            vchart_link = "http://192.168.16.207/validateCodeAction.do?"  # 示例路径，需根据实际页面调整
            vchart_response = session.get(vchart_link, headers=headers, verify=False)
            ocr = DdddOcr(show_ad=False, use_gpu=False)
            result = ocr.classification(vchart_response.content)
            payload = {
                'zjh': f'{self.zh}',
                'mm': f'{self.mm}',
                'v_yzm': f'{result}',
            }

            session.post(login_url, data=payload, headers=headers, verify=False)
            grades = session.get("http://192.168.16.207/gradeLnAllAction.do?type=ln&oper=qbinfo", verify=False)
            credits = session.get("http://192.168.16.207/gradeLnAllAction.do?oper=queryXfjd", verify=False)
        
        gradesTable = read_html(grades.text)   
        creditsTable = read_html(credits.text)

        # self.text +=  "--------------------------------------------\n"
        self.text += "{:\u3000<19}\t\t{:\u3000<5}\t{:\u3000<5}\n".format("课程名","学分","成绩")

        for index in range(10, len(gradesTable), 6):
            self.text +=  "--------------------------------------------\n"
            tmp_frame = gradesTable[index].iloc[:, [2, 4, 5, 7]]
            tmp_frame.columns = ['课程名', '学分', '课程属性', '成绩']  # 给列命名（你可改成你自己喜欢的顺序）

            for i, row in tmp_frame.iterrows():
               self.text += "{:\u3000<19}\t\t{:\u3000<5}\t{:\u3000<5}\n".format(row['课程名'], row['学分'] if len(str(row['学分'])) == 3 else str(row['学分']) + ".0", row['成绩'])
                
        self.text +=  "--------------------------------------------\n"
        self.text += "{:\u3000<19}\t{:\u3000<5}\t{:\u3000<5}\t{:\u3000<5}\n".format("学年学期","学分绩点","学位绩点", "加权总分")
        self.text +=  "--------------------------------------------\n"
        for i, row in creditsTable[11].iterrows():
            self.text += "{:\u3000<19}\t\t{:\u3000<5}\t{:\u3000<5}\t{:\u3000<5}\n".format(row['学年学期'], row['学分绩点'] if len(str(row['学分绩点'])) == 4 else str(row['学分绩点']) + "0", row['学位绩点'] if len(str(row['学位绩点'])) == 4 else str(row['学位绩点']) + "0",
                                                row['加权学分学位绩点'] if len(str(row['加权学分学位绩点'])) == 5 else str(row['加权学分学位绩点']) + "0")
        self.text +=  "--------------------------------------------\n"

        return self.text

    def run(self):

        now_time = time()
        
        if self.need_evaluate:
            up = urp_tools(self.zh, self.mm, self.mode, self.link)
            if self.link == 'https://223.112.21.198:6443':
                up.offline_preprocess()
            up.login()
            self.text = up.evaluation()
            self.process.emit(self.text)
        
        self.text = ''
        tmp_flag = 0
        while True:
            try:
                self.text = ''
                self.text = self.get_credit()
                break
            except:
                tmp_flag += 1
                self.text = ''
                self.text += f"验证码输入错误,正在重新识别, 还剩{10 - tmp_flag}次机会...\n"
                self.text += '-' * 20
                self.text += '\n'
                self.process.emit(self.text)
                if tmp_flag == 10:
                    self.text = ''
                    self.text += '请检查账号密码是否正确？网页是否维护中...\n'
                    break

        self.process.emit(self.text)
        new_time = time()
        self.process.emit("✅查询完毕,  🎉耗时：%.2f秒"%(new_time-now_time))
        self.finished.emit()

if __name__ == "__main__":
    
    app = QApplication(argv)
    window = MainWindow()
    window.show()

    app.exec()
    