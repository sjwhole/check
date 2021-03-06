import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import QIcon
from selenium import webdriver
import win32com.client as win32
import pandas as pd
import time, os, datetime


class info:
    def __init__(self, id, pw):
        self.id = id
        self.pw = pw

    def remove(self):
        # 이전파일 삭제
        for subject in self.code:
            file_xls = subject + '.xls'
            file_xlsx = subject + '.xlsx'
            if os.path.isfile(file_xls):
                os.remove(file_xls)
            if os.path.isfile(file_xlsx):
                os.remove(file_xlsx)

    def download(self):
        # Headless 옵션 적용
        options = webdriver.ChromeOptions()
        options.add_argument('headless')
        options.add_argument('window-size=1920x1080')
        options.add_argument("disable-gpu")

        # 로그인사이트접속
        path = "chromedriver.exe"
        driver = webdriver.Chrome(path, options=options)
        driver.maximize_window()
        driver.implicitly_wait(3)
        driver.get("https://learn.hanyang.ac.kr/ultra/institution-page")
        driver.find_element_by_xpath('//*[@id="entry-login-custom"]').click()
        time.sleep(1)

        # 로그인
        elem_login = driver.find_element_by_id("uid")
        elem_login.clear()
        elem_login.send_keys(self.id)
        elem_login = driver.find_element_by_id("upw")
        elem_login.clear()
        elem_login.send_keys(self.pw)
        driver.find_element_by_xpath('//*[@id="login_btn"]').click()
        time.sleep(5)

        # 학번 받아오기
        driver.get("https://learn.hanyang.ac.kr/ultra/profile")
        std_code = driver.find_element_by_xpath(
            "/html/body/div[1]/div[2]/bb-base-layout/div/main/div/div/div/div[1]/div[3]/div[1]/section[1]/ul/li[3]/div/div")
        self.std_code = std_code.text


        # 강좌명, 학수번호 받아오기
        driver.get("https://learn.hanyang.ac.kr/ultra/messages")
        self.code = {}
        i = 1
        while True:
            try:
                subject = driver.find_element_by_xpath(
                    "/html/body/div[1]/div[2]/bb-base-layout/div/main/div/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div/div[{}]/bb-course-conversations-summary/div/div/div[1]/h3".format(
                        i))
                hyn = driver.find_element_by_xpath(
                    "/html/body/div[1]/div[2]/bb-base-layout/div/main/div/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div/div[{}]/bb-course-conversations-summary/div/div/div[1]/div[1]/span/bdi".format(
                        i))
                subject_r = ''
                for j in subject.text:
                    if j.isalpha():
                        subject_r += j
                self.code[subject_r] = hyn.text
                i += 1
            except:
                break
        # 이전파일 삭제
        self.remove()

        # 다운로드 설정
        driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd': 'Page.setDownloadBehavior',
                  'params': {'behavior': 'allow', 'downloadPath': os.getcwd()}}
        command_result = driver.execute("send_command", params)
        # 엑셀파일 다운로드
        for subject, hyn in self.code.items():
            driver.get(
                'https://learn.hanyang.ac.kr/webapps/bbgs-OnlineAttendance-BB5a998b8c44671/excel?selectedUserId=' + self.std_code + '&crs_batch_uid=' + hyn + '&title=' + subject + '&column=%EC%82%AC%EC%9A%A9%EC%9E%90%EB%AA%85,%EC%9C%84%EC%B9%98,%EC%BB%A8%ED%85%90%EC%B8%A0%EB%AA%85,%ED%95%99%EC%8A%B5%ED%95%9C%EC%8B%9C%EA%B0%84,%ED%95%99%EC%8A%B5%EC%9D%B8%EC%A0%95%EC%8B%9C%EA%B0%84,%EC%BB%A8%ED%85%90%EC%B8%A0%EC%8B%9C%EA%B0%84,%EC%98%A8%EB%9D%BC%EC%9D%B8%EC%B6%9C%EC%84%9D%EC%A7%84%EB%8F%84%EC%9C%A8,%EC%98%A8%EB%9D%BC%EC%9D%B8%EC%B6%9C%EC%84%9D%EC%83%81%ED%83%9C(P/F)')
        time.sleep(1)
        driver.quit()

    def change(self):
        for subject in self.code:
            fname = os.getcwd() + '/' + subject + '.xls'
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(fname)

            wb.SaveAs(fname + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
            wb.Close()  # FileFormat = 56 is for .xls extension
            excel.Application.Quit()

    def F(self):
        with open("출결상황.txt", 'w') as f:
            now = datetime.datetime.now()
            f.write(str(now) + '\n\n')
            for subject in self.code:
                file = subject + '.xlsx'
                df = pd.read_excel(file)
                pd.set_option('display.max_colwidth', -1)
                F = (df[df['온라인출석상태(P/F)'] == 'F'])
                lecture = (F['컨텐츠명'])
                if len(lecture) == 0:
                    f.write(subject + '\n')
                    f.write("모두 수강했습니다.\n\n")
                else:
                    f.write(subject + '\n')
                    f.write(lecture.to_string(index=False) + '\n')

    def do(self):
        self.download()
        self.change()
        self.F()
        self.remove()


form_class = uic.loadUiType("check.ui")[0]


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowIcon(QIcon('hanyang.png'))
        self.pw.setEchoMode(QLineEdit.Password)
        self.pushButton.clicked.connect(self.btn_clicked)

    def btn_clicked(self):
        id = self.id.text()
        pw = self.pw.text()
        sj = info(id, pw)
        download = False
        try:
            sj.download()
            self.info.setText("로그인 성공")
            download = True
        except:
            self.info.setText("로그인 에러")
        if download:
            sj.change()
            sj.F()
            sj.remove()
            self.info.setText("실행 완료!")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    Window = MyWindow()
    Window.show()
    app.exec_()
