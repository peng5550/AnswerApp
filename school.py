# -*- coding: UTF-8 -*-
import re
import time
from tkinter.messagebox import showerror
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from openpyxl import load_workbook
import os
from mttkinter import mtTkinter as mtk
import threading
from tkinter import scrolledtext, filedialog, ttk
from datetime import datetime

# USERNAME = "202083290181"
# PASSWORD = "24696537.p"
# EXAMNAME = "南京信息工程大学《学生手册》在线测试"


# FILEPATH = os.path.join(os.path.expanduser("~"), '答题工具').replace("\\", "/")
# if not os.path.exists(FILEPATH):
#     os.makedirs(FILEPATH)


class Application(mtk.Frame):

    def __init__(self, master):
        super(Application, self).__init__()
        self.questionDict = {}
        self.root = master
        self.root.geometry("380x450")
        self.root.title("信达学习通 1.0")
        self.__creatUI()
        self.__creatBrowser()

    def __creatBrowser(self):
        firefox_options = webdriver.FirefoxOptions()
        # self.browser = webdriver.Chrome(executable_path=f"{FILEPATH}/geckodriver.exe", options=firefox_options)
        self.browser = webdriver.Chrome(executable_path=f"geckodriver.exe", options=firefox_options)
        self.wait = WebDriverWait(self.browser, 600)
        self.browser.get('http://stu.nuist.edu.cn/Mobile/login.aspx')
        self.logtext.insert(mtk.END, f"{datetime.now().strftime('%H:%M:%S')}\t初始化成功, 请加载题型.\n")

    def __creatUI(self):
        self.box = mtk.LabelFrame(self.root)
        self.box.place(x=20, y=20, width=340, height=400)

        self.user_name = mtk.Label(self.box, text="账  号：")
        self.user_name.place(x=20, y=20, width=60, height=25)
        self.user_name_text = mtk.Entry(self.box)
        self.user_name_text.place(x=80, y=20, width=100, height=25)

        self.pass_word = mtk.Label(self.box, text="密  码：")
        self.pass_word.place(x=20, y=70, width=60, height=25)
        self.password_text = mtk.Entry(self.box)
        self.password_text.place(x=80, y=70, width=100, height=25)

        self.verify_code = mtk.Label(self.box, text="验证码：")
        self.verify_code.place(x=20, y=120, width=60, height=25)
        self.verify_code_text = mtk.Entry(self.box)
        self.verify_code_text.place(x=80, y=120, width=100, height=25)

        self.exam_name = mtk.Label(self.box, text="考试名称：")
        self.exam_name.place(x=20, y=170, width=60, height=25)
        self.exam_name_text = mtk.Entry(self.box)
        self.exam_name_text.place(x=80, y=170, width=100, height=25)

        self.loadbtn = mtk.Button(self.box, text="加载题型", command=lambda: self.thread_it(self.__loadExcel))
        self.loadbtn.place(x=210, y=20, width=100, height=50)

        self.startbtn = mtk.Button(self.box, text="开始", command=lambda: self.thread_it(self.start))
        self.startbtn.place(x=210, y=85, width=100, height=50)

        self.stopbtn = mtk.Button(self.box, text="退出", command=lambda: self.thread_it(self.stop))
        self.stopbtn.place(x=210, y=150, width=100, height=50)

        self.logs = mtk.LabelFrame(self.box, text="信息", fg="blue")
        self.logs.place(x=20, y=210, width=300, height=170)
        self.logtext = scrolledtext.ScrolledText(self.logs, fg="green")
        self.logtext.place(x=20, y=10, width=270, height=130)

    def __loadExcel(self):
        excelPath = filedialog.askopenfilename(title=u'选择文件')
        if excelPath:
            try:
                wb = load_workbook(excelPath)
                ws = wb.active
                excelData = list(ws.values)[1:]
                for question in excelData:
                    question_title = re.sub(r"(，)|(。)|(（)|(）)|(\s*)|(\n)|(\n\s)|(\t)|\(|\)", "", question[0])
                    correctAnswer = question[1]
                    self.questionDict[question_title] = str(correctAnswer)

                self.logtext.insert(mtk.END, f"{datetime.now().strftime('%H:%M:%S')}\t题目加载成功, 请点击开始进行答题.\n")
            except Exception as e:
                print(e.args)
                showerror("提示信息", "题型格式有误.")
                return
        else:
            showerror("错误信息", "请导入文件!")


    def login(self):
        USERNAME = self.user_name_text.get()
        if not USERNAME.strip():
            showerror("错误信息", "请输入账号")
            return

        PASSWORD = self.password_text.get()
        if not PASSWORD:
            showerror("错误信息", "请输入密码")
            return

        captcha_code = self.verify_code_text.get()
        if not captcha_code:
            showerror("错误信息", "请输入验证码")
            return

        self.EXAMNAME = self.exam_name_text.get()
        if not self.EXAMNAME.strip():
            showerror("错误信息", "请输入考试名称")
            return

        username = self.wait.until(EC.presence_of_element_located((By.ID, 'userbh')))
        username.send_keys(USERNAME)
        password = self.wait.until(EC.presence_of_element_located((By.ID, 'pas1s')))
        password.send_keys(PASSWORD)
        captchaCode = self.wait.until(EC.presence_of_element_located((By.ID, 'vcode')))
        captchaCode.send_keys(captcha_code)

        loginBtn = self.wait.until(EC.presence_of_element_located((By.ID, 'save2')))
        loginBtn.click()

    def studyPage(self):
        # 切换到frame
        self.browser.switch_to_frame("r_3_3")
        studyCenter = self.wait.until(EC.presence_of_element_located((By.XPATH, '//span[contains(text(), "学业学风")]')))
        studyCenter.click()

        examStart = self.wait.until(EC.presence_of_element_located(
            (By.XPATH, '//div/a[contains(text(), "{}")]'.format(self.EXAMNAME))))
        examStart.click()

        btn = self.wait.until(EC.presence_of_element_located((By.XPATH, '//td/img[@src="/imagesal/xyb.gif"]')))
        btn.click()

    def doExam(self):
        time.sleep(5)
        self.logtext.insert(mtk.END, f"{datetime.now().strftime('%H:%M:%S')}\t开始答题.\n")

        self.wait.until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Mydatalist__ctl0_Mydatalist1"]/tbody/tr')))

        func = lambda a: map(lambda b: a[b:b + 3], range(0, len(a), 3))

        ques1 = self.browser.find_elements_by_xpath('//*[@id="Mydatalist__ctl0_Mydatalist1"]/tbody/tr')
        for _, question, select in func(ques1):

            question_title = re.sub(r"(，)|(。)|(（)|(）)|(\s*)|(\n)|(\n\s)|(\t)|\(|\)", "",
                                    question.find_element_by_css_selector("td > span:nth-child(1)").text)
            correctAnswer = self.questionDict.get(question_title, [])

            for answer in select.find_elements_by_css_selector("label"):
                webAnswer = re.sub(r"([ABCD])|(\(true\))|\(false\)|\(|\)", "", answer.text)
                if webAnswer in correctAnswer:
                    answer.click()

        self.logtext.insert(mtk.END, f"{datetime.now().strftime('%H:%M:%S')}\t【单选题】答题结束.\n")
        ques2 = self.browser.find_elements_by_xpath('//*[@id="Mydatalist__ctl0_Mydatalist2"]/tbody/tr')
        for _, question, select in func(ques2):

            question_title = re.sub(r"(，)|(。)|(（)|(）)|(\s*)|(\n)|(\n\s)|(\t)|\(|\)", "",
                                    question.find_element_by_css_selector("td > span:nth-child(1)").text)
            correctAnswer = self.questionDict.get(question_title, [])

            for answer in select.find_elements_by_css_selector("label"):
                webAnswer = re.sub(r"([ABCD])|(\(true\))|\(false\)|\(|\)", "", answer.text)
                if webAnswer in correctAnswer:
                    answer.click()
        self.logtext.insert(mtk.END, f"{datetime.now().strftime('%H:%M:%S')}\t【多选题】答题结束.\n")

        ques3 = self.browser.find_elements_by_xpath('//*[@id="Mydatalist__ctl0_Mydatalist3"]/tbody/tr')
        for _, question, select in func(ques3):

            question_title = re.sub(r"(，)|(。)|(（)|(）)|(\s*)|(\n)|(\n\s)|(\t)|\(|\)", "",
                                    question.find_element_by_css_selector("td > span:nth-child(1)").text)
            correctAnswer = self.questionDict.get(question_title, [])
            for answer in select.find_elements_by_css_selector("label"):
                webAnswer = re.sub(r"([ABCD])|(\(true\))|\(false\)|\(|\)", "", answer.text)
                if webAnswer in correctAnswer:
                    answer.click()
        self.logtext.insert(mtk.END, f"{datetime.now().strftime('%H:%M:%S')}\t【判断题】答题结束.\n")
        self.logtext.insert(mtk.END, f"{datetime.now().strftime('%H:%M:%S')}\t答题结束.\n")

    def start(self):
        self.login()
        self.studyPage()
        self.doExam()

    def stop(self):
        if self.browser:
            self.browser.quit()
        os._exit(0)

    @staticmethod
    def thread_it(func, *args):
        t = threading.Thread(target=func, args=args)
        t.setDaemon(True)
        t.start()


if __name__ == '__main__':
    root = mtk.Tk()
    Application(root)
    root.mainloop()
