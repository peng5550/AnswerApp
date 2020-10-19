# -*- coding: UTF-8 -*-
import re
import time
from tkinter.messagebox import showerror
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from fake_useragent import UserAgent
from openpyxl import load_workbook
import os
from mttkinter import mtTkinter as mtk
import threading
from tkinter import scrolledtext, filedialog, ttk

# USERNAME = "310230199109291458"
# PASSWORD = "cxh1991929."
# EXAMNAME = "9500题练习卷"
ANSWER = {"A": 1, "B": 2, "C": 3, "D": 4, "对": 5, "错": 5}

FILEPATH = os.path.join(os.path.expanduser("~"), '答题工具').replace("\\", "/")
if not os.path.exists(FILEPATH):
    os.makedirs(FILEPATH)

class Application(mtk.Frame):

    def __init__(self, master):
        super(Application, self).__init__()
        self.questionDict = {}
        self.root = master
        self.root.geometry("380x450")
        self.root.title("答题工具 1.0")
        self.__creatUI()
        self.__creatBrowser()


    def __creatBrowser(self):
        firefox_options = webdriver.FirefoxOptions()
        firefox_options.add_argument('user-agent="{}"'.format(UserAgent().random))
        self.browser = webdriver.Chrome(executable_path=f"{FILEPATH}/geckodriver.exe", options=firefox_options)
        self.wait = WebDriverWait(self.browser, 600)
        self.browser.get('http://learning.ceair.com/ilearn/en/learner/jsp/login.jsp')
        self.logtext.insert(mtk.END, "初始化成功, 请加载题型.\n")

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

        self.loginadd = mtk.Label(self.box, text="登录地：")
        self.loginadd.place(x=20, y=120, width=60, height=25)
        self.loginaddsel = ttk.Combobox(self.box)
        self.loginaddsel["values"] = [
            "上海总部", "上海物流", "甘肃", "河北",
            "西北", "山西", "浙江", "安徽",
            "山东", "江西", "武汉", "北京",
            "江苏", "四川", "广东"]
        self.loginaddsel.place(x=80, y=120, width=100, height=25)

        self.verify_code = mtk.Label(self.box, text="验证码：")
        self.verify_code.place(x=20, y=170, width=60, height=25)
        self.verify_code_text = mtk.Entry(self.box)
        self.verify_code_text.place(x=80, y=170,width=100, height=25)

        self.exam_name = mtk.Label(self.box, text="考试名称：")
        self.exam_name.place(x=20, y=220, width=60, height=25)
        self.exam_name_text = mtk.Entry(self.box)
        self.exam_name_text.place(x=80, y=220,width=100, height=25)

        self.loadbtn = mtk.Button(self.box, text="加载题型", command=lambda: self.thread_it(self.__loadExcel))
        self.loadbtn.place(x=210, y=20, width=100, height=50)

        self.startbtn = mtk.Button(self.box, text="开始", command=lambda: self.thread_it(self.start))
        self.startbtn.place(x=210, y=100, width=100, height=50)

        self.stopbtn = mtk.Button(self.box, text="退出", command=lambda: self.thread_it(self.stop))
        self.stopbtn.place(x=210, y=180, width=100, height=50)

        self.logs = mtk.LabelFrame(self.box, text="信息", fg="blue")
        self.logs.place(x=20, y=270, width=300, height=120)
        self.logtext = scrolledtext.ScrolledText(self.logs, fg="green")
        self.logtext.place(x=20, y=10, width=270, height=80)

    def __loadExcel(self):
        excelPath = filedialog.askopenfilename(title=u'选择文件')
        if excelPath:
            try:
                wb = load_workbook(excelPath)
                ws = wb.active
                excelData = list(ws.values)[1:]
                for question in excelData:
                    question_title = re.sub(r"(，)|(。)|(（)|(）)|(\s*)|(\n)|(\n\s)|(\t)", "", question[0])
                    answerList = re.findall(r'\w', question[5])
                    correctAnswer = []
                    for answer in answerList:
                        correctAnswer.append(question[ANSWER.get(answer)])
                    self.questionDict[question_title] = correctAnswer
                self.logtext.insert(mtk.END, "题目加载成功, 请点击开始进行答题.\n")
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

        self.loginAdd = self.loginaddsel.get()
        if not self.loginAdd.strip():
            showerror("错误信息", "请选择登陆地")
            return
        self.EXAMNAME = self.exam_name_text.get()
        if not self.EXAMNAME.strip():
            showerror("错误信息", "请输入考试名称")
            return

        username = self.wait.until(EC.presence_of_element_located((By.ID, 'unplay')))
        username.send_keys(USERNAME)
        password = self.wait.until(EC.presence_of_element_located((By.ID, 'pwplay')))
        password.send_keys(PASSWORD)
        loginAdd = self.wait.until(EC.presence_of_element_located((By.ID, 'content_domain')))
        loginAdd.click()
        loginAddSel = self.wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="content_domain"]/option[contains(text(), "{}")]'.format(self.loginAdd))))
        loginAddSel.click()
        captchaCode = self.wait.until(EC.presence_of_element_located((By.ID, 'vcode')))
        captchaCode.send_keys(captcha_code)

        loginBtn = self.wait.until(EC.presence_of_element_located((By.XPATH, '//li[@class="login"]/a')))
        loginBtn.click()

    def studyPage(self):
        studyCenter = self.wait.until(EC.presence_of_element_located((By.XPATH, '//li/a[contains(text(), "学习中心")]')))
        studyCenter.click()

        examCenter = self.wait.until(EC.presence_of_element_located((By.XPATH, '//li/a[contains(text(), "我的考试")]')))
        examCenter.click()

        examStart = self.wait.until(EC.presence_of_element_located(
            (By.XPATH, '//span[@title="{}"]/../a/span[contains(text(), "进入")]'.format(self.EXAMNAME))))
        examStart.click()
        window_1 = self.browser.current_window_handle
        # 获得打开的所有的窗口句柄
        windows = self.browser.window_handles
        # 切换到最新的窗口
        for current_window in windows:
            if current_window != window_1:
                self.browser.switch_to.window(current_window)
        start_btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "开始")]')))
        start_btn.click()

    def doExam(self):
        time.sleep(10)
        self.logtext.insert(mtk.END, "开始答题.\n")
        for question in self.browser.find_elements_by_xpath('//*[@id="paper_div"]/div[1]/div'):
            question_title = re.sub(r"(，)|(。)|(（)|(）)|(\s*)|(\n)|(\n\s)|(\t)", "", question.find_element_by_css_selector("h4 > span:nth-child(2)").text)
            correctAnswer = self.questionDict.get(question_title, [])
            for answer in question.find_elements_by_xpath("div[@class='test-margin']/div/label"):
                webAnswer = re.sub(r"([ABCD][\.、])|(\(true\))|\(false\)", "", answer.text)
                if webAnswer in correctAnswer:
                    answer.click()

        self.logtext.insert(mtk.END, "答题结束.\n")

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
