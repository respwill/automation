#Job 8:7
#Though your beginning was small, yet your latter end would greatly increase.
import sys

sys.path.append("D:\Python")
from HI_tool import emes_login
from selenium import webdriver
import pandas as pd

# # fields = ["Lot#", "Trace Code", "COO", "Date Code", "OPN"]
# fields = "C, F, J:M"
# test_df = pd.read_excel("target.xlsx", parse_cols=fields)
#
# for n in range(0,len(test_df)):
#     items = test_df.loc[n]
#     print(items["Lot#"],items["SO#"])


class exSch():
    # get log in info
    def __init__(self,user, pswd, badge):
        self.driver = webdriver.Chrome()
        # self.driver = webdriver.PhantomJS()
        self.driver.get("http://aak1ws01/eMES/user/login.do?user=" + user + "&password=" + pswd + "&badge=" + badge)
        fields = "C, F, J:M"
        self.target_df = pd.read_excel("target.xlsx", parse_cols=fields)

    def release(self,cust):
        self.driver.get("http://aak1ws01/eMES/sch/ReleasedPOList.jsp")
        self.fromDate = self.driver.find_element_by_name("fromDate")
        self.toDate = self.driver.find_element_by_name("toDate")

        for n in range(0,len(self.target_df)) :
            items = self.target_df.loc[n]
            print("working on {}".format(str(items["SO#"])))
            # clear so box and input so data from target file
            self.so = self.driver.find_element_by_name("salesOrderNo")
            self.so.clear()
            self.so.send_keys(str(items["SO#"]))

            # clear cust box and input cust code as 110 / 278 may be needed as well.
            self.custInput = self.driver.find_element_by_name("cust")
            self.custInput.clear()
            self.custInput.send_keys(cust)

            # click find key
            self.findButton = self.driver.find_element_by_name("find")
            self.findButton.click()

            # need to recognize what lot# i am working on.


            # self.driver.switch_to.window(self.driver.window_handles[0])
            # self.lot_check = self.driver.find_elements_by_xpath("//tr/td/input[@name='formcheckbox1']")
            # print("{} po(s) found".format(len(self.lot_check)))
            # for i, check in enumerate(self.lot_check):
            #     self.lot = self.driver.find_element_by_xpath("//tr/td/input[@name='formcheckbox1']")
            #     self.lot.click()
            #     self.schTestButton = self.driver.find_element_by_name("TESTSCH")
            #     self.schTestButton.click()
            #     # page change, 1 window
            #     self.executeButton1 = self.driver.find_element_by_xpath("//p/a[@href='javascript:executeDo()']")
            #     self.executeButton1.click()
            #     # page change, 2 windows
            #     self.driver.switch_to.window(self.driver.window_handles[0]) #index was 1
            #     self.coo = self.driver.find_element_by_xpath("//span/select[@name='coo']")
            #     self.coo.send_keys(coo)
            #     self.executeButton2 = self.driver.find_element_by_xpath("//tr/td/input[@value='  Execution  ']")
            #     self.executeButton2.click()
            #     # page2 change, 2 windows
            #     self.executeButton3 = self.driver.find_element_by_name("exeDo")
            #     self.executeButton3.click()
            #     # self.driver.switch_to.window(self.driver.window_handles[1])
            #     self.backButton = self.driver.find_element_by_xpath("//tr/td/input[@value='  Back  ']")
            #     self.backButton.click()
            #     # # close page2 and back to initial page
            #     self.driver.switch_to.window(self.driver.window_handles[0])
            #     print("{} out of {} po has done".format(i+1, len(self.lot_check)))

user = "parkhi"
pswd = "Phi12900"
badge = "374105"

altera = exSch(user, pswd, badge)
altera.release(110)
