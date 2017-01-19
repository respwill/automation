# Job 8:7
# Though your beginning was small, yet your latter end would greatly increase.

import sys
sys.path.append("D:\Python")

from selenium import webdriver
import pandas as pd
from HI_tool import CES_read, emes_parsing, emes_login
import os

class sch_check(emes_parsing.parser):
    # set data frame initially: excel information and column names
    def __init__(self, book, sheet, lot_column, pdl_column,device_column, po_column, datecode_column, tracecode_column, coo_column,
                 el_fg_column, be_fg_column, ship_column):
        # set excel information include column names.
        self.book = book
        self.sheet = sheet
        self.lot_column = lot_column
        self.device_column = device_column
        self.po_column = po_column
        self.datecode_column = datecode_column
        self.tracecode_column = tracecode_column
        self.el_fg_column = el_fg_column
        self.be_fg_column = be_fg_column
        self.ship_column = ship_column
        self.coo_column = coo_column
        self.pdl_column = pdl_column

        self.current_dir = os.getcwd()

        # creating Dataframe.
        self.EMES_df = pd.DataFrame(
            columns=[lot_column, device_column, po_column, datecode_column, tracecode_column, coo_column, el_fg_column,
                     be_fg_column,
                     ship_column])
        self.result_df = pd.DataFrame(
            columns=[lot_column, device_column, po_column, datecode_column, tracecode_column, coo_column, el_fg_column,
                     be_fg_column,
                     ship_column])

        # read lot# list from target excel file.
        self.CES_df = pd.read_excel(book, sheetname=sheet)

        # make target lot number list.
        self.target_lots = []

    # collecting target lot number and append to data fram and data base..
    def set_target(self):
        # collecting target lot number from BE scheduling column and ignore blank or column name.
        self.CES_df = self.CES_df.loc["Row"]
        self.CES_df.index = range(len(self.CES_df))
        for n in range(0, len(self.CES_df[self.lot_column])):
            if pd.isnull(self.CES_df[self.lot_column].loc[n]):
                break
            elif self.CES_df[self.lot_column].loc[n] == self.lot_column:
                break
            else:
                self.target_lots.append(str(self.CES_df[self.lot_column].loc[n]))
                CES_read.sch_db_input(74, self.CES_df[self.lot_column].loc[n], self.CES_df[self.pdl_column].loc[n],self.CES_df["EOH(D)"].loc[n],
                                      self.CES_df[self.el_fg_column].loc[n])

    # Compare emes and target excel file.
    def turnkey_checking(self):
        for n in range(0, len(self.EMES_df)):
            self.result_df[self.po_column][n] = CES_read.compare(emes_df=self.EMES_df[self.po_column][n],
                                                                 ces_df=int(self.CES_df[self.po_column][n]))
            self.result_df[self.tracecode_column][n] = CES_read.compare(emes_df=self.EMES_df[self.tracecode_column][n],
                                                                        ces_df=self.CES_df[self.tracecode_column][n])

            self.result_df[self.el_fg_column][n] = CES_read.elCompare(emes_df=self.EMES_df[self.el_fg_column][n],
                                                                   ces_df=self.CES_df[self.el_fg_column][n])

            self.result_df[self.be_fg_column][n] = CES_read.elCompare(emes_df=self.EMES_df[self.be_fg_column][n],
                                                                   ces_df=self.CES_df[self.be_fg_column][n])

            self.result_df[self.ship_column][n] = CES_read.shipCompare(emes_df=self.EMES_df[self.ship_column][n],
                                                                       ces_df=self.CES_df[self.ship_column][n],devide_char='-')
        if not "inspection result" in os.listdir(self.current_dir):
            os.mkdir("inspection result")

        writer = pd.ExcelWriter("{}/inspection result/Turnkey_Lot insp_result.xlsx".format(self.current_dir), engine="xlsxwriter")
        self.result_df.to_excel(writer, sheet_name=self.sheet)
        writer.close()
        print("\nChecking for Tunrkey lots has done")

    def only_checking(self):
        for n in range(0, len(self.EMES_df)):
            self.result_df[self.po_column][n] = CES_read.compare(emes_df=self.EMES_df[self.po_column][n],
                                                                 ces_df=self.CES_df[self.po_column][n])
            self.result_df[self.tracecode_column][n] = CES_read.compare(emes_df=self.EMES_df[self.tracecode_column][n],
                                                                        ces_df=self.CES_df[self.tracecode_column][n])
            self.result_df[self.coo_column][n] = CES_read.cooCompare(emes_df=self.EMES_df[self.coo_column][n],
                                                                      ces_df=self.CES_df[self.coo_column][n])
            self.result_df[self.datecode_column][n] = CES_read.compare(emes_df=self.EMES_df[self.datecode_column][n],
                                                                       ces_df=self.CES_df[self.datecode_column][n])
            self.result_df[self.el_fg_column][n] = CES_read.elCompare(emes_df=self.EMES_df[self.el_fg_column][n],
                                                                   ces_df=self.CES_df[self.el_fg_column][n])
            self.result_df[self.be_fg_column][n] = CES_read.elCompare(emes_df=self.EMES_df[self.be_fg_column][n],
                                                                   ces_df=self.CES_df[self.be_fg_column][n])
            self.result_df[self.ship_column][n] = CES_read.shipCompare(emes_df=self.EMES_df[self.ship_column][n],
                                                                       ces_df=self.CES_df[self.ship_column][n],devide_char='-')
        if "inspection result" not in os.listdir(self.current_dir):
            os.mkdir("inspection result")

        writer = pd.ExcelWriter("{}/inspection result/Only_Lot insp_result.xlsx".format(self.current_dir), engine="xlsxwriter")
        self.result_df.to_excel(writer, sheet_name=self.sheet)
        writer.close()
        print("\nChecking for Only lots has done")

# set driver
driver = webdriver.PhantomJS()

# create 'access' instance to log in to EMES.
login = emes_login.access(driver)
login.connecting()

# try:
# lot_column,device_column,po_column,datecode_column,tracecode_column,coo_column,fg_column,ship_column
turnkey = sch_check("Atmel(74,125,955)-TEST-CES V6.7.xlsm", "SCH(Turnkey)", "Lot# / Dcc", "P/D/L", "T device", "Test PO",
                    "Date", "Trace", "COO", "EL FG", "BE(Tstck) FG", "SHIP")
only = sch_check("Atmel(74,125,955)-TEST-CES V6.7.xlsm", "SCH(Only)", "Lot# / Dcc", "P/D/L", "T device", "Test PO", "Date",
                 "Trace", "COO", "EL FG", "BE(Tstck) FG", "SHIP")

#self.test_device, self.test_PO, self.dateCode, self.traceCode, self.coo, self.test_FG, self.ship_code

turnkey.set_target()
turnkey.parser(turnkey.target_lots,driver,turnkey.EMES_df,turnkey.result_df)
turnkey.run(['test_device', 'test_po', 'date_code', 'trace_code', 'coo', 'current_fg', 't_stock_fg', 'ship_code'])
turnkey.turnkey_checking()

only.set_target()
only.parser(only.target_lots,driver,only.EMES_df,only.result_df)
only.run(['test_device','test_po','date_code','trace_code','coo','current_fg','t_stock_fg','ship_code'])
only.only_checking()

driver.quit()
print("\nInspection complete")

# except PermissionError:
#     input("Please close result files. Proram needs to overwrite them.")
# except:
#     input("Unexpected error caused, please run it in pycharm to check error")





