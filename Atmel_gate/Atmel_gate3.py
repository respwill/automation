# Job 8:7
# Though your beginning was small, yet your latter end would greatly increase.

import sys
sys.path.append("D:\Python")
from HI_tool.schedule_check import sch_check

# lot_column,device_column,po_column,datecode_column,tracecode_column,coo_column,fg_column,ship_column
#self.test_device, self.test_PO, self.dateCode, self.traceCode, self.coo, self.test_FG, self.ship_code
turnkey = sch_check("Atmel(74,125,955)-TEST-CES V6.7.xlsm", "SCH(Turnkey)", "Lot# / Dcc", device_column="T device", po_column="Test PO",
                    tracecode_column="Trace", ship_column="SHIP", current_fg_column="EL FG", tstock_fg_column="BE(Tstck) FG", )
turnkey.set_target(74, "P/D/L", "EOH(D)")
turnkey.parser(turnkey.target_lots, turnkey.EMES_df, turnkey.result_df)
#list for 'get_info' method should be in order of parameter of 'sch_check' class except 'lot_column'
#'lot_column' will be added automatically.
turnkey.get_info(['test_device', 'test_po', 'trace_code', 'ship_code', 'current_fg', 't_stock_fg',]) #get_info
turnkey.comparing("Turnkey")
turnkey.quit_driver()


only = sch_check("Atmel(74,125,955)-TEST-CES V6.7.xlsm", "SCH(Only)", "Lot# / Dcc", device_column="T device", po_column="Test PO",
                 datecode_column="Date", tracecode_column="Trace", coo_column="COO", ship_column="SHIP", current_fg_column="EL FG", tstock_fg_column="BE(Tstck) FG", )
only.set_target(74, "P/D/L", "EOH(D)")
only.parser(only.target_lots, only.EMES_df, only.result_df)
only.get_info(['test_device','test_po','date_code','trace_code', 'coo', 'ship_code', 'current_fg','t_stock_fg',])
only.comparing("Only")
only.quit_driver()

print("\nInspection complete")

# except PermissionError:
#     input("Please close result files. Proram needs to overwrite them.")
# except:
#     input("Unexpected error caused, please run it in pycharm to check error")





