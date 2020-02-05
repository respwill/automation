#Job 8:7
#Though your beginning was small, yet your latter end would greatly increase.
import sys

sys.path.append("D:\Python")
from HI_tool.schedule_check import sch_check

# create sch_check class instance.
WF = sch_check("Verisilicon (941) - CES rev1.0.xlsm", "Turnkey schedule", "Lot#/DCC", device_column = "Test device", po_column = "PO#",  ship_column = "Ship To", current_fg_column="FG/PV",)

# collecting target lot number using 'set_target()' method in WF instance.
WF.set_target(941, "P/D/L", "EOH(D)")
WF.parser(WF.target_lots, WF.EMES_df, WF.result_df)
WF.get_info(["test_device", 'test_po', 'ship_code', 'current_fg', ])
WF.comparing("turnkey")
print("Inspection complete")