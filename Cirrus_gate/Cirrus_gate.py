#Job 8:7
#Though your beginning was small, yet your latter end would greatly increase.
import sys

sys.path.append("D:\Python")
from HI_tool.schedule_check import sch_check

# create sch_check class instance.
EL = sch_check("Cirrus(249) CES rev4.0.xlsm", "ProbeSCHD", "schedule Lot#", "dcc", "Tdevice", "T PO#", ship_column="선적지", current_fg_column="FG / PV",)

# collecting target lot number using 'set_target()' method in WF instance.
EL.set_target(249, "wafer", "Wfr")
EL.parser(EL.target_lots, EL.EMES_df, EL.result_df)
EL.get_info(['test_device', 'test_po', 'ship_code','current_fg'])
EL.comparing('Wafer')
print("Inspection complete")

