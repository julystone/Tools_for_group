import yaml

from library.xlwings_pd import OperatingExcelPd

YAML_PATH = "../config.yaml"

with open(YAML_PATH, 'r', encoding='utf-8') as f:
    conf = yaml.safe_load(f)
    # EXCEL_NAME = "../落花机密表_糙汉子版.xlsx"
    SHEET_NAME = conf['excel']['sheet']
    EXCEL_PATH = "../落花机密表_糙汉子版.xlsx"

ro = OperatingExcelPd(EXCEL_PATH, SHEET_NAME)

dp = ro.return_dp()

print(dp)
