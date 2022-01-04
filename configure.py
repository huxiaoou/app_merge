import os

TEMPLATES_FILE_LIST = {
    # "input_order_inside": "00_交易指令单_固收托管项目_YYYYMMDD.xlsx",
    # "input_order_outside": "01_投资管理总部交易指令单_固收托管项目_YYYYMMDD.xlsx",
    # "traded_order": "02_当日成交_固收托管项目_YYYYMMDD.xlsx",
    "traded_order_summary": "03_当日成交汇总_固收托管项目_YYYYMMDD.xlsx",
    "position_details": "04_持仓情况明细表_固收托管项目_YYYYMMDD.xlsx",
    "pnl_summary": "05_盈亏情况明细表_固收托管项目_YYYYMMDD.xlsx",
    # "risk_control": "06_风险限额监控表_固收托管项目_YYYYMMDD.xlsx",
    # "report_margin": "07_交易详情日报表_固收托管项目_YYYYMMDD_风控_Margin.xlsx",
    # "report_no_margin": "08_交易详情日报表_固收托管项目_YYYYMMDD_财务_NoMargin.xlsx",
}

account_configure = {
    "1001000016": {
        "src_dir": os.path.join("E:\\", "Works", "Trade", "Reports_Equity", "output"),
        "chs_name": "国轩高科托管项目",
    },
    "1003000010": {
        "src_dir": os.path.join("E:\\", "Works", "Trade", "Reports_Equity2", "output"),
        "chs_name": "小康股份托管项目",
    },
}

SAVE_NAME_START_IDX = 0
VERSION_TAG = ""
