from setup import *
from configure import *

report_date = sys.argv[1]
report_name = "traded_order_summary"
start_row_num = 4
check_and_mkdir(os.path.join(OUTPUT_DIR, report_date[0:4]))
check_and_mkdir(os.path.join(OUTPUT_DIR, report_date[0:4], report_date))
save_dir = os.path.join(OUTPUT_DIR, report_date[0:4], report_date)

loaded_account_data = {}
for account_id in account_configure:
    account_file = "03_当日成交汇总_股票可转债_{}_{}.xlsx".format(account_id, report_date)
    account_path = os.path.join(account_configure.get(account_id).get("src_dir"), report_date[0:4], report_date, account_file)
    account_df = pd.read_excel(account_path, header=2)
    loaded_account_data[account_id] = account_df
    if len(account_df) > 0:
        print(account_df)
    else:
        print("There is no traded data available for {} at {}".format(account_id, report_date))

# --- load report template
template_file = TEMPLATES_FILE_LIST[report_name]
template_path = os.path.join(TEMPLATES_DIR, template_file)
wb = xw.Book(template_path)
ws = wb.sheets["固收托管项目"]
ws.range("A1").value = "日期：" + date_format_converter_08_to_10(report_date)

s = start_row_num
amt_sum = 0
for account_id, account_df in loaded_account_data.items():
    for ti in account_df.index:
        if account_df.at[ti, "类别"] == "合计":
            pass
        else:
            # print("The following code has not been tested with real data, please be careful with the results")
            ws.range("A{}".format(s)).value = "{}({})".format(account_df.at[ti, "类别"], account_id)
            ws.range("B{}".format(s)).value = account_df.at[ti, "证券代码"]
            ws.range("C{}".format(s)).value = account_df.at[ti, "证券名称"]
            ws.range("D{}".format(s)).value = account_df.at[ti, "成交均价"]
            ws.range("E{}".format(s)).value = account_df.at[ti, "成交数量"]
            ws.range("F{}".format(s)).value = account_df.at[ti, "收付金额"]
            amt_sum += account_df.at[ti, "收付金额"]
            ws.range("G{}".format(s)).value = account_df.at[ti, "市场"]
            ws.range("H{}".format(s)).value = account_id
            s += 1
            ws.api.Rows(s).Insert()

ws.range("A{}".format(s)).value = "合计"
ws.range("B{}".format(s)).value = "--"
ws.range("C{}".format(s)).value = "--"
ws.range("D{}".format(s)).value = "--"
ws.range("E{}".format(s)).value = "--"
ws.range("F{}".format(s)).value = amt_sum
ws.range("G{}".format(s)).value = "--"
ws.range("H{}".format(s)).value = "--"

# --- save as xlsx
save_file = template_file[SAVE_NAME_START_IDX:].replace("YYYYMMDD", report_date + VERSION_TAG)
save_path = os.path.join(save_dir, save_file)
if os.path.exists(save_path):
    os.remove(save_path)
wb.save(save_path)
wb.close()
print("| {2} | {0} | {1} | generated |".format(report_name, save_file, dt.datetime.now()))
print("=" * 120)
