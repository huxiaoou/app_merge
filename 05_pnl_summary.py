from setup import *
from configure import *

report_date = sys.argv[1]
report_name = "pnl_summary"
start_row_num = 4
check_and_mkdir(os.path.join(OUTPUT_DIR, report_date[0:4]))
check_and_mkdir(os.path.join(OUTPUT_DIR, report_date[0:4], report_date))
save_dir = os.path.join(OUTPUT_DIR, report_date[0:4], report_date)

loaded_account_data = {}
for account_id in account_configure:
    account_file = "05_盈亏情况明细表_股票可转债_{}_{}.xlsx".format(account_id, report_date)
    account_path = os.path.join(account_configure.get(account_id).get("src_dir"), report_date[0:4], report_date, account_file)
    account_df: pd.DataFrame = pd.read_excel(account_path, header=2)
    account_df = account_df.dropna(axis=0, how="all")
    loaded_account_data[account_id] = account_df
    if len(account_df) > 0:
        print(account_df)
        print("=" * 120)
    else:
        print("There is no summary data available for {} at {}".format(account_id, report_date))

# --- load report template
template_file = TEMPLATES_FILE_LIST[report_name]
template_path = os.path.join(TEMPLATES_DIR, template_file)
wb = xw.Book(template_path)
ws = wb.sheets["固收托管项目"]
ws.range("A2").value = "日期：" + date_format_converter_08_to_10(report_date)

s = start_row_num
qty_sum = 0
mkt_val_sum = 0
unrealized_pnl_sum = 0
realized_pnl_cumsum = 0
tot_pnl_sum = 0
for account_id, account_df in loaded_account_data.items():
    account_chs_name = account_configure.get(account_id).get("chs_name")
    for ti in account_df.index:
        if account_df.at[ti, "类别"] == "合  计":
            pass
        else:
            if account_df.at[ti, "类别"].find("股票") >= 0:
                instrument_type = "股票"
            elif account_df.at[ti, "类别"].find("债券基金") >= 0:
                instrument_type = "债券基金"
            elif account_df.at[ti, "类别"].find("可转债、可交债") >= 0:
                instrument_type = "可转债、可交债"
            else:
                instrument_type = "其它"
            ws.range("A{}".format(s)).value = "{}（{}）".format(instrument_type, account_chs_name)
            ws.range("B{}".format(s)).value = account_df.at[ti, "本日持仓"]
            ws.range("C{}".format(s)).value = account_df.at[ti, "持仓市值"]
            ws.range("D{}".format(s)).value = account_df.at[ti, "持仓浮动盈亏"]
            ws.range("E{}".format(s)).value = account_df.at[ti, "累计已实现盈亏"]
            ws.range("F{}".format(s)).value = account_df.at[ti, "总盈亏"]

            qty_sum += account_df.at[ti, "本日持仓"]
            mkt_val_sum += account_df.at[ti, "持仓市值"]
            unrealized_pnl_sum += account_df.at[ti, "持仓浮动盈亏"]
            realized_pnl_cumsum += account_df.at[ti, "累计已实现盈亏"]
            tot_pnl_sum += account_df.at[ti, "总盈亏"]

            s += 1
            ws.api.Rows(s).Insert()

ws.api.Rows(s).Delete()
ws.range("B{}".format(s)).value = qty_sum
ws.range("C{}".format(s)).value = mkt_val_sum
ws.range("D{}".format(s)).value = unrealized_pnl_sum
ws.range("E{}".format(s)).value = realized_pnl_cumsum
ws.range("F{}".format(s)).value = tot_pnl_sum

# --- save as xlsx
save_file = template_file[SAVE_NAME_START_IDX:].replace("YYYYMMDD", report_date + VERSION_TAG)
save_path = os.path.join(save_dir, save_file)
if os.path.exists(save_path):
    os.remove(save_path)
wb.save(save_path)
wb.close()
print("| {2} | {0} | {1} | generated |".format(report_name, save_file, dt.datetime.now()))
print("=" * 120)
