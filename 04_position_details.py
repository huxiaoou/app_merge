from setup import *
from configure import *


def parse_desc(t_desc: str):
    parts = t_desc.split("，")
    r0 = float(re.sub("[^0-9.+-]", "", parts[0]))
    r1 = float(re.sub("[^0-9.+-]", "", parts[1]))
    r2 = float(re.sub("[^0-9.+-]", "", parts[2]))
    return r0, r1, r2


report_date = sys.argv[1]
report_name = "position_details"
start_row_num = 5
check_and_mkdir(os.path.join(OUTPUT_DIR, report_date[0:4]))
check_and_mkdir(os.path.join(OUTPUT_DIR, report_date[0:4], report_date))
save_dir = os.path.join(OUTPUT_DIR, report_date[0:4], report_date)

loaded_account_data = {}
for account_id in account_configure:
    account_file = "04_持仓情况明细表_股票可转债_{}_{}.xlsx".format(account_id, report_date)
    account_path = os.path.join(account_configure.get(account_id).get("src_dir"), report_date[0:4], report_date, account_file)
    account_df: pd.DataFrame = pd.read_excel(account_path, header=3)
    account_df = account_df.dropna(axis=0, how="all")
    loaded_account_data[account_id] = account_df
    if len(account_df) > 0:
        # print(account_df)
        pass
    else:
        print("There is no position data available for {} at {}".format(account_id, report_date))

# --- load report template
template_file = TEMPLATES_FILE_LIST[report_name]
template_path = os.path.join(TEMPLATES_DIR, template_file)
wb = xw.Book(template_path)
ws = wb.sheets["固收托管项目"]
ws.range("A2").value = "日期：" + date_format_converter_08_to_10(report_date)

s = start_row_num
qty_sum = 0
cost_val_sum = 0
mkt_val_sum = 0
unrealized_pnl_sum = 0
desc_mapper = {}
desc_data = []
for account_id, account_df in loaded_account_data.items():
    account_chs_name = account_configure.get(account_id).get("chs_name")
    for ti in account_df.index:
        if account_df.at[ti, "证券名称"] == "合计":
            pass
        elif account_df.at[ti, "证券名称"].find("年初至今") >= 0:
            account_desc = account_df.at[ti, "证券名称"]
            account_desc = account_desc.replace("注：年初至今，", account_chs_name)
            desc_mapper[account_id] = account_desc
            parsed_pnl = parse_desc(t_desc=account_desc)
            desc_data.append({"account": account_id, "realized": parsed_pnl[0], "unrealized": parsed_pnl[1], "tot": parsed_pnl[2]})
            # print(account_desc)
        else:
            ws.range("A{}".format(s)).value = account_df.at[ti, "证券名称"]
            ws.range("B{}".format(s)).value = account_df.at[ti, "证券代码"]
            ws.range("C{}".format(s)).value = account_df.at[ti, "持仓数量"]
            ws.range("D{}".format(s)).value = account_df.at[ti, "单位成本"]
            ws.range("E{}".format(s)).value = account_df.at[ti, "总成本"]
            ws.range("F{}".format(s)).value = account_df.at[ti, "收盘价"]
            ws.range("G{}".format(s)).value = account_df.at[ti, "证券市值"]
            ws.range("H{}".format(s)).value = account_df.at[ti, "浮动盈亏"]
            ws.range("I{}".format(s)).value = account_df.at[ti, "比例"]
            ws.range("J{}".format(s)).value = account_id

            qty_sum += account_df.at[ti, "持仓数量"]
            cost_val_sum += account_df.at[ti, "总成本"]
            mkt_val_sum += account_df.at[ti, "证券市值"]
            unrealized_pnl_sum += account_df.at[ti, "浮动盈亏"]

            s += 1
            ws.api.Rows(s).Insert()

s += 2
ws.range("B{}".format(s)).value = "--"
ws.range("C{}".format(s)).value = qty_sum
ws.range("D{}".format(s)).value = "--"
ws.range("E{}".format(s)).value = cost_val_sum
ws.range("F{}".format(s)).value = "--"
ws.range("G{}".format(s)).value = mkt_val_sum
ws.range("H{}".format(s)).value = unrealized_pnl_sum
ws.range("I{}".format(s)).value = unrealized_pnl_sum / cost_val_sum
ws.range("J{}".format(s)).value = "--"

s += 1
desc_by_account = "注：年初至今，\n" + ("\n".join(desc_mapper.values()))
print(desc_by_account)
desc_df = pd.DataFrame(desc_data)
print(desc_df)
desc_sum = "注：年初至今，固收托管项目实现盈利约{:.2f}万元，持仓盈亏{:.2f}万元，合计投资收益约{:.2f}万元。".format(
    desc_df["realized"].sum(), desc_df["unrealized"].sum(), desc_df["tot"].sum()
)
ws.range("A{}".format(s)).value = desc_sum
print(desc_sum)

# --- save as xlsx
save_file = template_file[SAVE_NAME_START_IDX:].replace("YYYYMMDD", report_date + VERSION_TAG)
save_path = os.path.join(save_dir, save_file)
if os.path.exists(save_path):
    os.remove(save_path)
wb.save(save_path)
wb.close()
print("| {2} | {0} | {1} | generated |".format(report_name, save_file, dt.datetime.now()))
print("=" * 120)
