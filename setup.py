import os
import sys
import numpy as np
import pandas as pd
import datetime as dt
import xlwings as xw
import re
from skyrim.winterhold import check_and_mkdir, date_format_converter_08_to_10

ROOT_DIR = os.path.join("E:\\", "Works", "Trade", "Reports_Merge")
OUTPUT_DIR = os.path.join(ROOT_DIR, "output")
TEMPLATES_DIR = os.path.join(ROOT_DIR, "templates")

src_output_dir = {
    "1001000016": os.path.join("E:", "Works", "Trade", "Reports_Equity", "output"),
    "1003000010": os.path.join("E:", "Works", "Trade", "Reports_Equity2", "output"),
}
