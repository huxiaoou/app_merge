$report_date = Read-Host -Prompt "Please input the report date to merge, format = [YYYYMMDD]`nOr you can hit ENTER key directly to use the default value, which is today"
if (-not($report_date)) {
    $report_date = Get-Date -Format yyyyMMdd
}

Write-Host "The date to merge is " $report_date
python 03_traded_order_summary.py $report_date
python 04_position_details.py $report_date
python 05_pnl_summary.py $report_date
Pause