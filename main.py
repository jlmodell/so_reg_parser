import pandas as pd
import os
from rich import print
import httpx
import re
from datetime import datetime
import win32com.client as win32

excel = win32.Dispatch("Excel.Application")

# Disable Excel alerts and visibility
excel.DisplayAlerts = False
excel.Visible = False

file = "C:\\temp\\SO REGISTER.xls"
secondary = "C:\\temp\\SO.REGISTER.xls"
output = "C:\\temp\\PARSED_SO.REGISTER.xlsx"

current_schedule: dict[str, list[str]] = {}

backorder_regex = re.compile(r"^2025-")


def parse_release_schedule() -> pd.DataFrame:
    global current_schedule

    url = "http://128.1.0.155:3637/api/bic/data?limit=2000"
    r = httpx.get(url)
    data = r.json()
    df = pd.DataFrame(data)

    # sort by run_date_time
    df = df.sort_values(by="run_date_time", ascending=True)

    # iterate through rows
    for _index, row in df.iterrows():
        date = datetime.strptime(row["run_date_time"], "%Y-%m-%d")
        if date < datetime.now():
            continue

        if row["item"] not in current_schedule:
            current_schedule[row["item"]] = [
                f"{row['run_date_time']}-{row['lot']}-{row['qty']}"
            ]
        else:
            current_schedule[row["item"]].append(
                f"{row['run_date_time']}-{row['lot']}-{row['qty']}"
            )

    print(df.head())

    df.to_csv("C:\\temp\\release_schedule.csv")

    return df


rs_df = parse_release_schedule()


def get_scheduled_jobs(row: pd.Series) -> str:
    global current_schedule

    if row["Backordered"] == False:
        return ""

    if row["Cat"] in current_schedule:
        print(row["Cat"])
        return " | ".join(current_schedule[row["Cat"]])
    else:
        return ""


def parse_file(file: str) -> pd.DataFrame:
    # assert os.path.exists(file), "File does not exist"

    df = pd.read_excel(file)

    # fill all nan with empty string
    df = df.fillna("")
    df["Cat"] = df["Cat"].astype(str)

    def is_row_backordered(row: pd.Series) -> bool:
        # if row["Reqstd"] is type datetime.datetime
        if type(row["Reqstd"]) == datetime:
            # if year is 2025, then it is backordered
            return row["Reqstd"].year == 2025
        else:
            return False

    df["Backordered"] = df.apply(is_row_backordered, axis=1)

    df["Scheduled"] = df.apply(get_scheduled_jobs, axis=1)

    print(df.head())

    df.to_excel(output, index=False)

    return df


def combine_reports(source: str = output, destination: str = file):
    try:
        # Open the source file
        source_file = excel.Workbooks.Open(output)
        source_sheet = source_file.Worksheets("Sheet1")

        # Open the destination file
        destination_file = excel.Workbooks.Open(file)
        destination_sheet = destination_file.Worksheets("SO Register")

        # Get the column range to copy from the source file
        source_column = source_sheet.Range("Q:Q")

        # Get the destination column range in the destination file
        destination_column = destination_sheet.Range("Q:Q")

        # Copy values from source column to destination column
        source_column.Copy(destination_column)

        # Auto fit the destination column
        destination_sheet.Columns("Q:Q").AutoFit()

        # Save and close the destination file
        destination_file.Save()
        destination_file.Close()

        # Close the source file without saving changes
        source_file.Close(False)
    except Exception:
        pass


if __name__ == "__main__":
    if not os.path.exists(file):
        file = secondary

    assert os.path.exists(file), "File does not exist"

    df = parse_file(file)
    combine_reports()
