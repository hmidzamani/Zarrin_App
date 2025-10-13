import snap7
import time
from openpyxl import Workbook, load_workbook

PLC_IP = "192.168.10.38"
RACK = 0
SLOT = 1
EXCEL_PATH = "plc_live_data.xlsx"

DBS = {
    "C1": 134,
    "C2": 135
}

TAGS = {
    "TotalWorktime": 8,
    "TotalProducts": 20,
    "TotalGoodProducts": 24,
    "TotalScrapProducts": 32,
    "MachineSpeed": 132
}

LABELS = [
    "TotalWorktime",
    "TotalProducts",
    "TotalGoodProducts",
    "TotalScrapProducts",
    "MachineSpeed",
    "Scrap_Percentage",
    "OEE"
]

def read_dint(db_data, offset):
    return int.from_bytes(db_data[offset:offset+4], byteorder='big', signed=True)

def read_tags(client, db_number):
    db_data = client.db_read(db_number, 0, 136)
    return {tag: read_dint(db_data, offset) for tag, offset in TAGS.items()}

def calculate_metrics(data):
    scrap = data["TotalScrapProducts"]
    total = data["TotalProducts"]
    good = data["TotalGoodProducts"]
    worktime = data["TotalWorktime"]

    scrap_pct = round((100 * scrap / total) if total else 0, 2)
    oee = round((100 * good / (900 * worktime)) if worktime else 0, 2)
    return scrap_pct, oee

def init_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "LiveData"

    # Write labels in column A
    for i, label in enumerate(LABELS, start=1):
        ws.cell(row=i, column=1, value=label)

    wb.save(EXCEL_PATH)

def write_to_excel(values):
    wb = load_workbook(EXCEL_PATH)
    ws = wb["LiveData"]

    for col_index, prefix in enumerate(DBS.keys(), start=2):  # B=2, C=3
        data = values[prefix]
        scrap_pct, oee = calculate_metrics(data)
        data["Scrap_Percentage"] = scrap_pct
        data["OEE"] = oee

        for row_index, label in enumerate(LABELS, start=1):
            value = data.get(label, "")
            ws.cell(row=row_index, column=col_index, value=value)

    wb.save(EXCEL_PATH)

def main():
    client = snap7.client.Client()
    client.connect(PLC_IP, RACK, SLOT)
    init_excel()

    while True:
        all_values = {}
        for prefix, db_num in DBS.items():
            raw = read_tags(client, db_num)
            all_values[prefix] = raw

            scrap_pct, oee = calculate_metrics(raw)
            print(f"{prefix} →", {k: raw[k] for k in TAGS})
            print(f"{prefix} Scrap%: {scrap_pct:.2f}, OEE: {oee:.2f}")

        write_to_excel(all_values)
        time.sleep(1)

if __name__ == "__main__":
    main()
