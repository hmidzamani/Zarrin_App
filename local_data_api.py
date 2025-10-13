from flask import Flask, jsonify, request
import pandas as pd

app = Flask(__name__)
EXCEL_PATH = "plc_live_data.xlsx"

@app.route("/data")
def data():
    machine = request.args.get("machine", "C1")
    try:
        df = pd.read_excel(EXCEL_PATH, header=None)
        col = 1 if machine == "C1" else 2
        values = {
            "TotalWorktime": df.at[1, col],
            "TotalProducts": df.at[2, col],
            "TotalGoodProducts": df.at[3, col],
            "TotalScrapProducts": df.at[4, col],
            "MachineSpeed": df.at[5, col],
            "scrapPercentage": df.at[6, col],
            "Online_OEE": df.at[7, col]
        }
    except Exception as e:
        values = {"error": str(e)}
    return jsonify(values)

if __name__ == "__main__":
    app.run(port=5050, debug=True)