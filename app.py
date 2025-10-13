from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import pandas as pd

app = Flask(__name__)
app.secret_key = "zarrinroya_secret"
EXCEL_PATH = "plc_live_data.xlsx"
USER_FILE = "Users.xlsx"

@app.route("/", methods=["GET", "POST"])
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = request.form["username"]
        pwd = request.form["password"]

        try:
            df = pd.read_excel(USER_FILE, header=0, usecols="A,B")
            credentials = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))
            if credentials.get(user) == pwd:
                session["user"] = user
                return redirect(url_for("dashboard"))
        except Exception as e:
            return render_template("login.html", error="User file error: " + str(e))

        return render_template("login.html", error="Invalid credentials")
    return render_template("login.html")

@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template("dashboard.html")

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

@app.route("/logout")
def logout():
    session.pop("user", None)
    return redirect(url_for("login"))

if __name__ == "__main__":
    app.run(port=5000, debug=True)
