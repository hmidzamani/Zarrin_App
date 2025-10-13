from git import Repo
from datetime import datetime
import os
import time

# Path to your local Git repo
repo_path = r"C:\Grafana_Dashboard_Files\Zarrin_App"
file_to_push = "plc_live_data.xlsx"

# Initialize repo
repo = Repo(repo_path)
assert not repo.bare

print("🔁 Starting ultra-fast PLC data push loop (every 5 seconds)...")

while True:
    try:
        # Stage the Excel file
        repo.git.add(file_to_push)

        # Commit with timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        repo.index.commit(f"Update PLC data {timestamp}")

        # Push to GitHub
        origin = repo.remote(name='origin')
        origin.push()

        print(f"✅ Pushed updated {file_to_push} at {timestamp}")
    except Exception as e:
        print(f"⚠️ Git push failed: {e}")

    # Wait 5 seconds before next push
    time.sleep(5)
