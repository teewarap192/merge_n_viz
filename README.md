# Last Wafer Defect Monitoring - Automation Tool

This tool automates the weekly data entry, formatting, and charting process for the Last Wafer Defect Monitoring database. It takes the new weekly data, seamlessly merges it into the master tracking file, highlights the new additions for easy comparison, and automatically updates the tracking chart to display the latest 80 lots.

## 🚀 Prerequisites & Setup (First Time Only)

Before using this tool for the first time, you need to set up your Python environment. Don't worry, it's just a few simple commands!

1. **Install Python:** Ensure you have Python installed on your computer. You can download it from [python.org](https://www.python.org/downloads/).
2. **Download this folder:** Clone this repository or download it as a ZIP file and extract it to a folder on your computer.
3. **Open your Terminal:** * *Mac:* Open the `Terminal` app.
   * *Windows:* Open `Command Prompt` or `PowerShell`.
4. **Navigate to the folder:** Use the `cd` command to enter the folder where you saved these files. (e.g., `cd Documents/LastWaferTool`)
5. **Install Required Libraries:** Run the following command to automatically install all the exact tools this script needs to run:
   ```bash
   pip install -r requirements.txt
   
(Note: If you are on a Mac, you may need to type pip3 instead of pip).


## 📁 File Setup

For the script to work properly, all files must be in the exact same folder.
1. Place the Python script (merge_n_viz.py) in your working folder.
2. Place your New Data File (e.g., file1.xlsx) in the same folder.
3. Place your Master Database File (e.g., file2.xlsx) in the same folder.

## 💻 How to Run the Tool

1. Open your Terminal or Command Prompt.
2. Navigate to the folder where you saved the files.
3. Run the script by typing:
   ```bash
   python merge_n_viz.py
   ```
(Again, Mac users may need to type python3 merge_n_viz.py)

## The Interactive Prompts

Once the script starts, it will ask you 4 simple questions. You do not need to type .xlsx at the end of your file names; the tool will add it automatically!

* Prompt 1: Enter the current week number (e.g., 35):
Type the week number. The script uses this to label the new data (e.g., LW35).
* Prompt 2: Enter the name of the NEW data file:
Type the name of the file containing this week's raw data (e.g., file1).
* Prompt 3: Enter the name of the MASTER data file:
Type the name of your historical database file (e.g., file2).
* Prompt 4: Enter the name to SAVE AS:
Type what you want the finished, updated file to be called (e.g., Updated_Master_Week35).

## ⚙️ What the Tool Does Automatically

Once you answer the prompts, the script will handle the rest in a matter of seconds:
* **Housekeeping**: It finds last week's highlighted staging columns, permanently merges their data into the historical database, and cleans up old formatting.
* **Data Staging**: It builds new columns on the far right for the current week, imports the new data, and highlights it yellow for immediate visual comparison.
* **Target Thresholds**: It automatically adds a 1 to the Target row for all new lots.
* **Visual Dashboards**: It generates a fresh, perfectly scaled bar chart underneath your data showing the defect rates for the latest 80 lots, overlaid with a solid red target threshold line.
