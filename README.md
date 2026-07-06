# Smart Excel Data Analyzer & Sorter

A professional, GUI-based software tool written in **Python**. This tool is designed to import large Excel files, validate the data, sort it automatically, and generate a statistical dashboard with integrated charts—all with a single click.

---

## 🚀 Key Features

* **Simple GUI:** No need to interact with code. Select your file and target column directly from the user-friendly interface.
* **Audit & Review System:** The tool scans data for errors (e.g., empty cells) and isolates records that require human intervention into a separate red-labeled worksheet for easy manual editing.
* **Smart Sorting:** Valid records are automatically sorted and distributed across multiple worksheets based on the category or department you choose.
* **Integrated Dashboard:** Generates a summary sheet containing the total number of records and total salaries/sales per category.
* **Automated Data Visualization:** Creates a Pie Chart showing category distributions and automatically inserts it into the final Excel file.

---

## 🛠️ Requirements & Installation

Make sure Python is installed on your system. Open your Command Prompt (CMD) or Terminal and install the required libraries by running:

```bash
pip install pandas openpyxl matplotlib
