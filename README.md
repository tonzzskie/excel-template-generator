# 🧾 Excel Template Generator

A simple Python utility to create Excel templates and populate them with dummy/sample data.

---

## 📁 Project Structure
```

Sample/
├── data_generator2.py       # Generates dummy/sample data for Excel
├── excel_formatter_all.py   # Formats and structures Excel templates
├── requirements.txt         # Python dependencies
├── venv/                    # Virtual environment (ignored in Git)
└── README.md

````

---

## ⚙️ Setup Instructions

### 1. Clone the Repository
```bash
git clone git@github.com:tonzzskie/excel-template-generator.git
cd excel-template-generator
````

### 2. Create and Activate a Virtual Environment

```bash
python -m venv venv
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

---

## ▶️ Run the Program

### Step 1: Format the Excel Template

```bash
python excel_formatter_all.py
```

### Step 2: Generate Dummy Data

```bash
python data_generator2.py
```

---

## 🧩 Notes

* The `venv/` folder and all `.xlsx` files are ignored via `.gitignore`.
* Always activate your virtual environment before running any script.
* Adjust parameters inside `excel_formatter_all.py` and `data_generator2.py` as needed to fit your data structure or format.

---

## 💡 Example Workflow

1. Create your environment and install requirements
2. Run `excel_formatter_all.py` to set up your Excel structure
3. Run `data_generator2.py` to fill in sample/dummy data
4. You’ll find the generated Excel file (e.g. `output.xlsx`) in your project folder

---

