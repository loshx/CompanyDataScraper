# Company Data Extractor  

A Python application with **Tkinter GUI + Selenium** that:
- Searches companies on [data2b.md](https://data2b.md) via query or Google
- Extracts company details (name, IDNO, address, contacts, financials, activities)
- Saves results into **Excel** with structured columns
- Allows extracting **company IDNOs from paginated URLs**

## 🛠️ Tech Stack
- Python 3.x
- Selenium WebDriver
- Tkinter / CustomTkinter
- Pandas
- OpenPyXL

## 🚀 Installation
Clone repo and install dependencies:
```bash
git clone https://github.com/YOUR_USERNAME/company-data-extractor.git
cd company-data-extractor
python -m venv .venv
.venv\Scripts\activate   # Windows
source .venv/bin/activate # Linux/Mac
pip install -r requirements.txt
