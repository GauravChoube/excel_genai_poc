# Excel Genai POC

## 📌 Overview  
This project provides a Python-based solution to convert the excel file into coding form which fetch data and macros from all sheet.Convert MAcors into desire coding language and map data.
  
---

## 🚀 Features  
- **Automated Data Extraction**: Reads all Excel sheets into a dictionary using `pandas`.  
- **VBA Macro Extraction**: Uses `win32com.client` to fetch VBA code from the Excel workbook.  
- **VBA to Python Conversion**: Translates basic VBA syntax to Python code using open ai[chat-gpt 4o].  
- **Macro Consolidation**: Merges all converted macros into a single Python function.  

---

## 🛠 Installation  
### **Prerequisites**  
Ensure you have Python installed (3.7+ recommended).  

Install dependencies:  
```bash
pip install pandas pywin32 openpyxl
```
---
## Getting Started
### Config File
Setup the config file with openAI credential into file 'config.json'.

### run
```
python main.py
```