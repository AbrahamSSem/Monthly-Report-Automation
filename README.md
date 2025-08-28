# Monthly-Report-Automation
Automated monthly reporting workflow using VBA + Power Query. From 3 days of manual work to just 2 clicks.
# Monthly Report Automation (VBA + PQ)

This is a short project that takes us through the automation process of the monthly reporting process for the job orders and profitability analysis for a small garage workshop.  

Instead of spending **2–3 days** manually compiling, copying and pasting data from multiple sheets, the workflow now runs in just a few seconds with just **2 clicks; Update and Refresh!** 🚀  

The automation uses a mixture of:  
- **VBA** (to loop through multiple sheets in each monthly file and extract data)  
- **Power Query** (to transform, clean and append everything in a single dataset).  

A simple report is then built on top to visualize gross profit over time.  

## Challenge – Before Automation

- Each month a file is received from the workshop via email (sometimes WhatsApp).  
- The file contains multiple sheets (3 sheets per work order) and a summary sheet.  
- Data had to be manually extracted, cleaned, and then compiled to send the report to the workshop manager.  
- Reports often took **2–3 days** to prepare depending on the number of work orders.  

## Solution – After Automation

### 1. VBA Script
- Locates the latest file dropped in the folder and opens it.  
- Checks for the summary sheet (**OVER**) and loops through all the sheets in that monthly file.  
- Extracts required data (e.g., **Car Model**) from each Work Order sheet.  
- Updates the summary sheet with the Car Model and Month in columns.  
- Saves and closes the file automatically.  
- Prompts Power Query to run.  

### 2. Power Query (M Script)
- Uses a **parameterized folder path** to make the query dynamic.  
- Cleans and transforms extracted data (e.g., fixing date formats and ensuring consistency).  
- Appends all files together into a single analysis-ready dataset.  

## Workflow

1. Drop the new monthly file into the designated folder.  
2. Open your report file and click **Update Files**.  
3. Click **Refresh All** (Data tab) to update your report.  

## Report Template

This is a simple report template that shows **Gross Profit over time**.  

## Getting Started

### Prerequisites
- Microsoft Excel with **VBA enabled**.  
- Basic knowledge of **Power Query**.  
- Folder structure: a designated folder where all monthly files will be stored.  

### Setup Instructions
1. Clone this repository or download the VBA and M script files.  
2. Copy the VBA script into a new **Excel Macro-Enabled Workbook** (`.xlsm`).  
3. Adjust the **folder path parameter** in Power Query to point to your designated folder.  
4. Save and close.  

## Notes

The report in this case is **secondary** — the main goal here is the **automation of the data process**.  

This project demonstrates how combining **VBA** and **Power Query** can drastically reduce manual workload.  
