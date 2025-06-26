# 🧱 H-Frame Carport Quote Tool

## 🔧 Overview  
A Python-based automation tool developed to streamline the quoting process for H-frame carport solar structures. This tool calculates structural component requirements based on project-specific parameters and outputs a CSV quote file compatible with SAGE Intacct or Excel-based workflows.

## 🚀 Features  
- **Inputs:**
  - Panel width and height  
  - Number of panels  
  - Row and bay configuration  
  - Tilt angle, bracing option, spacing type, and end bay preference  

- **Calculates:**
  - Beam lengths and spacing  
  - Purlins, brackets, baseplates, and anchor bolts  
  - End bay adjustments and row layout efficiency  
  - Total number of modules placed based on configuration  

- **Outputs:**
  - Timestamped CSV quote file with detailed bill of materials  
  - CSV is structured for direct import into **SAGE Intacct** or for editing in Excel  

## 🧠 Technologies Used  
- Python 3  
- `pandas` – for CSV creation and data handling  
- `math` – for layout and component calculations  
- `datetime` & `os` – for timestamped file generation and organization  

## 📈 Impact  
- Eliminated need for manual BOM calculation in Excel  
- Reduced quoting time from 30–60 minutes to under 15 minutes  
- Improved accuracy and consistency of component estimates  
- Enabled sales engineers to handle more quoting requests faster  

## 📷 Screenshots  
Add visuals here to showcase tool usage and output.

Example suggestions:
- Screenshot 1 – Script running with input prompts  
- Screenshot 2 – Output CSV opened in Excel  
- Screenshot 3 – Sample calculation block or component breakdown

## 🔒 Note  
This demo version has been adapted from an internal tool originally developed at **Lumax Energy**. It has been cleaned and generalized for public portfolio use. All pricing, proprietary data, and client-specific logic have been removed or replaced.
