# Project-IuG_OpenAI
Project-IuG_OpenAI


/Users/ranyacherara/Desktop/Gruppe1Bilder

AI in Museums - Prototype

This project is a prototype developed for the Deutsches Technikmuseum Berlin.  
Its goal is to automatically generate **neutral, factual descriptions** of museum objects based on photos, supported by existing metadata from Excel.

Features:
- Processes a **folder or ZIP file of object photos**.  
- Matches photos with museum inventory codes found in the Excel metadata.  
- Generates **short, objective descriptions** of objects using an AI model.  
- Writes results into a **new Excel file** (`descriptions_with_excel.xlsx`).  
- Ensures quality by:
  - Describing **only what is visible** (shape, color, material, markings).  
  - Avoiding guesses, history, or interpretation.  
  - Adding flags if image quality is low or the main object is unclear.  

Project Structure:
