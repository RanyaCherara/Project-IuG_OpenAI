# Project-IuG_OpenAI
Project-IuG_OpenAI

AI in Museums (Technisches Museum Berlin)
This repository contains a prototype that generates concise, catalogue-ready descriptions of museum objects by combining an Excel metadata file with a separate folder of images. The system preprocesses images, assembles a structured prompt, sends image + metadata to OpenAI GPT-4.0 Mini, and writes the results back to a new Excel file for curatorial review.
AI in Museums - Prototype

Features:
- Image–metadata fusion: links object photos to their Excel row via inventory numbers / filenames.
- Prompted generation: tightly controlled, catalogue-style descriptions (no fluff, no speculation).
- Structured output: standardized columns in an Excel file, easy to review and edit.
- Fast & cost-aware: uses GPT-4.0 Mini for rapid turnaround at lower cost.
- Reproducible workflow: deterministic preprocessing + clear configuration via .env.

How it works (System Overview): 
1. Selection & Assignment
- A script filters the images folder (or ZIP) to include only objects assigned to our group (by inventory number), creating a consistent dataset for experiments.

2. Preprocessing
- Images are standardized (format/size; optional brightness/contrast normalization) and encoded (base64) to ensure robust, comparable inputs. Prompted Description Generation
For each object, the system sends (a) the preprocessed image and (b) the corresponding Excel metadata to GPT-4.0 Mini with strict instructions for a concise, catalogue-compatible description and a fixed response schema.
Formatting & Evaluation
The script parses the model’s response and writes it into a new Excel file (descriptions_with_excel.xlsx). We perform spot checks against existing metadata to assess quality and feasibility.
