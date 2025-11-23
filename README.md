# VIP401A_LynnAriss

## Overview
This repository contains various scripts and notebooks related to my work during the semester, focusing on adapting and fine-tuning large language models (LLMs) for medical applications in the field of neuroradiology. Below is a description of each file, highlighting its role within the project's different phases:

## 1. **Lag_Llama.ipynb**
   This notebook contains the implementation and experiments with the **LAG-LLM model**, a time-series forecasting tool. It was tested during the final phase of the project (Phase 4), focusing on predicting trends based on patient visit data. The accuracy achieved was 32.14%, successfully predicting 18 out of 56 patients' visit outcomes.

## 2. **MRI_paragraph.py**
   This script processes and structures **MRI-related data** from the dataset in Phase 3. It filters out MRI-specific data points and generates clean text summaries related to the MRI findings, such as lesion counts, gadolinium use, and the presence of new lesions, to be used in the further analysis and testing phases.

## 3. **Clinical_paragraph.py**
   Similar to the MRI script, this file focuses on **clinical data** in Phase 3. It structures and processes data related to patient diagnosis, co-morbidities, treatment history, and other clinical factors relevant to Multiple Sclerosis (MS). This cleaned and structured clinical data was crucial for evaluating the dataset in Phase 3-2.

## 4. **CSF_paragraph.py**
   This script processes **Cerebrospinal Fluid (CSF) data** within Phase 3. It focuses on structuring CSF-related information, such as OCBS and IgG index levels, which were essential for understanding the patient's condition and for further clinical analysis.

## 5. **paragraphs.py**
   This Python script is responsible for the **initial cleaning and processing** of the data (Phase 1 and Phase 2). It generates structured text summarizing patient demographic information, disease progression, CSF, MRI, and other related data. The output is a collection of text summaries that were then used to create the Excel file for further analysis.

## 6. **Models.xlsx**
   This Excel file is associated with **Phase 2**, where I tested several LLM models using **Zero-shot, Few-shot, and Fine-tuning** techniques. It contains a record of the models evaluated, along with the results for each task. This file was key in documenting which models performed best during the adaptation phase.

## 7. **BioBERT.ipynb**
   This notebook focuses on the **fine-tuning of BioBERT**, a model designed for biomedical text. It was used to fine-tune BioBERT on the medical dataset and evaluate its performance in tasks such as clinical decision-making. The results from this notebook were used in creating the **Models.xlsx** file.

## 8. **MedPaLM_Flan-T5.ipynb**
   This notebook contains the fine-tuning process for **MedPaLM and Flan-T5** models, designed for medical applications. The code was used to train these models on our dataset and test their ability to handle medical queries. These models were also part of the results compiled in the **Models.xlsx** file.

## 9. **TIME-LLM.ipynb**
   This notebook experiments with the **Time-LLM model**, a time-series prediction tool. The model was adapted and fine-tuned to handle time-series data from patient visits, focusing on forecasting patient outcomes based on visit history. The notebook documents the adaptation process for fine-tuning.

## 10. **Lynn_finetunning.ipynb**
   This notebook is dedicated to the **fine-tuning** of various models during **Phase 4** (the trial phase). Specifically, it focuses on fine-tuning models for **NEDA classification** within our dataset. It includes details about the models used, adjustments made during training, and evaluation metrics.

---
Each file is associated with specific tasks throughout the project's phases, from initial data cleaning and model testing (Phase 1 and 2) to fine-tuning models for clinical predictions (Phase 3 and 4). Please refer to each script and notebook for more details on the implementation and results.
