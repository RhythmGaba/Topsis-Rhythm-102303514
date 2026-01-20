# TOPSIS Ranking Program  
**Author:** Rhythm Gaba

**Roll No.:** 102303514

**Course:** Predictive Analytics using Statistics

**Method Used:** TOPSIS (Multi-Criteria Decision Making)

---

## Overview  
The program applies the TOPSIS methodology to evaluate and rank different funds using an input Excel dataset. The implementation is designed as a command-line program and includes proper validation checks for inputs such as weights, impacts, and file format.

TOPSIS is a widely used multi-criteria decision-making (MCDM) approach that evaluates different criterias to make a decision using a score, based on which each criteria is ranked that further provides the result to make a decision.

---

## Methodology 

The following steps are performed:

- Construction of the decision matrix from the input Excel file  
- Normalization of criteria values using vector normalization technique  
- Apply user-defined weights to the normalized matrix  
- Identification of ideal best (V⁺) and ideal worst (V⁻) solutions based on impacts  
- Calculation of Euclidean distances from ideal best and ideal worst solutions  
- Calculate TOPSIS performance score for each fund and Rank them. 

---

## Dataset

- The dataset consists of 8 funds and 5 paramters to rank them 
- The input data is provided in Excel (.xlsx) format  
- The first column represents fund names and remaining columns represent criteria  
- All non-numeric and invalid inputs are handled through validation checks  

---

## File Structure  

- `Assignment_1_UCS654.ipynb` – Jupyter / Google Colab notebook containing the full implementation  
- `TOPSIS_code.py` – Python script implementing the TOPSIS algorithm  
- `data.xlsx` – Sample input file  
- `result.xlsx` – Output file containing TOPSIS scores and ranks  
- `README.md` – Project documentation  






