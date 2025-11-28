# Docx Bilingual Auditor 📝

A Python tool designed to automate the formatting comparison between **Chinese** and **English** versions of tender documents (or any bilingual contracts).

It extracts **Bold** and **Underlined** text from `.docx` files, aligns them section-by-section using structural analysis, and generates a side-by-side **HTML Report**.

## 🚀 Features

* **Smart Extraction**: Handles Word structural nuances (merges split runs due to spacing/punctuation).
* **Structure Aware**: Parses documents based on Table of Contents (TOC) logic.
* **Robust Matching**: Uses Fuzzy Matching (Levenshtein distance) to align sections even with minor naming differences.
* **Dual Check**: Audits both **Bold** and **Underline** styles independently.
* **Visual Reporting**: Generates a clean, readable HTML report for quick review.

## 🛠️ Installation

```bash
# Clone the repository
git clone [https://github.com/YOUR_USERNAME/docx-bilingual-auditor.git](https://github.com/YOUR_USERNAME/docx-bilingual-auditor.git)

# Install dependencies
pip install pandas python-docx thefuzz
