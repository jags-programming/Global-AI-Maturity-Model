**Advanced Excel Evaluation Toolkit Structure for GAIMM**

---

### 📁 File Location

Path in repo: `/evaluation/toolkit/GAIMM-Evaluation-Toolkit.xlsx`

---

### 📊 Sheet 1: Instructions

* **Purpose:** Guidance on how to use the toolkit.
* **Contents:**

  * Step-by-step instructions.
  * Explanation of maturity levels.
  * Color legend used for scoring.

---

### 📊 Sheet 2: Response Sheet (Diagnostic Input)

* **Columns:**

  | Pillar | Subdomain/Topic | Question | Response Options | Selected Response | Notes |
  | ------ | --------------- | -------- | ---------------- | ----------------- | ----- |
* **Features:**

  * Drop-down list in "Selected Response" with standardized answer choices.
  * Conditional formatting to color cells by maturity level.
  * Each question mapped back to one of the five levels (Initial, Emerging, Established, Advanced, Leading).

---

### 📊 Sheet 3: Scoring Logic

* **Columns:**
  \| Pillar | Question ID | Response | Maturity Level (Score 1-5) |
* **Purpose:**

  * Lookup table to assign a numeric maturity score (1-5) based on selected responses.
  * Used by other sheets via `VLOOKUP` or `XLOOKUP`.

---

### 📊 Sheet 4: Aggregated Scores

* **Columns:**
  \| Pillar | Total Questions | Score Sum | Average Score | Maturity Level |
* **Formulae:**

  * `Average Score = Score Sum / Total Questions`
  * `Maturity Level = IF(avg<=1.5,"Initial", IF(avg<=2.5,"Emerging", IF(avg<=3.5,"Established", IF(avg<=4.5,"Advanced","Leading"))))`
* **Features:**

  * Pillar-wise breakdown.
  * Conditional formatting for maturity level coloring.

---

### 📊 Sheet 5: Dashboard

* **Elements:**

  * Bar chart: Pillar vs. Average Score.
  * Radar chart: Spider view of all pillars.
  * Score distribution pie chart.
  * Summary banner with overall average maturity.
* **Dynamic visuals:**

  * Auto-update based on the input sheet.

---

### 💡 Optional Enhancements

* **Data Validation & Protection**

  * Lock scoring logic sheet to prevent accidental edits.
* **Macro (optional)**

  * Auto-clear and reset responses.
  * Export score summary as PDF.

