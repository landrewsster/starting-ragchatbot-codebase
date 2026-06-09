# Manual Data Cleaning Log
# CRC MDH Project — Cannabis Screening Survey

Record changes made directly to the exported REDCap CSV **before** running
`survey_frequencies.py`. Each time you pull a new datafile, re-apply every
entry below.

---

## How to use this log

When you make a change to the CSV:
1. Add a new entry at the top (most recent first).
2. Include enough detail to **find the record** (a unique text string from a
   free-text field works well) and **reproduce the change**.
3. Note the **reason** so future-you understands why.

---

## Change Log

---

### [DATE UNKNOWN] Recode free-text screening method for Record 247

| Field | Detail |
|-------|--------|
| **Type** | Value recode |
| **Record ID** | 247 (eligible) |
| **How to identify** | Column `Specify.3` contains: *"Part of rooming questions"* |
| **Change made** | Column `Specify.3` → `Standardized assessment tool/questionnaire, care team administered. Please specify:` |
| **Reason** | Free-text clarifies this is a standardized care-team-administered tool |

---

### [DATE UNKNOWN] Recode free-text screening method for Record 143

| Field | Detail |
|-------|--------|
| **Type** | Value recode |
| **Record ID** | 143 (eligible) |
| **How to identify** | Column `Specify.3` contains: *"pt questionnaire in Epic"* |
| **Change made** | Column `Specify.3` → `Standardized assessment tool/questionnaire, patient-administered (e.g., patient fills out form). Please specify:` |
| **Reason** | Free-text clarifies this is a patient-administered questionnaire in Epic |

---

### [DATE UNKNOWN] Recode screening method for Record 217 *(recode uncertain — confirm)*

| Field | Detail |
|-------|--------|
| **Type** | Value recode |
| **Record ID** | 217 (eligible) |
| **How to identify** | Column `Specify.1` contains: *"targeted screening based on patient admitting substance use"* |
| **Change made** | Column `Specify.1` → `Patients with a self-reported history of drug abuse` *(recode marked uncertain — verify before applying)* |
| **Reason** | Free-text description most closely matches the "self-reported history" response option |

---

### [DATE UNKNOWN] Recoded screening question for Record 128

| Field | Detail |
|-------|--------|
| **Type** | Value recode |
| **Record ID** | 128 (eligible) |
| **How to identify** | Column `Specify.1` contains: *"moving to the next slide clarified that 'screen' includes questioning if use. We ask 100% about any and all drugs used. 100%."* |
| **Change made** | **Q: Do you or others in your practice screen all patients who are pregnant or breastfeeding for cannabis use?** → `Yes, all patients (i.e. universal screening)` |
| **Also changed** | Cleared `Specify.1` free-text (no longer applicable after recode) |
| **Reason** | Respondent clarified mid-survey they screen 100% of patients; initial "No" response did not reflect actual practice |

---

### [DATE UNKNOWN] Deleted 4 test responses

| Field | Detail |
|-------|--------|
| **Type** | Row deletion |
| **Records affected** | 4 test/pilot submissions |
| **How to identify** | County free-text field contained `"TEST"` (or similar test indicator) |
| **Change made** | Deleted all 4 rows from the CSV |
| **Reason** | Test submissions, not real respondents |

---

## Template for new entries

Copy and paste this block for each new change:

```
### [DATE] Brief description

| Field | Detail |
|-------|--------|
| **Type** | Row deletion / Value recode / Column edit |
| **Record ID** | XXX (eligible / ineligible) |
| **How to identify** | Column `[col]` contains: *"unique text string"* |
| **Change made** | Column `[col]`  Old value: `[X]`  →  New value: `[Y]` |
| **Reason** | … |
```
