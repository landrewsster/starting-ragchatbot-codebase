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

### [DATE UNKNOWN] Deleted 4 test responses

| Field | Detail |
|-------|--------|
| **Type** | Row deletion |
| **Records affected** | 4 test/pilot submissions |
| **How to identify** | *(Add identifying info here — e.g., record IDs, names, or a unique text string from their responses)* |
| **Change made** | Deleted all 4 rows from the CSV |
| **Reason** | Test submissions, not real respondents |

---

### [DATE UNKNOWN] Recoded screening question for 1 respondent

| Field | Detail |
|-------|--------|
| **Type** | Value recode |
| **Records affected** | 1 respondent |
| **How to identify** | Free-text response containing: *"moving to the next slide clarified that 'screen' includes questioning if use. We ask 100% about any and all drugs used."* |
| **Change made** | **Q: Do you or others in your practice screen all patients who are pregnant or breastfeeding for cannabis use?** — changed response to `Yes, all patients (i.e. universal screening)` |
| **Reason** | Respondent clarified mid-survey they screen 100% of patients; initial response did not reflect actual practice |
| **Also changed** | Cleared their free-text response to "Why don't you or others in your practice screen all patients…" (no longer applicable after recode) |

---

### [DATE UNKNOWN] Other recodes (details unknown)

| Field | Detail |
|-------|--------|
| **Type** | Unknown |
| **Records affected** | Unknown |
| **How to identify** | *(Fill in when you remember or rediscover these)* |
| **Change made** | *(Fill in)* |
| **Reason** | *(Fill in)* |

---

## Template for new entries

Copy and paste this block for each new change:

```
### [DATE] Brief description

| Field | Detail |
|-------|--------|
| **Type** | Row deletion / Value recode / Column edit |
| **Records affected** | N respondent(s) |
| **How to identify** | Record ID: XXX  —OR—  unique text: "…" |
| **Change made** | Column: [column name]  Old value: [X]  New value: [Y] |
| **Reason** | … |
```
