# README for Excel Macros

This README explains the functionality and usage of the two macros: `ExtractColumnsToFilteredWithConditions` and `FilterAndCopyAllRows`. These macros are designed to extract and filter data from the "NSF IPPT" sheet into a "filtered" sheet and then process it further in the "Training Phase" sheet based on specific conditions.

---

## 1. ExtractColumnsToFilteredWithConditions

### Purpose:
The `ExtractColumnsToFilteredWithConditions` macro extracts and filters data from the "NSF IPPT" sheet to the "filtered" sheet based on specific conditions. It checks for conditions in multiple columns and, if they are met, copies relevant information to the "filtered" sheet.

### Conditions:
- **Column M** must have values "Year 1" or "Year 2".
- **Column U** must be "IPPT".
- **Column V** must be "FAIL" or "NOT TAKEN".

### Data to be Copied:
The macro copies the following columns from "NSF IPPT" to "filtered":
- **Column A** -> ID
- **Column B** -> NAME
- **Column C** -> DEPT
- **Column D** -> AGE
- **Column E** -> ADDRESS
- **Column F** -> PHONE
- **Column G** -> EMAIL
- **Column M** -> YEAR TYPE
- **Column I** -> START DATE
- **Column J** -> END DATE
- **Column V** -> REMARKS
- **Column U** -> STATUS

### How to Use:
1. Ensure that your Excel workbook contains a sheet named **"NSF IPPT"** with the appropriate data.
2. The **"filtered"** sheet should already exist in the workbook, or the macro will create it.
3. Run the macro `ExtractColumnsToFilteredWithConditions`.
4. The data will be extracted and filtered into the "filtered" sheet based on the specified conditions.

---

## 2. FilterAndCopyAllRows

### Purpose:
The `FilterAndCopyAllRows` macro processes data in the "filtered" sheet and creates a new sheet named "Training Phase." It adjusts the value in **Column K** based on certain conditions and then copies the relevant rows to the new sheet.

### Process:
1. **Year 1** rows: If the date in **Column I** (Start Date) is earlier than the date in cell **M2** on the "master" sheet, the year type in **Column K** is updated to "Year 2."
2. **Year 2** rows: If the date in **Column J** (End Date) is earlier than the date in cell **M2** on the "master" sheet, the year type in **Column K** is updated to "ORD."

### How to Use:
1. Ensure that your Excel workbook contains a sheet named **"Filtered"** with the appropriate data.
2. Ensure that the **"master"** sheet contains a date value in **cell M2** for comparison.
3. The macro checks for the existence of a sheet named **"Training Phase"**. If it doesn't exist, it will create one.
4. Run the macro `FilterAndCopyAllRows`.
5. The relevant data will be copied from the "filtered" sheet to the "Training Phase" sheet, with the updated **Column K** values.
6. The **Column I** (Start Date) and **Column J** (End Date) will be deleted from the "Training Phase" sheet.

### Summary of Changes:
- A new sheet named **"Training Phase"** is created.
- The value in **Column K** is updated based on conditions.
- Columns **I** and **J** are deleted in the "Training Phase" sheet after the copy.

---

## Notes:
- Make sure that the column references in the code match the actual columns in your sheets.
- The macro assumes that data starts from row 2 in both sheets (with headers in row 1).
- If a sheet required by the macro doesn't exist (e.g., "Training Phase"), it will be created automatically.

---

For any questions or issues, feel free to ask!