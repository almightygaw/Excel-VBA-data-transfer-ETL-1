# VBA-biz_plan_temp_xfr
copy formulas and hard-coded values from one workbook to another

Business Plan Template Transfer VBA Macro

This macro copies data from a given FPO Business Development business plan template and pastes it to a given NYP business plan template. 

Macro instructions:
1.	Create a copy of the NYP Business Plan Template Excel file in desired directory folder. Make a note of this location, as it will be required for step 5 below.
2.	Open any Excel file (a blank file, the source data file, the destination data file, or any other Excel file).
3.	Run the macro: Developer tab > Macros > PERSONAL.XLSB!biz_plan_temp_xfr.biz_plan_temp_xfr.
4.	You will see a prompt: "Source file path:". Copy the file path for the FPO Bus Dev department file, paste it in the text box, and click OK.
5.	You will see a second prompt: "Destination file path:". Copy the file path for the NYP template file you created in step 1 above, paste it in the text box, and click OK.

Known issues:
1. Expense Schedule tab: for Westchester business plans, there are known issues regarding how the business rules differ from other sites, and how these are reflected in the Business Plan Template. Manual review of data on this tab that is specific to Westchester is recommended.
