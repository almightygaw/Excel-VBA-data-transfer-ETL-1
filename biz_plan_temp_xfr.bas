Attribute VB_Name = "biz_plan_temp_xfr"
' ---------- Business Plan Template Transfer Macro ----------

' This macro copies data from a given FPO Business Development business plan template and pastes it to a given NYP business plan template.

' Macro instructions:
' 1.  Create a copy of the NYP Business Plan Template Excel file in desired directory folder. Make a note of this location, as it will be required for step 5 below.
' 2.  Open any Excel file (a blank file, the source data file, the destination data file, or any other Excel file).
' 3.  Run the macro: Developer tab > Macros > PERSONAL.XLSB!biz_plan_temp_xfr.biz_plan_temp_xfr
' 4.  You will see a prompt: "Source file path:". Copy the file path for the FPO Bus Dev department file, paste it in the text box, and click OK.
' 5.  You will see a second prompt: "Destination file path:". Copy the file path for the NYP template file you created in step 1 above, paste it in the text box, and click OK.

' Known issues:
' * Expense Schedule tab: for Westchester business plans, there are known issues regarding how the business rules differ from other sites, and how these are reflected in the Business Plan Template.
'   Manual review of data on this tab that is specific to Westchester is recommended.
' ------------------------------------------------------------

Sub biz_plan_temp_xfr()
  
  ' Source workbook
  Dim sourceWbPath As String
  sourceWbPath = InputBox("Source file path: ") ' enter source file name
  Dim Source As Workbook
  Set Source = Workbooks.Open(sourceWbPath)
  
  ' Destination workbook (blank template workbook)
  Dim destWbName As String
  destWbName = InputBox("Destination file path: ")    ' enter destination file name
  Application.AskToUpdateLinks = False
  Application.DisplayAlerts = False
  Workbooks.Open (destWbName)
  Dim Destination As Workbook
  Set Destination = Workbooks.Open(destWbName)

  ' ---------------------------------------------------------------------------------------------------
  ' Proposal Package ----------------------------------------------------------------------------------
  Dim PropPack As String
  PropPack = "Proposal Package"
  
  Dim PropPack_Source As Worksheet
  Set PropPack_Source = Source.Worksheets(PropPack)
  Dim PropPack_Destination As Worksheet
  Set PropPack_Destination = Destination.Worksheets(PropPack)
  
  ' cell values to be copied over
  Dim PropPackRange As Variant
  PropPackRange = Array("C12", "C13", "C15", "C18", "C21", "C30", "C31", "C32", "C33", "C34", _
                        "C36", "C37", "C38", "C39", "C41", "C43", "C44", "C46", "C47", "C48", _
                        "C51", "C54", "C57", "C58", "C59", "C60", "C63", "C64", "C67", "C69", _
                        "C71", "B76", "F23", "F24", "F25", "F26", "G23", "G24", "G25", "G26", _
                        "F31", "F32", "F33", "F34", "F36", "E58", "E59", "E60", "F58", "F59", _
                        "F60", "G58", "G59", "G60", "F64", "F65", "G101", "G106", "E111", "G118")
  
  ' loop through cells, assign values
  For Each i In PropPackRange
'    If PropPack_Source.Range(i).HasFormula Then                                  ' no formulas on this tab
'      PropPack_Destination.Range(i).Formula = PropPack_Source.Range(i).Formula
'    Else
      PropPack_Destination.Range(i).Value2 = PropPack_Source.Range(i).Value2
'    End If
  Next i
  ' ---------------------------------------------------------------------------------------------------
  ' Proposal Package & Support Req --------------------------------------------------------------------
  Dim SupportReq As String
  SupportReq = "Proposal Package & Support Req"
  
  Dim SupportReq_Source As Worksheet: Set SupportReq_Source = Source.Worksheets(SupportReq)
  Dim SupportReq_Destination As Worksheet: Set SupportReq_Destination = Destination.Worksheets(SupportReq)
  
  ' cell values to be copied over
  Dim SupportReqRange As Variant
  SupportReqRange = Array("G107", "G112", "G117", "G128")
  
  ' loop through cells, assign values
  For Each i In SupportReqRange
    If SupportReq_Source.Range(i).HasFormula = True Then
      SupportReq_Destination.Range(i).Formula = SupportReq_Source.Range(i).Formula
    Else
      SupportReq_Destination.Range(i).Value2 = SupportReq_Source.Range(i).Value2
    End If
  Next i
  ' ---------------------------------------------------------------------------------------------------
  ' Payor Mix -----------------------------------------------------------------------------------------
  Dim PayorMix As String
  PayorMix = "Payor Mix"
  
  Dim PayorMix_Source As Worksheet: Set PayorMix_Source = Source.Worksheets(PayorMix)
  Dim PayorMix_Destination As Worksheet: Set PayorMix_Destination = Destination.Worksheets(PayorMix)
  
  Dim PayorMixRange As Variant
  PayorMixRange = Array("D7", "E7", "D11", "D12", "D13", "D14")
  
  For Each i In PayorMixRange
'    If PayorMix_Source.Range(i).HasFormula = True Then                             ' returns error because references
'      PayorMix_Destination.Range(i).Formula2 = PayorMix_Source.Range(i).Formula2   ' tab not included in Destination wb
'    Else
      PayorMix_Destination.Range(i).Value2 = PayorMix_Source.Range(i).Value2
 '   End If
  Next i
  ' ---------------------------------------------------------------------------------------------------
  ' Professional RVU Schedule -------------------------------------------------------------------------
  Dim ProfRVUSched As String
  ProfRVUSched = "Professional RVU Schedule"
  
  Dim ProfRVUSched_Source As Worksheet: Set ProfRVUSched_Source = Source.Worksheets(ProfRVUSched)
  Dim ProfRVUSched_Destination As Worksheet: Set ProfRVUSched_Destination = Destination.Worksheets(ProfRVUSched)
  
  Dim profRVUSchedRange As Variant
  profRVUSchedRange = Array("F2", "D5", "E11", "D15", "E15", "F15", "G15", "H15")
  
  For Each i In profRVUSchedRange
'    If ProfRVUSched_Source.Range(i).HasFormula = True Then                              ' returns error because references
'      ProfRVUSched_Destination.Range(i).Formula = ProfRVUSched_Source.Range(i).Formula  ' tab not included in Destination wb
'    Else
      ProfRVUSched_Destination.Range(i).Value2 = ProfRVUSched_Source.Range(i).Value2
'    End If
  Next i
  ' ---------------------------------------------------------------------------------------------------
  ' Professional Revenue Schedule ---------------------------------------------------------------------
  Dim profRevSched As String
  profRevSched = "Professional Revenue Schedule"
  
  Dim ProfRevSched_Source As Worksheet: Set ProfRevSched_Source = Source.Worksheets(profRevSched)
  Dim ProfRevSched_Destination As Worksheet: Set ProfRevSched_Destination = Destination.Worksheets(profRevSched)
  
  Dim ProfRevSchedRange As Variant
  ProfRevSchedRange = Array("C11", "D11", "E11", "F11", "G11")
  
  For Each i In ProfRevSchedRange
'    If ProfRevSched_Source.Range(i).HasFormula = True Then                             ' returns blank because references
'      ProfRevSched_Destination.Range(i).Formula = ProfRevSched_Source.Range(i).Formula ' cells that are empty in Destination wb
'    Else
      ProfRevSched_Destination.Range(i).Value2 = ProfRevSched_Source.Range(i).Value2
'    End If
  Next i
  ' ---------------------------------------------------------------------------------------------------
  ' Expense Schedule ----------------------------------------------------------------------------------
  Dim expenseSched As String
  expenseSched = "Expense Schedule"
  
  Dim ExpenseSched_Source As Worksheet: Set ExpenseSched_Source = Source.Worksheets(expenseSched)
  Dim ExpenseSched_Destination As Worksheet: Set ExpenseSched_Destination = Destination.Worksheets(expenseSched)
  
  Dim ExpenseSchedRange1 As Variant
  ExpenseSchedRange1 = Array("B34:B45", "C68:C71", "C90", "C102:C110", "C118:C120", _
                             "E10:E15", "E21:E24", "E34:F45", "G21:K24", "G71:G72", _
                             "G74:K75", "G77:K82", "G84:G87", "G114:K115", "I71", _
                             "J10", "K71", "M27", "M34:Q45", "M65:N121", "S31:S46", _
                             "H90:K90", "G68", "G21:K24")
  
  For Each i In ExpenseSchedRange1
    If ExpenseSched_Source.Range(i).HasFormula = True Then
      ExpenseSched_Destination.Range(i).Formula2 = ExpenseSched_Source.Range(i).Formula2
    Else
      ExpenseSched_Destination.Range(i).Value2 = ExpenseSched_Source.Range(i).Value2
    End If
  Next i
  
  
  ' 1111 Westchester condition 2: if cell in row M contains "1111", paste hardcoded values in A:L, else ignore
  Dim ExpenseSchedRange4 As Variant
  ExpenseSchedRange4 = Array("M1:M121")
  
  For Each i In ExpenseSchedRange4
    If InStr(i, "1111") > 0 Then
      For j = -11 To -1
        ExpenseSched_Destination.Range(i).Offset(0, j).Value2 = ExpenseSched_Source.Range(i).Offset(0, j).Value2
      Next j
    End If
  Next i
  
  
  ' hardcoded values, per Lindsay
  Dim ExpenseSchedRange5 As Variant
  ExpenseSchedRange5 = Array("E16:E17", "G90", "C73:N73", "C83:N83")
  
  For Each i In ExpenseSchedRange5
    ExpenseSched_Destination.Range(i).Value2 = ExpenseSched_Source.Range(i).Value2
  Next i
  

  ' 1111 Westchester condition 1: if B34 like "1111", paste formulas, else paste values
  Dim ExpenseSchedRange2 As Variant
  ExpenseSchedRange2 = Array("G34:K45")
  
  For Each i In ExpenseSchedRange2
    If InStr(ExpenseSched_Source.Range("B34"), "1111") > 0 Then
      ExpenseSched_Destination.Range(i).Value2 = ExpenseSched_Source.Range(i).Value2
    Else
      ExpenseSched_Destination.Range(i).Formula2 = ExpenseSched_Source.Range(i).Formula2
    End If
  Next i
  
  ' 1111 Westchester condition 1 applied to SUPPORT STAFF SALARY FRINGE section
  Dim ExpenseSchedRange3 As Variant
  ExpenseSchedRange3 = Array("G50:K61")
  
  For Each i In ExpenseSchedRange3
    If InStr(ExpenseSched_Source.Range("B50"), "1111") > 0 Then
      ExpenseSched_Destination.Range(i).Value2 = ExpenseSched_Source.Range(i).Value2
    Else
      ExpenseSched_Destination.Range(i).Formula2 = ExpenseSched_Source.Range(i).Formula2
    End If
  Next i
  
  
  ' --------------------------------------------------------------------------------------------------
  ' Columbia Data (wRVU, Rev, MGMA) ------------------------------------------------------------------
  ' Q:\FPO Business Development\Business Plans\Templates\Business Plan Template Password.docx --------
  ' --------------------------------------------------------------------------------------------------
  Dim columbiaData As String
  columbiaData = "Columbia Data (wRVU, Rev, MGMA)"
  
  Dim columbiaData_Destination As Worksheet: Set columbiaData_Destination = Destination.Worksheets(columbiaData)
  
  Dim pw As String
  pw = "churchbell"
  
  columbiaData_Destination.Activate
  
  ' unprotect worksheet
  columbiaData_Destination.Unprotect Password:=pw
  
  ' unhide all columns
  columbiaData_Destination.Columns.EntireColumn.Hidden = False
  
  ' remove autofilter
  columbiaData_Destination.ShowAllData
  
  ' E11:E13, D16:H18, D22:H24 = sourceWb.Worksheets("Professional RVU Schedule").Range(E24:E26, D29:D31, D35:H37)
  Dim columbiaDataArray1 As Variant
  columbiaDataArray1 = Array("E11", "E12", "E13", _
                             "D16", "D17", "D18", "E16", "E17", "E18", "F16", "F17", "F18", "G16", "G17", "G18", "H16", "H17", "H18", _
                             "D22", "D23", "D24", "E22", "E23", "E24", "F22", "F23", "F24", "G22", "G23", "G24", "H22", "H23", "H24")
                             
  For Each i In columbiaDataArray1
'    If ProfRVUSched_Source.Range(i).Offset(13, 0).HasFormula = True Then
'      columbiaData_Destination.Range(i).Formula = ProfRVUSched_Source.Range(i).Offset(13, 0).Formula
'    Else
      columbiaData_Destination.Range(i).Value2 = ProfRVUSched_Source.Range(i).Offset(13, 0).Value2
'    End If
  Next i
  
  
  ' D30:H32 = sourceWb.Worksheets("Professional Revenue Schedule").Range(C22:G24)
  ' D36:H38 = sourceWb.Worksheets("Professional Revenue Schedule").Range(C28:G30)
  Dim columbiaDataArray2 As Variant
  columbiaDataArray2 = Array("D30", "D31", "D32", "E30", "E31", "E32", "F30", "F31", "F32", _
                             "G30", "G31", "G32", "H30", "H31", "H32")
                             
  For Each i In columbiaDataArray2
'    If ProfRevSched_Source.Range(i).HasFormula = True Then
'      columbiaData_Destination.Range(i).Formula = ProfRevSched_Source.Range(i).Offset(-8, -1).Formula
'    Else
      columbiaData_Destination.Range(i).Value2 = ProfRevSched_Source.Range(i).Offset(-8, -1).Value2
'    End If
  Next i
  
  ' MGMA BENCHMARKING: unhide everything, unprotect
  ' entire section from sourceWb.Worksheets("MGMA Benchmarking") For FPO Use (Autopopulated) section
  ' row 47 headers also need to be copied
  Dim columbiaDataArray3() As String
  ReDim Preserve columbiaDataArray3(0)
  
  ' create array of data to be copied
  For Each i In Array("B", "C", "D", "E", "F", "G", "H", "I")
    For j = 46 To 101
      ReDim Preserve columbiaDataArray3(UBound(columbiaDataArray3) + 1)
      If VarType(Source.Worksheets("MGMA Benchmarking").Range(i & j)) = 10 Then
        columbiaDataArray3(UBound(columbiaDataArray3)) = ""
      Else
        columbiaDataArray3(UBound(columbiaDataArray3)) = Source.Worksheets("MGMA Benchmarking").Range(i & j)
      End If
    Next j
  Next i
  
  ' paste array elements to destination range
  Dim k As Long
  k = 1
  
  For Each i In Array("C", "D", "E", "F", "G", "H", "I", "J")
    For j = 42 To 97
      columbiaData_Destination.Range(i & j).Value2 = columbiaDataArray3(k)
      k = k + 1
    Next j
  Next i
  
  ' replace autofilter
  columbiaData_Destination.Range("A:A").AutoFilter Field:=1, Criteria1:="Yes"
  
  ' hide column A
  columbiaData_Destination.Range("A:A").EntireColumn.Hidden = True
  
  ' re-protect worksheet
  columbiaData_Destination.Protect Password:=pw

  
  ' ---------------------------------------------------------------------------------------------------
End Sub




