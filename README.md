# Excel-VBA-Data-Quality-Checker
An Excel VBA-based Data Quality Checker that automatically identifies duplicate records, missing values, invalid entries, and calculation errors. Designed to improve data accuracy, reduce manual checking, save processing time, and enhance overall data quality through automated validation.
#Project Overview:
The Data Quality Checker is an Excel VBA automation tool designed to identify data inconsistencies, improve accuracy, and reduce manual validation efforts.
Key Features:
•	Detects duplicate records
•	Identifies missing or blank values
•	Checks invalid data entries
•	Detects incorrect calculations
•	Automates data validation
•	Reduces manual checking time
•	Improves data accuracy and reliability
Technologies Used:
•	Microsoft Excel
•	Visual Basic for Applications (VBA)
•	Excel Macros
•	Data Validation
How to Use:
1.	Download the Excel macro-enabled workbook (.xlsm).
2.	Open the file in Microsoft Excel.
3.	Enable macros if the workbook is trusted.
4.	Enter or import the data to be validated.
5.	Run the Data Quality Checker macro.
6.	Review the identified data quality issues.
Business Benefits:
•	Reduces repetitive manual work
•	Improves operational efficiency
•	Minimizes human errors
•	Supports accurate reporting
•	Enhances quality control processes
Project Category:
Excel VBA | Data Quality | Process Automation | Data Validation
Author
Rabi Das

VBA Code 

Option Explicit

Public Sub RunDataQualityChecker()

    Dim ws As Worksheet, rpt As Worksheet, dash As Worksheet
    Dim headers As Object, seen As Object, badRows As Object
    Dim required As Variant, fields As Variant, h As Variant
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long
    Dim reportRow As Long, errors As Long, validCount As Long
    Dim duplicateCount As Long, missingCount As Long
    Dim invalidCount As Long, calcCount As Long
    Dim tx As String, country As String
    Dim status As String, email As String
    Dim qty As Variant, price As Variant
    Dim total As Variant, dateValue As Variant
    Dim oldCalc As XlCalculation
    Dim success As Boolean

    On Error GoTo Fail

    Set ws = ThisWorkbook.Worksheets("Raw_Data")
    Set dash = ThisWorkbook.Worksheets("Dashboard")

    Set headers = CreateObject("Scripting.Dictionary")
    headers.CompareMode = vbTextCompare

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Set badRows = CreateObject("Scripting.Dictionary")

    fields = Array("Transaction_ID", "Customer_ID", _
        "Customer_Name", "Country", "Transaction_Date", _
        "Product", "Quantity", "Unit_Price", _
        "Total_Amount", "Email", "Payment_Status")

    required = Array("Transaction_ID", "Customer_ID", _
        "Customer_Name", "Country", "Transaction_Date", _
        "Quantity", "Unit_Price", "Total_Amount", _
        "Email", "Payment_Status")

    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If Len(Trim$(CStr(ws.Cells(1, c).Value))) > 0 Then
            headers(Trim$(CStr(ws.Cells(1, c).Value))) = c
        End If
    Next c

   
    For Each h In fields
        If Not headers.Exists(CStr(h)) Then
            MsgBox "Missing header: " & CStr(h), vbExclamation
            Exit Sub
        End If
    Next h

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

  
    lastRow = 1

    For Each h In fields
        r = ws.Cells(ws.Rows.Count, _
            CLng(headers(CStr(h)))).End(xlUp).Row
        If r > lastRow Then lastRow = r
    Next h

   
    If lastRow >= 2 Then
        For Each h In fields
            ws.Range(ws.Cells(2, headers(CStr(h))), _
                ws.Cells(lastRow, headers(CStr(h)))).Interior.Pattern = xlNone
        Next h
    End If

    
    Set rpt = GetOrCreateSheet("Error_Report")
    rpt.Cells.Clear

    rpt.Range("A1:F1").Value = Array( _
        "Excel_Row", "Transaction_ID", "Field", _
        "Issue_Type", "Explanation", "Severity")

    reportRow = 2

   
    For r = 2 To lastRow

        tx = Trim$(CStr(ws.Cells(r, _
            headers("Transaction_ID")).Value))

      
        For Each h In required
            If IsBlankValue(ws.Cells(r, headers(CStr(h))).Value) Then

                AddIssue ws, rpt, badRows, r, _
                    headers(CStr(h)), tx, CStr(h), _
                    "Missing", "Mandatory field is blank", _
                    "High", reportRow

                missingCount = missingCount + 1
            End If
        Next h

        
        If Len(tx) > 0 Then
            If seen.Exists(tx) Then

                AddIssue ws, rpt, badRows, r, _
                    headers("Transaction_ID"), tx, _
                    "Transaction_ID", "Duplicate", _
                    "ID already exists in row " & seen(tx), _
                    "High", reportRow

                duplicateCount = duplicateCount + 1

            Else
                seen.Add tx, r
            End If
        End If

      
        country = Trim$(CStr(ws.Cells(r, _
            headers("Country")).Value))

        If Len(country) > 0 Then
            Select Case UCase$(country)
                Case "INDIA", "USA", "CANADA", "UK", "AUSTRALIA"
                Case Else
                    AddIssue ws, rpt, badRows, r, _
                        headers("Country"), tx, "Country", _
                        "Invalid", "Country not approved", _
                        "Medium", reportRow
                    invalidCount = invalidCount + 1
            End Select
        End If

   
        dateValue = ws.Cells(r, headers("Transaction_Date")).Value

        If Not IsBlankValue(dateValue) Then
            If Not ValidISODate(dateValue) Then

                AddIssue ws, rpt, badRows, r, _
                    headers("Transaction_Date"), tx, _
                    "Transaction_Date", "Invalid", _
                    "Invalid date. Expected YYYY-MM-DD", _
                    "High", reportRow

                invalidCount = invalidCount + 1
            End If
        End If

        
        qty = ws.Cells(r, headers("Quantity")).Value

        If Not IsBlankValue(qty) Then
            If Not IsPositiveInteger(qty) Then

                AddIssue ws, rpt, badRows, r, _
                    headers("Quantity"), tx, "Quantity", _
                    "Invalid", "Quantity must be positive integer", _
                    "High", reportRow

                invalidCount = invalidCount + 1
            End If
        End If

       
        price = ws.Cells(r, headers("Unit_Price")).Value

        If Not IsBlankValue(price) Then
            If Not IsNonnegativeNumber(price) Then

                AddIssue ws, rpt, badRows, r, _
                    headers("Unit_Price"), tx, "Unit_Price", _
                    "Invalid", "Unit price must be nonnegative number", _
                    "High", reportRow

                invalidCount = invalidCount + 1
            End If
        End If

       
        total = ws.Cells(r, headers("Total_Amount")).Value

        If Not IsBlankValue(total) Then

            If Not IsNonnegativeNumber(total) Then

                AddIssue ws, rpt, badRows, r, _
                    headers("Total_Amount"), tx, "Total_Amount", _
                    "Invalid", "Total must be nonnegative number", _
                    "High", reportRow

                invalidCount = invalidCount + 1

            ElseIf IsPositiveInteger(qty) And _
                   IsNonnegativeNumber(price) Then

                If Abs(CDbl(total) - CDbl(qty) * CDbl(price)) > 0.01 Then

                    AddIssue ws, rpt, badRows, r, _
                        headers("Total_Amount"), tx, "Total_Amount", _
                        "Calculation mismatch", _
                        "Total differs from Quantity x Unit_Price", _
                        "High", reportRow

                    calcCount = calcCount + 1
                End If
            End If
        End If

        
        email = Trim$(CStr(ws.Cells(r, headers("Email")).Value))

        If Len(email) > 0 Then
            If Not ValidEmail(email) Then

                AddIssue ws, rpt, badRows, r, _
                    headers("Email"), tx, "Email", _
                    "Invalid", "Invalid email format", _
                    "Medium", reportRow

                invalidCount = invalidCount + 1
            End If
        End If

        
        status = Trim$(CStr(ws.Cells(r, _
            headers("Payment_Status")).Value))

        If Len(status) > 0 Then
            Select Case UCase$(status)
                Case "PAID", "PENDING", "FAILED"
                Case Else

                    AddIssue ws, rpt, badRows, r, _
                        headers("Payment_Status"), tx, _
                        "Payment_Status", "Invalid", _
                        "Unapproved payment status", _
                        "Medium", reportRow

                    invalidCount = invalidCount + 1
            End Select
        End If

    Next r

   
    errors = reportRow - 2
    validCount = Application.Max(0, lastRow - 1 - badRows.Count)

    rpt.Rows(1).Font.Bold = True
    rpt.Rows(1).Interior.Color = RGB(31, 78, 121)
    rpt.Rows(1).Font.Color = vbWhite
    rpt.Columns("A:F").AutoFit
    rpt.Range("A1:F1").AutoFilter

    
    dash.Range("A1").Value = "DATA QUALITY CHECKER - RESULTS"
    dash.Range("A2:B12").ClearContents

    dash.Range("A2:B2").Value = Array("Metric", "Value")
    dash.Range("A3:B3").Value = Array("Total Records", lastRow - 1)
    dash.Range("A4:B4").Value = Array("Records With Errors", badRows.Count)
    dash.Range("A5:B5").Value = Array("Valid Records", validCount)
    dash.Range("A6:B6").Value = Array("Duplicate IDs", duplicateCount)
    dash.Range("A7:B7").Value = Array("Missing Values", missingCount)
    dash.Range("A8:B8").Value = Array("Invalid Entries", invalidCount)
    dash.Range("A9:B9").Value = Array("Calculation Mismatches", calcCount)

    dash.Range("A10").Value = "Data Quality Score"

    If lastRow > 1 Then
        dash.Range("B10").Value = validCount / (lastRow - 1)
    Else
        dash.Range("B10").Value = 0
    End If

    dash.Range("B10").NumberFormat = "0.0%"

    dash.Range("A12").Value = "Total Issues Detected"
    dash.Range("B12").Value = errors

    dash.Range("A2:B2").Font.Bold = True
    dash.Columns("A:B").AutoFit

    success = True

CleanUp:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True

    If success Then
        MsgBox "Validation Complete!" & vbCrLf & _
            "Total Issues: " & errors & vbCrLf & _
            "Records With Errors: " & badRows.Count, vbInformation
    End If

    Exit Sub

Fail:
    MsgBox "Validation failed: " & Err.Description, vbCritical
    Resume CleanUp

End Sub


Private Sub AddIssue(ByVal ws As Worksheet, _
    ByVal rpt As Worksheet, ByVal badRows As Object, _
    ByVal r As Long, ByVal c As Long, _
    ByVal tx As String, ByVal field As String, _
    ByVal issue As String, ByVal explanation As String, _
    ByVal severity As String, ByRef reportRow As Long)

    rpt.Cells(reportRow, 1).Resize(1, 6).Value = _
        Array(r, tx, field, issue, explanation, severity)

    ws.Cells(r, c).Interior.Color = RGB(255, 199, 206)

    If Not badRows.Exists(CStr(r)) Then
        badRows.Add CStr(r), True
    End If

    reportRow = reportRow + 1

End Sub


Private Function IsBlankValue(ByVal v As Variant) As Boolean

    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        IsBlankValue = True
    Else
        IsBlankValue = (Len(Trim$(CStr(v))) = 0)
    End If

End Function


Private Function IsPositiveInteger(ByVal v As Variant) As Boolean

    If IsBlankValue(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function

    On Error GoTo InvalidValue
    IsPositiveInteger = (CDbl(v) > 0 And CDbl(v) = Fix(CDbl(v)))

InvalidValue:
End Function


Private Function IsNonnegativeNumber(ByVal v As Variant) As Boolean

    If IsBlankValue(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function

    On Error GoTo InvalidValue
    IsNonnegativeNumber = (CDbl(v) >= 0)

InvalidValue:
End Function


Private Function ValidISODate(ByVal v As Variant) As Boolean

    Dim s As String
    Dim y As Long, m As Long, d As Long
    Dim dt As Date

    On Error GoTo InvalidDate

    If VarType(v) = vbDate Then
        ValidISODate = True
        Exit Function
    End If

    s = Trim$(CStr(v))

    If Len(s) <> 10 Then Exit Function
    If Mid$(s, 5, 1) <> "-" Then Exit Function
    If Mid$(s, 8, 1) <> "-" Then Exit Function

    If Not (Left$(s, 4) Like "####" And _
        Mid$(s, 6, 2) Like "##" And _
        Right$(s, 2) Like "##") Then Exit Function

    y = CLng(Left$(s, 4))
    m = CLng(Mid$(s, 6, 2))
    d = CLng(Right$(s, 2))

    dt = DateSerial(y, m, d)

    ValidISODate = (Year(dt) = y And _
        Month(dt) = m And Day(dt) = d)

InvalidDate:
End Function


Private Function ValidEmail(ByVal s As String) As Boolean

    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")

    re.Pattern = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$"
    re.IgnoreCase = True

    ValidEmail = re.Test(s)

End Function


Private Function GetOrCreateSheet( _
    ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
            ThisWorkbook.Worksheets.Count))

        GetOrCreateSheet.Name = sheetName
    End If

End Function


