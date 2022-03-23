Attribute VB_Name = "Devo_Macros"
Sub PREP_FILE_FOR_CSV_SAVE()
Attribute PREP_FILE_FOR_CSV_SAVE.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Prep_CSV_File Macro
'
    Range("A1").Select
    Selection.End(xlToRight).Select
    Selection.Offset(0, 1).Select
    Selection.EntireColumn.Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
 
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Selection.EntireRow.Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    Range("A1").Select
    
End Sub
Sub PROCESS_CLASSY_REPORTS()

    Application.ScreenUpdating = False
    'Application.DisplayAlerts = False 'Turn off alerts
    
'    'Change ScreenUpdating, Calculation and EnableEvents
'    With Application
'        CalcMode = .Calculation
'        .Calculation = xlCalculationManual
'        .ScreenUpdating = False
'        .EnableEvents = False
'    End With

    Dim RawPayoutFileSpec As String
    Dim RawDetailsFileSpec As String
    
    RawPayoutFileSpec = ""
    RawDetailsFileSpec = ""
    
'
'   Get Payout File's FileSpec
'
    Set RawPayoutFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With RawPayoutFileDialog
        .AllowMultiSelect = False
        .Title = "SELECT **PAYOUT** REPORT (.CSV FILE)"
        .ButtonName = "Select PAYOUT Report"
        .Filters.Clear
        .Filters.Add ".csv files", "*.csv"
        '.Filters.Add ".xls* files", "*.xls*"
        
        '.Show
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                RawPayoutFileSpec = vrtSelectedItem
            Next vrtSelectedItem
        Else
            'MsgBox "User Pressed Cancel"
            'RawPayoutFileSpec will be ""
        End If
    End With

    If RawPayoutFileSpec = "" Then
        MsgBox "You pressed Cancel. Please try again later."
        Exit Sub
    End If

'
'   Get Details File's FileSpec
'
    Set RawDetailsFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With RawPayoutFileDialog
        .AllowMultiSelect = False
        .Title = "SELECT **DETAILS** REPORT (.CSV FILE)"
        .ButtonName = "Select DETAILS Report"
        .Filters.Clear
        .Filters.Add ".csv files", "*.csv"
        '.Filters.Add ".xls* files", "*.xls*"
        
        '.Show
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                RawDetailsFileSpec = vrtSelectedItem
            Next vrtSelectedItem
        Else
            'MsgBox "User Pressed Cancel"
            'RawDetailsFileSpec will be ""
        End If
    End With

    If RawDetailsFileSpec = "" Then
        MsgBox "You pressed Cancel. Please try again later."
        Exit Sub
    End If
    
    'MsgBox "Payout FileSpec is " & RawPayoutFileSpec
    'MsgBox "Details FileSpec is " & RawDetailsFileSpec
    
'    RawPayoutFileSpec = "C:\Users\MichaelM\Downloads\TEST STRIPE PAYOUT REPORT.csv"
'    RawDetailsFileSpec = "C:\Users\MichaelM\Downloads\TEST STRIPE DETAILS REPORT WITH PAYMENT PROCESSOR ID.csv"
    
    Call ProcessClassyReports(RawPayoutFileSpec, RawDetailsFileSpec)
    
End Sub
Private Sub ProcessClassyReports(PayoutFileSpec As String, DetailsFileSpec As String)

'
' OPEN THE FILES
'
    Dim PayoutWorkbook As Workbook
    Dim DetailsWorkbook As Workbook
    
    ' Open the Payout workbook
    Set PayoutWorkbook = Workbooks.Open(PayoutFileSpec)
    PREP_NEW_PAYOUT_REPORT
    
    Set DetailsWorkbook = Workbooks.Open(DetailsFileSpec)
    PREP_NEW_DETAILS_REPORT

    'Application.ScreenUpdating = True
    
    Dim objCurrPPRID As Variant 'the Payment Processor Reference ID value for the current row in the Details report
    Dim lMatchingRow As Long    'the row (if any) that contains the matching PPRID in the Payout report
    Dim lDetailsRow As Long
    Dim objRangeMatch As Range
    Dim SourceRange As Range
    
    'Activate the Details Workbook
    DetailsWorkbook.Activate
    
    ' Get the last row in the Payment Processor Reference ID column in the Details workbook (Column I)
    lLastRowInPPRIDCol = GetLastRowInColumn("I")
    
    ' Catch the non-founds in the for loop
    On Error Resume Next
    
    ' Iterate through the PPRID's
    For Each objCurrPPRID In DetailsWorkbook.ActiveSheet.Range("I2:I" & lLastRowInPPRIDCol)

        'MsgBox "objCurrPPRID.Value is " & objCurrPPRID.Value
        lDetailsRow = objCurrPPRID.Row
        'MsgBox "current row in DETAILS file is " & lDetailsRow
        
        'See if the current PPRID matches a PPRID in the Payout file
        PayoutWorkbook.Activate
        
        ' Note that the .Row **MUST** be included in the Find() for some reason.
        ' If I don't add the .Row and just save the returned Range object, the Row
        ' value of the returned Range object is **NOT** the row that we matched on!
        ' In fact I think it doesn't even SET the .Row parameter, so that if I say
        ' objRangeMatch = Cell.Find(...) and then lMatchingRow = objRangeMatch.Row,
        ' then lMatchingRow = 0! (not found, even though it's found!)
        '
        ' The below code DOESN'T work! (But should!)
        '
        '        objRangeMatch = Cells.Find(what:=objCurrPPRID.Value, _
        '                    After:=Range("R2"), _
        '                    LookAt:=xlWhole, _
        '                    LookIn:=xlValues, _
        '                    SearchOrder:=xlByRows, _
        '                    SearchDirection:=xlNext, _
        '                    MatchCase:=True)
        '         lMatchingRow = objRangeMatch.Row

        lMatchingRow = 0
        lMatchingRow = PayoutWorkbook.ActiveSheet.Cells.Find(what:=objCurrPPRID.Value, _
                    After:=Range("R2"), _
                    LookAt:=xlWhole, _
                    LookIn:=xlValues, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=True).Row
    
        ' If lMatchingRow is not 0, then we found a match; otherwise no match
        If lMatchingRow <> 0 Then
            
            'MsgBox "lMatchingRow is " & lMatchingRow
            
            'Copy from the Details report to the Payout report
            Set SourceRange = DetailsWorkbook.ActiveSheet.Range("F" & lDetailsRow & ":H" & lDetailsRow)
            SourceRange.Copy
            PayoutWorkbook.ActiveSheet.Range("T" & lMatchingRow & ":V" & lMatchingRow).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Else
'            MsgBox "No Match for Details PPRID of """ & objCurrPPRID.Value & """ on row " & lDetailsRow
'            MsgBox "Err.Source = " & Err.Source & ", Err.Number = " & Err.Number
'            MsgBox "Err.Description is " & Err.Description
'            MsgBox "Matching Row: " & lMatchingRow
        End If
        
    Next objCurrPPRID
    
    'Exit error checking
    On Error GoTo 0
    
    'Wrap the text in the Payout report's Reference field (Column V)
    Columns("V:V").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True  'Wrap the text. There can be long text in this column's cells.
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    ' Close the details workbook
    DetailsWorkbook.Close False

'
' Display "Save this file" message
'
    PayoutWorkbook.Activate
    Range("A2").Select
    Application.ScreenUpdating = True
    MsgBox "Now save this file with a meaningful name as an Excel file (.xls or .xlsx) to preserve formatting.", vbOKOnly, "SAVE THIS FILE"
    
End Sub

Private Sub PREP_NEW_PAYOUT_REPORT()

' The two payout report types -- Stripe and PayPal -- have their columns in different orders, so that's
' one way to tell (and that's the way I'm currently using). So they both have the columns "Payment Processor Reference ID" and "Payment Type"
' but they're in different columns in each report. Also, it appears that for PayPal Payout Reports, Payment Type is always "paypal" but for
' Stripe the values can be "credit_card" or "ach". But I haven't had that much PayPal data yet so it's hard to be certain that "paypal" is
' the only value that can show up there.
'
' This subroutine uses the "column location" method to determine whether we're a PayPal or Stripe payout report.
'
    If Range("C1").Value = "Payment Processor Reference ID" Then
        'MsgBox ("Payout Report Type is PayPal")
        Prep_New_PayPal_Payout_Report
    Else
        If Range("E1").Value = "Payment Processor Reference ID" Then
            'MsgBox ("Payout Report Type is Stripe")
            Prep_New_Stripe_Payout_Report
        Else
            MsgBox ("Invalid Payout File Format. Please verify this is truly a Payout File in CSV format. Exiting.")
            End
        End If
    End If

End Sub
Private Sub Prep_New_Stripe_Payout_Report()
'
' It just deletes unnecessary columns in the Stripe-type payout report
'
    Columns("AU:AV").Select
    Selection.Delete Shift:=xlToLeft

    Columns("AP:AQ").Select
    Selection.Delete Shift:=xlToLeft

    Columns("AN:AN").Select
    Selection.Delete Shift:=xlToLeft

    Columns("AI:AL").Select
    Selection.Delete Shift:=xlToLeft

    Columns("AB:AG").Select
    Selection.Delete Shift:=xlToLeft

    Columns("Y:Y").Select
    Selection.Delete Shift:=xlToLeft

    Columns("M:O").Select
    Selection.Delete Shift:=xlToLeft

    Columns("F:K").Select
    Selection.Delete Shift:=xlToLeft

    Columns("C:D").Select
    Selection.Delete Shift:=xlToLeft

    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft

'Finish Up

    Finish_Payout_Report ("Stripe")

End Sub
Private Sub Prep_New_PayPal_Payout_Report()
'
' It just deletes unnecessary columns in the Paypal-type payout report
'

' Delete unnecessary columns
'
    Columns("AQ:AR").Select
    Selection.Delete Shift:=xlToLeft

    Columns("AL:AL").Select
    Selection.Delete Shift:=xlToLeft

    Columns("AG:AJ").Select
    Selection.Delete Shift:=xlToLeft

    Columns("Z:AE").Select
    Selection.Delete Shift:=xlToLeft

    Columns("W:W").Select
    Selection.Delete Shift:=xlToLeft

    Columns("K:M").Select
    Selection.Delete Shift:=xlToLeft

    Columns("D:I").Select
    Selection.Delete Shift:=xlToLeft

    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    
'Finish Up

    Finish_Payout_Report ("PayPal")

End Sub

Private Sub Finish_Payout_Report(PaymentTypeStr)

'
' Reorder some columns
'

' Move Payment Processor Reference ID column to the end (It's now at Column B)

    ' Cut Payment Processor Reference ID column
    Columns("B:B").Select
    Selection.Cut
    
    'Move to last column and paste Payment Processor Reference ID column *AFTER* it
    Range("B1").Select
    Selection.End(xlToRight).Select
    Selection.Offset(0, 1).Select
    ActiveCell.EntireColumn.Select
    Selection.Insert Shift:=xlToRight

' Move Transaction Create Date

    Columns("S:S").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    
' Move Supporter Name

    Columns("M:M").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight

'
' Format the column headers so far
'
    Range("A1").Select
    Header_Setup
    Range("A1").Select
    Highlight_Selection_Yellow
      
'
' Add some column headers at the end
'
    'Move to the column just AFTER the last used column
    Range("A1").Select
    Selection.End(xlToRight).Select
    
    Selection.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "Donor is Anon (only if Yes)"

    Selection.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "Donor Phone Number"

    Selection.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "Raiser's Edge Reference Field"

'
' Color the added column headers with the correct colors
'
    'Color these fields' headers orange
   Range("U1:W1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
'
' Format the added column headers
'
    Range("U1").Select
    Header_Setup
    
'
' Format all the column data
'
    'Format Transaction Created Date and Payout Created Date columns
    Columns("A:B").Select
    Selection.NumberFormat = "mm/dd/yy;@"
    
    'Format money columns with Accounting format
    Columns("C:D").Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
'
' COMBINE "Billing Address" and "Billing Address 2" columns
'
    'Get the last used cell in the "Billing Address" column. We assume here that if there
    'is nothing in the "Billing Address" column, there is likewise nothing in the
    '"Billing Address 2" column, and thus there is no reason to find the last used cell
    'in THAT column.
    lLastRow = GetLastRowInColumn("I")
    For Each cell In Range("I2:I" & lLastRow)
        cell.Value = cell.Value & " " & cell.Offset(0, 1).Value
    Next cell
    
    'Delete "Billing Address 2" column (because it's no longer needed)
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    
'
' Format zip-code column (Column L). The problem is that some zip codes begin with a zero, and
' by default Excel handles these poorly (it automatically treats "00009" like a number and so imports
' it as simply "9". We'll fix that now.
'
    ' Format the zip-code column with the zip-code type. Note that simply formatting the whole column as Special->Zip + 4
    ' simply makes Excel DISPLAY the leading-zero zip codes correctly but the underlying data (a number) remains wrong.
    ' Thus if the user goes to copy that data out of Excel, the leading zeroes will still be stripped. That's why we
    ' can't just do that.
    
    ' NOTE also that this code ASSUMES that Classy outputs a "00009" zip code as "00009" without any apostrophe (all of which
    ' appears to be the case in all of the data files I've looked at. Classy never outputs an apostrophe (probably because not
    ' all CSV-processing programs require it like Excel), and without that apostrophe Excel will treat it as a number, not text.
    ' If Excel sees an "un-apostrophied" five digit zip with any leading zeroes, it will strip the leading zeroes (so that "00009"
    ' becomes "9"). However, if Excel sees a DASH in there for a zip plus five, it will blessedly leave it alone (so "00009-0007" is left as-is).
    ' NOTE HOWEVER that with no padding, a zip code like "9-9857" IRRETRIEVABLY becomes a date code
    ' (something like "9 Sept 2057" which translated back becomes "2775326" or something like that and the ORIGINAL VALUE IS PERMANENTLY
    ' LOST. There is literally NOTHING we can do (as the VBA writer) to fix that because Excel does that automatically upon opening
    ' the file. The only thing I could do would be to write VBA code to use Excel's "Get Data" and "Transform" functions, which introduces
    ' more coding and more processing time and so I will put that off unless it's absolutely necessary. So far, Classy CSV zips seem to always
    ' have ALL required digits in them, which is cool.
    '
    ' I have seen one more kind of zip code -- from Canada -- but it has letters in it and so is treated like text by default. Here's the
    ' format (from Wikipedia's "Postal codes in Canada" entry):
    '
    '       Like British, Irish and Dutch postcodes, Canada's postal codes are alphanumeric. They are in the format A1A 1A1,
    '       where A is a letter and 1 is a digit, with a space separating the third and fourth characters.
    
    ' Get last used row in zip code column
    lLastRow = GetLastRowInColumn("L")
    
    For Each cell In Range("L2:L" & lLastRow)
    
        ' We only care if Excel has converted the zip code to a number (and thus is not treated as text) AND if that
        ' converted number has fewer than 5 digits
        If Not WorksheetFunction.IsText(cell.Value) Then
        
            'If the number has fewer than 5 digits, make it text and pad it as necesary to get 5 digits
            If cell.Value < 10000 Then
                'MsgBox ("cell value is less than 10000")
                cell.NumberFormat = "@" 'MUST do this before the call to Text()
                cell.Value = WorksheetFunction.Text(cell, "00000")
            End If
        End If
    Next cell
    
'
' FORMAT COLUMN WIDTHS
'
' AutoFit and check that minimum column widths have been met. AutoFit sets column widths based only
' on the cell of maximum length in the DATA rows, not the headers. It doesn't care if
' words inside the headers get wrapped around mid-word. We don't want that.
'
    ' AutoFit the column data
    Range("A1:Z1001").Select  'Assumes Z is the last used column and 1001 is the last used row
    Selection.Columns.AutoFit

    ' Adjust "Gross Transaction Amount" column width
    Columns("C:C").Select
    If (Selection.ColumnWidth < 10.43) Then Selection.ColumnWidth = 10.43
    
    ' Adjust "Payment Frequency" column width
    Columns("Q:Q").Select
    If (Selection.ColumnWidth < 10.29) Then Selection.ColumnWidth = 10.29

    ' Adjust "Payment Processor Reference ID" column width
    Columns("S:S").Select
    If (Selection.ColumnWidth < 10) Then Selection.ColumnWidth = 10

    ' Adjust "Donor Phone Number" column width
    Columns("U:U").Select
    Selection.ColumnWidth = 14
    
    ' Adjust "Raiser's Edge Reference Field" column width
    Columns("V:V").Select
    If (Selection.ColumnWidth < 80) Then Selection.ColumnWidth = 80

   
'
' SORT FIRST BY "Payment Create Date" AND SECOND BY "Transaction Create Date"
'
    Call Double_Sort_In_Ascending_Order("B", "A")
    
'
' GROUP ROWS BY PAYOUT CREATED DATE (COLUMN B)
'
    Range("B2").Select
    Group_Rows_By_Date

'
' SUM THE GROUPED ROWS' GROSS AND NET AMOUNTS
'
    Sum_Grouped_Totals

'
' Move cursor to A2
'
    Range("A2").Select 'There's still the Details report to process, so this probably doesn't matter
    
End Sub
Private Sub Sum_Grouped_Totals()

' Add totals for the gross and net amounts, highlight with a green line

    Dim SumStartRow As Long
    Dim SumEndRow As Long
    Dim LastRowInColumn As Long
    Dim FindSum As Boolean
    Dim CurrentRow As Long
    Dim strGrossAmtRange As String
    Dim strNetAmtRange As String
    FindingSum = True
    SumStartRow = 2
    SumEndRow = 2
    LastRowInColumn = 2
    CurrentRow = 1
    strGrossAmtRange = ""
    strNetAmtRange = ""
    
' Gross amount starts at C2, net amount starts at D2
' Start by getting last used row in Column C. That should also be the last
' used row in Column D.
    LastRowInColumn = GetLastRowInColumn("C")

    For Each cell In Range("C2:C" & LastRowInColumn + 1)
        
        CurrentRow = CurrentRow + 1 'we're at the next row
        
        If IsEmpty(cell) Then
            If FindingSum Then 'we hit an empty cell while looking for the next sum location
            
                'Set the locations of the gross sum and net sum
                strGrossAmtRange = "C" & CurrentRow & ":C" & CurrentRow
                strNetAmtRange = "D" & CurrentRow & ":D" & CurrentRow
                
                'Sum and boldface the gross amount
                Range(strGrossAmtRange).Select
                'Selection.Value = WorksheetFunction.Sum(Range("C" & SumStartRow & ":C" & SumEndRow))
                Selection.Formula = "= SUM(C" & SumStartRow & ":C" & SumEndRow & ")"
                Selection.Font.Bold = True
                
                'Sum and boldface the net amount
                Range(strNetAmtRange).Select
                'Selection.Value = WorksheetFunction.Sum(Range("D" & SumStartRow & ":D" & SumEndRow))
                Selection.Formula = "= SUM(D" & SumStartRow & ":D" & SumEndRow & ")"
                Selection.Font.Bold = True
                
                'Highlight the entire row in green
                Selection.EntireRow.Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 5287936
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                
                'Sum row found and created. Now we're looking for a new sum start
                SumStartRow = SumEndRow = 0
                FindingSum = False
                
            Else 'we found an empty cell, but we were looking for the *start* of a new sum, so do nothing
            
            End If
            
        Else 'This cell contains data
            
            'This cell contains data. If we're not already looking for a sum, than this is our start row
            If Not FindingSum Then
                SumStartRow = CurrentRow
                SumEndRow = CurrentRow
                FindingSum = True
                
            Else 'We found a data containing cell, and we're already looking for a sum, so increment the end row
                SumEndRow = CurrentRow
                
            End If
            
        End If
        'Exit Sub
   
    Next cell
    
    'Autofit the totals columns
    Columns("C:D").Select
    Selection.Columns.AutoFit

End Sub
 
Private Function Add_Formulas(PaymentType)
'
'
' NOTE!!!! HORRIBLE MICROSOFT STRIKES AGAIN! ***IN YOUR FORMULA, THE EQUALS SIGN MUST BE THE VERY FIRST CHARACTER AFTER THE
' QUOTE MARK OR MICROSOFT WILL INSERT ITS OWN APOSTROPHE AND MARK THE WHOLE CELL AS TEXT!!!! NOWHERE IS THIS DOCUMENTED. A SINGLE
' SOLITARY SPACE BEFORE YOUR EQUAL SIGN WILL CAUSE MICROSOFT TO INTERPRET THE WHOLE STRING -- EVEN THOUGH THE COMMAND ITSELF SAYS
' ***FORMULA!!!!**** -- AS TEXT. IT DOESN'T DO THAT INTERACTIVELY -- ONLY IN VBA. INTERACTIVELY, I CAN ADD AS MANY SPACES AS I WANT
' BEFORE THE EQUAL SIGN.
'
' EXAMPLE OF THE LUNACY:
'
'   ActiveCell.Formula = "= R2"  : YOU GET A FORMULA, AS INTENDED. CELL CONTAINS [= R2]
'   ActiveCell.Formula = " = R2" : YOU GET !@#$ TEXT!!!! CELL CONTAINS [' = R2] !!!!

    Range("W2").Select
    'ActiveCell.Formula = "= ""Classy "" & R2 & "" gift via Stripe."" & IF(NOT(ISBLANK(X2)), "" "" & X2 & """", """") & IF(NOT(ISBLANK(Y2)), "" "" & Y2 & ""."", """") & IF(NOT(ISBLANK(Z2)), "" "" & Z2 & ""."", """") & IF(NOT(ISBLANK(AA2)), "" "" & AA2 & ""."", """")"
    ActiveCell.Formula = "= ""Classy "" & R2 & "" gift via " & PaymentType & "."" & IF(NOT(ISBLANK(X2)), "" "" & X2 & """", """") & IF(NOT(ISBLANK(Y2)), "" "" & Y2 & ""."", """") & IF(NOT(ISBLANK(Z2)), "" "" & Z2 & ""."", """") & IF(NOT(ISBLANK(AA2)), "" "" & AA2 & ""."", """")"
    
' Copy the formula down the column
'
    Range("W2").Select
    Selection.Copy
    'Range("W3").Select
    'Range(Selection, Selection.End(xlDown)).Select
    Range("W3:W1000").Select
    ActiveSheet.Paste
   
End Function

Private Function Add_Details_Report_Reference_Field(lStartRow As Long, strTargetCol As String, strFreqCol As String, strPymtProcCol As String, _
                                             strDedTypeCol As String, strDedNameCol As String, _
                                             strDedMsgCol As String, strDonorCmtCol As String)
    
    Dim strFreqCell As String
    Dim strPymtProcCell As String
    Dim strDedTypeCell As String
    Dim strDedNameCell As String
    Dim strDedMsgCell As String
    Dim strDonorCmtCell As String
    Dim lEndingRow As Long
    
    ' Get the ending row for our target column
    lEndingRow = GetLastRowInColumn(strFreqCol)
    
    ' Selecting the starting cell for our target column
    ' Then we'll work our way down
    For lCurrRow = lStartRow To lEndingRow
    
        ' get relevant cell names
        strFreqCell = strFreqCol & lCurrRow
        strPymtProcCell = strPymtProcCol & lCurrRow
        strDedTypeCell = strDedTypeCol & lCurrRow
        strDedNameCell = strDedNameCol & lCurrRow
        strDedMsgCell = strDedMsgCol & lCurrRow
        strDonorCmtCell = strDonorCmtCol & lCurrRow
        strCurrCell = strTargetCol & lCurrRow 'get name of current cell that we're populating
        
        Range(strCurrCell).Select 'select the current cell
    
        '
        ' Add the initial description: "Classy one-time/recurring gift via Stripe/PayPal. "
        '
        ActiveCell.Value = "Classy " & Range(strFreqCell).Value & " gift via " & Range(strPymtProcCell).Value
        
        '
        ' Add the dedication, if any
        '
        If Not (IsEmpty(Range(strDedTypeCell).Value)) Then
            ActiveCell.Value = ActiveCell.Value & " " & Range(strDedTypeCell).Value & " " & Range(strDedNameCell).Value & "."
            
            If Not (IsEmpty(Range(strDedMsgCell).Value)) Then ActiveCell.Value = ActiveCell.Value & " " & Range(strDedMsgCell).Value & "."
        Else: ActiveCell.Value = ActiveCell.Value & "."
        End If
        
        '
        ' Add the donor's comment, if any
        '
        If Not (IsEmpty(Range(strDonorCmtCell).Value)) Then
            ActiveCell.Value = ActiveCell.Value & " " & Range(strDonorCmtCell).Value & "."
        End If
        
    Next lCurrRow
        
End Function

Sub BASIC_HEADER_SETUP()
Attribute BASIC_HEADER_SETUP.VB_ProcData.VB_Invoke_Func = "j\n14"

' Keyboard Shortcut: Ctrl+j
'
' Sets up header row from current selection point and colors it yellow
'

' Save selected starting cell
    Dim StartingCell As Range
    Set StartingCell = Selection
    
' Setup the header
    Header_Setup
    
' Highlight the same cells in yellow
    StartingCell.Select
    Highlight_Selection_Yellow
    
End Sub

Private Sub Header_Setup()

' Save selected starting cell

    Dim StartingCell As Range
    Set StartingCell = Selection

' Set header cell text to bold
'
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
   
' Put full borders around header cells
'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
'
' Set row height to 100
'
    Selection.EntireRow.Select
    Selection.RowHeight = 100
    StartingCell.Select
    
' Autofit the column widths
'
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireColumn.Select
    Selection.Columns.AutoFit
    StartingCell.Select
'
' Freeze Panes
'
    StartingCell.Select
    Selection.Offset(1, 0).Select
    With ActiveWindow
       .SplitColumn = 0
       .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    StartingCell.Select

End Sub

Private Sub Highlight_Selection_Yellow()

' Save selected starting cell

    Dim StartingCell As Range
    Set StartingCell = Selection

' Set cell from this point to ending cell (presumably a header or portion thereof) to yellow background color
'
    Range(Selection, Selection.End(xlToRight)).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    StartingCell.Select

End Sub

Private Sub PREP_NEW_DETAILS_REPORT()
'
' Prep_New_Details_Report Macro
' Takes a downloaded details report and removes unnecessary columns and reorders the remaining columns
' and formats everything nicely

Dim strPhoneNum As String
Dim iStrLength As Integer

'
'  FORMAT COLUMNS
'
    'FORMAT TRANSACTION DATE COLUMN (COLUMN A)
    Columns("A:A").Select
    Selection.NumberFormat = "mm/dd/yy;@"
    
    'FORMAT FREQUENCY COLUMN (COLUMN D)
    ' Just need to convert all occurrences of "one_time" to "one-time"
    Columns("D:D").Select
    Selection.Replace "one_time", "one-time"
    
    'FORMAT MONEY COLUMNS WITH ACCOUNTING FORMAT
    Columns("E:F").Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    '
    ' FORMAT PHONE NUMBER COLUMN
    '
        ' Format the whole column as phone numbers
        Columns("K:K").Select
        Selection.NumberFormat = "[<=9999999]###-####;(###) ###-####"
    
        ' Get last used row in "Donor Phone Number" column
        lLastRow = GetLastRowInColumn("K")
        
        Dim objRange As Range
        
        For Each objRange In Range("K2:K" & lLastRow)
            objRange.Value = Replace(objRange.Value, " ", "", , , vbBinaryCompare)
            objRange.Value = Replace(objRange.Value, "-", "", , , vbBinaryCompare)
            objRange.Value = Replace(objRange.Value, ".", "", , , vbBinaryCompare)
            
            ' Remove the leading "1", if any
            If Left(objRange.Value, 1) = "1" Then
                iStrLength = Len(objRange.Value)
                objRange.Value = Right(objRange.Value, iStrLength - 1)
            End If
        Next objRange
        
    ' FORMAT THE "Dedication Message" COLUMN (COLUMN I)
    
    'Remove junk, trim leading & trailing whitespace, and remove any final period
    ' We assume data begins at I2 but we need to find out where it ends
        lLastRow = GetLastRowInColumn("I")
        For Each cell In Range("I2:I" & lLastRow)
            cell.Value = CleanString(cell.Value)
            cell.Value = Trim(cell.Value)
            If Right(cell.Value, 1) = "." Then cell.Value = Left(cell.Value, Len(cell.Value) - 1)
        Next cell
    
    ' FORMAT THE "Donor Is Anonymous" COLUMN (COLUMN J)
    ' Just need to remove all "FALSE" and convert all "TRUE" to "Yes"
    Columns("J:J").Select
    
    'No Need for false; we don't trust Classy to be current for a false answer
    '(turning anonymity off). If a donor's who is currently marked as anonymous
    'wants to be public again (a very rare occurrence), they'll contact us
    'personally
    Selection.Replace "FALSE", ""
    
    'Convert all "TRUE" to "Yes". This is just for ImportOMatic. A "Yes" is what
    'ImportOMatic expects in this field.
    Selection.Replace "TRUE", "Yes"
    
    ' FORMAT THE "Donor's Comment" COLUMN (COLUMN L)
    
    'Remove junk, trim leading & trailing whitespace, and remove any final period
    ' We assume data begins at L2 but we need to find out where it ends
    lLastRow = GetLastRowInColumn("L")
    For Each cell In Range("L2:L" & lLastRow)
        cell.Value = CleanString(cell.Value)
        cell.Value = Trim(cell.Value)
        If Right(cell.Value, 1) = "." Then cell.Value = Left(cell.Value, Len(cell.Value) - 1)
    Next cell
    
    ' FORMAT THE "Payment Processor" COLUMN (COLUMN M)

    ' We assume data begins at M2 but we need to find out where it ends
    lLastRow = GetLastRowInColumn("M")
    For Each cell In Range("M2:M" & lLastRow)
        ' Reformat "classy_pay_powered_by_stripe" to "Stripe"
        ' and reformat "paypal_commerce" to "PayPal"
        If cell.Value = "classy_pay_powered_by_stripe" Then
            cell.Value = "Stripe"
        ElseIf cell.Value = "paypal_commerce" Then
            cell.Value = "PayPal"
        End If
    Next cell
    
'
' ADD A COLUMN FOR OUR Raiser's Edge "Reference" FIELD, WHICH WE'LL NOW GENERATE
'
    Range("O1").Select
    ActiveCell.FormulaR1C1 = _
        "Raiser's Edge Reference Field (combo of ""Classy"" + Frequency + ""gift via"" + ""PayPal"" or ""Stripe"" + Dedication Type, Dedication Name, Dedication Msg, and Donor's Comment)"

    Call Add_Details_Report_Reference_Field(2, "O", "D", "M", "G", "H", "I", "L")
    
'
' Format all column headers with boldface, borders, and background color
'
    Range("A1").Select
    Header_Setup
    Range("A1").Select
    Highlight_Selection_Yellow
    
'
' Check that minimum column widths have been met. AutoFit sets column widths based only
' on the cell of maximum length in the data rows, not the headers. It doesn't care if
' words inside the headers get wrapped around mid-word. We don't want that.
'
    ' Adjust "Transaction Date" column width
    Columns("A:A").Select
    If (Selection.ColumnWidth < 10.14) Then Selection.ColumnWidth = 10.14
    
    ' Adjust "Transaction Status" column width
    Columns("C:C").Select
    If (Selection.ColumnWidth < 10.29) Then Selection.ColumnWidth = 10.29
    
    ' Adjust "Frequency" column width
    Columns("D:D").Select
    If (Selection.ColumnWidth < 9.43) Then Selection.ColumnWidth = 9.43
    
    ' Adjust "Gross Transaction Amount" column width
    Columns("E:E").Select
    If (Selection.ColumnWidth < 10.43) Then Selection.ColumnWidth = 10.43
    
    ' Adjust "Net Transaction Amount" column width
    Columns("F:F").Select
    If (Selection.ColumnWidth < 10.43) Then Selection.ColumnWidth = 10.43
    
    ' Adjust "Dedication Message" column width
    Columns("I:I").Select
    If (Selection.ColumnWidth < 36.86) Then Selection.ColumnWidth = 30
    
    ' Adjust "Donor is Anonymous" column width
    Columns("J:J").Select
    If (Selection.ColumnWidth < 10.86) Then Selection.ColumnWidth = 10.86
    
    ' Adjust "Donor's Comment" column width
    Columns("L:L").Select
    If (Selection.ColumnWidth < 36.86) Then Selection.ColumnWidth = 30

    ' Adjust "Payment Processor" column width
    Columns("M:M").Select
    If (Selection.ColumnWidth < 9) Then Selection.ColumnWidth = 9

    ' Adjust "Raiser's Edge Reference Field" column width
    Columns("O:O").Select
    If (Selection.ColumnWidth > 80) Then Selection.ColumnWidth = 80
    ' Turn on Word Wrap
    Columns("O:O").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

'
' REMOVE "PRE-COMBINED" COLUMNS NOW THAT THEY'VE ALL BEEN COMBINED INTO THE
' REFERENCE COLUMN
'
' Remove "Donor's Comment" and "Payment Processor"
    Columns("L:M").Select
    Selection.Delete Shift:=xlToLeft
'
' Remove "Dedication Type", "Dedication Name", and "Dedication Message"
    Columns("G:I").Select
    Selection.Delete Shift:=xlToLeft
'
' Remove "Frequency"
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft

'
' Move "Reference" column from Column I to Column H
'
    Columns("I:I").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight

'
' SORT BY TRANSACTION DATE (COLUMN A) IN OLDEST-TO-NEWEST ORDER
'
    Sort_In_Ascending_Order ("A")

'
' GROUP ROWS BY TRANSACTION DATE (COLUMN A)
' Disabled for now because we're just going to copy a few columns from this worksheet
' and paste into the Payout report.
'
'    Range("A2").Select
'    Group_Rows_By_Date
'
'
' End with cursor on A2
'
 Range("A2").Select
    
End Sub
Private Function Sort_In_Ascending_Order(SortColumn As String)
'
' NOTE: THIS METHOD ASSUMES THAT THE FIRST USED DATA ROW IS ROW 2 (BECAUSE IT ASSUMES THAT ROW 1 IS A HEADER ROW
'       AND THAT THE DATA FOLLOWS IMMEDIATELY AFTER)
'
    Dim LastCol As Long
    Dim LastRow As Long
    
    LastCol = GetLastCol(ActiveSheet)
    LastRow = GetLastRow(ActiveSheet)
    FirstKeyCell = SortColumn & "2"
    LastKeyCell = SortColumn & LastRow
    
    ColumnLetter = GetColumnLetterFromColumnNumber(LastCol)
    LastRangeCell = ColumnLetter & LastRow
    
    Application.ActiveSheet.Sort.SortFields.Clear
    
    Application.ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        FirstKeyCell & ":" & LastKeyCell), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal

    With Application.ActiveSheet.Sort
        .SetRange Range("A1:" & LastRangeCell)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Function
Private Function Double_Sort_In_Ascending_Order(FirstSortColumn As String, SecondSortColumn As String)
'
' NOTE: THIS METHOD ASSUMES THAT THE FIRST USED DATA ROW IS ROW 2 (BECAUSE IT ASSUMES THAT ROW 1 IS A HEADER ROW
'       AND THAT THE DATA FOLLOWS IMMEDIATELY AFTER)
'
    Dim LastCol As Long
    Dim LastRow As Long
    
    LastCol = GetLastCol(ActiveSheet)
    LastRow = GetLastRow(ActiveSheet)
    FirstSort_StartingKeyCell = FirstSortColumn & "2"      'Here's that assumption discussed above
    FirstSort_EndingKeyCell = FirstSortColumn & LastRow
    SecondSort_StartingKeyCell = SecondSortColumn & "2"    'Here's that assumption discussed above
    SecondSort_EndingKeyCell = SecondSortColumn & LastRow
    
    ColumnLetter = GetColumnLetterFromColumnNumber(LastCol)
    LastRangeCell = ColumnLetter & LastRow
    
    ' Clear the last sort settings
    Application.ActiveSheet.Sort.SortFields.Clear
  
    ' Set first sort key
    Application.ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        FirstSort_StartingKeyCell & ":" & FirstSort_EndingKeyCell), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
  
    ' Set second sort key
    Application.ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        SecondSort_StartingKeyCell & ":" & SecondSort_EndingKeyCell), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
        
    With Application.ActiveSheet.Sort
        .SetRange Range("A1:" & LastRangeCell)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
End Function

Private Sub Group_Rows_By_Date()

' Group Rows by Date (but just the date portion, not the time

    Dim GroupDate As Double
    Dim CurrentDate As Double
    Dim ContinueOn As Boolean
    ContinueOn = True
    
    If IsDate(Selection.Value) Then
    
        GroupDate = Int(Selection.Value)

        Do While ContinueOn
            Selection.Offset(1, 0).Select  'Move to next transaction date
            If IsDate(Selection.Value) Then
                CurrentDate = Int(Selection.Value)
                If CurrentDate <> GroupDate Then
                    GroupDate = CurrentDate 'New group date
                    Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove 'Insert a line before this row
                    Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove 'Insert another line before this row
                    Selection.Offset(1, 0).Select  'Move down one line
                End If
            Else: ContinueOn = False
            End If
        Loop
    End If
End Sub

Private Sub Delete_And_Move_Details_Report_Columns()

'
' NOTE: THIS SUBROUTINE IS NO LONGER NEEDED! WE NOW DOWNLOAD DETAILS REPORTS FROM CLASSY WITH *ONLY*
'       THE COLUMNS WE NEED, *AND* IN THE RIGHT ORDER ALREADY!!!
'

' DELETE UNNECESSARY COLUMNS

' Delete "Program Designation" Column (already have it in the Payout report)
    Columns("AQ:AQ").Select
    Selection.Delete Shift:=xlToLeft

' Delete "Donor Supporter ID" Column (don't need it)
    Columns("AO:AO").Select
    Selection.Delete Shift:=xlToLeft

' Delete all of the following columns. Some of them are not needed;
' others we already have in the Payout report.
'
' W: Dedication Last Name
' X: Dedication Contact State
' Y: Dedication First Name
' Z: Authorization Code
' AA: Billing Address
' AB: Billing Address 2
' AC: Billing City
' AD: Billing Country
' AE: Billing First Name
' AF: Billing Last Name
' AG: Billing Postal Code
' AH: Billing State
' AI: CC Exp Date
' AJ: Credit Card Type
' AK: Last Four
' AL: Donor Email
    Columns("W:AL").Select
    Selection.Delete Shift:=xlToLeft

' Delete all of the following columns. Some of them are not needed;
' others we already have in the Payout report.
'
' L: Dedication Contact Address
' M: Dedication Contact City
' N: Dedication Contact Country
' O: Dedication Contact Email
' P: Dedication Contact First Name
' Q: Dedication Contact Last Name
' R: Dedication Contact Name
' S: Dedication Contact Postal Code
    Columns("L:S").Select
    Selection.Delete Shift:=xlToLeft

' Delete all of the following columns. Some of them are not needed;
' others we already have in the Payout report.
'
' H: Campaign Name
' I: How do you prefer <nonprofit name> communicates the impact of your support?
' J: Total Fees
    Columns("H:J").Select
    Selection.Delete Shift:=xlToLeft

' Delete "Display Name" column (unneeded)
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft

' Delete "Transaction ID" column (unneeded)
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft

' MOVE COLUMNS
'
' Move "Transaction Date" column to column A
'
    Columns("E:E").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    
End Sub
'
' "get last row in column" and "get last row in sheet" methods by **Jon Acampora** at Excel Campus
' https://www.excelcampus.com/vba/find-last-row-column-cell/
'
Private Function GetLastRowInColumn(strColLtr As String)
'Finds the last non-blank cell in a single row or column
'I modified it to use the column letter (string) as opposed to the column number

Dim lLastRow As Long
'Dim lLastCol As Long
    
    'Get the column number for this column letter
    lColNumber = GetColumnNumberFromColumnLetter(strColLtr)
    
    'Find the last non-blank cell in given column
    lLastRow = Cells(Rows.Count, lColNumber).End(xlUp).Row
    
    GetLastRowInColumn = lLastRow
    
    'Find the last non-blank cell in row 1
    'lLastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    'MsgBox "Last Row: " & lRow & vbNewLine & _
    '        "Last Column: " & lCol
  
End Function
Private Function Unused_Get_Last_Row_Function() 'again by Jon Acampora (see prior function)
'Finds the last non-blank cell on a sheet/range.

    Dim lRow As Long
    Dim lCol As Long
    
    lRow = Cells.Find(what:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    
    MsgBox "Last Row: " & lRow

End Function

'
' FIND LAST USED ROW AND COLUMN SOLUTION BY **Stephan Vierkant** ON STACK OVERFLOW
' https://stackoverflow.com/questions/38882321/better-way-to-find-last-used-row
'
' Problems with normal methods
'
'Account for Blank Rows / Columns - If you have blank rows or columns at the beginning of your data
'then methods like UsedRange.Rows.Count and UsedRange.Columns.Count will skip over these blank rows
'(although they do account for any blank rows / columns that might break up the data), so if you refer
'to ThisWorkbook.Sheets(1).UsedRange.Rows.Count you will skip lines in cases where there are blank rows
'at the top of your sheet:
'
'     <Shows sheet with one blank row at the top followed by 11 lines>
'
'This will skip the top row from the count and return only 11:
'
'       ThisWorkbook.Sheets(1).UsedRange.Rows.Count (returns incorrect used rows of 11 instead of 12)
'
'This code will include the blank row and return 12 instead:
'
'       ThisWorkbook.Sheets(1).UsedRange.Cells(ThisWorkbook.Sheets(1).UsedRange.Rows.Count, 1).Row
'
'The same issue applies to columns.
'

'Solution

'Note that this is targeted at finding the last "Used" Row or Column on an entire sheet;
'this doesn't work if you just want the last cell in a specific range.

' HIS TWO FUNCTIONS ARE BELOW
'
'Examples of calling these Functions:
'
'    'Define the Target Worksheet  we're interested in:
'       Dim Sht1 As Worksheet: Set Sht1 = ThisWorkbook.Sheets(1)
'    'Print the last row and column numbers:
'       Debug.Print "Last Row = "; GetLastRow(Sht1)
'       Debug.Print "Last Col = "; GetLastCol(Sht1)

Private Function GetLastRow(Sheet As Worksheet)
    'Gets last used row # on sheet.
    GetLastRow = Sheet.UsedRange.Cells(Sheet.UsedRange.Rows.Count, 1).Row
End Function

Private Function GetLastCol(Sheet As Worksheet)
    'Gets last used column # on sheet.
    GetLastCol = Sheet.UsedRange.Cells(1, Sheet.UsedRange.Columns.Count).Column
End Function

' From "brettdj" on StackOverflow
' https://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter
'
' See also TheSpreadsheetGuru's take on this for an equivalent (and nearly identical) solution:
' https://www.thespreadsheetguru.com/the-code-vault/vba-code-to-convert-column-number-to-letter-or-letter-to-number

'This function returns the column letter for a given column number.
'
Private Function GetColumnLetterFromColumnNumber(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    GetColumnLetterFromColumnNumber = vArr(0)
End Function
' From TheSpreadsheetGuru:
' https://www.thespreadsheetguru.com/the-code-vault/vba-code-to-convert-column-number-to-letter-or-letter-to-number
'
Private Function GetColumnNumberFromColumnLetter(strColLtr As String)
    Dim ColumnNumber As Long
    
    'Convert To Column Number
    ColumnNumber = Range(strColLtr & 1).Column
      
    GetColumnNumberFromColumnLetter = ColumnNumber
    
End Function
' From Nigel Foster on Stack Overflow:
' https://stackoverflow.com/questions/15723672/how-to-remove-all-non-alphanumeric-characters-from-a-string-except-period-and-sp
'
Private Function CleanString(str As String) As String
    Dim i As Integer
    
    For i = 1 To Len(str)
        If InStr(1, "01234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz. !@#$%&*()^-_=+\/'""?<>[]{},~", Mid(str, i, 1)) Then CleanString = CleanString & Mid(str, i, 1)
    Next
    
End Function
'From https://www.automateexcel.com/vba/array-length-size/
Private Function GetArrayLength(myArray() As String) As Long
   
   Dim result As Long
   
   If IsEmpty(myArray) Then
      result = 0
   Else
      result = UBound(myArray) - LBound(myArray) + 1
   End If
   
   GetArrayLength = result
   
   'MsgBox ("At end of GetArrayLength(), return value is " & GetArrayLength)
   
End Function

