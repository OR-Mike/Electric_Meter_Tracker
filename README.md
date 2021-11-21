# Electric_Meter_Tracker
  This Excel Macro Enabled VBA tool built for tracking electric meters. The tool is user friendly with instructions on how to use the tool on the "Control Panel" tab. The code can be modified to fit other tracking needs. Feel free to make suggestions. Enjoy!

  The Electric Meter Tracker was intended to be an Upwork project. Unfortunately, my proposal never received a response. As a result, I figured I share the unused code with other codes more experienced than I. The coding isn't great to look at. I was lazy with the coding, by not commenting. I will fix this soon to ensure I keep up with good coding practices.
  The tracker allows the user to start a new sheet for each fiscal year. This button is called "Add New Billing Year". The second button, "Add Billing Cycle Information", allows user to input all associated data required to keep track tenant who use electric meters. The third button, "Clear Data" is self-explanatory. This allows the user to delete desired sheets.
  
If you're new to adding macros, please see below on how to add the developer tab to access VBA macro page.

Adding macros back into your Electric Meter Tool:

1)	Enable the Developer tab, if it isn’t already enabled.
a.	If not enabled, please follow steps below:
i.	First, we want to right-click on any of the existing tabs on our ribbon.
ii.	This opens a menu of options, and we want to select Customize the Ribbon
iii.	Then, select the Developer checkbox and click OK
iv.	The Developer tab is now visible.
v.	Ref: https://www.excelcampus.com/vba/enable-developer-tab/
vi.	Be sure to save the excel file as a macro enabled file (.xlsm)
2)	To add "Add New Biling Year" Macro
a.	Right click on  "Add New Billing Year"
b.	Assign Macro
c.	Name Macro
d.	"Add_new_billing_year"
e.	Click new
f.	Copy and paste the below:

Option Explicit

Dim SheetName As String
Dim BillCycle As String
Dim DueDate As String
Dim tenantName As String
Dim OvBillTot As Double
Dim kwh As Double
Dim PreviousVal As Double
Dim CurrentVal As Double
Dim UsageVal As Double
Dim MultVal As Double
Dim SubTotVal As Double
Dim RateVal As Double
Dim GrossTotalVal As Double
Dim GrossTotSubMetVal As Double
Dim NetTotVal As Double
Dim BillTotVal As Double
Dim answer As Integer
Dim meters As Integer
Dim submeterBeg As Double
Dim submeterEnd As Double
Dim submeterUse As Double
Dim submeterMult As Double
Dim submeterSubTot As Double
Dim submeterRate As Double
Dim submeterGrossTot As Double
Dim TotAllSubmeters As Double
Dim pctVal As Double
Dim i As Long

Sub Add_New_Billing_year()

SheetName = InputBox("Input Billing Year", "Add Sheet")

Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = SheetName
MsgBox "The sheet " & (SheetName) & " is successfuly made", , "Result"

Worksheets(SheetName).Cells(1, 1).Value = "Billing Cycle"
Worksheets(SheetName).Cells(1, 1).Font.Bold = True

Worksheets(SheetName).Cells(1, 2).Value = "Payment Due Date"
Worksheets(SheetName).Cells(1, 2).Font.Bold = True

Worksheets(SheetName).Cells(1, 3).Value = "Tenant"
Worksheets(SheetName).Cells(1, 3).Font.Bold = True

Worksheets(SheetName).Cells(1, 4).Value = "Overall Billing Total"
Worksheets(SheetName).Cells(1, 4).Font.Bold = True

Worksheets(SheetName).Cells(1, 5).Value = "kWh"
Worksheets(SheetName).Cells(1, 5).Font.Bold = True

Worksheets(SheetName).Cells(1, 6).Value = "Beginning"
Worksheets(SheetName).Cells(1, 6).Font.Bold = True

Worksheets(SheetName).Cells(1, 7).Value = "End"
Worksheets(SheetName).Cells(1, 7).Font.Bold = True

Worksheets(SheetName).Cells(1, 8).Value = "Usage"
Worksheets(SheetName).Cells(1, 8).Font.Bold = True

Worksheets(SheetName).Cells(1, 9).Value = "Mult"
Worksheets(SheetName).Cells(1, 9).Font.Bold = True

Worksheets(SheetName).Cells(1, 10).Value = "Sub Total"
Worksheets(SheetName).Cells(1, 10).Font.Bold = True

Worksheets(SheetName).Cells(1, 11).Value = "Rate"
Worksheets(SheetName).Cells(1, 11).Font.Bold = True

Worksheets(SheetName).Cells(1, 12).Value = "Gross Total"
Worksheets(SheetName).Cells(1, 12).Font.Bold = True

Worksheets(SheetName).Cells(1, 13).Value = "Gross Total of Sub Meters"
Worksheets(SheetName).Cells(1, 13).Font.Bold = True

Worksheets(SheetName).Cells(1, 14).Value = "Net Total"
Worksheets(SheetName).Cells(1, 14).Font.Bold = True

Worksheets(SheetName).Cells(1, 15).Value = "Bill Total"
Worksheets(SheetName).Cells(1, 15).Font.Bold = True

Worksheets(SheetName).Cells(1, 16).Value = "Percentage"
Worksheets(SheetName).Cells(1, 16).Font.Bold = True

Worksheets(SheetName).Cells(1, 17).Value = "Amount Due"
Worksheets(SheetName).Cells(1, 17).Font.Bold = True

Worksheets(SheetName).Columns("A:Z").AutoFit

End Sub

3)	To add "Add Billing Cycle Information" Macro
a.	Right click on "Add Billing Cycle Information"
b.	Assign Macro
c.	Name Macro
d.	"Billing_Cycle_Information"
e.	Click new
f.	Copy and paste the below:

Sub Billing_Cycle_Information()

Dim tgtSheetName As String
Dim LastCell As Long

tgtSheetName = InputBox("Enter Sheet Name", "Billing Year")
BillCycle = InputBox("Enter Billing Cycle", "Bill Cycle")

Worksheets(tgtSheetName).Activate
Range("A" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.Value = BillCycle

DueDate = InputBox("Enter Payment Due Date", "Bill Cycle")
Range("B" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.Value = DueDate

tenantName = InputBox("Enter Tenant Name", "Bill Cycle")
Range("C" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.Value = tenantName

OvBillTot = InputBox("Enter Overall Bill Total", "Bill Cycle")
Range("D" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.00"
ActiveCell.Value = OvBillTot

kwh = InputBox("Enter kWh", "Bill Cycle")
Range("E" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.00"
ActiveCell.Value = kwh

PreviousVal = InputBox("Enter Beginning Meter Value", "Bill Cycle")
Range("F" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.00"
ActiveCell.Value = PreviousVal

CurrentVal = InputBox("Enter End Meter Value", "Bill Cycle")
Range("G" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.00"
ActiveCell.Value = CurrentVal


UsageVal = CurrentVal - PreviousVal
Range("H" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.00"
ActiveCell.Value = UsageVal

MultVal = InputBox("Enter Mult Value", "Bill Cycle")
Range("I" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.00"
ActiveCell.Value = MultVal

SubTotVal = MultVal * UsageVal
Range("J" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.00"
ActiveCell.Value = SubTotVal

RateVal = OvBillTot / kwh
Range("K" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.0000"
ActiveCell.Value = RateVal

GrossTotalVal = SubTotVal * RateVal
Range("L" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.00"
ActiveCell.Value = GrossTotalVal

answer = MsgBox("Does this meter have sub-meters?", vbQuestion + vbYesNo + vbDefaultButton2, "Parent-Child Meters")
If answer = vbYes Then
  meters = InputBox("How many sub-meters does this meter have?", "Parent-Child meters")
  Range("M" & Rows.Count).End(xlUp).Offset(1).Select
    ActiveCell.NumberFormat = "0.00"
    ActiveCell.Value = 0#
  For i = 1 To meters
    submeterBeg = InputBox("What is beginning meter value for sub-meter " & i & "?", "Parent-Child Meters")
    submeterEnd = InputBox("What is End meter value for sub-meter " & i & "?", "Parent-Child Meters")
    submeterUse = submeterEnd - submeterBeg
    submeterMult = InputBox("What is the Mult value for sub-meter " & i & "?", "Parent-Child Meters")
    submeterSubTot = submeterMult * submeterUse
    submeterGrossTot = submeterSubTot * RateVal
    ActiveCell.Value = ActiveCell.Value + submeterGrossTot
 Next i
Else
    Range("M" & Rows.Count).End(xlUp).Offset(1).Select
    ActiveCell.NumberFormat = "0.00"
    ActiveCell.Value = 0#
End If

NetTotVal = GrossTotalVal - ActiveCell.Value
Range("N" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.00"
ActiveCell.Value = NetTotVal

BillTotVal = InputBox("What is the Bill Total?", "Parent-Child Meters")
Range("O" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.00"
ActiveCell.Value = BillTotVal

pctVal = InputBox("What is the percentage?", "Parent-Child Meters") / 100
Range("P" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.000"
ActiveCell.Value = pctVal

Range("Q" & Rows.Count).End(xlUp).Offset(1).Select
ActiveCell.NumberFormat = "0.00"
ActiveCell.Value = NetTotVal * pctVal

Worksheets(tgtSheetName).Columns("A:Z").AutoFit

End Sub

4)	To add "Clear Data" Macro
a.	Right click on "Clear Data"
b.	Assign Macro
c.	Name Macro
d.	"Clear_Contents"
e.	Click new
f.	Copy and paste the below:
Sub Clear_Contents()

Dim tgtSheet As String
Dim answer2 As Integer

tgtSheet = InputBox("Which sheet do you want to delete?", "Clear Contents")
Worksheets(tgtSheet).Activate

answer = MsgBox("Are you sure you want to delete sheet: " & tgtSheet & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Contents")
If answer = vbYes Then
    Sheets(tgtSheet).Delete
Else
  MsgBox "No"
End If

End Sub

