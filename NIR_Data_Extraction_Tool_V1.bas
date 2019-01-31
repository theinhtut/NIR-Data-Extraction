Attribute VB_Name = "Module11"
Sub NIR_Data_Extraction_Tool()
' NIR_Data_Extraction_Tool is Excel VBA macro for extracting Near-Infrared(NIR) spectroscopy data from
' Texas Instruments DLP NIRNano Scan device.
' In this code, only Absorbance(B23 to B250) is extreacted to transpose for data pre-processing. The range can be changed as desired.
' Updated 31 January 2019


' Declearing variables using Dim
' csvFilePath       - User input .csv file location
' xRng              - User entered range for output destination
' userSelectedWb    - User selected workbook
' currentWb         - Current active workbook
' AC                - Active Cell address as Range
' acAddress         - Active Cell Address as String
' i                 - Looping for user selected multiple files.

On Error GoTo errorHandler
Dim csvFilePath As Variant
Dim xRng, AC As Range
Dim acAddress As String
Dim userSelectedWb, currentWb As Workbook
Dim i As Long
Set currentWb = Application.ActiveWorkbook
titleId = "NIR Data Extraction Tool by 2019 Thein Htut"

' Browse File to import data
csvFilePath = Application.GetOpenFilename("CSV File (*.csv), *.csv", , "Select file to import | " & titleId, MultiSelect:=True)
'Set userSelectedWb = Workbooks.Open(csvFilePath)
currentWb.Activate

' Getting current active cell address
Set AC = ActiveCell
acAddress = AC.Address

' Select destination cell and transpose the data
Set xRng = Application.InputBox(prompt:="Select destination cell", Title:=titleId, Default:=acAddress, Type:=8)



If IsArray(csvFilePath) Then
For i = LBound(csvFilePath) To UBound(csvFilePath)
Set userSelectedWb = Workbooks.Open(Filename:=csvFilePath(i))

' This data is Absorbance from range B23:B250
userSelectedWb.Sheets(1).Range("B23:B250").Copy
xRng.PasteSpecial Transpose:=True

Application.DisplayAlerts = False
userSelectedWb.Close savechanges:=False

ActiveCell.Offset(1, 0).Select
Set xRng = ActiveCell
Next i
End If

'Successful Import Data Message
MsgBox "You have successfully import" & vbNewLine & vbNewLine & vbNewLine & "Thank you for using NIR Data Extraction Tool", , "Import Successful"
Exit Sub

'Error Handler for errors.
errorHandler:
MsgBox ("Error! Please run it again.")


End Sub
