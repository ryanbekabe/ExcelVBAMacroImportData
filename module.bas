Sub Upload()
Dim sImportFile As String, sFile As String
Dim sThisBk As Workbook
Dim vfilename As Variant
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set sThisBk = ActiveWorkbook
sImportFile = Application.GetOpenFilename( _
FileFilter:="Microsoft Excel Workbooks, *.xls; *.xlsx", Title:="Open Workbook")
If sImportFile = "False" Then
MsgBox "Belum ada File yang anda pilih!"
Exit Sub
End If

Application.Workbooks.Open Filename:=sImportFile
Set sThisBk = Workbooks.Open(sImportFile)
sThisBk.Sheets("NILAI").Range("E13:N52").Copy
sThisBk.Close (False)
Range("e13:n52").Select
ActiveSheet.Paste
MsgBox "ANDA YAKIN MENGISI DATA NILAI DISINI KLIK OK, MAKA ALHAMDULILLAH DATA NILAI SUDAH DISALIN PADA HALAMAN INI"
End Sub


