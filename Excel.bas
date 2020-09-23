Attribute VB_Name = "Excel"
Option Explicit

Public Const ex_Str = 0
Public Const ex_Num = 1
    

Public RndArray()  As Double    '(0 To 1000, 0 To 1) As Single
Public ExelTextArray()  As String
Public oExcel As Object 'Excel.Application
Public bNewExcelObjectCreated As Boolean
Public oWorkBook1 As Object ' Excel.Workbook
Public oWorkBook2 As Object ' Excel.Workbook
Public Function OpenExcel(ByRef bObjectCreated As Boolean) As Boolean
  On Error Resume Next
  Err.Clear
  If oExcel Is Nothing Then
    bObjectCreated = False
    Set oExcel = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
      Err.Clear
      Set oExcel = CreateObject("excel.application")
      oExcel.Visible = True
      If Err.Number <> 0 Then
        MsgBox "Cannot create excel object", vbCritical
        Exit Function
      End If
      bObjectCreated = True
    End If
  End If
  OpenExcel = True
End Function
Public Sub CloseExcel()
  If oExcel Is Nothing Then Exit Sub
  If Not oWorkBook1 Is Nothing Then
    oWorkBook1.Close False
  End If
  If Not oWorkBook2 Is Nothing Then
    oWorkBook2.Close False
  End If
  If bNewExcelObjectCreated Then
    oExcel.Quit
  End If
  Set oExcel = Nothing
End Sub
Public Function AddWorkBook(oExcelApplication As Object) As Object
  Set AddWorkBook = oExcelApplication.Workbooks.Add()
End Function
Public Function GetWorkSheet(oWorkBook As Object, n As Integer) As Object
  Set GetWorkSheet = oWorkBook.Sheets(n)
End Function

Public Sub FillSheetCellByCell(oSheet As Object, aryData)
  Dim n As Integer, m As Integer
  For n = LBound(aryData, 1) To UBound(aryData, 1)
    For m = LBound(aryData, 2) To UBound(aryData, 2)
      oSheet.Cells(n + 1, m + 1) = aryData(n, m)
    Next
  Next
End Sub
Public Sub FillSheetUsingRange(oSheet As Object, aryData)
  oSheet.Range("A1", "A1").Resize(UBound(aryData, 1) - LBound(aryData, 1) + 1, UBound(aryData, 2) - LBound(aryData, 2) + 1).Value = aryData
End Sub
Public Sub InitDataArray()
  Dim n As Integer, m As Integer
  
  For n = 0 To 99
    For m = 0 To 99
      RndArray(n, m) = n * 100 + m
    Next
  Next
  
End Sub

Public Sub ExcelSaveArrayStp()
  Screen.MousePointer = vbHourglass
  
  If Not OpenExcel(bNewExcelObjectCreated) Then Exit Sub
  Set oWorkBook1 = AddWorkBook(oExcel)
  FillSheetCellByCell GetWorkSheet(oWorkBook1, 1), RndArray
  
  Screen.MousePointer = vbDefault
End Sub

Public Sub ExcelSaveArray(A As Byte)
  Screen.MousePointer = vbHourglass
  
  If Not OpenExcel(bNewExcelObjectCreated) Then Exit Sub
  Set oWorkBook2 = AddWorkBook(oExcel)
  If A = ex_Str Then
      FillSheetUsingRange GetWorkSheet(oWorkBook2, 1), ExelTextArray
  ElseIf A = ex_Num Then
      FillSheetUsingRange GetWorkSheet(oWorkBook2, 1), RndArray
  End If
  Screen.MousePointer = vbDefault
End Sub


