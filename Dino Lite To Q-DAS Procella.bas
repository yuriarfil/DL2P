Attribute VB_Name = "Module1"
Option Explicit
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long

Sub DinoLiteToQDAS()
Dim windNme As Long
Dim i As Integer
Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Sheet1")
Dim ThisRng As Range
Dim arrY() As Variant

windNme = FindWindow(vbEmpty, "procella ®")

'//check if window Procella exist?
If windNme = 0 Then
    MsgBox "Could not find window Procella", vbOKOnly & vbExclamation
Exit Sub
End If

'//check if window Procella is Minimize
If IsIconic(windNme) Then
    MsgBox "Please try again" & vbCrLf & "Procela is Miinimized !!!", vbExclamation
Exit Sub
End If

'//ignore run time error due to inputbox = cancel
On Error Resume Next

Set ThisRng = Application.InputBox("Select a range", "Get Range", Type:=8)
If ThisRng Is Nothing Then
    Exit Sub
End If

arrY = ThisRng.Value2
Application.ScreenUpdating = False

'//check if no data selected
If arrY(1, 1) = "" Then
    MsgBox "No Data Selected", vbExclamation
Exit Sub
End If
    
BringWindowToTop windNme
Sleep (1000)
    i = 1
    Do Until i = UBound(arrY) + 1
        SendKeys arrY(i, 1), True
        SendKeys "~", True
        i = i + 1
    Loop
    
SendKeys "{NUMLOCK}", True
Application.ScreenUpdating = True
End Sub
