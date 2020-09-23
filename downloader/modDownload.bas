Attribute VB_Name = "modDownload"
Option Explicit

Global strSvrURL As String
Global Url As String
Global RESUMEFILE As Boolean
Global FilePathName As String
Global Filename As String
Global FileLength As Single
Global Sec%, Min%, Hr%
Public Const SW_NORMAL = 1
Public strFormLoaded As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOW = 5

Public Function File_ByteConversion(NumberOfBytes As Single) As String
On Error Resume Next
    If NumberOfBytes < 1024 Then 'checks to see if its so small that it cant be converted into larger grouping
        File_ByteConversion = NumberOfBytes & " Bytes"
    
    End If
    
    If NumberOfBytes > 1024 Then  'Checks to see if file is big enough to convert into KB
        File_ByteConversion = Format(NumberOfBytes / 1024, "0.00") & " KB"
    
    End If
    
    If NumberOfBytes > 1024000 Then 'Checks to see if its big enough to convert into MB
        File_ByteConversion = Format(NumberOfBytes / 1024000, "###,###,##0.00") & " MB"
    
    End If
    
End Function

Public Function UpdateProgress(pb As Control, ByVal Percent)
'Replacement for progress bar..looks nicer also
Dim Num$ 'use percent
    If Not pb.AutoRedraw Then 'picture in memory ?
        pb.AutoRedraw = -1 'no, make one
    
    End If
    
    pb.Cls 'clear picture in memory
    pb.ScaleWidth = 100 'new sclaemodus
    pb.DrawMode = 10 'not XOR Pen Modus
    Num$ = Format$(Percent, "###") + "%"
    pb.CurrentX = 50 - pb.TextWidth(Num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(Num$)) / 2
    pb.Print Num$ 'print percent
    pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
    pb.Refresh 'show differents
    
End Function

Public Function FileCheck(Path$) As Boolean

    FileCheck = True 'Assume Success
    On Error Resume Next
    Dim Disregard As Long
    Disregard = FileLen(Path)
        If Err <> 0 Then
            FileCheck = False
        End If
    
End Function

Public Function GETDATAHEAD(Data As Variant, ToRetrieve As String)
    On Error Resume Next
        If Data = "" Then Exit Function
        Dim EndBYTES%, a$, LENGTHEND%, PART%, Part2%, RetrieveLength%
            If InStr(Data, ToRetrieve) > 0 Then
                LENGTHEND = Len(Data)
                PART = InStr(Data, ToRetrieve)
                RetrieveLength = Len(ToRetrieve)
                a = Right(Data, LENGTHEND - PART - RetrieveLength)
                LENGTHEND = Len(a)
                If InStr(a, vbCrLf) > 0 Then
                Part2 = InStr(a, vbCrLf)
                a = Left(a, Part2 - 1)
            End If
            
        GETDATAHEAD = a
        
        End If
End Function

Public Function OpenIt(Frm As Form, ToOpen As String)
    ShellExecute Frm.hwnd, "Open", ToOpen, &O0, &O0, SW_NORMAL

End Function

