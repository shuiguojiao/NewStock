Attribute VB_Name = "Module1"
Option Explicit

Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public Const BM_CLICK = &HF5

'button handle，最多支持100个同时打新
Public arrayWnd(100) As Long
'total handle num
Public arrayIndex As Long

'不同窗口（交易软件）的个数
Public lFormIndex As Long
'windows+button label组合来识别交易软件
Public sFormName(10), sLabelName(10) As String




'find button handle
Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean

    Dim s As String
    Dim t As String
    Dim i As Integer
    Dim tSubWnd As Long
    
    s = String(1024, Chr(0))
    
    
    Call GetWindowText(hWnd, s, 1024)
    t = Trim(Left(s, InStr(1, s, Chr(0)) - 1))
    
    If hWnd = &H121504 Then
       tSubWnd = 1
    End If
    
    If hWnd = &HD0454 Then
       tSubWnd = 2
    End If
    
    For i = 0 To lFormIndex - 1
        If t = "" Then  'window label is null, and tile is in subwindows
            tSubWnd = FindWindowEx(hWnd, 0, vbNullString, sFormName(i))
            
            If tSubWnd > 0 Then
                Form1.List1.AddItem (arrayIndex + 1) & ":" & sFormName(i) & " - " & "s" & Hex(hWnd)
                'find button
                tSubWnd = FindWindowEx(hWnd, 0, vbNullString, sLabelName(i))
                If tSubWnd > 0 Then 'find button
                    Form1.List1.AddItem sLabelName(i)
                    arrayWnd(arrayIndex) = tSubWnd
                    arrayIndex = arrayIndex + 1
                    Exit For
                End If
            End If
        ElseIf StrComp(sFormName(i), t, vbTextCompare) = 0 Then  'windows label + button label
            Form1.List1.AddItem (arrayIndex + 1) & ":" & t & " - " & Hex(hWnd)
            
            'find button
            tSubWnd = FindWindowEx(hWnd, 0, vbNullString, sLabelName(i))
            If tSubWnd > 0 Then 'find button
                Form1.List1.AddItem sLabelName(i)
                arrayWnd(arrayIndex) = tSubWnd
                arrayIndex = arrayIndex + 1
                Exit For
            End If
        End If
        
        
        'if
    Next i
    EnumWindowsProc = True
    
End Function


Public Function ReadIniFile(ByVal strIniFileName As String) As Integer
    On Error GoTo GetIniStrErr
    
    Dim strLine As String
    Dim iFormNum, pos, index As Integer
    Dim bSwitch As Boolean
    
    ReadIniFile = True
    iFormNum = 0
    index = 0
    bSwitch = False
    
    Open strIniFileName For Input As #1
    While Not EOF(1)
        Line Input #1, strLine
        If strLine <> "" Then
        
            'Read Form Number
            If iFormNum = 0 Then
                pos = InStr(strLine, "FormNumber")
                If pos > 0 Then
                    pos = InStr(strLine, "=")
                    If pos > 0 Then
                     iFormNum = Val(Trim(Mid(strLine, pos + 1)))
                    End If
                End If
            Else
                If index < iFormNum Then
                    If Not bSwitch Then
                        pos = InStr(strLine, "FormName" & (index + 1))
                        If pos > 0 Then
                            pos = InStr(strLine, "=")
                            If pos > 0 Then
                                sFormName(index) = Trim(Mid(strLine, pos + 1))
                                bSwitch = Not bSwitch
                            End If
                        End If
                    Else
                        pos = InStr(strLine, "LabelName" & (index + 1))
                        If pos > 0 Then
                            pos = InStr(strLine, "=")
                            If pos > 0 Then
                                sLabelName(index) = Trim(Mid(strLine, pos + 1))
                                bSwitch = Not bSwitch
                                index = index + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Wend
    Close #1
    ReadIniFile = index
    Exit Function
GetIniStrErr:
    Err.Clear
    ReadIniFile = 0
End Function
