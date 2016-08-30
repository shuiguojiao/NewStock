VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "打新助手"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   Icon            =   "auto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   4500
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "打新"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "列出"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   6480
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   5100
      ItemData        =   "auto.frx":058A
      Left            =   240
      List            =   "auto.frx":058C
      TabIndex        =   0
      Top             =   480
      Width           =   3885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "auto.ini"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    'enum windows and display in list
    arrayIndex = 0
    List1.Clear
    Call EnumWindows(AddressOf EnumWindowsProc, 0)
    
End Sub


Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim i As Integer
    'Dim tSubWnd As Long
    'Dim sLabelName As String
    
    'Find subwindow and send BM_CLICK
    For i = 0 To arrayIndex - 1
        'sLabelName = Form1.List1.List(i * 2 + 1)
        'tSubWnd = FindWindowEx(arrayWnd(i), 0, vbNullString, sLabelName)
        'If tSubWnd > 0 Then
        Call PostMessage(arrayWnd(i), BM_CLICK, 0, 0)
        Call PostMessage(arrayWnd(i), BM_CLICK, 0, 0)
        'End If
    Next i
End Sub

Private Sub Form_Load()
    'Form1.Caption = "打新助手 V1.02"
    'Command1.Caption = "开始"
    'Command2.Caption = "关闭"
    List1.Clear
    
    lFormIndex = ReadIniFile("auto.ini")
    If lFormIndex = 0 Then
        Label1.Caption = "未找到配置文件，使用默认值"
        
        lFormIndex = 3
        sFormName(0) = "买入交易确认"
        sLabelName(0) = "买入确认"
        sFormName(1) = "委托买入"
        sLabelName(1) = "确定(Y)"
        sFormName(2) = "委托确认"
        sLabelName(2) = "是(Y)"
        
    End If
End Sub

