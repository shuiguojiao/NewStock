VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��������"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   Icon            =   "auto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   4500
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�г�"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
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
    'Form1.Caption = "�������� V1.02"
    'Command1.Caption = "��ʼ"
    'Command2.Caption = "�ر�"
    List1.Clear
    
    lFormIndex = ReadIniFile("auto.ini")
    If lFormIndex = 0 Then
        Label1.Caption = "δ�ҵ������ļ���ʹ��Ĭ��ֵ"
        
        lFormIndex = 3
        sFormName(0) = "���뽻��ȷ��"
        sLabelName(0) = "����ȷ��"
        sFormName(1) = "ί������"
        sLabelName(1) = "ȷ��(Y)"
        sFormName(2) = "ί��ȷ��"
        sLabelName(2) = "��(Y)"
        
    End If
End Sub

