VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "p76-2统计分数"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7980
   FillColor       =   &H0080FF80&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   7980
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "tip"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      MaskColor       =   &H0000FFFF&
      TabIndex        =   1
      Top             =   5520
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "开始统计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cls
Dim a(10)
Title = "输入"
n = InputBox("请输入班级人数（不少于100人）", Title)
If Not IsNumeric(n) Or n < 100 Then
  MsgBox "请在此输入不小于100的数字", 5
End If
  For i = 1 To n
    f = Int(101 * Rnd)
    Select Case f
    Case 0 To 9
      a(0) = a(0) + 1
    Case 10 To 19
      a(1) = a(1) + 1
    Case 20 To 29
      a(2) = a(2) + 1
    Case 30 To 39
      a(3) = a(3) + 1
    Case 40 To 49
      a(4) = a(4) + 1
    Case 50 To 59
      a(5) = a(5) + 1
    Case 60 To 69
      a(6) = a(6) + 1
    Case 70 To 79
      a(7) = a(7) + 1
    Case 80 To 89
      a(8) = a(8) + 1
    Case 90 To 99
      a(9) = a(9) + 1
    Case 100
      a(10) = a(10) + 1
   End Select
   Next i
   FontSize = 20
Print "各个分数段的人数 该班级共有" & n & "人 其中"
Print "0到9分共" & a(0) & "人"
Print "10到19分共" & a(1) & "人"
Print "20到29分共" & a(2) & "人"
Print "30到39分共" & a(3) & "人"
Print "40到49分共" & a(4) & "人"
Print "50到59分共" & a(5) & "人"
Print "60到69分共" & a(6) & "人"
Print "70到79分共" & a(7) & "人"
Print "80到89分共" & a(8) & "人"
Print "90到99分共" & a(9) & "人"
Print "100分共" & a(10) & "人"
End Sub

Private Sub Command2_Click()
FontSize = 10
Print "请单击‘开始统计’按钮开始进行统计，在弹出框中输入班级人数"
End Sub

