VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "p76-2ͳ�Ʒ���"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7980
   FillColor       =   &H0080FF80&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   7980
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "tip"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ʼͳ��"
      BeginProperty Font 
         Name            =   "����"
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
Title = "����"
n = InputBox("������༶������������100�ˣ�", Title)
If Not IsNumeric(n) Or n < 100 Then
  MsgBox "���ڴ����벻С��100������", 5
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
Print "���������ε����� �ð༶����" & n & "�� ����"
Print "0��9�ֹ�" & a(0) & "��"
Print "10��19�ֹ�" & a(1) & "��"
Print "20��29�ֹ�" & a(2) & "��"
Print "30��39�ֹ�" & a(3) & "��"
Print "40��49�ֹ�" & a(4) & "��"
Print "50��59�ֹ�" & a(5) & "��"
Print "60��69�ֹ�" & a(6) & "��"
Print "70��79�ֹ�" & a(7) & "��"
Print "80��89�ֹ�" & a(8) & "��"
Print "90��99�ֹ�" & a(9) & "��"
Print "100�ֹ�" & a(10) & "��"
End Sub

Private Sub Command2_Click()
FontSize = 10
Print "�뵥������ʼͳ�ơ���ť��ʼ����ͳ�ƣ��ڵ�����������༶����"
End Sub

