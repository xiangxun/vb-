VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "p76-3杨辉三角"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   9465
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "开始"
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
      Left            =   7800
      MaskColor       =   &H0000FF00&
      TabIndex        =   3
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   240
      ScaleHeight     =   3915
      ScaleWidth      =   7035
      TabIndex        =   2
      Top             =   240
      Width           =   7095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "输入行号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   7770
      TabIndex        =   1
      Top             =   1200
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture1.Cls
n = Text1.Text
If Not IsNumeric(n) Then
  MsgBox "请输入数字", 1, "try again"
  Text1.Text = ""
Else
  n = Text1.Text
'End If
   ReDim a(1 To n, 1 To n)
   For i = 1 To n
     For j = 1 To i
       If j = 1 Or i = j Then
          a(i, j) = 1
       Else
          a(i, j) = a(i - 1, j) + a(i - 1, j - 1)
       End If
     Next j
   Next i
   For i = 1 To n
     For j = 1 To i
       Picture1.Print Tab(j * 5); a(i, j);
     Next j
     Picture1.Print
   Next i
End If
End Sub
