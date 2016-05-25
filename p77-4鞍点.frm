VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   6840
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
FontSize = 5
Dim a(1 To 4, 1 To 5)
For i = 1 To 4
  For j = 1 To 5
   a(i, j) = Int(10 * Rnd)
   Print a(i, j);
  Next j
  Print
Next i
For i = 1 To 4
Max = a(i, 1)
maxj = 1
For j = 1 To 5
 If Max < a(i, j) Then
  Max = a(i, j)
  maxj = j
 End If
Next j
Min = a(1, maxj): mini = 1
For m = 1 To 4
 If Min > a(m, maxj) Then
  Min = a(m, maxj)
  mini = m
 End If
Next m
Next i
If Max = Min And mini = i Then
Print "鞍点在第" & i & "行" & "第" & maxj & "列，鞍点为" & Max;
Else
Print "无鞍点"
End If
End Sub
