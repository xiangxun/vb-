VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "p77-6����"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5130
   ForeColor       =   &H00FFFF00&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   5130
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
FontSize = 13
Dim a(4, 4)
  For i = 0 To 4
      Print
    For m = 0 To 4
      If i > 0 And i < 4 And m > 0 And m < 4 Then
        a(i, m) = Int(4 + 76 * Rnd)
      Else
        a(i, m) = 1
      End If
    Next m
  Next i
'----------------------------------------------------------
  Print "����"
  For i = 0 To 4
      Print
    For m = 0 To 4
      Print Tab(5 * m); a(i, m);
    Next m
  Next i
'----------------------------------------------------------
  For i = 0 To 4
    For m = i To 4
       If i = m Then
       z = z + a(i, m)
       ElseIf i + m = 5 Then
       c = c + a(i, m)
       End If
    Next m
  Next i
  Print
  Print "���Խ��ߺ�Ϊ��" & z
  Print "�ζԽ��ߺ�Ϊ��" & c
'----------------------------------------------------------
  For i = 0 To 4
    For m = 0 To 4
       If i < m Then
       s = s + a(i, m)
       ElseIf i > m Then
       x = x + a(i, m)
       End If
    Next m
  Next i
  Print "������Ԫ�غ�Ϊ��" & s
  Print "������Ԫ�غ�Ϊ��" & x
'----------------------------------------------------------
  Print
  Print "������"
  For i = 0 To 4
      Print
    For m = i To 4
      Print Tab(5 * m); a(i, m);
    Next m
  Next i
'----------------------------------------------------------
  Print
  Print "������"
  For i = 0 To 4
      Print
    For m = 0 To i
      Print Tab(5 * m); a(i, m);
    Next m
  Next i
End Sub
