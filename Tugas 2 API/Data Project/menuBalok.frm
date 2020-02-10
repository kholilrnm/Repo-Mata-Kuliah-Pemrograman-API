VERSION 5.00
Begin VB.Form menuBalok 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "menuBalok"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11505
   Icon            =   "menuBalok.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuBalok.frx":2AA22
   ScaleHeight     =   6450
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton backBtn 
      BackColor       =   &H0080FFFF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   28
      Width           =   960
   End
   Begin VB.TextBox t_balok 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   7
      Top             =   2700
      Width           =   2000
   End
   Begin VB.TextBox l_balok 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   6
      Top             =   1960
      Width           =   2000
   End
   Begin VB.TextBox p_balok 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   5
      Top             =   1200
      Width           =   2000
   End
   Begin VB.CommandButton btnVolBalok 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Volume"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   509
      Left            =   2450
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3570
      Width           =   1550
   End
   Begin VB.CommandButton btnLPBalok 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Area"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   509
      Left            =   4100
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3580
      Width           =   1550
   End
   Begin VB.CommandButton btnDiagonalBalok 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Side One"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   509
      Left            =   2450
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4250
      Width           =   1550
   End
   Begin VB.CommandButton btnKelilingBalok 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Around"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   509
      Left            =   4100
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4260
      Width           =   1550
   End
   Begin VB.TextBox hasil_balok 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   5090
      Width           =   2000
   End
End
Attribute VB_Name = "menuBalok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBtn_Click()
    Me.Hide
    menuPilihan.Show
    
End Sub

Private Sub btnDiagonalBalok_Click()
Dim h_diagonal_sisi_balok As New rumus

    If (p_balok = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (l_balok = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (t_balok = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        hasil_balok = h_diagonal_sisi_balok.luas_1sisi_balok(p_balok, l_balok, t_balok)
    End If
End Sub


Private Sub btnKelilingBalok_Click()
Dim keliling_balok As New rumus

    If (p_balok = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (l_balok = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (t_balok = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        hasil_balok = keliling_balok.keliling_balok(p_balok, l_balok, t_balok)
    End If

End Sub

Private Sub btnLPBalok_Click()
Dim h_luas_balok As New rumus

    If (p_balok = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (l_balok = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (t_balok = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        hasil_balok = h_luas_balok.luas_balok(p_balok, l_balok, t_balok)
        
    End If

End Sub

Private Sub btnVolBalok_Click()
Dim h_vol_balok As New rumus

    If (p_balok = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (l_balok = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (t_balok = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        hasil_balok = h_vol_balok.vol_balok(p_balok, l_balok, t_balok)
    End If
End Sub

Private Sub Picture1_Click()

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, unloadMode As Integer)
    Select Case unloadMode
        Case 1, 2, 3
            Cancel = False
            Unload Me
        Case Else
            Cancel = True
            End
    End Select
End Sub

Private Sub p_balok_Change()
Dim textval As String
Dim numval As String

textval = p_balok.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    p_balok.Text = CStr(numval)
  End If
  
End Sub

Private Sub l_balok_Change()
Dim textval As String
Dim numval As String

textval = l_balok.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    l_balok.Text = CStr(numval)
  End If
  
End Sub

Private Sub t_balok_Change()
Dim textval As String
Dim numval As String

textval = t_balok.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    t_balok.Text = CStr(numval)
  End If
  
End Sub

