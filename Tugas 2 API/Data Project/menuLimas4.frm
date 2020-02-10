VERSION 5.00
Begin VB.Form menuLimas4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pyramid 4 Menu"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11505
   Icon            =   "menuLimas4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuLimas4.frx":2AA22
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
      TabIndex        =   6
      Top             =   28
      Width           =   960
   End
   Begin VB.TextBox hasil_limas4 
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4775
      Width           =   2000
   End
   Begin VB.CommandButton btnLP_4 
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
      Left            =   4250
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1550
   End
   Begin VB.CommandButton btnVolP_4 
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
      Left            =   2600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1550
   End
   Begin VB.TextBox p_limas4 
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
      Left            =   3840
      TabIndex        =   2
      Top             =   1420
      Width           =   2000
   End
   Begin VB.TextBox t_alas_limas4 
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
      Left            =   3840
      TabIndex        =   1
      Top             =   2180
      Width           =   2000
   End
   Begin VB.TextBox t_limas4 
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
      Left            =   3840
      TabIndex        =   0
      Top             =   2920
      Width           =   2000
   End
End
Attribute VB_Name = "menuLimas4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBtn_Click()
    Me.Hide
    menuPilihan.Show
    
End Sub

Private Sub btnLP_4_Click()
    Dim luas_limas4 As New rumus

    Dim sisi_diagonal_alas As Integer
    Dim sisi_diagonal_tgh As Integer
        
    If p_limas4 = "" Then
        MsgBox ("Length of Surface Can't Empty!")
    ElseIf (t_alas_limas4 = "") Then
        MsgBox ("Width of Surface Can't Empty!")
    ElseIf (t_limas4 = "") Then
        MsgBox ("Height of Pyramid Can't Empty!")
    Else
        
        hasil_limas4 = luas_limas4.luas_limas4(p_limas4, t_alas_limas4, t_limas4)
        
    End If
End Sub

Private Sub btnVolP_4_Click()
Dim vol_limas4 As New rumus

    If p_limas4 = "" Then
        MsgBox ("Length of Surface Can't Empty!")
    ElseIf (t_alas_limas4 = "") Then
        MsgBox ("Width of Surface Can't Empty!")
    ElseIf (t_limas4 = "") Then
        MsgBox ("Height of Pyramid Can't Empty!")
    Else
        hasil_limas4 = vol_limas4.vol_limas4(p_limas4, t_alas_limas4, t_limas4)
        
    End If
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

Private Sub Picture1_Click()

End Sub

Private Sub p_limas4_Change()
Dim textval As String
Dim numval As String

textval = p_limas4.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    p_limas4.Text = CStr(numval)
  End If
  
End Sub

Private Sub t_alas_limas4_Change()
Dim textval As String
Dim numval As String

textval = t_alas_limas4.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    t_alas_limas4.Text = CStr(numval)
  End If
  
End Sub

Private Sub t_limas4_Change()
Dim textval As String
Dim numval As String

textval = t_limas4.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    t_limas4.Text = CStr(numval)
  End If
  
End Sub

