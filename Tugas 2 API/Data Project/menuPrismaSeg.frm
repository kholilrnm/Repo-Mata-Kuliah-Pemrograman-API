VERSION 5.00
Begin VB.Form menuPrismaSeg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prism Menu"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11505
   Icon            =   "menuPrismaSeg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuPrismaSeg.frx":2AA22
   ScaleHeight     =   6603.571
   ScaleMode       =   0  'User
   ScaleWidth      =   11656.98
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
      Height          =   440
      Left            =   118
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   27
      Width           =   960
   End
   Begin VB.TextBox hasil_prism 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4800
      Width           =   1974
   End
   Begin VB.CommandButton btnLPPrisma 
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
      Height          =   498
      Left            =   4096
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1530
   End
   Begin VB.CommandButton btnVolPrisma 
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
      Height          =   508
      Left            =   2448
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1530
   End
   Begin VB.TextBox p_prism 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   440
      Left            =   3711
      TabIndex        =   2
      Top             =   1440
      Width           =   1974
   End
   Begin VB.TextBox t_alas_prism 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3711
      TabIndex        =   1
      Top             =   2198
      Width           =   1974
   End
   Begin VB.TextBox t_prism 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3711
      TabIndex        =   0
      Top             =   2950
      Width           =   1974
   End
End
Attribute VB_Name = "menuPrismaSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBtn_Click()
    Me.Hide
    menuPilihan.Show
    
End Sub
Private Sub btnLPPrisma_Click()
Dim luas_prism As New rumus

    If p_prism = "" Then
        MsgBox ("Length of Surface Can't Empty!")
    ElseIf (t_alas_prism = "") Then
        MsgBox ("Height of Surface Can't Empty!")
    ElseIf (t_prism = "") Then
        MsgBox ("Height of Prism Can't Empty!")
    Else
        hasil_prism = luas_prism.luas_prismaseg(p_prism, t_alas_prism, t_prism)
    End If
End Sub

Private Sub btnVolPrisma_Click()
Dim vol_prism As New rumus

    If p_prism = "" Then
        MsgBox ("Length of Surface Can't Empty!")
    ElseIf (t_alas_prism = "") Then
        MsgBox ("Height of Surface Can't Empty!")
    ElseIf (t_prism = "") Then
        MsgBox ("Height of Prism Can't Empty!")
    Else
        hasil_prism = vol_prism.vol_prismaseg(p_prism, t_alas_prism, t_prism)
        
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

Private Sub p_prism_Change()
Dim textval As String
Dim numval As String

textval = p_prism.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    p_prism.Text = CStr(numval)
  End If
  
End Sub

Private Sub t_alas_prism_Change()
Dim textval As String
Dim numval As String

textval = t_alas_prism.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    t_alas_prism.Text = CStr(numval)
  End If
  
End Sub

Private Sub t_prism_Change()
Dim textval As String
Dim numval As String

textval = t_prism.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    t_prism.Text = CStr(numval)
  End If
  
End Sub
