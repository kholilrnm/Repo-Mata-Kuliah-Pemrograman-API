VERSION 5.00
Begin VB.Form menuKerucut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cone Menu"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
   Icon            =   "menuKerucut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuKerucut.frx":2AA22
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
      TabIndex        =   5
      Top             =   28
      Width           =   960
   End
   Begin VB.TextBox hasil_kerucut 
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
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4320
      Width           =   2000
   End
   Begin VB.CommandButton btnLPKer 
      BackColor       =   &H000080FF&
      Caption         =   "Hitung Luas Permukaan"
      Height          =   509
      Left            =   3940
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3280
      Width           =   1550
   End
   Begin VB.CommandButton btnVolKer 
      BackColor       =   &H000080FF&
      Caption         =   "Hitung Volume"
      Height          =   509
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3280
      Width           =   1550
   End
   Begin VB.TextBox jari_kerucut 
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
      Left            =   3480
      TabIndex        =   1
      Top             =   1585
      Width           =   2000
   End
   Begin VB.TextBox tinggi_kerucut 
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
      Left            =   3480
      TabIndex        =   0
      Top             =   2480
      Width           =   2000
   End
End
Attribute VB_Name = "menuKerucut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBtn_Click()
    Me.Hide
    menuPilihan.Show
    
End Sub

Private Sub btnLPKer_Click()
Dim luas_kerucut As New rumus

  If jari_kerucut = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    ElseIf (tinggi_kerucut = "") Then
        MsgBox ("Height of Cone Can't Empty!")
    Else
        hasil_kerucut = luas_kerucut.luas_kerucut(jari_kerucut, tinggi_kerucut)
    End If
End Sub

Private Sub btnVolKer_Click()
Dim vol_kerucut As New rumus

  If jari_kerucut = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    ElseIf (tinggi_kerucut = "") Then
        MsgBox ("Height of Cone Can't Empty!")
    Else
        hasil_kerucut = vol_kerucut.vol_kerucut(jari_kerucut, tinggi_kerucut)
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


Private Sub jari_kerucut_Change()
Dim textval As String
Dim numval As String

textval = jari_kerucut.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    jari_kerucut.Text = CStr(numval)
  End If
  
End Sub

Private Sub tinggi_kerucut_Change()
Dim textval As String
Dim numval As String

textval = tinggi_kerucut.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    tinggi_kerucut.Text = CStr(numval)
  End If
  
End Sub

