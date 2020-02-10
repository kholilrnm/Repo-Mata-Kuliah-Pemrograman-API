VERSION 5.00
Begin VB.Form menuBola 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ball Menu"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
   Icon            =   "menuBola.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuBola.frx":2AA22
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
      TabIndex        =   4
      Top             =   28
      Width           =   960
   End
   Begin VB.TextBox jari_bola 
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
      Left            =   3860
      TabIndex        =   3
      Top             =   1820
      Width           =   2000
   End
   Begin VB.CommandButton btnVolBola 
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
      TabIndex        =   2
      Top             =   2680
      Width           =   1550
   End
   Begin VB.CommandButton btnLPBola 
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
      Left            =   4240
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2680
      Width           =   1550
   End
   Begin VB.TextBox hasil_bola 
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
      Left            =   3860
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3885
      Width           =   2000
   End
End
Attribute VB_Name = "menuBola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBtn_Click()
    Me.Hide
    menuPilihan.Show
    
End Sub

Private Sub btnLPBola_Click()
Dim luas_bola As New rumus
     If jari_bola = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    Else
        hasil_bola = luas_bola.luas_bola(jari_bola)
    End If
End Sub

Private Sub btnVolBola_Click()
Dim vol_bola As New rumus

    If jari_bola = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    Else
        hasil_bola = vol_bola.vol_bola(jari_bola)
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


Private Sub Label4_Click()

End Sub

Private Sub jari_bola_Change()
Dim textval As String
Dim numval As String

textval = jari_bola.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    jari_bola.Text = CStr(numval)
  End If
  
End Sub

