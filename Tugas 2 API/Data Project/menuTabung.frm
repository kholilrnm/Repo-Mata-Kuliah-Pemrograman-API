VERSION 5.00
Begin VB.Form menuTabung 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tube Menu"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11505
   Icon            =   "menuTabung.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuTabung.frx":2AA22
   ScaleHeight     =   6450
   ScaleMode       =   0  'User
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
      Height          =   440
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   28
      Width           =   960
   End
   Begin VB.TextBox t_tabung 
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
      TabIndex        =   4
      Top             =   2490
      Width           =   2000
   End
   Begin VB.TextBox jari_tabung 
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
      TabIndex        =   3
      Top             =   1590
      Width           =   2000
   End
   Begin VB.CommandButton btnVolKubus 
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
      Height          =   495
      Left            =   2300
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3300
      Width           =   1540
   End
   Begin VB.CommandButton btnLPTabung 
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
      Height          =   495
      Left            =   3950
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3290
      Width           =   1540
   End
   Begin VB.TextBox hasil_tabung 
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
      TabIndex        =   0
      Top             =   4330
      Width           =   2000
   End
End
Attribute VB_Name = "menuTabung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub backBtn_Click()
    Me.Hide
    menuPilihan.Show
    
End Sub

Private Sub btnLPTabung_Click()
Dim luas_tabung As New rumus

  If jari_tabung = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    ElseIf (t_tabung = "") Then
        MsgBox ("Tube of Height Can't Empty!")
    Else
         hasil_tabung = luas_tabung.luas_tabung(jari_tabung, t_tabung)
    End If
End Sub

Private Sub btnVolKubus_Click()
Dim vol_tabung As New rumus

  If jari_tabung = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    ElseIf (t_tabung = "") Then
        MsgBox ("Tube of Height Can't Empty!")
    Else
        hasil_tabung = vol_tabung.vol_tabung(jari_tabung, t_tabung)
        
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


Private Sub jari_tabung_Change()
Dim textval As String
Dim numval As String

textval = jari_tabung.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    jari_tabung.Text = CStr(numval)
  End If
  
End Sub

Private Sub t_tabung_Change()
Dim textval As String
Dim numval As String

textval = t_tabung.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    t_tabung.Text = CStr(numval)
  End If
  
End Sub
