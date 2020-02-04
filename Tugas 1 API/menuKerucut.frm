VERSION 5.00
Begin VB.Form menuKerucut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cone Menu"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuKerucut.frx":0000
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
   Begin VB.TextBox Text2 
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
   Begin VB.TextBox Text1 
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
   Begin VB.TextBox Text3 
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
  If Text1 = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    ElseIf (Text3 = "") Then
        MsgBox ("Height of Cone Can't Empty!")
    Else
        Text2 = ((22 / 7 * Text1 * Text1) + (22 / 7 * Text1 * (Sqr((Text1 * Text1) + (Text3 * Text3)))))
    End If
End Sub

Private Sub btnVolKer_Click()
  If Text1 = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    ElseIf (Text3 = "") Then
        MsgBox ("Height of Cone Can't Empty!")
    Else
        Text2 = 1 / 3 * 22 / 7 * Text1 * Text1 * Text3
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


Private Sub Text1_Change()
Dim textval As String
Dim numval As String

textval = Text1.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    Text1.Text = CStr(numval)
  End If
  
End Sub

Private Sub Text3_Change()
Dim textval As String
Dim numval As String

textval = Text3.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    Text3.Text = CStr(numval)
  End If
  
End Sub

