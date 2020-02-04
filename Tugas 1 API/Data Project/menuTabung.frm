VERSION 5.00
Begin VB.Form menuTabung 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tube Menu"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuTabung.frx":0000
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
      TabIndex        =   4
      Top             =   2490
      Width           =   2000
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
  If Text1 = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    ElseIf (Text3 = "") Then
        MsgBox ("Tube of Height Can't Empty!")
    Else
        Text2 = ((2 * 22 / 7 * Text1 * Text1) + (2 * 22 / 7 * Text1 * Text3))
    End If
End Sub

Private Sub btnVolKubus_Click()
  If Text1 = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    ElseIf (Text3 = "") Then
        MsgBox ("Tube of Height Can't Empty!")
    Else
        Text2 = 22 / 7 * Text1 * Text1 * Text3
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
