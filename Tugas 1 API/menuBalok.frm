VERSION 5.00
Begin VB.Form menuBalok 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "menuBalok"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuBalok.frx":0000
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
      Left            =   3720
      TabIndex        =   7
      Top             =   2700
      Width           =   2000
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
      Left            =   3720
      TabIndex        =   6
      Top             =   1960
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
   Begin VB.TextBox Text4 
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
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
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
    If (Text1 = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (Text2 = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (Text3 = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        Text4 = Sqr((Text1 * Text1) + (Text2 * Text2) + (Text3 * Text3))
    End If
End Sub


Private Sub btnKelilingBalok_Click()
    If (Text1 = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (Text2 = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (Text3 = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        Text4 = 4 * (Text1 * Text2 * Text3)
    End If

End Sub

Private Sub btnLPBalok_Click()
    If (Text1 = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (Text2 = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (Text3 = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        Text4 = 2 * ((Text1 * Text2) + (Text2 * Text3) + (Text1 * Text3))
    End If

End Sub

Private Sub btnVolBalok_Click()
    If (Text1 = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (Text2 = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (Text3 = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        Text4 = Text1 * Text2 * Text3
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

Private Sub Text2_Change()
Dim textval As String
Dim numval As String

textval = Text2.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    Text2.Text = CStr(numval)
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

