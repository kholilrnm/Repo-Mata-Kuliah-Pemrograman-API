VERSION 5.00
Begin VB.Form menuLimas3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pyramid 3 Menu"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuLimas3.frx":0000
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
      Height          =   450
      Left            =   3840
      TabIndex        =   5
      Top             =   2920
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
      Left            =   3840
      TabIndex        =   4
      Top             =   2180
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
      Left            =   3840
      TabIndex        =   3
      Top             =   1420
      Width           =   2000
   End
   Begin VB.CommandButton btnVolPrisma3 
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
      Top             =   3720
      Width           =   1550
   End
   Begin VB.CommandButton btnLP_3 
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
      Top             =   3720
      Width           =   1550
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   4780
      Width           =   2000
   End
End
Attribute VB_Name = "menuLimas3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBtn_Click()
    Me.Hide
    menuPilihan.Show
    
End Sub

Private Sub btnLP_3_Click()
 Dim sisi_diagonal_alas As Integer
    Dim sisi_diagonal_tgh As Integer
        
    If Text1 = "" Then
        MsgBox ("Length of Surface Can't Empty!!")
    ElseIf (Text3 = "") Then
        MsgBox ("Height of Surface Can't Empty!")
    ElseIf (Text4 = "") Then
        MsgBox ("Height of Pyramid Can't Empty!")
    Else
        sisi_diagonal_tgh = Sqr((Text1 * Text1) + (Text3 * Text3))
        setengah_alas_seg3 = 1 / 2 * sisi_diagonal_tgh
        sisi_diagonal_samping = Sqr((Text1 * Text1) + (Text4 * Text4))
        tinggi_alas_seg3 = Sqr((setengah_alas_seg3 * setengah_alas_seg3) + (sisi_diagonal_samping * sisi_diagonal_samping))
        Text2 = ((1 / 2 * Text1 * Text3) + (1 / 2 * Text1 * Text4) + (1 / 2 * Text3 * Text4)) + (1 / 2 * sisi_diagonal_rgh * tinggi_alas_seg3)
    End If
End Sub

Private Sub btnVolPrisma3_Click()
    If Text1 = "" Then
        MsgBox ("Length of Surface Can't Empty!!")
    ElseIf (Text3 = "") Then
        MsgBox ("Height of Surface Can't Empty!")
    ElseIf (Text4 = "") Then
        MsgBox ("Height of Pyramid Can't Empty!")
    Else
        Text2 = 1 / 3 * 1 / 2 * Text1 * Text3 * Text4
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
Private Sub Text4_Change()
Dim textval As String
Dim numval As String

textval = Text4.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    Text4.Text = CStr(numval)
  End If
  
End Sub


