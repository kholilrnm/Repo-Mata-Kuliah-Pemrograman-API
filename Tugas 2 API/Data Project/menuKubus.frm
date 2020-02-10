VERSION 5.00
Begin VB.Form menuKubus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cube Menu"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
   Icon            =   "menuKubus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuKubus.frx":2AA22
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
   Begin VB.TextBox hasil_kubus 
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
      Top             =   4040
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4250
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3060
      Width           =   1550
   End
   Begin VB.CommandButton btnKelilingKubus 
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
      Left            =   2600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3060
      Width           =   1550
   End
   Begin VB.CommandButton btnLPKubus 
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
      TabIndex        =   2
      Top             =   2350
      Width           =   1550
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
      Height          =   509
      Left            =   2600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2380
      Width           =   1550
   End
   Begin VB.TextBox sisi_kubus 
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
      Top             =   1520
      Width           =   2000
   End
End
Attribute VB_Name = "menuKubus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBtn_Click()
    Me.Hide
    menuPilihan.Show
    
End Sub

Private Sub btnKelilingKubus_Click()
Dim keliling_kubus As New rumus

    If sisi_kubus = "" Then
        MsgBox ("Length of Side Can't Empty!")
    Else
        hasil_kubus = keliling_kubus.keliling_kubus(sisi_kubus)
        
    End If
    
End Sub

Private Sub btnLPKubus_Click()
Dim luas_kubus As New rumus

    If sisi_kubus = "" Then
        MsgBox ("Length of Side Can't Empty!")
    Else
        hasil_kubus = luas_kubus.luas_kubus(sisi_kubus)
    End If
    
End Sub

Private Sub btnVolKubus_Click()
Dim vol_kubus As New rumus

    If sisi_kubus = "" Then
        MsgBox ("Length of Side Can't Empty!")
    Else
        hasil_kubus = vol_kubus.vol_kubus(sisi_kubus)
    End If
End Sub

Private Sub Command1_Click()
Dim sisi_satu As New rumus

    If sisi_kubus = "" Then
        MsgBox ("Length of Side Can't Empty!")
    Else
        hasil_kubus = sisi_satu.kel_1sisi_kubus(sisi_kubus)
        
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

Private Sub sisi_kubus_Change()
Dim textval As String
Dim numval As String

textval = sisi_kubus.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    sisi_kubus.Text = CStr(numval)
  End If
  
End Sub

