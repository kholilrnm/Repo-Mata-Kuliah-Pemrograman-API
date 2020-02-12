VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form menuBalok 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "menuBalok"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11505
   Icon            =   "menuBalok.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuBalok.frx":2AA22
   ScaleHeight     =   6450
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox t_balok 
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
   Begin VB.TextBox l_balok 
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
   Begin VB.TextBox p_balok 
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
   Begin VB.TextBox hasil_balok 
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
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   5090
      Width           =   2000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   540
      Left            =   -50
      TabIndex        =   8
      Top             =   0
      Width           =   11590
      _ExtentX        =   20426
      _ExtentY        =   953
      _Version        =   393216
      MousePointer    =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Menu Balok"
      TabPicture(0)   =   "menuBalok.frx":3741E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Menu Bola"
      TabPicture(1)   =   "menuBalok.frx":3743A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Menu Kerucut"
      TabPicture(2)   =   "menuBalok.frx":37456
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Menu Kubus"
      TabPicture(3)   =   "menuBalok.frx":37472
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Menu Limas 3"
      TabPicture(4)   =   "menuBalok.frx":3748E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Menu Limas 4"
      TabPicture(5)   =   "menuBalok.frx":374AA
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Menu Prisma"
      TabPicture(6)   =   "menuBalok.frx":374C6
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Menu Tabung"
      TabPicture(7)   =   "menuBalok.frx":374E2
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
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
Dim h_diagonal_sisi_balok As New rumus

    If (p_balok = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (l_balok = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (t_balok = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        hasil_balok = h_diagonal_sisi_balok.luas_1sisi_balok(p_balok, l_balok, t_balok)
    End If
End Sub


Private Sub btnKelilingBalok_Click()
Dim keliling_balok As New rumus

    If (p_balok = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (l_balok = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (t_balok = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        hasil_balok = keliling_balok.keliling_balok(p_balok, l_balok, t_balok)
    End If

End Sub

Private Sub btnLPBalok_Click()
Dim h_luas_balok As New rumus

    If (p_balok = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (l_balok = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (t_balok = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        hasil_balok = h_luas_balok.luas_balok(p_balok, l_balok, t_balok)
        
    End If

End Sub

Private Sub btnVolBalok_Click()
Dim h_vol_balok As New rumus

    If (p_balok = "") Then
        MsgBox ("Length of Block Can't Empty!")
    ElseIf (l_balok = "") Then
        MsgBox ("Width of Block Can't Empty!")
    ElseIf (t_balok = "") Then
        MsgBox ("Height of Block Can't Empty!")
    Else
        hasil_balok = h_vol_balok.vol_balok(p_balok, l_balok, t_balok)
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
    SSTab1.Tab = Null
End Sub

Private Sub p_balok_Change()
Dim textval As String
Dim numval As String

textval = p_balok.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    p_balok.Text = CStr(numval)
  End If
  
End Sub

Private Sub l_balok_Change()
Dim textval As String
Dim numval As String

textval = l_balok.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    l_balok.Text = CStr(numval)
  End If
  
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
Case 0
menuBalok.Show
Case 1
menuBola.Show
Case 2
menuKerucut.Show
Case 3
menuKubus.Show
Case 4
menuLimas3.Show
Case 5
menuLimas4.Show
Case 6
menuPrismaSeg.Show
Case 7
menuTabung.Show
End Select
End Sub

Private Sub t_balok_Change()
Dim textval As String
Dim numval As String

textval = t_balok.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    t_balok.Text = CStr(numval)
  End If
  
End Sub

