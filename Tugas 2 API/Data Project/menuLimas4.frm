VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form menuLimas4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pyramid 4 Menu"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11505
   Icon            =   "menuLimas4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuLimas4.frx":2AA22
   ScaleHeight     =   6450
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox hasil_limas4 
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
      Top             =   4775
      Width           =   2000
   End
   Begin VB.CommandButton btnLP_4 
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
      TabIndex        =   4
      Top             =   3720
      Width           =   1550
   End
   Begin VB.CommandButton btnVolP_4 
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
      TabIndex        =   3
      Top             =   3720
      Width           =   1550
   End
   Begin VB.TextBox p_limas4 
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
      TabIndex        =   2
      Top             =   1420
      Width           =   2000
   End
   Begin VB.TextBox t_alas_limas4 
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
      TabIndex        =   1
      Top             =   2180
      Width           =   2000
   End
   Begin VB.TextBox t_limas4 
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
      Top             =   2920
      Width           =   2000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   540
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11590
      _ExtentX        =   20452
      _ExtentY        =   953
      _Version        =   393216
      MousePointer    =   1
      Tabs            =   8
      Tab             =   5
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Menu Balok"
      TabPicture(0)   =   "menuLimas4.frx":3562A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Menu Bola"
      TabPicture(1)   =   "menuLimas4.frx":35646
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Menu Kerucut"
      TabPicture(2)   =   "menuLimas4.frx":35662
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Menu Kubus"
      TabPicture(3)   =   "menuLimas4.frx":3567E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Menu Limas 3"
      TabPicture(4)   =   "menuLimas4.frx":3569A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Menu Limas 4"
      TabPicture(5)   =   "menuLimas4.frx":356B6
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Menu Prisma"
      TabPicture(6)   =   "menuLimas4.frx":356D2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Menu Tabung"
      TabPicture(7)   =   "menuLimas4.frx":356EE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
   End
End
Attribute VB_Name = "menuLimas4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBtn_Click()
    Me.Hide
    menuPilihan.Show
    
End Sub

Private Sub btnLP_4_Click()
    Dim luas_limas4 As New rumus

    Dim sisi_diagonal_alas As Integer
    Dim sisi_diagonal_tgh As Integer
        
    If p_limas4 = "" Then
        MsgBox ("Length of Surface Can't Empty!")
    ElseIf (t_alas_limas4 = "") Then
        MsgBox ("Width of Surface Can't Empty!")
    ElseIf (t_limas4 = "") Then
        MsgBox ("Height of Pyramid Can't Empty!")
    Else
        
        hasil_limas4 = luas_limas4.luas_limas4(p_limas4, t_alas_limas4, t_limas4)
        
    End If
End Sub

Private Sub btnVolP_4_Click()
Dim vol_limas4 As New rumus

    If p_limas4 = "" Then
        MsgBox ("Length of Surface Can't Empty!")
    ElseIf (t_alas_limas4 = "") Then
        MsgBox ("Width of Surface Can't Empty!")
    ElseIf (t_limas4 = "") Then
        MsgBox ("Height of Pyramid Can't Empty!")
    Else
        hasil_limas4 = vol_limas4.vol_limas4(p_limas4, t_alas_limas4, t_limas4)
        
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

Private Sub p_limas4_Change()
Dim textval As String
Dim numval As String

textval = p_limas4.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    p_limas4.Text = CStr(numval)
  End If
  
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 5
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

Private Sub t_alas_limas4_Change()
Dim textval As String
Dim numval As String

textval = t_alas_limas4.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    t_alas_limas4.Text = CStr(numval)
  End If
  
End Sub

Private Sub t_limas4_Change()
Dim textval As String
Dim numval As String

textval = t_limas4.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    t_limas4.Text = CStr(numval)
  End If
  
End Sub

