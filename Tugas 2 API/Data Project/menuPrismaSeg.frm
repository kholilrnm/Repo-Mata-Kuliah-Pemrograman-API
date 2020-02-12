VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form menuPrismaSeg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prism Menu"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11505
   Icon            =   "menuPrismaSeg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuPrismaSeg.frx":2AA22
   ScaleHeight     =   6603.571
   ScaleMode       =   0  'User
   ScaleWidth      =   11656.98
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox hasil_prism 
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
      TabIndex        =   5
      Top             =   4800
      Width           =   1974
   End
   Begin VB.CommandButton btnLPPrisma 
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
      Height          =   498
      Left            =   4096
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1530
   End
   Begin VB.CommandButton btnVolPrisma 
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
      Height          =   508
      Left            =   2448
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1530
   End
   Begin VB.TextBox p_prism 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   440
      Left            =   3711
      TabIndex        =   2
      Top             =   1440
      Width           =   1974
   End
   Begin VB.TextBox t_alas_prism 
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
      Left            =   3711
      TabIndex        =   1
      Top             =   2198
      Width           =   1974
   End
   Begin VB.TextBox t_prism 
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
      Left            =   3711
      TabIndex        =   0
      Top             =   2950
      Width           =   1974
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
      Tab             =   6
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Menu Balok"
      TabPicture(0)   =   "menuPrismaSeg.frx":352D5
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Menu Bola"
      TabPicture(1)   =   "menuPrismaSeg.frx":352F1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Menu Kerucut"
      TabPicture(2)   =   "menuPrismaSeg.frx":3530D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Menu Kubus"
      TabPicture(3)   =   "menuPrismaSeg.frx":35329
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Menu Limas 3"
      TabPicture(4)   =   "menuPrismaSeg.frx":35345
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Menu Limas 4"
      TabPicture(5)   =   "menuPrismaSeg.frx":35361
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Menu Prisma"
      TabPicture(6)   =   "menuPrismaSeg.frx":3537D
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Menu Tabung"
      TabPicture(7)   =   "menuPrismaSeg.frx":35399
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
   End
End
Attribute VB_Name = "menuPrismaSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBtn_Click()
    Me.Hide
    menuPilihan.Show
    
End Sub
Private Sub btnLPPrisma_Click()
Dim luas_prism As New rumus

    If p_prism = "" Then
        MsgBox ("Length of Surface Can't Empty!")
    ElseIf (t_alas_prism = "") Then
        MsgBox ("Height of Surface Can't Empty!")
    ElseIf (t_prism = "") Then
        MsgBox ("Height of Prism Can't Empty!")
    Else
        hasil_prism = luas_prism.luas_prismaseg(p_prism, t_alas_prism, t_prism)
    End If
End Sub

Private Sub btnVolPrisma_Click()
Dim vol_prism As New rumus

    If p_prism = "" Then
        MsgBox ("Length of Surface Can't Empty!")
    ElseIf (t_alas_prism = "") Then
        MsgBox ("Height of Surface Can't Empty!")
    ElseIf (t_prism = "") Then
        MsgBox ("Height of Prism Can't Empty!")
    Else
        hasil_prism = vol_prism.vol_prismaseg(p_prism, t_alas_prism, t_prism)
        
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

Private Sub p_prism_Change()
Dim textval As String
Dim numval As String

textval = p_prism.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    p_prism.Text = CStr(numval)
  End If
  
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 6
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

Private Sub t_alas_prism_Change()
Dim textval As String
Dim numval As String

textval = t_alas_prism.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    t_alas_prism.Text = CStr(numval)
  End If
  
End Sub

Private Sub t_prism_Change()
Dim textval As String
Dim numval As String

textval = t_prism.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    t_prism.Text = CStr(numval)
  End If
  
End Sub
