VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   540
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11590
      _ExtentX        =   20452
      _ExtentY        =   953
      _Version        =   393216
      MousePointer    =   1
      Tabs            =   8
      Tab             =   7
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Menu Balok"
      TabPicture(0)   =   "menuTabung.frx":34629
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Menu Bola"
      TabPicture(1)   =   "menuTabung.frx":34645
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Menu Kerucut"
      TabPicture(2)   =   "menuTabung.frx":34661
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Menu Kubus"
      TabPicture(3)   =   "menuTabung.frx":3467D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Menu Limas 3"
      TabPicture(4)   =   "menuTabung.frx":34699
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Menu Limas 4"
      TabPicture(5)   =   "menuTabung.frx":346B5
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Menu Prisma"
      TabPicture(6)   =   "menuTabung.frx":346D1
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Menu Tabung"
      TabPicture(7)   =   "menuTabung.frx":346ED
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).ControlCount=   0
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
Private Sub Form_Load()
    SSTab1.Tab = 7
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
