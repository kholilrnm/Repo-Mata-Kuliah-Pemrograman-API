VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form menuKerucut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cone Menu"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
   Icon            =   "menuKerucut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuKerucut.frx":2AA22
   ScaleHeight     =   6450
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox hasil_kerucut 
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
   Begin VB.TextBox jari_kerucut 
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
   Begin VB.TextBox tinggi_kerucut 
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
      Tab             =   2
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Menu Balok"
      TabPicture(0)   =   "menuKerucut.frx":347D8
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Menu Bola"
      TabPicture(1)   =   "menuKerucut.frx":347F4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Menu Kerucut"
      TabPicture(2)   =   "menuKerucut.frx":34810
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Menu Kubus"
      TabPicture(3)   =   "menuKerucut.frx":3482C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Menu Limas 3"
      TabPicture(4)   =   "menuKerucut.frx":34848
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Menu Limas 4"
      TabPicture(5)   =   "menuKerucut.frx":34864
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Menu Prisma"
      TabPicture(6)   =   "menuKerucut.frx":34880
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Menu Tabung"
      TabPicture(7)   =   "menuKerucut.frx":3489C
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
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
Dim luas_kerucut As New rumus

  If jari_kerucut = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    ElseIf (tinggi_kerucut = "") Then
        MsgBox ("Height of Cone Can't Empty!")
    Else
        hasil_kerucut = luas_kerucut.luas_kerucut(jari_kerucut, tinggi_kerucut)
    End If
End Sub

Private Sub btnVolKer_Click()
Dim vol_kerucut As New rumus

  If jari_kerucut = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    ElseIf (tinggi_kerucut = "") Then
        MsgBox ("Height of Cone Can't Empty!")
    Else
        hasil_kerucut = vol_kerucut.vol_kerucut(jari_kerucut, tinggi_kerucut)
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


Private Sub jari_kerucut_Change()
Dim textval As String
Dim numval As String

textval = jari_kerucut.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    jari_kerucut.Text = CStr(numval)
  End If
  
End Sub
Private Sub Form_Load()
    SSTab1.Tab = 2
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

Private Sub tinggi_kerucut_Change()
Dim textval As String
Dim numval As String

textval = tinggi_kerucut.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    tinggi_kerucut.Text = CStr(numval)
  End If
  
End Sub

