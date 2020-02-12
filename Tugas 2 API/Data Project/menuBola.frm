VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form menuBola 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ball Menu"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
   Icon            =   "menuBola.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuBola.frx":2AA22
   ScaleHeight     =   6450
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox jari_bola 
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
      Left            =   3860
      TabIndex        =   3
      Top             =   1820
      Width           =   2000
   End
   Begin VB.CommandButton btnVolBola 
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
      Top             =   2680
      Width           =   1550
   End
   Begin VB.CommandButton btnLPBola 
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
      Top             =   2680
      Width           =   1550
   End
   Begin VB.TextBox hasil_bola 
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
      Left            =   3860
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3885
      Width           =   2000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   540
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11590
      _ExtentX        =   20452
      _ExtentY        =   953
      _Version        =   393216
      MousePointer    =   1
      Tabs            =   8
      Tab             =   1
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Menu Balok"
      TabPicture(0)   =   "menuBola.frx":346A3
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Menu Bola"
      TabPicture(1)   =   "menuBola.frx":346BF
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Menu Kerucut"
      TabPicture(2)   =   "menuBola.frx":346DB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Menu Kubus"
      TabPicture(3)   =   "menuBola.frx":346F7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Menu Limas 3"
      TabPicture(4)   =   "menuBola.frx":34713
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Menu Limas 4"
      TabPicture(5)   =   "menuBola.frx":3472F
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Menu Prisma"
      TabPicture(6)   =   "menuBola.frx":3474B
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Menu Tabung"
      TabPicture(7)   =   "menuBola.frx":34767
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
   End
End
Attribute VB_Name = "menuBola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBtn_Click()
    Me.Hide
    menuPilihan.Show
    
End Sub

Private Sub btnLPBola_Click()
Dim luas_bola As New rumus
     If jari_bola = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    Else
        hasil_bola = luas_bola.luas_bola(jari_bola)
    End If
End Sub

Private Sub btnVolBola_Click()
Dim vol_bola As New rumus

    If jari_bola = "" Then
        MsgBox ("Radius of Surface Can't Empty!")
    Else
        hasil_bola = vol_bola.vol_bola(jari_bola)
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
    SSTab1.Tab = Null

End Sub
Private Sub Form_Load()
    SSTab1.Tab = 1
End Sub

Private Sub Label4_Click()

End Sub

Private Sub jari_bola_Change()
Dim textval As String
Dim numval As String

textval = jari_bola.Text
If IsNumeric(textval) Then
    numval = textval
  Else
    jari_bola.Text = CStr(numval)
  End If
  
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

