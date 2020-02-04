VERSION 5.00
Begin VB.Form menuPilihan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choice Menu"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuPilihan.frx":0000
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
      Height          =   440
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   28
      Width           =   960
   End
   Begin VB.CommandButton btnMenuBola 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Ball"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4225
      Width           =   1575
   End
   Begin VB.CommandButton btnMenuKerucut 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Cone"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3510
      Width           =   1575
   End
   Begin VB.CommandButton btnMenuTabung 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Tube"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2770
      Width           =   1575
   End
   Begin VB.CommandButton btnMenuLimas3 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Pyramid 3"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2050
      Width           =   1560
   End
   Begin VB.CommandButton btnMenuLimas4 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Pyramid 4"
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
      Left            =   6590
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4210
      Width           =   1575
   End
   Begin VB.CommandButton btnMenuPrisma 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Prism"
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
      Left            =   6590
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3490
      Width           =   1575
   End
   Begin VB.CommandButton btnMenuBalok 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Block"
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
      Left            =   6590
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2770
      Width           =   1575
   End
   Begin VB.CommandButton btnMenuKubus 
      BackColor       =   &H000080FF&
      Caption         =   "Calc Cube"
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
      Left            =   6590
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2060
      Width           =   1560
   End
End
Attribute VB_Name = "menuPilihan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub backBtn_Click()
    Me.Hide
    menuUtama.Show
    
End Sub

Private Sub btnMenuBalok_Click()
    Me.Hide
    menuBalok.Show
    
End Sub

Private Sub btnMenuBola_Click()
    Me.Hide
    menuBola.Show
    
End Sub

Private Sub btnMenuKerucut_Click()
    Me.Hide
    menuKerucut.Show
    
End Sub

Private Sub btnMenuKubus_Click()
    Me.Hide
    menuKubus.Show
    
End Sub

Private Sub btnMenuLimas3_Click()
    Me.Hide
    menuLimas3.Show
    
End Sub

Private Sub btnMenuLimas4_Click()
    Me.Hide
    menuLimas4.Show
    
End Sub

Private Sub btnMenuPrisma_Click()
    Me.Hide
    menuPrismaSeg.Show
    
End Sub

Private Sub btnMenuTabung_Click()
    Me.Hide
    menuTabung.Show
    
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

