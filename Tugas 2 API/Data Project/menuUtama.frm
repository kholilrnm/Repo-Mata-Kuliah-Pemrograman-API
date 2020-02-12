VERSION 5.00
Begin VB.Form menuUtama 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calc Apps - Build of Shape"
   ClientHeight    =   6450
   ClientLeft      =   150
   ClientTop       =   195
   ClientWidth     =   11505
   Icon            =   "menuUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menuUtama.frx":2AA22
   ScaleHeight     =   6450
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnMainMenu 
      BackColor       =   &H000080FF&
      Caption         =   "Main Menu"
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
      Left            =   3670
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton btnTentang 
      BackColor       =   &H000080FF&
      Caption         =   "About"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton btnKeluar 
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
End
Attribute VB_Name = "menuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnKeluar_Click()
 If MsgBox("Apa anda yakin ingin keluar ?", vbExclamation + vbYesNo) = vbNo Then
        Cancel = 1
 Else
    Unload Me
 End If

End Sub

Private Sub btnMainMenu_Click()
    Me.Hide
    menuBalok.Show
    
End Sub

Private Sub btnTentang_Click()
    Me.Hide
    menuTentang.Show
    
End Sub

Private Sub Command1_Click()
menuPilihan.Show

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

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
menuPilihan.Visible = True

End Sub

