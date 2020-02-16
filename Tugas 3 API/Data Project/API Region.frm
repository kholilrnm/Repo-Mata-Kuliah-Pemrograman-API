VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18765
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   18765
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


Private Sub Form_Load()
k1 = CreateRectRgn(120, 155, 140, 55)

k2 = CreateRectRgn(130, 75, 160, 95)
k4 = CreateRectRgn(160, 55, 180, 75)

k3 = CreateRectRgn(130, 111, 160, 131)
k5 = CreateRectRgn(160, 130, 180, 155)

h1 = CreateRectRgn(210, 155, 230, 55)
h2 = CreateRectRgn(250, 155, 270, 55)
h3 = CreateRectRgn(230, 95, 250, 115)

o1 = CreateEllipticRgn(300, 155, 365, 55)
o2 = CreateEllipticRgn(320, 130, 345, 80)

l1 = CreateRectRgn(390, 155, 410, 55)
l2 = CreateRectRgn(410, 155, 450, 135)

i1 = CreateRectRgn(485, 155, 505, 55)

l3 = CreateRectRgn(545, 155, 565, 55)
l4 = CreateRectRgn(565, 155, 605, 135)

r1 = CreateRoundRectRgn(120, 215, 180, 275, 10, 10)
r2 = CreateRectRgn(140, 235, 160, 255)
r3 = CreateRectRgn(120, 215, 140, 315)

r4 = CreateRectRgn(170, 270, 189, 290)
r5 = CreateRectRgn(182, 290, 216, 315)

a1 = CreateRoundRectRgn(235, 215, 300, 315, 10, 10)
a2 = CreateRectRgn(255, 235, 280, 255)
a3 = CreateRectRgn(255, 275, 280, 314)

c1 = CreateRoundRectRgn(335, 215, 405, 316, 10, 10)
c2 = CreateRoundRectRgn(355, 235, 405, 296, 10, 10)

h4 = CreateRectRgn(445, 215, 465, 316)
h5 = CreateRectRgn(465, 255, 490, 276)
h6 = CreateRectRgn(490, 215, 510, 316)

m1 = CreateRoundRectRgn(545, 215, 635, 316, 10, 10)
m2 = CreateRectRgn(565, 245, 580, 315)
m3 = CreateRectRgn(600, 245, 615, 315)

a4 = CreateRoundRectRgn(675, 215, 745, 315, 10, 10)
a5 = CreateRectRgn(695, 235, 725, 255)
a6 = CreateRectRgn(695, 275, 725, 314)

n1 = CreateRoundRectRgn(785, 215, 875, 316, 10, 10)
n2 = CreateRectRgn(810, 245, 850, 315)

n3 = CreateRoundRectRgn(120, 375, 210, 476, 10, 10)
n4 = CreateRectRgn(145, 405, 185, 475)

u1 = CreateRoundRectRgn(240, 375, 310, 476, 10, 10)
u2 = CreateRectRgn(260, 375, 290, 445)

r6 = CreateRoundRectRgn(335, 375, 405, 432, 10, 10)
r7 = CreateRectRgn(335, 375, 355, 476)
r8 = CreateRectRgn(355, 395, 385, 410)

r9 = CreateRectRgn(390, 430, 415, 450)
r10 = CreateRectRgn(400, 450, 425, 476)

m4 = CreateRoundRectRgn(545, 375, 635, 476, 10, 10)
m5 = CreateRectRgn(565, 405, 580, 475)
m6 = CreateRectRgn(600, 405, 615, 475)

a7 = CreateRoundRectRgn(675, 375, 745, 475, 10, 10)
a8 = CreateRectRgn(695, 385, 725, 415)
a9 = CreateRectRgn(695, 435, 725, 474)

n5 = CreateRoundRectRgn(785, 375, 875, 476, 10, 10)
n6 = CreateRectRgn(810, 405, 850, 475)

a10 = CreateRoundRectRgn(915, 375, 985, 475, 10, 10)
a11 = CreateRectRgn(935, 385, 965, 415)
a12 = CreateRectRgn(935, 435, 965, 474)

b1 = CreateRectRgn(1020, 375, 1105, 475)
b2 = CreateRectRgn(1040, 375, 1105, 415)
b3 = CreateRectRgn(1040, 435, 1080, 455)


CombineRgn k1, k2, k1, 2
CombineRgn k1, k4, k1, 2
CombineRgn k1, k3, k1, 2
CombineRgn k1, k5, k1, 2

CombineRgn k1, h1, k1, 2
CombineRgn k1, h2, k1, 2
CombineRgn k1, h3, k1, 2

CombineRgn k1, o1, k1, 2
CombineRgn k1, o2, k1, 3

CombineRgn k1, l1, k1, 2
CombineRgn k1, l2, k1, 2

CombineRgn k1, i1, k1, 2

CombineRgn k1, l3, k1, 2
CombineRgn k1, l4, k1, 2

CombineRgn k1, r1, k1, 2
CombineRgn k1, r2, k1, 3
CombineRgn k1, r3, k1, 2

CombineRgn k1, r4, k1, 2
CombineRgn k1, r5, k1, 2

CombineRgn k1, a1, k1, 2
CombineRgn k1, a2, k1, 3
CombineRgn k1, a3, k1, 3

CombineRgn k1, c1, k1, 2
CombineRgn k1, c2, k1, 3

CombineRgn k1, h4, k1, 2
CombineRgn k1, h5, k1, 2
CombineRgn k1, h6, k1, 2

CombineRgn k1, m1, k1, 2
CombineRgn k1, m2, k1, 3
CombineRgn k1, m3, k1, 3

CombineRgn k1, a4, k1, 2
CombineRgn k1, a5, k1, 3
CombineRgn k1, a6, k1, 3

CombineRgn k1, n1, k1, 2
CombineRgn k1, n2, k1, 3

CombineRgn k1, n3, k1, 2
CombineRgn k1, n4, k1, 3

CombineRgn k1, u1, k1, 3
CombineRgn k1, u2, k1, 3

CombineRgn k1, r6, k1, 2
CombineRgn k1, r7, k1, 2
CombineRgn k1, r8, k1, 3

CombineRgn k1, r9, k1, 2
CombineRgn k1, r10, k1, 2

CombineRgn k1, m4, k1, 2
CombineRgn k1, m5, k1, 3
CombineRgn k1, m6, k1, 3

CombineRgn k1, a7, k1, 2
CombineRgn k1, a8, k1, 3
CombineRgn k1, a9, k1, 3

CombineRgn k1, n5, k1, 2
CombineRgn k1, n6, k1, 3

CombineRgn k1, a10, k1, 2
CombineRgn k1, a11, k1, 3
CombineRgn k1, a12, k1, 3

CombineRgn k1, b1, k1, 2
CombineRgn k1, b2, k1, 3
CombineRgn k1, b3, k1, 3


SetWindowRgn Me.hWnd, k1, True
End Sub

