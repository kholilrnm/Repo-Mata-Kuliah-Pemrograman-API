VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rumus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function vol_balok(ByVal p As Integer, ByVal l As Integer, ByVal t As Integer)
vol_balok = p * l * t
End Function

Public Function luas_balok(ByVal p As Integer, ByVal l As Integer, ByVal t As Integer)
luas_balok = 2 * ((p * l) + (l * t) + (p * t))
End Function

Public Function luas_1sisi_balok(ByVal p As Integer, ByVal l As Integer, ByVal t As Integer)
luas_1sisi_balok = Sqr((p * p) + (l * l) + (t * t))
End Function

Public Function keliling_balok(ByVal p As Integer, ByVal l As Integer, ByVal t As Integer)
keliling_balok = 4 * (p + l + t)
End Function

Public Function vol_bola(ByVal r As Integer)
vol_bola = 4 / 3 * 22 / 7 * r * r * r
End Function

Public Function luas_bola(ByVal r As Integer)
luas_bola = 4 * 22 / 7 * r * r
End Function

Public Function vol_kerucut(ByVal r As Integer, ByVal t As Integer)
vol_kerucut = 1 / 3 * 22 / 7 * r * r * t
End Function

Public Function luas_kerucut(ByVal r As Integer, ByVal t As Integer)
luas_kerucut = ((22 / 7 * r * r) + (22 / 7 * r * (Sqr((r * r) + (t * t)))))
End Function

Public Function vol_kubus(ByVal s As Integer)
vol_kubus = s * s * s
End Function

Public Function luas_kubus(ByVal s As Integer)
luas_kubus = 6 * s * s
End Function

Public Function keliling_kubus(ByVal s As Integer)
keliling_kubus = 12 * s
End Function

Public Function kel_1sisi_kubus(ByVal s As Integer)
kel_1sisi_kubus = s * s
End Function

Public Function vol_limas3(ByVal p_alas3 As Integer, ByVal t_alas_limas3 As Integer, ByVal t_limas3 As Integer)
vol_limas3 = 1 / 3 * 1 / 2 * p_alas3 * t_alas_limas3 * t_limas3

End Function

Public Function luas_limas3(ByVal p_alas3 As Integer, ByVal t_alas_limas3 As Integer, ByVal t_limas3 As Integer)

luas_limas3 = ((1 / 2 * p_alas3 * t_alas_limas3) + (1 / 2 * p_alas3 * t_limas3) + (1 / 2 * t_alas_limas3 * t_limas3)) + (1 / 2 * Sqr((p_alas3 * p_alas3) + (t_alas_limas3 * t_alas_limas3)) * Sqr((1 / 2 * Sqr((p_alas3 * p_alas3) + (t_alas_limas3 * t_alas_limas3)) * 1 / 2 * Sqr((p_alas3 * p_alas3) + (t_alas_limas3 * t_alas_limas3))) + (Sqr((p_alas3 * p_alas3) + (t_limas3 * t_limas3)) * Sqr((p_alas3 * p_alas3) + (t_limas3 * t_limas3)))))

End Function

Public Function vol_limas4(ByVal p_alas4 As Integer, ByVal t_alas_limas4 As Integer, ByVal t_limas4 As Integer)

vol_limas4 = 1 / 3 * p_alas4 * t_alas_limas4 * t_limas4

End Function

Public Function luas_limas4(ByVal p_alas4 As Integer, ByVal t_alas_limas4 As Integer, ByVal t_limas4 As Integer)
sisi_diagonal_alas = 1 / 2 * p_alas4
sisi_diagonal_tgh = Sqr((sisi_diagonal_alas * sisi_diagonal_alas) + (t_limas4 * t_limas4))
luas_limas4 = (p_alas4 * t_alas_limas4) + (4 * (1 / 2 * p_alas4 * sisi_diagonal_tgh))

End Function

Public Function luas_prismaseg(ByVal p_prism4 As Integer, ByVal t_alas_prism As Integer, ByVal t_prism4 As Integer)

luas_prismaseg = (3 * (p_prism4) * t_alas_prism) + (2 * 1 / 2 * p_prism4 * t_prism4)

End Function

Public Function vol_prismaseg(ByVal p_prism4 As Integer, ByVal t_alas_prism As Integer, ByVal t_prism4 As Integer)
vol_prismaseg = p_prism4 * t_alas_prism * t_prism4

End Function

Public Function vol_tabung(ByVal jari_tabung As Integer, ByVal t_tabung As Integer)

vol_tabung = 22 / 7 * jari_tabung * jari_tabung * t_tabung
End Function

Public Function luas_tabung(ByVal jari_tabung As Integer, ByVal t_tabung As Integer)

luas_tabung = ((2 * 22 / 7 * jari_tabung * jari_tabung) + (2 * 22 / 7 * jari_tabung * t_tabung))
End Function

