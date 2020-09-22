Attribute VB_Name = "XRC4"
'RC4 Standardised Encryption Algorithm
Option Explicit

Public Function RC4_Crypt(ByVal txt As String, _
                           ByVal KeyWord As String) As String

  
  Dim rc4Next_Rand As Long
  Dim RC4S(255)    As Long
  Dim RC4K(255)    As Long
  Dim RC4I         As Long
  Dim RC4J         As Long
  Dim RC4T         As Long
  Dim x            As Long
  Dim keypos       As Long
  Dim t            As Long
  Dim j            As Long
  Dim KeyWordLen   As Long

    KeyWordLen = Len(KeyWord)
    keypos = 0
    For x = 0 To 255
        keypos = keypos + 1
        If keypos >= KeyWordLen Then
            keypos = 1
        End If
        RC4S(x) = x
        RC4K(keypos - 1) = Asc(Mid$(KeyWord, keypos, 1))
    Next x
    KeyWord = ""
    For x = 0 To 255
        j = (j + RC4S(x) + RC4K(x)) Mod 256
        t = RC4S(x)
        RC4S(x) = RC4S(j)
        RC4S(j) = t
    Next x
    RC4I = 0
    RC4J = 0
    RC4T = 0
    For t = 1 To Len(txt)
        RC4I = (RC4I + 1) Mod 256
        RC4J = (RC4J + RC4S(RC4I)) Mod 256
        RC4T = RC4S(RC4I)
        RC4S(RC4I) = RC4S(RC4J)
        RC4S(RC4J) = RC4T
        RC4T = (RC4S(RC4I) + RC4S(RC4J)) Mod 256
        rc4Next_Rand = RC4S(RC4T)
        Mid$(txt, t, 1) = Chr$(Asc(Mid$(txt, t, 1)) Xor rc4Next_Rand)
    Next t
    RC4_Crypt = txt

End Function

Public Function RC4_UnCrypt(ByVal txt As String, _
                             ByVal KeyWord As String) As String


    RC4_UnCrypt = RC4_Crypt(txt, KeyWord)

End Function

Public Function RC4_CryptV(ByVal txt As String, ByVal KeyWord As String) As String

    RC4_CryptV = RC4_Crypt("XRC4-01.00", "xrc4") + RC4_Crypt(txt, KeyWord)

End Function


