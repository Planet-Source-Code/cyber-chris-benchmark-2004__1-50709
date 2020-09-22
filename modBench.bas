Attribute VB_Name = "modBench"
'(C) Copyright by Cyber Chris
' Email: cyber_chris235@gmx.net

Option Explicit

Public Sub HeapSort(vArray As Variant, _
                    Optional Ascending As Boolean = True)   'A complex sort algorithm

  
  Dim EndIdx  As Long
  Dim l       As Long
  Dim rEndIdx As Long
  Dim j       As Long
  Dim i       As Long

  Dim Flag    As Boolean
  Dim TempNew As Variant
    
    If Not IsArray(vArray) Then
        Exit Sub
    End If
    EndIdx = UBound(vArray)
    l = EndIdx \ 2 + 1
    rEndIdx = EndIdx
    If Ascending Then
        Do While l > 1
            l = l - 1
            TempNew = vArray(l)
            j = l
            Flag = False
            Do While Flag = False
                i = j
                j = j * 2
                If j < rEndIdx Then
                    If vArray(j) < vArray(j + 1) Then
                        j = j + 1
                    End If
                End If
                If j > rEndIdx Then
                    vArray(i) = TempNew
                    Flag = True
                 Else
                    If TempNew > vArray(j) Then
                        vArray(i) = TempNew
                        Flag = True
                    End If
                    If Not TempNew Then
                        vArray(i) = vArray(j)
                    End If
                End If
            Loop
        Loop
        Do While rEndIdx > 2
            TempNew = vArray(rEndIdx)
            vArray(rEndIdx) = vArray(1)
            rEndIdx = rEndIdx - 1
            j = l
            Flag = False
            Do While Flag = False
                i = j
                j = j * 2
                If j < rEndIdx Then
                    If vArray(j) < vArray(j + 1) Then
                        j = j + 1
                    End If
                End If
                If j > rEndIdx Then
                    vArray(i) = TempNew
                    Flag = True
                 Else
                    If TempNew > vArray(j) Then
                        vArray(i) = TempNew
                        Flag = True
                     Else
                        vArray(i) = vArray(j)
                    End If
                End If
            Loop
        Loop
        TempNew = vArray(2)
        vArray(2) = vArray(1)
        vArray(1) = TempNew
     Else
        Do While l > 1
            l = l - 1
            TempNew = vArray(l)
            j = l
            Flag = False
            Do While Flag = False
                i = j
                j = j * 2
                If j < rEndIdx Then
                    If vArray(j) > vArray(j + 1) Then
                        j = j + 1
                    End If
                End If
                If j > rEndIdx Then
                    vArray(i) = TempNew
                    Flag = True
                End If
                If Not j Then
                    If TempNew < vArray(j) Then
                        vArray(i) = TempNew
                        Flag = True
                     Else
                        vArray(i) = vArray(j)
                    End If
                End If
            Loop
        Loop
        Do While rEndIdx > 2
            TempNew = vArray(rEndIdx)
            vArray(rEndIdx) = vArray(1)
            rEndIdx = rEndIdx - 1
            j = l
            Flag = False
            Do While Flag = False
                i = j
                j = j * 2
                If j < rEndIdx Then
                    If vArray(j) > vArray(j + 1) Then
                        j = j + 1
                    End If
                End If
                If j > rEndIdx Then
                    vArray(i) = TempNew
                    Flag = True
                End If
                If Not j Then
                    If TempNew < vArray(j) Then
                        vArray(i) = TempNew
                        Flag = True
                     Else
                        vArray(i) = vArray(j)
                    End If
                End If
            Loop
        Loop
        TempNew = vArray(2)
        vArray(2) = vArray(1)
        vArray(1) = TempNew
    End If

End Sub


