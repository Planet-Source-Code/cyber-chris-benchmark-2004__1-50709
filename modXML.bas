Attribute VB_Name = "modXML"
'(C) Copyright by Cyber Chris
' Email: cyber_chris235@gmx.net


Option Explicit
Private Xml                    As String
Private Const Version          As String = "2004"
Private Const Constructor      As String = "CSBenchmark"

Private Sub AddXML(strData As String)    'Just to make programming easier :-)

    Xml = Xml & strData

End Sub

Public Function BuildXML() As String    'This builds the "main" XML Data

  'Note: This is my interpretation of an XML Datafile including program
  'information and the main "BODY"

    Xml = ""        'Reset
    AddXML "<XML>"
    AddXML "<Header>"   'Build header data
    AddXML "<App>"      'Add the App information
    AddXML Construct("Constructor", Construct("Name", Constructor) & Construct("Version", Version)) 'Everything about this program
    AddXML Construct("Created", Now)    'Creation time
    AddXML "</App>"
    AddXML "</Header>"
    AddXML "<Body>"     'Build the Body with the Benchmark results
    With frmResults
        AddXML Construct("Counter_Benchmark", .Label1.Caption)
        AddXML Construct("Extended_Counter_Benchmark", .Label2.Caption)
        AddXML Construct("Drawing_Benchmark", .Label3.Caption)
        AddXML Construct("Timer_Benchmark", .Label4.Caption)
        AddXML Construct("Cryption_Benchmark", .Label5.Caption)
        AddXML Construct("Sort_Benchmark", .Label6.Caption)
    End With
    AddXML "</Body>"
    AddXML "</XML>"
    BuildXML = Xml

End Function

Private Function Construct(strName As String, _
                           strData As String) As String    'Also to make programming more easy

    Construct = "<" & strName & ">" & strData & "</" & strName & ">"

End Function

'This Function isn't really needed in this program,
'but I'd added it just for completation
'Public Function FindItem(Xml As String, item As String) As String   'XML Parser
'Dim temp  As Long
'Dim loop2 As Long
'For temp = 1 To Len(Xml)
'If Mid$(Xml, temp, Len(item) + 2) = "<" & item & ">" Then
'For loop2 = temp + Len(item) + 2 To temp + Len(item) + 514
'If Mid$(Xml, loop2, Len(item) + 3) <> "</" & item & ">" Then
'FindItem = FindItem & Mid$(Xml, loop2, 1)
'Else
'Exit Function
'End If
'Next loop2
'End If
'Next temp
'End Function


