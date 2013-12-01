VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RallyQueryResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    ' Class Attributes

Private prvStartIndex As Long
Private prvTotalResultCount As Long
Private prvRawResults As Object
Private prvResults As Object
' Class Constructor
Private Sub Class_Initialize()
    ' Default project scoping to flat
    prvStartIndex = 1
    prvTotalResultCount = 0

End Sub

' Setters and Getters
' In VBA this is done with Property Get and Property Let functions (or Property Set for object type variables).

''''''''''''''''''''''
' StartIndex property
''''''''''''''''''''''
Public Property Get StartIndex() As Long
    StartIndex = prvStartIndex
End Property

Public Property Let StartIndex(value As Long)
    prvStartIndex = value
End Property
''''''''''''''''''''''
' TotalResultCount property
''''''''''''''''''''''
Public Property Get TotalResultCount() As Long
    TotalResultCount = prvTotalResultCount
End Property
''''''''''''''''''''''
' Results collection property
''''''''''''''''''''''
Public Property Set RawResults(ByRef value As Object)
    Dim QueryResultCollection As Object
    Dim TotalResultCount As Long
    Set prvRawResults = value
    Set QueryResultCollection = prvRawResults("QueryResult")
    TotalResultCount = QueryResultCollection("TotalResultCount")
    prvTotalResultCount = TotalResultCount
    Set prvResults = QueryResultCollection("Results")
End Property
Public Property Get Results() As Object
    Set Results = prvResults
End Property
