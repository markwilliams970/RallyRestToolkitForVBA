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
Public Property Get totalResultCount() As Long
    totalResultCount = prvTotalResultCount
End Property
''''''''''''''''''''''
' Results collection property
''''''''''''''''''''''
Public Property Set RawResults(ByRef value As Object)
    Dim QueryResultCollection As Object
    Dim totalResultCount As Long
    Set prvRawResults = value
    Set QueryResultCollection = prvRawResults("QueryResult")
    totalResultCount = QueryResultCollection("TotalResultCount")
    prvTotalResultCount = totalResultCount
    Set prvResults = QueryResultCollection("Results")
    
    Set QueryResultCollection = Nothing
    
End Property
Public Property Get Results() As Object
    Set Results = prvResults
End Property

''''''''''''''''''''''
' Add to Results collection property
''''''''''''''''''''''
Public Sub AddRawResults(ByRef value As Object)
    Dim QueryResultCollection As Object
    Dim theseResults As Object
    Dim K As Variant
    Dim Keys As Variant
    Dim blah As String
    Set QueryResultCollection = value("QueryResult")
    Set theseResults = QueryResultCollection("Results")
    
    Call concatenateResults(theseResults)
End Sub

Private Sub concatenateResults(ByVal value As Object)

    Dim thisResult As Object
    Dim i As Long
    Dim newResultsCollection As Collection
    Dim existSize As Long, addedSize As Long, newResultCount As Long
    
    ' Added Size
    addedSize = value.Count
    
    ' Existing Size
    existSize = prvResults.Count
    
    ' Size of existing result set plus those being added
    newResultCount = existSize + addedSize
    
    Set newResultsCollection = New Collection
       
    ' Concatenate Collections
    For Each result In prvResults
        Set thisResult = result
        newResultsCollection.Add thisResult
    Next
    For Each result In value
        Set thisResult = result
        newResultsCollection.Add thisResult
    Next

    Set prvResults = newResultsCollection
    Set newResultsCollection = Nothing
    
End Sub

