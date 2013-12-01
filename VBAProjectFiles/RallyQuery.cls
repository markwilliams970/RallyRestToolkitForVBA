VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RallyQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Attributes

Private prvQueryString As String

' Class Constructor
Private Sub Class_Initialize()

End Sub

' Setters and Getters
' In VBA this is done with Property Get and Property Let functions (or Property Set for object type variables).

''''''''''''''''''''''
' QueryString property
''''''''''''''''''''''
Public Property Get queryString() As String
    queryString = prvQueryString
End Property

Public Property Let queryString(value As String)
    prvQueryString = value
End Property
' Adds an AND condition to the Query
Public Function AddAnd(newCondition As String)
    Dim newQueryString As String
    newQueryString = "(" & prvQueryString & " AND " & newCondition & ")"
    prvQueryString = newQueryString
End Function

' Adds an OR condition to the Query
Public Function AddOr(newCondition As String)
    Dim newQueryString As String
    newQueryString = "(" & prvQueryString & " OR " & newCondition & ")"
    prvQueryString = newQueryString
End Function
