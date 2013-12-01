VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RallyRestApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Attributes

Private prvRallyConnection As RallyConnection
Private prvRallyRequest As RallyRequest
Private prvRallyQueryUrl As String

' Class Constructor
Private Sub Class_Initialize()

End Sub

' Setters and Getters
' In VBA this is done with Property Get and Property Let functions (or Property Set for object type variables).

''''''''''''''''''''''
' RequestType property
''''''''''''''''''''''
Public Property Get webServiceUrl() As String
    webServiceUrl = prvRallyConnection.RallyUrl & "/webservice/" & prvRallyConnection.WsapiVersion
End Property
Public Property Get RallyConnection() As RallyConnection
    Set RallyConnection = prvRallyConnection
End Property
Public Property Let RallyConnection(value As RallyConnection)
    Set prvRallyConnection = value
End Property

Public Property Let RallyRequest(value As RallyRequest)
    Set prvRallyRequest = value
End Property
Public Property Get RallyRequest() As RallyRequest
    Set RallyRequest = prvRallyRequest
End Property

Public Function Query(value As RallyRequest) As RallyQueryResult

    Dim queryResponse As Object
    Dim queryUrl As String
    Dim myQueryResult As RallyQueryResult
    
    ' Set private reference to RallyRequest
    Set prvRallyRequest = value
    
    ' Build Query URL
    queryUrl = buildQueryUrl()
    
    ' Query Rally
    Set queryResponse = prvRallyConnection.rallyGET(queryUrl)
    
    Set myQueryResult = New RallyQueryResult
    Set myQueryResult.RawResults = queryResponse
    
    Set Query = myQueryResult
    Set queryResponse = Nothing
End Function

Private Function buildQueryUrl() As String

    Dim queryString As String
    Dim myRallyRequest As RallyRequest
    Dim myRallyQuery As RallyQuery
    Dim urlEncodedQueryUrl As String
    Dim blah As String
    
    Set myRallyRequest = RallyRequest()
    Set myRallyQuery = myRallyRequest.Query
    queryString = myRallyQuery.queryString
    
    prvRallyQueryUrl = webServiceUrl() & "/" & prvRallyRequest.ArtifactName & "?query=" & _
        prvRallyRequest.Query.queryString
        
    If Not IsNull(prvRallyRequest.Fetch) Then
        prvRallyQueryUrl = prvRallyQueryUrl & "&fetch=" & prvRallyRequest.Fetch
    End If
    If Not IsNull(prvRallyRequest.Workspace) Then
        prvRallyQueryUrl = prvRallyQueryUrl & "&workspace=" & prvRallyRequest.Workspace
    End If
    If Not IsNull(prvRallyRequest.Project) Then
        prvRallyQueryUrl = prvRallyQueryUrl & "&project=" & prvRallyRequest.Project
    End If
    If Not IsNull(prvRallyRequest.PageSize) Then
        prvRallyQueryUrl = prvRallyQueryUrl & "&pagesize=" & prvRallyRequest.PageSize
    End If
    
    ' urlEncodedQueryUrl = URLEncode(prvRallyQueryUrl)
    buildQueryUrl = prvRallyQueryUrl
End Function
