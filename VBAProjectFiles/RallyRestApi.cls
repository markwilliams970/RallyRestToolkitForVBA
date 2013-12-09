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
Private prvRallyQueryUrl As String
Private prvRallyRequest As RallyRequest
Private JSON As JSONLib

' Class Constructor
Private Sub Class_Initialize()

    Set JSON = New JSONLib

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
    Dim resultCollection As Object
    Dim queryUrl As String
    Dim myRallyRequest As RallyRequest
    Dim myQueryResult As RallyQueryResult
    Dim myTotalResultCount As Long, nPages As Long, i As Long
    
    ' Set private reference to RallyRequest
    Set myRallyRequest = value
    
    ' Build Query URL
    queryUrl = buildQueryUrl(myRallyRequest)
    
    ' Query Rally
    Set queryResponse = prvRallyConnection.rallyGET(queryUrl)
    
    ' Populate Result object
    Set myQueryResult = New RallyQueryResult
    Set myQueryResult.RawResults = queryResponse
    
    ' Find out how many pages are in result set
    Set resultCollection = queryResponse("QueryResult")
    myTotalResultCount = resultCollection("TotalResultCount")
    nPages = myTotalResultCount / myRallyRequest.pageSize
    
    ' If we have multiple pages, page through and collect entire result set
    If nPages > 1 Then
        For i = 2 To nPages
            ' Set Start to next page
            myRallyRequest.Start = myRallyRequest.pageSize * (i - 1) + 1
            
            ' Build Query URL
            queryUrl = buildQueryUrl(myRallyRequest)
            
            ' Query Rally
            Set queryResponse = prvRallyConnection.rallyGET(queryUrl)
    
            ' Update Result object
            Call myQueryResult.AddRawResults(queryResponse)
        Next i
        Set queryResponse = Nothing
    End If
    
    Set Query = myQueryResult

End Function

Private Function buildQueryUrl(value As RallyRequest) As String

    Dim queryString As String
    Dim myRallyRequest As RallyRequest
    Dim myRallyQuery As RallyQuery
    Dim myRallyQueryUrl As String
    
    Set myRallyRequest = value
    
    myRallyQueryUrl = webServiceUrl() & "/" & myRallyRequest.ArtifactName & "?query=" & _
        myRallyRequest.Query.queryString
        
    If Not IsNull(myRallyRequest.Fetch) Then
        myRallyQueryUrl = myRallyQueryUrl & "&fetch=" & myRallyRequest.Fetch
    End If
    If Not IsNull(myRallyRequest.Workspace) Then
        myRallyQueryUrl = myRallyQueryUrl & "&workspace=" & myRallyRequest.Workspace
    End If
    If Not IsNull(myRallyRequest.Project) Then
        myRallyQueryUrl = myRallyQueryUrl & "&project=" & myRallyRequest.Project
    End If
    If Not IsNull(myRallyRequest.pageSize) Then
        myRallyQueryUrl = myRallyQueryUrl & "&pagesize=" & myRallyRequest.pageSize
    End If
    If Not IsNull(myRallyRequest.Start) Then
        myRallyQueryUrl = myRallyQueryUrl & "&start=" & myRallyRequest.Start
    End If
        
    urlEncodedQueryUrl = URLEncode(myRallyQueryUrl)
    buildQueryUrl = myRallyQueryUrl
End Function

Public Function Create(ArtifactType As String, workspaceRef As String, value As RallyObject) As String
    Dim myRallyCreateUrl As String
    Dim myRallyObjectJSON As Object
    Dim myCreateObject As Object
    Dim myCreateJSON As String, myURLEncodedJSON As String
    Dim myCreateResponse As String
    
    Set myRallyObject = value.JSONObject
    
    Set myCreateObject = CreateObject("Scripting.Dictionary")
    myCreateObject.Add key:=ArtifactType, Item:=myRallyObject
    
    myCreateJSON = JSON.toString(myCreateObject)
    myURLEncodedJSON = URLEncode(myCreateJSON)
    
    myRallyCreateUrl = buildCreateUrl(ArtifactType)
    
    myCreateResponse = prvRallyConnection.rallyPOST(myRallyCreateUrl, workspaceRef, myCreateJSON)
    Create = myCreateResponse

End Function

Private Function buildCreateUrl(ArtifactType As String) As String
    Dim myRallyCreateUrl As String
    
    myRallyCreateUrl = webServiceUrl() & "/" & ArtifactType & "/create"
    buildCreateUrl = myRallyCreateUrl
End Function
