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
Public Property Let RallyConnection(Value As RallyConnection)
    Set prvRallyConnection = Value
End Property

Public Property Let RallyRequest(Value As RallyRequest)
    Set prvRallyRequest = Value
End Property
Public Property Get RallyRequest() As RallyRequest
    Set RallyRequest = prvRallyRequest
End Property
Public Function findWorkspace(WorkspaceName As String) As Object

    Dim queryResponse As Object
    Dim subscriptionResult As Object
    Dim workspaceCollection As Object, workspacesRef As String, workspacesQueryUrl As String
    Dim queryUrl As String
    Dim mySubscriptionRequest As RallyRequest, myWorkspacesRequest As RallyRequest
    Dim myWorkspacesResults As Object, foundWorkspace As Object

    Dim myQueryResult As RallyQueryResult
    Dim myTotalResultCount As Long, nPages As Long, i As Long, pageSize As Long, start As Long
    
    ' Query Subscription Object
    Set mySubscriptionRequest = New RallyRequest
    mySubscriptionRequest.ArtifactName = "subscription"
    
    ' Build Query URL
    queryUrl = buildQueryUrl(mySubscriptionRequest)
    
    ' Query Rally
    Set queryResponse = prvRallyConnection.rallyGET(queryUrl)

    ' Get Subscription Object
    Set subscriptionResult = queryResponse("Subscription")
    
    ' Get Workspaces Collection Object
    Set workspaceCollection = subscriptionResult("Workspaces")
    
    ' Get Workspaces Ref
    workspacesRef = workspaceCollection("_ref")
    
    ' Query URL
    workspacesQueryUrl = workspacesRef & "?fetch=true"
    
    ' Request Workspaces Collection
    Set queryResponse = prvRallyConnection.rallyGET(workspacesQueryUrl)
    
    ' Populate Result object
    Set myQueryResult = New RallyQueryResult
    Set myQueryResult.RawResults = queryResponse
    
    ' Find out how many pages are in result set
    pageSize = 20
    Set resultCollection = queryResponse("QueryResult")
    myTotalResultCount = resultCollection("TotalResultCount")
    nPages = CalculateNumberOfPages(myTotalResultCount, pageSize)
    
    ' If we have multiple pages, page through and collect entire result set
    If nPages > 1 Then
        For i = 2 To nPages
            ' Set Start to next page
            start = pageSize * (i - 1) + 1
            
            ' Build Query URL
            workspacesQueryUrl = workspacesRef & "?fetch=true&start=" & start
            
            ' Query Rally
            Set queryResponse = prvRallyConnection.rallyGET(workspacesQueryUrl)
    
            ' Update Result object
            Call myQueryResult.AddRawResults(queryResponse)
        Next i
        Set queryResponse = Nothing
    End If
    
    Set myWorkspacesResults = myQueryResult.Results
     
    For Each workspaceResult In myWorkspacesResults
        If workspaceResult("Name") = WorkspaceName Then
            Set foundWorkspace = workspaceResult
            Exit For
        End If
    Next
    
    Set findWorkspace = foundWorkspace

End Function

Public Function findProject(RallyWorkspace As Object, ProjectName As String) As Object

    Dim queryResponse As Object
    Dim projectsCollection As Object, projectsRef As String, projectsQueryUrl As String
    Dim queryUrl As String
    Dim myProjectsResults As Object, foundProject As Object

    Dim myQueryResult As RallyQueryResult
    Dim myTotalResultCount As Long, nPages As Long, i As Long, pageSize As Long, start As Long
    
    ' Get Projects Collection Object
    Set projectsCollection = RallyWorkspace("Projects")
    
    ' Get Projects Ref
    projectsRef = projectsCollection("_ref")
    
    ' Query URL
    projectsQueryUrl = projectsRef & "?fetch=true"
    
    ' Request Projects Collection
    Set queryResponse = prvRallyConnection.rallyGET(projectsQueryUrl)
    
    ' Populate Result object
    Set myQueryResult = New RallyQueryResult
    Set myQueryResult.RawResults = queryResponse
    
    ' Find out how many pages are in result set
    pageSize = 20
    Set resultCollection = queryResponse("QueryResult")
    myTotalResultCount = resultCollection("TotalResultCount")
    nPages = CalculateNumberOfPages(myTotalResultCount, pageSize)
    
    ' If we have multiple pages, page through and collect entire result set
    If nPages > 1 Then
        For i = 2 To nPages
            ' Set Start to next page
            start = pageSize * (i - 1) + 1
            
            ' Build Query URL
            projectsQueryUrl = projectsRef & "?fetch=true&start=" & start
            
            ' Query Rally
            Set queryResponse = prvRallyConnection.rallyGET(projectsQueryUrl)
    
            ' Update Result object
            Call myQueryResult.AddRawResults(queryResponse)
        Next i
        Set queryResponse = Nothing
    End If
    
    Set myProjectsResults = myQueryResult.Results
     
    For Each projectResult In myProjectsResults
        If projectResult("Name") = ProjectName Then
            Set foundProject = projectResult
            Exit For
        End If
    Next
    
    Set findProject = foundProject

End Function

Public Function Query(Value As RallyRequest) As RallyQueryResult

    Dim queryResponse As Object
    Dim resultCollection As Object
    Dim queryUrl As String
    Dim myRallyRequest As RallyRequest
    Dim myQueryResult As RallyQueryResult
    Dim myTotalResultCount As Long, nPages As Long, i As Long, remainder As Long, lastPage As Long
    
    ' Set private reference to RallyRequest
    Set myRallyRequest = Value
    
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
    nPages = CalculateNumberOfPages(myTotalResultCount, myRallyRequest.pageSize)
    
    ' If we have multiple pages, page through and collect entire result set
    If nPages > 1 Then
        For i = 2 To nPages
            ' Set Start to next page
            myRallyRequest.start = myRallyRequest.pageSize * (i - 1) + 1
            
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

Private Function buildQueryUrl(Value As RallyRequest) As String

    Dim queryString As String
    Dim myRallyRequest As RallyRequest
    Dim myRallyQuery As RallyQuery
    Dim myRallyQueryUrl As String
    Dim queryParams() As String, paramsString As String
    Dim nParams As Integer, i As Integer
    
    Set myRallyRequest = Value
    
    myRallyQueryUrl = webServiceUrl() & "/" & myRallyRequest.ArtifactName
        
    nParams = 0
    If myRallyRequest.HasQuery() Then
        If Not IsNull(myRallyRequest.Query.queryString) And myRallyRequest.Query.queryString <> "" Then
            thisParam = "query=" & myRallyRequest.Query.queryString
            ReDim Preserve queryParams(nParams) As String
            queryParams(nParams) = thisParam
            nParams = nParams + 1
        End If
    End If
    
    If Not IsNull(myRallyRequest.Fetch) And myRallyRequest.Fetch <> "" Then
        thisParam = "fetch=" & myRallyRequest.Fetch
        ReDim Preserve queryParams(nParams) As String
        queryParams(nParams) = thisParam
        nParams = nParams + 1
    End If
    
    If Not IsNull(myRallyRequest.Workspace) And myRallyRequest.Workspace <> "" Then
        thisParam = "workspace=" & myRallyRequest.Workspace
        ReDim Preserve queryParams(nParams) As String
        queryParams(nParams) = thisParam
        nParams = nParams + 1
    End If
    
    If Not IsNull(myRallyRequest.Project) And myRallyRequest.Project <> "" Then
        thisParam = "project=" & myRallyRequest.Project
        ReDim Preserve queryParams(nParams) As String
        queryParams(nParams) = thisParam
        nParams = nParams + 1
    End If
    
    If Not IsNull(myRallyRequest.pageSize) Then
        thisParam = "pagesize=" & CStr(myRallyRequest.pageSize)
        ReDim Preserve queryParams(nParams) As String
        queryParams(nParams) = thisParam
        nParams = nParams + 1
    End If
    
    If Not IsNull(myRallyRequest.start) Then
        thisParam = "start=" & CStr(myRallyRequest.start)
        ReDim Preserve queryParams(nParams) As String
        queryParams(nParams) = thisParam
        nParams = nParams + 1
    End If
    
    i = 0
    For Each param In queryParams
        If i = 0 Then
            paramsString = queryParams(i)
        Else
            paramsString = paramsString & "&" & queryParams(i)
        End If
        i = i + 1
    Next
    
    If nParams > 0 Then
        myRallyQueryUrl = myRallyQueryUrl & "?" & paramsString
    End If
        
    urlEncodedQueryUrl = URLEncode(myRallyQueryUrl)
    buildQueryUrl = myRallyQueryUrl
End Function

Public Function Create(ArtifactType As String, workspaceRef As String, Value As RallyObject) As RallyCreateResult
    Dim myRallyCreateUrl As String
    Dim myRallyObject As Object
    Dim myCreateObject As Object
    Dim myCreateJSON As String, myURLEncodedJSON As String, _
        authFailure As String, myErrorString As String
    Dim myCreateResponse As Object
    Dim myCreateResult As RallyCreateResult
    Dim myCreateErrors As Object
    
    Set myRallyObject = Value.JSONObject
    
    Set myCreateObject = CreateObject("Scripting.Dictionary")
    myCreateObject.Add key:=ArtifactType, Item:=myRallyObject
    
    myCreateJSON = JSON.toString(myCreateObject)
    myURLEncodedJSON = URLEncode(myCreateJSON)
    
    myRallyCreateUrl = buildCreateUrl(ArtifactType)
    
    Set myCreateResponse = prvRallyConnection.rallyPOST(myRallyCreateUrl, workspaceRef, myCreateJSON)
        
    ' Populate Result object
    Set myCreateResult = New RallyCreateResult
    Set myCreateResult.RawResult = myCreateResponse
    
    ' Check to see if we failed to authenticate - if so, re-authenticate and retry create
    authFailure = "Not authorized to perform action: Invalid key"
    
    If myCreateResult.WasSuccessful <> True Then
        Set myCreateErrors = myCreateResult.Errors
        myErrorString = ""
        For Each thisError In myCreateErrors
            myErrorString = myErrorString & thisError
        Next
        If myErrorString = authFailure Then
        
            Set myCreateResponse = Nothing
            Set myCreateResult = Nothing
            
            prvRallyConnection.Authenticate
            
            Set myCreateResponse = prvRallyConnection.rallyPOST(myRallyCreateUrl, workspaceRef, myCreateJSON)
        
            ' Populate Result object
            Set myCreateResult = New RallyCreateResult
            Set myCreateResult.RawResult = myCreateResponse
        End If
    End If
    
    Set Create = myCreateResult
    
    Set myRallyObject = Nothing
    Set myCreateObject = Nothing
    Set myCreateResponse = Nothing
    Set myCreateResult = Nothing


End Function

Private Function buildCreateUrl(ArtifactType As String) As String
    Dim myRallyCreateUrl As String
    
    myRallyCreateUrl = webServiceUrl() & "/" & ArtifactType & "/create"
    buildCreateUrl = myRallyCreateUrl
End Function

Private Function buildUpdateUrl(ArtifactType As String, ArtifactObjectID As String) As String
    Dim myRallyUpdateUrl As String
    
    myRallyUpdateUrl = webServiceUrl() & "/" & ArtifactType & "/" & ArtifactObjectID
    buildUpdateUrl = myRallyUpdateUrl
End Function

Public Function Update(ArtifactType As String, ArtifactObjectID As String, Value As RallyObject) As RallyOperationResult
    Dim myRallyUpdateUrl As String
    Dim myRallyObject, myUpdateObject As Object
    Dim myUpdateJSON As String, myURLEncodedJSON As String, _
        authFailure As String, myErrorString As String
    Dim myUpdateResponse As Object
    Dim myUpdateResult As RallyOperationResult
    Dim myUpdateErrors As Object
        
    Set myRallyObject = Value.JSONObject
    
    ' i.e., we receive fields as key value pairs only. These are wrapped up
    ' in the artifact type JSON: {"task":{"Name":"My Task"}}
    
    Set myUpdateObject = CreateObject("Scripting.Dictionary")
    myUpdateObject.Add key:=ArtifactType, Item:=myRallyObject
    
    myUpdateJSON = JSON.toString(myUpdateObject)
    myURLEncodedJSON = URLEncode(myUpdateJSON)
    
    ' UPDATE URL FUNCTION
    myRallyUpdateUrl = buildUpdateUrl(ArtifactType, ArtifactObjectID)
    
    Set myUpdateResponse = prvRallyConnection.rallyUpdatePOST(myRallyUpdateUrl, myUpdateJSON)
        
    ' Populate Result object
    Set myUpdateResult = New RallyOperationResult
    Set myUpdateResult.RawResult = myUpdateResponse
    
    ' Check to see if we failed to authenticate - if so, re-authenticate and retry create
    authFailure = "Not authorized to perform action: Invalid key"
    
    If myUpdateResult.WasSuccessful <> True Then
        Set myUpdateErrors = myUpdateResult.Errors
        myErrorString = ""
        For Each thisError In myUpdateErrors
            myErrorString = myErrorString & thisError
        Next
        If myErrorString = authFailure Then
        
            Set myUpdateResponse = Nothing
            Set myUpdateResult = Nothing
            
            prvRallyConnection.Authenticate
            
            Set myUpdateResponse = prvRallyConnection.rallyUpdatePOST(myRallyUpdateUrl, myUpdateJSON)
        
            ' Populate Result object
            Set myUpdateResult = New RallyOperationResult
            Set myUpdateResult.RawResult = myUpdateResponse
        End If
    End If
    
    Set Update = myUpdateResult
    
    Set myRallyObject = Nothing
    Set myUpdateObject = Nothing
    Set myUpdateResponse = Nothing
    Set myUpdateResult = Nothing

End Function

Private Function GetTypeFromRef(Ref As String)

End Function

Private Function CalculateNumberOfPages(resultCount As Long, pageSize As Long)
    Dim numberPages As Long, remainder As Long, lastPage As Long
    lastPage = 0
    
    ' Calculate remainder
    remainder = resultCount Mod pageSize
    
    ' Add a page if we have a remainder
    If remainder <> 0 Then
        lastPage = 1
    End If
    
    numberPages = (resultCount / pageSize) + lastPage
    CalculateNumberOfPages = numberPages

End Function
