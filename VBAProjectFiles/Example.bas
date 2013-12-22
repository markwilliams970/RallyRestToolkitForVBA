Attribute VB_Name = "Test"
Public Sub TestRest()
    Dim myRallyRestApi As RallyRestApi
    Dim myRallyUsername As String, myRallyPassword As String, myWSAPIVersion As String
    Dim myRallyURL As String
    Dim myRallyConnection As RallyConnection
    Dim myRallyQuery As RallyQuery
    Dim myRallyRequest As RallyRequest
    Dim myWorkspaceRef As String
    Dim myResponseString As String
    Dim myFormattedID As String
    Dim myRallyAuthKey As String
    Dim myRallySessionCookie As String
    Dim myQueryResult As RallyQueryResult
    Dim myQueryResultObject As Object, myQueryResultString As String
    Dim myCreateResult As RallyCreateResult
    Dim myCreateResultObject As Object, myCreatedRef As String, myCreatedObjectID As String
    Dim myCreateResultString As String
    Dim myResults As Object
    Dim myResultString As String
    Dim myNewDefect As RallyObject

    Dim currentDateTime As Date, currentDateTimeString As String
    
    Dim blah As String
    
    ' Personal Settings
    myRallyURL = "https://rally1.rallydev.com/slm"
    myRallyUsername = "user@company.com"
    myRallyPassword = "topsecret"
    myWSAPIVersion = "v2.0"
    myWorkspaceRef = "/workspace/12345678910"

    ' Instantiate RallyConnection
    Set myRallyConnection = New RallyConnection
    myRallyConnection.UserID = myRallyUsername
    myRallyConnection.Password = myRallyPassword
    myRallyConnection.WsapiVersion = myWSAPIVersion
    myRallyConnection.RallyUrl = myRallyURL
    
    ' Instantiate RallyRestApi
    Set myRallyRestApi = New RallyRestApi
    myRallyRestApi.RallyConnection = myRallyConnection
    
    ' Formulate a Query
    Set myRallyQuery = New RallyQuery
    myFormattedID = addEscapedDoubleQuotes("US100")
    myRallyQuery.queryString = "(FormattedID > " & myFormattedID & ")"
    myRallyQuery.AddAnd ("(CreationDate > 2012-01-01)")
    
    ' Create a RallyRequest
    Set myRallyRequest = New RallyRequest
    myRallyRequest.ArtifactName = "hierarchicalrequirement"
    myRallyRequest.Fetch = "Name,FormattedID,Description,PlanEstimate"
    myRallyRequest.pageSize = 20
    Set myRallyRequest.Query = myRallyQuery
    myRallyRequest.Order = "FormattedID Asc"
    myRallyRequest.ProjectScopeDown = True
    
    ' Authenticate To Rally
    isAuthenticated = myRallyConnection.Authenticate()
    myRallyAuthKey = myRallyConnection.SecurityToken
    myRallySessionCookie = myRallyConnection.SessionCookie
    
    ' Execute Query
    myRallyRestApi.RallyRequest = myRallyRequest
    Set myQueryResult = myRallyRestApi.Query(myRallyRequest)
    Set myResults = myQueryResult.Results
     
    myResultsString = ""
    For Each result In myResults
        myResultsString = myResultsString & result("FormattedID") & ": " & result("Name") & _
            "; PlanEstimate: " & result("PlanEstimate") & _
            vbCr & vbLf
            
    Next
    
    MsgBox myResultsString
    
    ' Get Current Time
    currentDateTime = Now()
    currentDateTimeString = Format(currentDateTime, "yyyy-MM-ddThh:mm")
    
    Set myNewDefect = New RallyObject
    Call myNewDefect.AddProperty("Name", "My Defect from VBA: " & currentDateTimeString)
    Call myNewDefect.AddProperty("Severity", "Major Problem")
    Call myNewDefect.AddProperty("Priority", "Resolve Immediately")
    
    Set myCreateResult = myRallyRestApi.Create("defect", myWorkspaceRef, myNewDefect)
    If myCreateResult.WasSuccessful Then
        Set myCreateResultObject = myCreateResult.CreatedItem
        myCreatedRef = myCreateResult.Ref
        myCreatedObjectID = myCreateResult.ObjectID
    End If
    
    blah = "blah"

End Sub