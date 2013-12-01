Attribute VB_Name = "ExampleCRUD"
Public Sub ExampleCRUD()
    Dim myRallyRestApi As RallyRestApi
    Dim myRallyUsername As String, myRallyPassword As String, myWSAPIVersion As String
    Dim myRallyURL As String
    Dim myRallyConnection As RallyConnection
    Dim myRallyQuery As RallyQuery
    Dim myRallyRequest As RallyRequest
    Dim myResponseString As String
    Dim myFormattedID As String
    Dim myRallyAuthKey As String
    Dim myRallySessionCookie As String
    Dim myQueryResult As RallyQueryResult
    Dim myQueryResultObject As Object
    Dim myResults As Object
    Dim myResultString As String    
    
    ' Personal Settings
    myRallyURL = "https://rally1.rallydev.com/slm"
    myRallyUsername = "user@company.com"
    myRallyPassword = "topsecret"
    myWSAPIVersion = "v2.0"

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
    myRallyQuery.AddAnd ("(CreationDate > 2013-01-01)")
    
    ' Create a RallyRequest
    Set myRallyRequest = New RallyRequest
    myRallyRequest.ArtifactName = "hierarchicalrequirement"
    myRallyRequest.Fetch = "Name,FormattedID,Description"
    myRallyRequest.PageSize = 200
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
    For Each Result In myResults
        myResultsString = myResultsString & Result("FormattedID") & ": " & Result("Name") & _
            vbCr & vbLf
    Next
    
   MsgBox myResultsString

End Sub
