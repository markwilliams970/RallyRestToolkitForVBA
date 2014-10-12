Attribute VB_Name = "Examples"

Public Sub CreateDefects()
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
    Dim myNewDefect As RallyObject
    Dim defectSuffixes(0 To 10) As String
    Dim i As Long

    Dim currentDateTime As Date, currentDateTimeString As String
    
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
    
    ' Authenticate To Rally
    isAuthenticated = myRallyConnection.Authenticate()
    myRallyAuthKey = myRallyConnection.SecurityToken
    myRallySessionCookie = myRallyConnection.SessionCookie
      
    ' Get Current Time
    currentDateTime = Now()
    currentDateTimeString = Format(currentDateTime, "yyyy-MM-ddThh:mm")

    ' Initials array of suffixes
    defectSuffixes(0) = "A"
    defectSuffixes(1) = "B"
    defectSuffixes(2) = "C"
        
    For i = 0 To 2
        Set myNewDefect = New RallyObject
        Call myNewDefect.AddProperty("Name", "My Defect from VBA: " & _
            currentDateTimeString & defectSuffixes(i))
        Call myNewDefect.AddProperty("Severity", "Major Problem")
        Call myNewDefect.AddProperty("Priority", "Resolve Immediately")
            
        Set myCreateResult = myRallyRestApi.Create("defect", myWorkspaceRef, myNewDefect)
        If myCreateResult.WasSuccessful Then
            MsgBox "Create Succeeded"
            Set myCreateResultObject = myCreateResult.CreatedItem
            myCreatedRef = myCreateResult.Ref
            myCreatedObjectID = myCreateResult.ObjectID
        End If
        
        MsgBox "Created: " & myCreatedRef
        Set myNewDefect = Nothing
        Set myCreateResult = Nothing

    Next i
    
End Sub

