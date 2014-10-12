VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GetStoriesForm 
   Caption         =   "Rally Authentication"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   OleObjectBlob   =   "GetStoriesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GetStoriesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RallyUserID As String, RallyPassword As String

Private Sub AuthFormOKButton_Click()

End Sub

Private Sub GetStoriesFormCancelButton_Click()

    GetStoriesForm.Hide

End Sub

Private Sub GetStoriesFormOKButton_Click()

    Dim RallyUserName As String, RallyPassword As String
    Dim RallyWorkspace As String, RallyProject As String
    Dim SampleQuery As String

    RallyUserName = Me.txtBoxUserID.Value
    RallyPassword = Me.txtBoxPassword.Value
    
    RallyWorkspace = Me.txtBoxWorkspaceName.Value
    RallyProject = Me.textBoxProjectName.Value
    SampleQuery = Me.textBoxSampleQuery.Value

    If RallyUserName = "" Or RallyPassword = "" Then
        MsgBox "Please enter both Rally Username and Password."
    Else
        If RallyWorkspace = "" Or RallyPassword = "" Or SampleQuery = "" Then
            MsgBox "Please enter all of these: Workspace Name, Project Name, and a Sample Query."
        Else
            Call QueryStories(RallyUserName, RallyPassword, RallyWorkspace, RallyProject, SampleQuery)
        End If
    End If

    Me.Hide

End Sub
Public Sub QueryStories(RallyUserName As String, RallyPassword As String, RallyWorkspace As String, _
    RallyProject As String, SampleQuery As String)

    Dim myRallyRestApi As RallyRestApi
    Dim myRallyUsername As String, myRallyPassword As String, myWSAPIVersion As String
    Dim myRallyURL As String
    Dim myRallyConnection As RallyConnection
    Dim myRallyQuery As RallyQuery
    Dim myRallyRequest As RallyRequest, mySubscriptionRequest As RallyRequest
    Dim myWorkspace As Object
    Dim myWorkspaceRef As String
    Dim myResponseString As String
    Dim myFormattedID As String
    Dim myRallyAuthKey As String
    Dim myRallySessionCookie As String
    Dim myQueryResult As RallyQueryResult
    Dim myQueryResultObject As Object, myQueryResultString As String
    Dim totalResultCount As Long
    Dim myCreateResult As RallyCreateResult
    Dim myCreateResultObject As Object, myCreatedRef As String, myCreatedObjectID As String
    Dim myCreateResultString As String
    Dim myResults As Object
    Dim myResultString As String
    Dim i As Long, currentRow As Long
    Dim FormattedIDRange As String, NameRange As String, _
        ScheduleStateRange As String, PlanEstimateRange As String
    
    Dim currentDateTime As Date, currentDateTimeString As String
      
    ' Personal Settings
    myRallyURL = "https://rally1.rallydev.com/slm"
    myRallyUsername = RallyUserName
    myRallyPassword = RallyPassword
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
    
    ' Authenticate To Rally
    isAuthenticated = myRallyConnection.Authenticate()
    
    If Not isAuthenticated Then
        MsgBox "Failed to authenticate as: " & myRallyUsername
        Exit Sub
    End If
    
    myRallyAuthKey = myRallyConnection.SecurityToken
    myRallySessionCookie = myRallyConnection.SessionCookie
    
    ' Lookup Workspace
    Set myWorkspace = myRallyRestApi.findWorkspace(RallyWorkspace)
    
    ' Check to see if we found workspace of interest
    If myWorkspace Is Nothing Then
        MsgBox "Could not find Workspace named: " & RallyWorkspace
        Exit Sub
    End If
    
    ' Lookup Project
    Set myProject = myRallyRestApi.findProject(myWorkspace, RallyProject)
    
    ' Check to see if we found project of interest
    If myProject Is Nothing Then
        MsgBox "Could not find Project named: " & RallyProject
        Exit Sub
    End If
    
    ' Formulate a Query
    Set myRallyQuery = New RallyQuery
    
    ' Additional Sample Query Syntax in commented section below
    ' For now, we'll take query passed in via dialog box
    myRallyQuery.queryString = SampleQuery
    
    ' myFormattedID = addEscapedDoubleQuotes("US100")
    ' myRallyQuery.queryString = "(FormattedID < " & myFormattedID & ")"
    ' myRallyQuery.AddAnd ("(CreationDate > 2012-01-01)")
    
    ' Create a RallyRequest
    Set myRallyRequest = New RallyRequest
    myRallyRequest.ArtifactName = "hierarchicalrequirement"
    myRallyRequest.Fetch = "Name,FormattedID,ScheduleState,PlanEstimate"
    myRallyRequest.Workspace = myWorkspace("_ref")
    myRallyRequest.Project = myProject("_ref")
    myRallyRequest.pageSize = 20
    Set myRallyRequest.Query = myRallyQuery
    myRallyRequest.Order = "FormattedID Asc"
    myRallyRequest.ProjectScopeDown = True
    
    ' Execute Query
    myRallyRestApi.RallyRequest = myRallyRequest
    Set myQueryResult = myRallyRestApi.Query(myRallyRequest)
    Set myResults = myQueryResult.Results
    
    totalResultCount = myQueryResult.totalResultCount
    
    If totalResultCount = 0 Then
        MsgBox "No Stories found matching query: " & SampleQuery
        Exit Sub
    End If
     
    currentRow = 5
    For Each result In myResults
        
        ' Cell References
        FormattedIDRange = "A" & currentRow
        NameRange = "B" & currentRow
        ScheduleStateRange = "C" & currentRow
        PlanEstimateRange = "D" & currentRow
        
        ' Set values
        Worksheets("Sheet1").Range(FormattedIDRange).Value = result("FormattedID")
        Worksheets("Sheet1").Range(NameRange).Value = result("Name")
        Worksheets("Sheet1").Range(ScheduleStateRange).Value = result("ScheduleState")
        Worksheets("Sheet1").Range(PlanEstimateRange) = result("PlanEstimate")
        currentRow = currentRow + 1
    Next
    
    MsgBox "Finished Querying Rally"

End Sub
