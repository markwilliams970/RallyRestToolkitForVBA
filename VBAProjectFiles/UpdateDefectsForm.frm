VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UpdateDefectsForm 
   Caption         =   "Rally Authentication"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   OleObjectBlob   =   "UpdateDefectsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UpdateDefectsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RallyUserID As String, RallyPassword As String


Private Sub UpdateDefectsFormCancelButton_Click()

    UpdateDefectsForm.Hide

End Sub

Private Sub UpdateDefectsFormOKButton_Click()

    Dim RallyUserName As String, RallyPassword As String
    Dim RallyWorkspace As String

    RallyUserName = Me.txtBoxUserID.Value
    RallyPassword = Me.txtBoxPassword.Value
    
    RallyWorkspace = Me.txtBoxWorkspaceName.Value

    If RallyUserName = "" Or RallyPassword = "" Then
        MsgBox "Please enter both Rally Username and Password."
    Else
        If RallyWorkspace = "" Then
            MsgBox "Please enter Workspace Name."
        Else
            Call UpdateDefects(RallyUserName, RallyPassword, RallyWorkspace)
        End If
    End If

    Me.Hide

End Sub
Public Sub UpdateDefects(RallyUserName As String, RallyPassword As String, RallyWorkspace As String)

    Dim myRallyRestApi As RallyRestApi
    Dim myRallyUsername As String, myRallyPassword As String, myWSAPIVersion As String
    Dim myRallyURL As String
    Dim myRallyConnection As RallyConnection
    Dim myRallyQuery As RallyQuery
    Dim myRallyRequest As RallyRequest, mySubscriptionRequest As RallyRequest
    Dim myWorkspace As Object
    Dim myWorkspaceRef As String, myProjectRef As String
    Dim myFormattedID As String
    Dim myUpdatedDefect As RallyObject
    Dim myUpdateResult As RallyOperationResult
    Dim myResults As Object, firstResult As Object
    Dim myResultString As String
    Dim i As Long, currentRow As Long
    Dim startRow As Long, endRow As Long
    Dim FormattedIDRange As String, NameRange As String, SeverityRange As String, _
        PriorityRange As String, StateRange As String
    Dim formattedIDValue As String, nameValue As String, severityValue As String, _
        priorityValue As String, stateValue As String, objectIDValue As String
      
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
       
    ' Lookup Workspace
    Set myWorkspace = myRallyRestApi.findWorkspace(RallyWorkspace)
    
    ' Check to see if we found workspace of interest
    If myWorkspace Is Nothing Then
        MsgBox "Could not find Workspace named: " & RallyWorkspace
        Exit Sub
    End If
    
    myWorkspaceRef = myWorkspace("_ref")
         
    startRow = 4
    endRow = 6
    For currentRow = startRow To endRow
        
        ' Cell References
        FormattedIDRange = "A" & currentRow
        NameRange = "B" & currentRow
        SeverityRange = "C" & currentRow
        PriorityRange = "D" & currentRow
        StateRange = "E" & currentRow
        
        ' Get FormattedID
        formattedIDValue = Worksheets("UpdateDefects").Range(FormattedIDRange).Value
        
        ' Formulate a Query
        Set myRallyQuery = New RallyQuery
        myFormattedID = addEscapedDoubleQuotes(formattedIDValue)
        myRallyQuery.queryString = "(FormattedID = " & myFormattedID & ")"
        
        ' Create a RallyRequest
        Set myRallyRequest = New RallyRequest
        myRallyRequest.ArtifactName = "defect"
        myRallyRequest.Fetch = "FormattedID,ObjectID,Name,Severity,Priority,State"
        myRallyRequest.Workspace = myWorkspace("_ref")
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
            MsgBox "No Defects found matching query: " & myRallyQuery.queryString
        Else
            ' Get ObjectID
            Set firstResult = myResults(1)
            objectIDValue = firstResult("ObjectID")
            
            ' Retrieve values from cells and use them
            ' Set values for update
            Set myUpdatedDefect = New RallyObject
            Call myUpdatedDefect.AddProperty("Name", Worksheets("UpdateDefects").Range(NameRange).Value)
            Call myUpdatedDefect.AddProperty("Severity", Worksheets("UpdateDefects").Range(SeverityRange).Value)
            Call myUpdatedDefect.AddProperty("Priority", Worksheets("UpdateDefects").Range(PriorityRange).Value)
            Call myUpdatedDefect.AddProperty("State", Worksheets("UpdateDefects").Range(StateRange))
                
            Set myUpdateResult = myRallyRestApi.Update("defect", objectIDValue, myUpdatedDefect)
            If myUpdateResult.WasSuccessful Then
                MsgBox "Update Succeeded: " & formattedIDValue
            End If
            
            Set myUpdatedDefect = Nothing
            Set myUpdateResult = Nothing
        End If
    Next
    
    MsgBox "Finished Updating Defects In Rally"

End Sub
