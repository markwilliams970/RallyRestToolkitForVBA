VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UploadDefectsForm 
   Caption         =   "Rally Authentication"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   OleObjectBlob   =   "UploadDefectsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UploadDefectsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RallyUserID As String, RallyPassword As String


Private Sub UploadDefectsFormCancelButton_Click()

    UploadDefectsForm.Hide

End Sub

Private Sub UploadDefectsFormOKButton_Click()

    Dim RallyUserName As String, RallyPassword As String
    Dim RallyWorkspace As String, RallyProject As String

    RallyUserName = Me.txtBoxUserID.Value
    RallyPassword = Me.txtBoxPassword.Value
    
    RallyWorkspace = Me.txtBoxWorkspaceName.Value
    RallyProject = Me.textBoxProjectName.Value

    If RallyUserName = "" Or RallyPassword = "" Then
        MsgBox "Please enter both Rally Username and Password."
    Else
        If RallyWorkspace = "" Or RallyPassword = "" Then
            MsgBox "Please enter all of these: Workspace Name, Project Name."
        Else
            Call UploadDefects(RallyUserName, RallyPassword, RallyWorkspace, RallyProject)
        End If
    End If

    Me.Hide

End Sub
Public Sub UploadDefects(RallyUserName As String, RallyPassword As String, RallyWorkspace As String, _
    RallyProject As String)

    Dim myRallyRestApi As RallyRestApi
    Dim myRallyUsername As String, myRallyPassword As String, myWSAPIVersion As String
    Dim myRallyURL As String
    Dim myRallyConnection As RallyConnection
    Dim myRallyQuery As RallyQuery
    Dim myRallyRequest As RallyRequest, mySubscriptionRequest As RallyRequest
    Dim myWorkspace As Object
    Dim myWorkspaceRef As String, myProjectRef As String
    Dim myResponseString As String
    Dim myFormattedID As String
    Dim myRallyAuthKey As String
    Dim myRallySessionCookie As String
    Dim myNewDefect As RallyObject
    Dim myCreateResult As RallyCreateResult
    Dim myCreateResultObject As Object, myCreatedRef As String, myCreatedObjectID As String
    Dim myCreateResultString As String
    Dim myResults As Object
    Dim myResultString As String
    Dim i As Long, currentRow As Long
    Dim startRow As Long, endRow As Long
    Dim NameRange As String, SeverityRange As String, _
        PriorityRange As String, StateRange As String
    Dim nameValue As String, severityValue As String, _
        priorityValue As String, stateValue As String
      
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
    
    myWorkspaceRef = myWorkspace("_ref")
    
    ' Lookup Project
    Set myProject = myRallyRestApi.findProject(myWorkspace, RallyProject)
    
    ' Check to see if we found project of interest
    If myProject Is Nothing Then
        MsgBox "Could not find Project named: " & RallyProject
        Exit Sub
    End If
    
    myProjectRef = myProject("_ref")
         
    startRow = 4
    endRow = 6
    For currentRow = startRow To endRow
        
        ' Cell References
        NameRange = "A" & currentRow
        SeverityRange = "B" & currentRow
        PriorityRange = "C" & currentRow
        StateRange = "D" & currentRow
        
        ' Retrieve values from cells
        ' Set values
        Set myNewDefect = New RallyObject
        Call myNewDefect.AddProperty("Name", Worksheets("CreateDefects").Range(NameRange).Value)
        Call myNewDefect.AddProperty("Severity", Worksheets("CreateDefects").Range(SeverityRange).Value)
        Call myNewDefect.AddProperty("Priority", Worksheets("CreateDefects").Range(PriorityRange).Value)
        Call myNewDefect.AddProperty("State", Worksheets("CreateDefects").Range(StateRange))
        Call myNewDefect.AddProperty("Project", myProjectRef)
            
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
    Next
    
    MsgBox "Finished Uploading Defects To Rally"

End Sub


