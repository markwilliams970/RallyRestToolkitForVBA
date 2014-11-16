VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteDefectsForm 
   Caption         =   "Rally Authentication"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   OleObjectBlob   =   "DeleteDefectsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteDefectsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RallyUserID As String, RallyPassword As String


Private Sub DeleteDefectsFormCancelButton_Click()

    DeleteDefectsForm.Hide

End Sub

Private Sub DeleteDefectsFormOKButton_Click()

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
            Call DeleteDefects(RallyUserName, RallyPassword, RallyWorkspace)
        End If
    End If

    Me.Hide

End Sub
Public Sub DeleteDefects(RallyUserName As String, RallyPassword As String, RallyWorkspace As String)

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
    Dim myDeleteResult As RallyOperationResult
    Dim myResults As Object, firstResult As Object, myDeleteErrors As Object, _
        Error As Variant, myErrorString As String
    Dim myResultString As String
    Dim i As Long, currentRow As Long
    Dim startRow As Long, endRow As Long
    Dim FormattedIDRange As String, formattedIDValue As String, objectIDValue As String
    Dim nameValue As String, confirmMsg As String, confirmDelete As VbMsgBoxResult
    
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

        ' Get FormattedID
        formattedIDValue = Worksheets("DeleteDefects").Range(FormattedIDRange).Value
        
        ' Formulate a Query
        Set myRallyQuery = New RallyQuery
        myFormattedID = addEscapedDoubleQuotes(formattedIDValue)
        myRallyQuery.queryString = "(FormattedID = " & myFormattedID & ")"
        
        ' Create a RallyRequest
        Set myRallyRequest = New RallyRequest
        myRallyRequest.ArtifactName = "defect"
        myRallyRequest.Fetch = "FormattedID,ObjectID,Name"
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
            nameValue = firstResult("Name")
            
            confirmMsg = "Really Delete " & myFormattedID & ": " & nameValue & "?"
            confirmDelete = MsgBox(confirmMsg, vbYesNo, "Really Delete?")
            
            If confirmDelete = vbYes Then
                
                Set myDeleteResult = myRallyRestApi.Delete("defect", objectIDValue)
                If myDeleteResult.WasSuccessful Then
                    MsgBox "Deleted: " & formattedIDValue
                Else
                    Set myDeleteErrors = myDeleteResult.Errors
                    myErrorString = ""
                    For Each Error In myDeleteErrors
                        myErrorString = myErrorString & ", " & Error
                    Next
                    MsgBox "Problem Deleting: " & formattedIDValue & myErrorString
                End If
            Else
                MsgBox myFormattedID & " NOT Deleted."
            End If
            Set myDeleteResult = Nothing
        End If
    Next
    
    MsgBox "Finished Deleting Defects In Rally"

End Sub

