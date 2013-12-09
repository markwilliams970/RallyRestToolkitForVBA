VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RallyRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Attributes

Private prvArtifactName As String
Private prvFetch As String
Private prvProjectScopeUp As Boolean, prvProjectScopeDown As Boolean
Private prvPageSize As Integer
Private prvRecordLimit As Long
Private prvStart As Long
Private prvOrder As String
Private prvQuery As RallyQuery
Private prvWorkspace As String
Private prvProject As String

' Class Constructor
Private Sub Class_Initialize()
    ' Default project scoping to flat
    prvProjectScopeUp = False
    prvProjectScopeDown = False
    prvPageSize = 200
    prvRecordLimit = 99999
    prvStart = 1
End Sub

' Setters and Getters
' In VBA this is done with Property Get and Property Let functions (or Property Set for object type variables).

''''''''''''''''''''''
' ArtifactName property
''''''''''''''''''''''
Public Property Get ArtifactName() As String
    ArtifactName = prvArtifactName
End Property

Public Property Let ArtifactName(value As String)
    prvArtifactName = value
End Property
''''''''''''''''''''''
' Query property
''''''''''''''''''''''
Public Property Get Query() As RallyQuery
    Set Query = prvQuery
End Property

Public Property Set Query(ByRef value As RallyQuery)
    Set prvQuery = value
End Property

''''''''''''''''''''''
' Fetch property
''''''''''''''''''''''
Public Property Get Fetch() As String
    Fetch = prvFetch
End Property

Public Property Let Fetch(value As String)
    prvFetch = value
End Property

''''''''''''''''''''''
' ProjectScopeUp property
''''''''''''''''''''''
Public Property Get ProjectScopeUp() As Boolean
    ProjectScopeUp = prvProjectScopeUp
End Property

Public Property Let ProjectScopeUp(value As Boolean)
    prvProjectScopeUp = value
End Property
''''''''''''''''''''''
' ProjectScopeDown property
''''''''''''''''''''''
Public Property Get ProjectScopeDown() As Boolean
    ProjectScopeDown = prvProjectScopeDown
End Property

Public Property Let ProjectScopeDown(value As Boolean)
    prvProjectScopeDown = value
End Property
''''''''''''''''''''''
' PageSize property
''''''''''''''''''''''
Public Property Get pageSize() As Integer
    pageSize = prvPageSize
End Property

Public Property Let pageSize(value As Integer)
    prvPageSize = value
End Property
''''''''''''''''''''''
' Project property
''''''''''''''''''''''
Public Property Get Project() As String
    Project = prvProject
End Property

Public Property Let Project(value As String)
    prvProject = value
End Property

''''''''''''''''''''''
' RecordLimit property
''''''''''''''''''''''
Public Property Get RecordLimit() As Integer
    RecordLimit = prvRecordLimit
End Property

Public Property Let RecordLimit(value As Integer)
    prvRecordLimit = value
End Property
''''''''''''''''''''''
' Order property
''''''''''''''''''''''
Public Property Get Order() As String
    Order = prvOrder
End Property

Public Property Let Order(value As String)
    prvOrder = value
End Property

''''''''''''''''''''''
' Workspace property
''''''''''''''''''''''
Public Property Get Workspace() As String
    Workspace = prvWorkspace
End Property

Public Property Let Workspace(value As String)
    prvWorkspace = value
End Property

''''''''''''''''''''''
' Start property
''''''''''''''''''''''
Public Property Get Start() As Long
    Start = prvStart
End Property

Public Property Let Start(value As Long)
    prvStart = value
End Property
