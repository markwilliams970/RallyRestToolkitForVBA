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
Private prvHasQuery As Boolean

' Class Constructor
Private Sub Class_Initialize()
    ' Default project scoping to flat
    prvProjectScopeUp = False
    prvProjectScopeDown = False
    prvPageSize = 200
    prvRecordLimit = 99999
    prvStart = 1
    prvHasQuery = False
End Sub

' Setters and Getters
' In VBA this is done with Property Get and Property Let functions (or Property Set for object type variables).

''''''''''''''''''''''
' ArtifactName property
''''''''''''''''''''''
Public Property Get ArtifactName() As String
    ArtifactName = prvArtifactName
End Property

Public Property Let ArtifactName(Value As String)
    prvArtifactName = Value
End Property
''''''''''''''''''''''
' Query property
''''''''''''''''''''''
Public Property Get Query() As RallyQuery
    Set Query = prvQuery
End Property

Public Property Set Query(ByRef Value As RallyQuery)
    Set prvQuery = Value
    prvHasQuery = True
End Property

''''''''''''''''''''''
' Fetch property
''''''''''''''''''''''
Public Property Get Fetch() As String
    Fetch = prvFetch
End Property

Public Property Let Fetch(Value As String)
    prvFetch = Value
End Property

''''''''''''''''''''''
' ProjectScopeUp property
''''''''''''''''''''''
Public Property Get ProjectScopeUp() As Boolean
    ProjectScopeUp = prvProjectScopeUp
End Property

Public Property Let ProjectScopeUp(Value As Boolean)
    prvProjectScopeUp = Value
End Property
''''''''''''''''''''''
' ProjectScopeDown property
''''''''''''''''''''''
Public Property Get ProjectScopeDown() As Boolean
    ProjectScopeDown = prvProjectScopeDown
End Property

Public Property Let ProjectScopeDown(Value As Boolean)
    prvProjectScopeDown = Value
End Property
''''''''''''''''''''''
' PageSize property
''''''''''''''''''''''
Public Property Get pageSize() As Integer
    pageSize = prvPageSize
End Property

Public Property Let pageSize(Value As Integer)
    prvPageSize = Value
End Property
''''''''''''''''''''''
' Project property
' This should be a ref of the form "/project/12345678910"
''''''''''''''''''''''
Public Property Get Project() As String
    Project = prvProject
End Property

Public Property Let Project(Value As String)
    prvProject = Value
End Property

''''''''''''''''''''''
' RecordLimit property
''''''''''''''''''''''
Public Property Get RecordLimit() As Integer
    RecordLimit = prvRecordLimit
End Property

Public Property Let RecordLimit(Value As Integer)
    prvRecordLimit = Value
End Property
''''''''''''''''''''''
' Order property
''''''''''''''''''''''
Public Property Get Order() As String
    Order = prvOrder
End Property

Public Property Let Order(Value As String)
    prvOrder = Value
End Property

''''''''''''''''''''''
' Workspace property
' This should be a ref of the form "/workspace/12345678911"
''''''''''''''''''''''
Public Property Get Workspace() As String
    Workspace = prvWorkspace
End Property

Public Property Let Workspace(Value As String)
    prvWorkspace = Value
End Property

''''''''''''''''''''''
' Start property
''''''''''''''''''''''
Public Property Get start() As Long
    start = prvStart
End Property

Public Property Let start(Value As Long)
    prvStart = Value
End Property

''''''''''''''''''''''''''''
' HasQuery function
''''''''''''''''''''''
Public Function HasQuery()
    HasQuery = prvHasQuery
End Function
