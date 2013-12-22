VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RallyCreateResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Attributes
Dim prvErrors As Object
Dim prvWarnings As Object
Dim prvCreatedObject As Object
Dim prvRawResult As Object
Dim prvCreateResult As Object
Dim prvCreatedItem As Object
Dim prvRef As String, prvObjectID As String
Dim prvWasSuccessful As Boolean

' Class Constructor
Private Sub Class_Initialize()

    prvWasSuccessful = False

End Sub

' Setters and Getters
' In VBA this is done with Property Get and Property Let functions (or Property Set for object type variables).

''''''''''''''''''''''
' Errors property
''''''''''''''''''''''
Public Property Get Errors() As Object
    Set Errors = prvErrors
End Property

Public Property Let Errors(Value As Object)
    Set prvErrors = Value
End Property

''''''''''''''''''''''
' Warnings property
''''''''''''''''''''''
Public Property Get Warnings() As Object
    Set Warnings = prvWarnings
End Property

Public Property Let Warnings(Value As Object)
    Set prvWarnings = Value
End Property

''''''''''''''''''''''
' Results collection property
''''''''''''''''''''''
Public Property Set RawResult(ByRef Value As Object)
    Dim CreateResultCollection As Object
    Dim Errors As Object
    Dim Warnings As Object
    
    Set prvRawResult = Value
    Set CreateResultCollection = prvRawResult("CreateResult")
    Set Errors = CreateResultCollection("Errors")
    Set Warnings = CreateResultCollection("Warnings")
    If Errors.Count = 0 Then
        Set prvCreatedItem = CreateResultCollection("Object")
        prvRef = prvCreatedItem("_ref")
        prvObjectID = prvCreatedItem("ObjectID")
        prvWasSuccessful = True
    End If
        
    Set prvCreateResult = CreateResultCollection
    
    Set CreateResultCollection = Nothing
    Set Errors = Nothing
    Set Warnings = Nothing
    
End Property
Public Property Get CreatedItem() As Object
    Set CreatedItem = prvCreatedItem
End Property
Public Property Get WasSuccessful() As Boolean
    WasSuccessful = prvWasSuccessful
End Property
Public Property Get Ref() As String
    Ref = prvRef
End Property
Public Property Get ObjectID() As String
    ObjectID = prvObjectID
End Property


