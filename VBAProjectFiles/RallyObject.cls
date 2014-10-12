VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RallyObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Attributes
Private prvJSONObject As Object

' Class Constructor
Private Sub Class_Initialize()
    Set prvJSONObject = CreateObject("Scripting.Dictionary")
End Sub

''''''''''''''''''''''
' JSONObject property
''''''''''''''''''''''
Public Property Get JSONObject() As Object
    Set JSONObject = prvJSONObject
End Property

Public Sub AddProperty(key As String, Value As String)
    ' add key/value pair
    prvJSONObject.Add key:=key, Item:=Value
End Sub

Public Sub AddObjectProperty(key As String, Value As Object)
    ' add key/value pair
    prvJSONObject.Add key:=key, Item:=Value
End Sub
