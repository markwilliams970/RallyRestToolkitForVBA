VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RallyConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Attributes

Private prvRallyUrl  As String
Private prvUserID    As String
Private prvPassword  As String
Private prvWsapiVersion   As String
Private prvSecurityToken As String
Private prvSessionCookie As String
Private prvQuery As String
Private JSON As JSONLib

' Class Constructor
Private Sub Class_Initialize()
    prvRallyUrl = "https://rally1.rallydev.com/slm"
    prvWsapiVersion = "v2.0"
    Set JSON = New JSONLib
End Sub

' Setters and Getters
' In VBA this is done with Property Get and Property Let functions (or Property Set for object type variables).

''''''''''''''''''''''
' RallyUrl property
''''''''''''''''''''''
Public Property Get RallyUrl() As String
    RallyUrl = prvRallyUrl
End Property
Public Property Let RallyUrl(value As String)
    prvRallyUrl = value
End Property

''''''''''''''''''''''
' prvUserID property
''''''''''''''''''''''
Public Property Get UserID() As String
    UserID = prvUserID
End Property
Public Property Let UserID(value As String)
    prvUserID = value
End Property

''''''''''''''''''''''
' prvPassword property
''''''''''''''''''''''
Public Property Get Password() As String
    RallyPassword = prvPassword
End Property
Public Property Let Password(value As String)
    prvPassword = value
End Property

''''''''''''''''''''''
' prvWsapiVersion property
''''''''''''''''''''''
Public Property Get WsapiVersion() As String
    WsapiVersion = prvWsapiVersion
End Property
Public Property Let WsapiVersion(value As String)
    prvWsapiVersion = value
End Property
''''''''''''''''''''''
' prvRallyAuthKey property
''''''''''''''''''''''
Public Property Get SecurityToken() As String
    SecurityToken = prvSecurityToken
End Property
Public Property Let SecurityToken(value As String)
    prvSecurityToken = value
End Property

Private Function getRallyAuthenticationURL()
    getRallyAuthenticationURL = "https://rally1.rallydev.com/slm/webservice/" & _
        prvWsapiVersion & "/security/authorize"
End Function
''''''''''''''''''''''
' prvRallySessionCookie property
''''''''''''''''''''''
Public Property Get SessionCookie() As String
    SessionCookie = prvSessionCookie
End Property
Public Property Let SessionCookie(value As String)
    prvSessionCookie = value
End Property

Public Function Authenticate() As Boolean
    Dim rallyAuthUrl As String
    Dim responseBody As String
    Dim responseJSON As Object, operationResult As Object, operationErrors As Object
    Dim numberErrors As Integer
    Dim SecurityToken As String
    Dim SessionCookie As String
    Dim isAuthenticated As Boolean
    
    ' Initialize isAuthenticated to false
    isAuthenticated = False
    
    ' Get Rally Authorization URL endpoint
    rallyAuthUrl = getRallyAuthenticationURL()
    
    ' Issue a GET request against the URL using specified credentials
    responseBody = httpGET(rallyAuthUrl, prvUserID, prvPassword, "application/json", True)
    
    ' Parse the resulting response
    Set responseJSON = JSON.parse(responseBody)
    Set operationResult = responseJSON("OperationResult")
    Set operationErrors = operationResult("Errors")
    
    numberErrors = operationErrors.Count
    SecurityToken = operationResult("SecurityToken")
    
    ' Save the security token
    prvSecurityToken = SecurityToken
    
    ' Set the SubBucketID
    Call setSubBucketID
        
    Set JSON = Nothing
    Set responseJSON = Nothing
    Set operationResult = Nothing
    Set operationErrors = Nothing
    If numberErrors = 0 Then
        isAuthenticated = True
    End If
    Authenticate = isAuthenticated
    
End Function
' http://ramblings.mcpher.com/Home/excelquirks/json/rest
Public Function httpPost(fn As String, _
        Optional data As String = vbNullString) As String
    pHtml = fn
    Dim oHttp As Object
    Set oHttp = CreateObject("MSXML2.ServerXMLHTTP")
    With oHttp
        .Open "PUT", pHtml, False
        .SetRequestHeader "Content-Type", "application/json"
        .Send data
        httpPost = .responseText
        pStatus = .Status
    End With
    Set oHttp = Nothing
End Function

' http://ramblings.mcpher.com/Home/excelquirks/json/rest
Public Function httpGET(fn As String, _
        Optional authUser As String = vbNullString, _
        Optional authPass As String = vbNullString, _
        Optional accept As String = vbNullString, _
        Optional isAuthRequest As Boolean) As String
    pHtml = fn
    Dim responseCookie As String
    Dim oHttp As Object
    Set oHttp = CreateObject("MSXML2.ServerXMLHTTP")
    Call oHttp.Open("GET", pHtml, False)
    If (authUser <> vbNullString) Then
    ' need to do basic authentication
    ' acknowledgement to http://pastie.org/1192157
        oHttp.SetRequestHeader "Content-Type", "application/json"
        oHttp.SetRequestHeader "Accept", "application/json"
        oHttp.SetRequestHeader "Authorization", "Basic " + _
            Base64Encode(authUser + ":" + authPass)
    End If
    
    ' some times we need to set the accept header
    If accept <> vbNullString Then
        oHttp.SetRequestHeader "Accept", accept
    End If
    Call oHttp.Send("")
    httpGET = oHttp.responseText
    pStatus = oHttp.Status
    If isAuthRequest = True Then
        responseCookie = oHttp.getResponseHeader("set-cookie")
        ' Save the session cookie
        prvSessionCookie = responseCookie
    End If
    Set oHttp = Nothing
End Function

Public Function rallyGET(fn As String) As Object
    Dim myJsonLib As JSONLib
    Dim getResponse As String
    Dim getJSON As Object
    
    Set myJsonLib = New JSONLib
   
    getResponse = httpGET(fn, prvUserID, prvPassword)
    Set getJSON = myJsonLib.parse(getResponse)
    Set rallyGET = getJSON
    Set myJsonLib = Nothing
    Set getJSON = Nothing
End Function

Private Sub setSubBucketID()
    Dim mySubscriptionQueryURL As String
    Dim mySubscriptionResponse As Object
    Dim mySubscriptionResponseString As String
    Dim QueryResultCollection As Object
    Dim Results As Object
    Dim mySubscription As Object
    Dim myRallyHost As String
    Dim SubID As String
    Dim SubBucketIDCookie As String
    
    mySubscriptionQueryURL = _
        prvRallyUrl & "/webservice/" & _
        prvWsapiVersion & _
        "/subscription?query=&start=1&pagesize=20&fetch=Name,SubscriptionID"
    
    Set mySubscriptionResponse = rallyGET(mySubscriptionQueryURL)
    Set QueryResultCollection = mySubscriptionResponse("QueryResult")
    TotalResultCount = QueryResultCollection("TotalResultCount")
    Set Results = QueryResultCollection("Results")
    Set mySubscription = Results(1)
    
    SubID = mySubscription("SubscriptionID")
    
    myRallyHost = Split(prvRallyUrl, "/")(2)
    
    SubBucketIDCookie = ";SUBBUCKETID=" & SubID & ";Path=/;Domain=" & myRallyHost & ";Secure;HttpOnly"
    prvSessionCookie = prvSessionCookie & SubBucketIDCookie
    
End Sub

Public Function TestQuery() As Object

    Dim myQueryUrl As String
    Dim myResponse As Object
    myQueryUrl = "https://rally1.rallydev.com/slm/webservice/" & _
        prvWsapiVersion & "/hierarchicalrequirement?query=((FormattedID > " & addEscapedDoubleQuotes("US23") & ") AND (CreationDate > 2013-01-01))&fetch=Name,FormattedID,Description&workspace=&project=&pagesize=200"
    ' Issue a GET request against the URL using specified credentials
    myResponse = rallyGET(myQueryUrl)
    TestQuery = myResponse
     
End Function
