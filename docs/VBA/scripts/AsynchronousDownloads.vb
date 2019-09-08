'Usage:
'Set References (menu extra/references) to “Microsoft ActiveData Objects” and “Microsoft WinHTTP Services”
'Add two classes to your VBA-project, name them as indicated above and fill in the code.
'In your module/class/form paste at the top:
	'Private WithEvents pDownloader As HTTPDownloader
	'And make sure you create/destroy the downloader properly (for forms this is best done in the form_load/form_unload events):
	'Set pDownloader = New HTTPDownloader
	'Set pDownloader = Nothing

'Now you can download files asynchronuosly with
'pDownloader.DownloadAsync "http://whatever.com/file.zip","c:/on_my_machine/file.zip"

'Events:
'When you choose pDownloader from the left dropdown at the top of your IDE, the right dropdown gives you the Events of the downoader. 
'The following events are available:
	'OnResponseStart: you won’t probably use this
	'OnResponseDataAvailable: triggers, when the first batch of data is coming in. Attention: this is triggered A LOT – use with care
	'OnResponseFinished: If you didn’t give a Filename to save, you have to get your data here (http.responsetext).
	'OnError: check here whats gone wrong.

'Hope this is of use to someone.
'I will probably update this soon to accomodate XML, get rid of the ADODB requirement and let the same object do synchronuous downloads. So keep looking;)


'HTTP Downloader

Option Explicit

Private pHTTPRequestCollection As New Collection
Private pCounter As Integer

Public Event OnResponseStart(http As WinHttpRequest, ByVal Status As Long, _
       ByVal ContentType As String, ByVal Tag As String)
Public Event OnResponseDataAvailable(http As WinHttpRequest, data() As Byte, _
       ByVal Tag As String)
Public Event OnResponseFinished(http As WinHttpRequest, ByVal Tag As String)
Public Event OnError(http As WinHttpRequest, ByVal ErrorNumber As Long, _
       ByVal ErrorDescription As String, ByVal Tag As String)
Public Event OnQueueEmpty()

Private Sub Class_Initialize()
  pCounter = 0
End Sub

Sub DownloadAsync(URL As String, Optional DestPath As String = "", Optional Tag As String = "")
  Dim httpx As HTTPRequest
  Dim key As String
  key = "key" & pCounter
  pCounter = pCounter + 1
  Set httpx = New HTTPRequest
  pHTTPRequestCollection.Add httpx, key  
  httpx.Download key, URL, DestPath, Tag, "CallBack", Me
End Sub

Sub CallBack(ByVal objID, ByVal cbtype)
  With getObj(objID)
    Select Case cbtype
    Case 1:
      RaiseEvent OnResponseStart(.WinHttp, .Status, .ContentType, .Tag)
    Case 2:
      RaiseEvent OnResponseDataAvailable(.WinHttp, .data, .Tag)
    Case 3:
      RaiseEvent OnResponseFinished(.WinHttp, .Tag)
      RemoveFromCollection objID
    Case 4:
      RaiseEvent OnError(.WinHttp, .ErrorNumber, .ErrorDescription, .Tag)
      RemoveFromCollection objID
    End Select
  End With
End Sub

Private Sub RemoveFromCollection(ByVal objID As String)
  pHTTPRequestCollection.Remove objID
  If pHTTPRequestCollection.Count = 0 Then RaiseEvent OnQueueEmpty
End Sub

Private Function getObj(ByVal key As String) As HTTPRequest
  Set getObj = pHTTPRequestCollection(key)
End Function

----------------------------------------------------------------------------------
'HTTPRequest

Option Explicit

Private WithEvents http As WinHttpRequest
Private ADStream As ADODB.Stream
Private p_id As String
Private pDestPath As String
Private pTag As String
Private pCallBack As String
Private pCallbackObj As Object
Private pBytes() As Byte
Private pErrorNumber As Long
Private pErrorDescription As String
Private pStatus As Long
Private pContenTtype As String

Property Get ErrorNumber() As Long
  ErrorNumber = pErrorNumber
End Property
Property Get ErrorDescription() As String
  ErrorDescription = pErrorDescription
End Property
Property Get Status() As Long
  Status = pStatus
End Property
Property Get ContentType() As String
  ContentType = pContenTtype
End Property
Property Get data() As Byte()
  data = pBytes
End Property
Property Get Id() As String
  Id = p_id
End Property
Property Get Tag() As String
  Tag = pTag
End Property
Property Get WinHttp() As WinHttpRequest
  Set WinHttp = http
End Property

Private Sub Class_Initialize()
  Set http = New WinHttpRequest
  Set ADStream = New ADODB.Stream
End Sub

Private Sub Class_Terminate()
  Set http = Nothing
  Set ADStream = Nothing
End Sub

Private Sub http_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
  pStatus = Status
  pContenTtype = ContentType
  Trigger 1
End Sub

Private Sub http_OnResponseDataAvailable(data() As Byte)
  pBytes = data
  Trigger 2
End Sub

Private Sub HTTP_OnResponseFinished()
  If pDestPath <> "" Then
    With ADStream
      .Type = adTypeBinary
      .Open
      .Write http.responseBody
      .SaveToFile pDestPath, adSaveCreateOverWrite
      .Close
    End With
  End If
  Trigger 3
End Sub

Private Sub HTTP_OnError(ByVal ErrorNumber As Long, _
    ByVal ErrorDescription As String)
  pErrorNumber = ErrorNumber
  pErrorDescription = ErrorDescription
  Trigger 4
End Sub

Sub Download(Id As String, URL As String, _
    DestPath As String, Tag As String, _
    CallBack As String, CB_obj As Object)
  p_id = Id
  pDestPath = DestPath
  pTag = Tag
  pCallBack = CallBack
  Set pCallbackObj = CB_obj
  http.Open "GET", URL, True
  http.send
End Sub

Private Sub Trigger(ByVal EventType As Integer)
  CallByName pCallbackObj, pCallBack, VbMethod, p_id, EventType
End Sub