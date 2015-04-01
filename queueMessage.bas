Attribute VB_Name = "queueMessage"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       28th March 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       Add a reference to C:\windows\system32\winhttp.dll
'   References  :       https://issues.apache.org/jira/browse/AMQ-5579
'   References  :       Winhttp code is cut-and-paste with simplifications from
'               :       http://www.808.dk/?code-simplewinhttprequest
'
'   To replicate:       1.  Download this code
'                       2.  Create a new 64bit Excel 2013 workbook running on Windows 8.1 64 bit
'                       3.  Save the workbook as a .xlsm workbook with macros enabled
'                       4.  I run activemq locally so there is clearly no issue with firewalls or anything
'                       5.  Install ActiveMQ at cd "c:\Program Files (x86)\apache-activemq-5.10.0"
'                       6.  From a command prompt running as administrator execute cd "c:\Program Files (x86)\apache-activemq-5.10.0"
'                       7.  From that command prompt execute bin\win64\activemq.bat
'                       8.  Have a look at C:\Program Files (x86)\apache-activemq-5.10.0\data
'                       8.  If ActiveMQ has started you'll see a line like this:
'                       jvm 1    |  INFO | Apache ActiveMQ 5.10.0 (localhost, ID:apple-50841-1427852327092-0:1) started
'                       9.  Run an instance of http://localhost:8161/admin/queues.jsp in a browser on the machine
'                       10. In http://localhost:8161/admin/queues.jsp create a queue called TEST.A
'                       11. Now, run DoStuffWorking in the VBA code.
'                       12. In http://localhost:8161/admin/queues.jsp you will see a message has arrived on the queue
'                       13. In the VBA debug window you will see "queueMessage.DoStuffWorking Message sent"
'                       14. Now, run DoStuffNotWorking in the VBA code.
'                       15. In the VBA debug window you will see "queueMessage.DoStuffNotWorking HTTP 500 STREAMED"
'                       16. In http://localhost:8161/admin/queues.jsp you will see a message has not arrived on the queue
'                       17. I created a variable called lngFlavourOfWorkAroundToTry.  You can see a few workaround I have tried.
'                       18. Default for lngFlavourOfWorkAroundToTry is ZERO
'============================================================================================================================
'Constants
Const mstrTargetURL As String = "http://localhost:8161/api/message?destination=queue://TEST.A"
Const mstrPayload = "Hello World"
Dim varReturnFromWinHttp As Variant
'Variables

Function DoStuffWorking() As Boolean
'Constants
Const strMethodName As String = "queueMessage.DoStuffWorking "

varReturnFromWinHttp = GetDataFromURL(mstrTargetURL, "POST", "")

Debug.Print strMethodName & CStr(varReturnFromWinHttp)

End Function

Function DoStuffNotWorking() As Boolean
'Constants
Const strMethodName As String = "queueMessage.DoStuffNotWorking "

varReturnFromWinHttp = GetDataFromURL(mstrTargetURL, "POST", mstrPayload)

Debug.Print strMethodName & CStr(varReturnFromWinHttp)

End Function

Function GetDataFromURL(strURL, strMethod, strPostData)
'Constants
Const strMethodName As String = "queueMessage.GetDataFromURL "
'Variables
Dim lngTimeout
Dim strUserAgentString
Dim intSslErrorIgnoreFlags
Dim blnEnableRedirects
Dim blnEnableHttpsToHttpRedirects
Dim strHostOverride
Dim strLogin
Dim strPassword
Dim strResponseText
Dim objWinHttp As WinHttp.WinHttpRequest

lngTimeout = 1000
strUserAgentString = "http_requester/0.1"
intSslErrorIgnoreFlags = 13056 ' 13056: ignore all err, 0: accept no err
blnEnableRedirects = True
blnEnableHttpsToHttpRedirects = True
strHostOverride = ""
strLogin = "admin"
strPassword = "admin"
   
Set objWinHttp = New WinHttp.WinHttpRequest

objWinHttp.SetTimeouts lngTimeout, lngTimeout, lngTimeout, lngTimeout
objWinHttp.Open strMethod, strURL

If strMethod = "POST" Then
  objWinHttp.SetRequestHeader "Content-type", _
    "application/x-www-form-urlencoded"
 End If

'If strHostOverride <> "" Then
'    objWinHttp.SetRequestHeader "Host", strHostOverride
'End If

objWinHttp.Option(0) = strUserAgentString
objWinHttp.Option(4) = intSslErrorIgnoreFlags
objWinHttp.Option(6) = blnEnableRedirects
objWinHttp.Option(12) = blnEnableHttpsToHttpRedirects

If (strLogin <> "") And (strPassword <> "") Then
  objWinHttp.SetCredentials strLogin, strPassword, 0
End If

On Error Resume Next

Const lngFlavourOfWorkAroundToTry As Long = 0

If Len(strPostData) > 0 Then
    Select Case lngFlavourOfWorkAroundToTry
        Case 0
            objWinHttp.Send strPostData
        Case 1
            objWinHttp.Send ("Body=" & strPostData)
        Case 2
            objWinHttp.Send "OneWord"
        Case Else
    End Select
Else
    objWinHttp.Send
End If


If Err.Number = 0 Then
    If objWinHttp.Status = "200" Then
        GetDataFromURL = objWinHttp.ResponseText
    Else
        GetDataFromURL = "HTTP " & objWinHttp.Status & " " & objWinHttp.StatusText
    End If
Else
    GetDataFromURL = "Error " & Err.Number & " " & Err.Source & " " & Err.Description
End If

On Error GoTo 0

Set objWinHttp = Nothing

End Function

