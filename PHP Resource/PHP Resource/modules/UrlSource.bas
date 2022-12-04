Attribute VB_Name = "UrlSource"
Option Explicit

'Unfortunately, not my script but i've kept it in my archive as it does come in
'use a lot when i decide to program in vb.  I'd give credit if i knew who did it
    Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
    Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
    Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
    Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
    Public Const IF_FROM_CACHE = &H1000000
    Public Const IF_MAKE_PERSISTENT = &H2000000
    Public Const IF_NO_CACHE_WRITE = &H4000000
    Private Const BUFFER_LEN = 256

Public Function GetUrlSource(sURL As String) As String
    
    'the commenting below is not mine, so i'll leave it be and let it be the way it was.
    
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sdata As String
    Dim hInternet As Long, hSession As Long, lReturn As Long

    'get the handle of the current internet connection
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    'get the handle of the url
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    'if we have the handle, then start reading the web page
    If hInternet Then
        'get the first chunk & buffer it.
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sdata = sBuffer
        'if there's more data then keep reading it into the buffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sdata = sdata + Mid(sBuffer, 1, lReturn)
        Loop
    End If
   
    'close the URL
    iResult = InternetCloseHandle(hInternet)

    GetUrlSource = sdata
End Function
