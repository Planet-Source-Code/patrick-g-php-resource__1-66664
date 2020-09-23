Attribute VB_Name = "Utilities"
Option Explicit
'My regular declare functions and varibles.  I have these saved for future uses.
'you'll probably notice the same declars in some of my programs that have no use
'for that program, but i put them in anyway
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Declare Function GetTickCount& Lib "kernel32" ()
    Declare Function GetComputername Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Const HWND_NOTOPMOST = -2
    Public Const HWND_TOPMOST = -1
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
    Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Global r%
    Global InIPath$
    Global entry$

Function GetTok(ByVal strVal As String, intIndex As Integer, strDelimiter As String) As String
    'Not sure if you've ever used mIRC scripting but i got this idea from the old fasion
    'scripting days of mIRC.  gettok(string, where, what) to put easier if you had a file
    'that had strings like this var1%var2%var3
    'this function will allow you to read var2 just by gettok(string, 2, %)
        Dim strSubString() As String
        Dim intIndex2 As Integer
        Dim i As Integer
        Dim intDelimitLen As Integer
        intIndex2 = 1
        i = 0
        intDelimitLen = Len(strDelimiter)
        Do While intIndex2 > 0
            ReDim Preserve strSubString(i + 1)
            intIndex2 = InStr(1, strVal, strDelimiter)
            If intIndex2 > 0 Then
                strSubString(i) = Mid(strVal, 1, (intIndex2 - 1))
                strVal = Mid(strVal, (intIndex2 + intDelimitLen), Len(strVal))
            Else
                strSubString(i) = strVal
            End If
            i = i + 1
        Loop
        If intIndex > (i + 1) Or intIndex < 1 Then
            GetTok = ""
        Else
            GetTok = strSubString(intIndex - 1)
        End If
End Function
Public Sub ontop(FormName As Form)
    'This will put whatever form i want ontop of all others
        Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub Notontop(FormName As Form)
    'this will disable the ONTOP of all others
        Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Function FileResource(ByVal editor As String, ByVal sFile As String)
    'set varibles
        Dim info As String

    'lets verify the file exists so we avoid errors
        If FileExists(App.Path & "\db\source\" & sFile & ".phpr") = False Then
            MsgBox "I am unable to open that file for you.  File does not exist.", vbCritical, "Missing File"
            Exit Function
        End If
        
    'now we figure out what editor is selected and use it
        If editor = "wordpad" Then
            Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe " & Chr(34) & App.Path & "\db\source\" & sFile & ".phpr" & Chr(34), vbNormalFocus
            Exit Function
        End If
    
        If editor = "notepad" Then
            Shell "notepad " & App.Path & "\db\source\" & sFile & ".phpr", vbNormalFocus
            Exit Function
        End If
        
    'if the editor is the internal editor (viewer) then we'll use the one i made
        If editor = "phpresource" Then
            FrmView.Show
            FrmView.Caption = "PHP Resource Viewer -  " & sFile & ".phpr"
                Open App.Path & "/db/source/" & sFile & ".phpr" For Input As #1
                    Do While Not EOF(1)
                        Line Input #1, info
                        FrmView.Text1.Text = FrmView.Text1.Text & vbCrLf & info
                    Loop
            Close #1
        
        End If
End Function

Public Function FileExists(FullFileName As String) As Boolean
    'verify the file exists
        On Error GoTo iend
            Open FullFileName For Input As #1
            Close #1
        'If the file exists we'll set to true.  If it doesn't it'll get an error and thats
        'where the on error comes into play to tell us its false
            FileExists = True
            Exit Function
iend:
        'file doesn't exist give an error
            FileExists = False
            Exit Function
End Function

Function GetFromINI(AppName$, KeyName$, FileName$) As String
    'this is the function to read an ini file
        Dim RetStr As String
            RetStr = String(255, Chr(0))
            GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function
Public Function SaveToINI(ByVal vntValue As Variant, ByVal sSection As String, ByVal sKey As String, SFilename As String)
    'this is the function to save to the ini file
        #If Win32 Then
            Dim xRet          As Long
        #Else
            Dim xRet          As Integer
        #End If
            xRet = WritePrivateProfileString(sSection, sKey, CStr(vntValue), SFilename)
End Function
Public Function Uptime()
    'setting our veribles, sure know there are a lot of them
        Dim Secs, Mins, Hours, Days
        Dim TotalMins, TotalHours, TotalSecs, TempSecs
        Dim CaptionText
    
    'now going to do my mad calculating to figure out the uptime by reading the tickcount
        TotalSecs = Int(GetTickCount / 1000)
        Days = Int(((TotalSecs / 60) / 60) / 24)
        TempSecs = Int(Days * 86400)
        TotalSecs = TotalSecs - TempSecs
        TotalHours = Int((TotalSecs / 60) / 60)
        TempSecs = Int(TotalHours * 3600)
        TotalSecs = TotalSecs - TempSecs
        TotalMins = Int(TotalSecs / 60)
        TempSecs = Int(TotalMins * 60)
        TotalSecs = (TotalSecs - TempSecs)
            If TotalHours > 23 Then
                Hours = (TotalHours - 23)
            Else
                Hours = TotalHours
            End If
           
            If TotalMins > 59 Then
                Mins = (TotalMins - (Hours * 60))
            Else
                Mins = TotalMins
            End If

    'Now that we've done our calculating we'll now give our result
        CaptionText = Days & " Days, " & Hours & " Hours, " & Mins & " Minutes, " & TotalSecs & " seconds"
        Uptime = CaptionText
        Exit Function
End Function

Public Function Computername() As String
    'this is how you get the computer name.  Kinda complicated just to get a stupid name
        Dim lpBuff   As String * 25
        Dim retval   As Long
            retval = GetComputername(lpBuff, 25)
            Computername = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function


Public Function Encryption(Text, Types)
    'this is my very, very insecure encryption to mask your password if you select
    'it to be saved (remembered) when logging in to add resources
    'May be a simple encryption but it does mask the password at least

        Dim Cnt%, Var$, Pro$
        For Cnt = 1 To Len(Text)
            If Types = 0 Then
                Var$ = Asc(Mid(Text, Cnt, 1)) - (Len(Text) * 2) \ 2
            Else
                Var$ = Asc(Mid(Text, Cnt, 1)) + (Len(Text) * 2) \ 2
            End If

                Pro$ = Pro$ & Chr(Var$)
        Next Cnt
            
            'we now display encryption result
                Encryption = Pro$
End Function

