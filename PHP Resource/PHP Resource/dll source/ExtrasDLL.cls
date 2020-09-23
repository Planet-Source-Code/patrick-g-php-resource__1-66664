VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Extras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Below i will explain how to use each function.
'
'I will be adding many many more functions.
'Only if i see that people like my idea and find it can be useful
'I am trying to win on PSC one last time,
'but i do not want to win unless i deserve it!
'So let me know how my DLL can help you! =)
'
'To load the DLL in to your project do the following!
'
'Goto Project Menu > Refrences then hit browse
'Look for the MultiFuncDLL.dll file on your computer and hit ok!
'Then hit OK again
'
'Now in your project i'll use a "Form Load EG"
'Dim the following into your project
'
'dim Dll as multifuncdll.extras
'
'public sub form_load()
'set dll = new extras
'
'From there you can call any function like this
'dll.osuptime (And whatever else is in the selection menu
'
'end sub

Option Explicit
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset

Public Function MathAdd(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
'Useage: text1.text = dll.mathadd("1", "2") will give you a result with "3"
'Useage: msgbox can be used as well along with everything else
On Error GoTo fail:
    MathAdd = FirstNumber + SecondNumber
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "math +"
End Function

Public Function MathSubtract(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
'Useage: text1.text = dll.mathsubtract("2", "1") will give you a result with "1"
'Useage: msgbox can be used as well along with everything else
On Error GoTo fail:
    MathSubtract = FirstNumber - SecondNumber
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "math -"
End Function

Public Function MathDivide(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
'Useage: text1.text = dll.mathdivide("2", "1") will give you a result with "1"
'Useage: msgbox can be used as well along with everything else
On Error GoTo fail:
    MathDivide = FirstNumber / SecondNumber
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Math /"
End Function

Public Function MathMultiply(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
'Useage: text1.text = dll.mathMultiply("12", "12") will give you a result with "144"
'Useage: msgbox can be used as well along with everything else
On Error GoTo fail:
    MathMultiply = FirstNumber * SecondNumber
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "math X"
End Function

Public Function RandomNum(ByVal PickaNumber As Long)
'Useage: text1.text = dll.RandomNum("Any number") will pick a random number between your value
'You can also just leave it blank and just use dll.RandomNum and it will pick a number between 8000
On Error GoTo fail:
    If RandomNum = "" Then
        RandomNum = Int(Rnd * 8001)
    Else
        RandomNum = Int(Rnd * PickaNumber)
    End If
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "random number"
End Function

Public Function Computername() As String
'Useage: text1.text = dll.computername will return the network name of the computer
'Useage: msgbox can be used as well along with everything else
On Error GoTo fail:
  Dim lpBuff   As String * 25
  Dim retval   As Long
  retval = GetComputername(lpBuff, 25)
  Computername = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
  Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "computername"
End Function

Public Function SaveToINI(ByVal vntValue As Variant, ByVal sSection As String, ByVal sKey As String, SFilename As String)
'Useage: text1.text = dll.savetoini("Keyname", "Section", text1.text, appdir & "test.ini")
On Error GoTo fail:
    #If Win32 Then
        Dim xRet          As Long
    #Else
        Dim xRet          As Integer
    #End If
    
    xRet = WritePrivateProfileString(sSection, sKey, CStr(vntValue), SFilename)
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "savetoini"
End Function
Public Function GetFromINI(ByVal sSection As String, ByVal sKey As String, SFilename As String) As Variant
'Useage: text1.text = dll.getfromini("Keyname", "Section", "String", appdir & "test.ini")
'Useage: msgbox can be used as well along with everything else
On Error GoTo fail:
    #If Win32 Then
        Dim xRet          As Long
    #Else
        Dim xRet          As Integer
    #End If

Dim sReturnStr    As String
Dim nStringLen    As Integer

    nStringLen = 255
    sReturnStr = String(nStringLen, Chr$(0))
    xRet = GetPrivateProfileString(sSection, sKey, "", sReturnStr, nStringLen, SFilename)
    GetFromINI = Left(sReturnStr, xRet)
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "getfromini"
End Function

Public Function ExitMessage(ByVal Message As String, ByVal title As String, ByVal IfYesMessage As String, ByVal IfNoMessage As String, Optional YesNo As Boolean)
'Info: Exit message is something that just saves time, isn't absolutely useful but sufficiant
'Useage: To get Yes or No message use,
'Useage: call dll.exitmessage("Enter Exit Message", "Msgbox Title", "IF answer is yes message", "If answer is no message", true)
'Useage: To show just a regular exit message use,
'Useage: call dll.exitmessage("Enter Exit Message", "Msgbox Title", "", "", false)
On Error GoTo fail:
    If YesNo = True Then
        Dim Answer
        Answer = MsgBox(Message, vbYesNo Or vbInformation, title)
            If Answer = vbYes Then
               MsgBox IfYesMessage, vbInformation, title
            Else
                If IfNoMessage = "end" Then
                    Dim Form As Form
                    Unload Form
                    Set Form = Nothing
                Else
                    Dim Cancel As Integer
                    MsgBox IfNoMessage, vbInformation, title
                    Cancel = 1
                    Exit Function
                End If
            End If
    Else
        MsgBox Message, vbInformation, title
    End If
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "exitmessage"
End Function

Public Function AppIsRunning() As Boolean
'Useage: Use the follow for eg on how to use this
'
'if dll.appisrunning = true then
'   msgbox "App is already running", vbinformation, "App"
'end if
On Error GoTo fail:
    If (App.PrevInstance = True) Then
        AppIsRunning = True
    Else
        AppIsRunning = False
    End If
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "appoisrunning"
End Function

Public Function AppVersion()
'Useage: very self explanitory!!!
'Msgbox dll.appversion
    AppVersion = VB.App.Major & "." & VB.App.Major & "." & VB.App.Revision
End Function

Public Function FileCheck(Path$) As Boolean
'Info: Gives ability to check if a file exists or not
'Useage: call dll.filecheck("C:\windows\command.com")

    FileCheck = True
    
    Dim Disregard As Long
    
    Disregard = FileLen(Path)
    
    If Disregard <> 0 Then
        FileCheck = False
        Exit Function
    End If

End Function

Function OsExitWindows(BootMode As Integer)
'Info: Allows you to shutdown or reboot your computer with small little command
'Useage: dll.osexitwindows("1") Shutdown windows
'Useage: dll.osexitwindows("2") reboots windows

Dim bootans As Integer, bootvalue As Long
On Error GoTo fail:

Select Case BootMode
    Case 1 'Shutdown Windows
        bootans = vbYes
            If bootans = vbYes Then
                bootvalue = ExitWindowsEx(EWX_SHUTDOWN, 0&)
            End If
    Case 2 ' Reboot Windows
        bootans = MsgBox("Are you sure you want to reboot windows?", vbQuestion Or vbYesNo, "Reboot Windows")
            If bootans = vbYes Then
                bootvalue = ExitWindowsEx(EWX_REBOOT, 0&)
            End If
    Case Else
End Select
Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "osexitwindows"
End Function

Function FreeProcess()
'Info: Unfreezes a locked loop or subroutine
'Useage: dll.freeprocess
On Error GoTo fail:
Dim Process%
Do: DoEvents
    Process = Process + 1
    If Process = 50 Then Exit Do
Loop
Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "freeprocess"
End Function

Public Function FileDirCheck(ByVal sDirName As String) As Boolean
'Info: Returns true if path exists else it returns false
'Useage: dll.filedircheck("Filename") Will return true or false
Dim sDir As String
On Error GoTo fail:
    FileDirCheck = False
    sDir = Dir$(sDirName, vbDirectory)
        If (Len(sDir) > 0) And (Err = 0) Then
            FileDirCheck = True
        End If
        Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "filedircheck"
End Function

Public Function FileSize(FilePathName As String) As String
'Info: Returns size of a file
'USAGE: Label1.Caption = FileSize("C:\Stuff.exe")
On Error GoTo fail:
Dim Str As Single
    Str = FileLen(FilePathName)
    FileSize = File_ByteConversion(Str)
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "filesize"
End Function

Public Function OsUptime()
'Info: Returns current uptime of your personal computer
'Useage: text1.text dll.osuptime
Dim Secs, Mins, Hours, Days
Dim TotalMins, TotalHours, TotalSecs, TempSecs
Dim CaptionText
On Error GoTo fail:
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

    CaptionText = Days & " Days, " & Hours & " Hours, " & Mins & " Minutes, " & TotalSecs & " seconds"
    OsUptime = CaptionText
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Os uptime"
End Function

Public Function Show()
'Info: Just displays the about dialog
'Useage: dll.show
    frmAbout.Show
End Function

Public Function SaveRegistry(ByVal StrFolder As String, ByVal StrSubFolder As String, ByVal Key As String, ByVal Str As String)
'Info: Gives ability to save program settings in registry.
'Info: I know its not the best way to do it, but it works does it not?
'Useage: dll.saveregistry "Folder name", "Subfolder like options folder", "Password", text1.text
On Error GoTo fail:
    SaveSetting StrFolder, StrSubFolder, Key, Str
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Save Reg"
End Function

Public Function GetRegistry(ByVal StrFolder As String, ByVal StrSubFolder As String, ByVal Key As String) As String
'Info: Gives ability to save program settings in registry.
'Info: I know its not the best way to do it, but it works does it not?
'Useage: dll.getregistry "Folder name", "Subfolder like options folder", "Key name"
On Error GoTo fail:
    GetRegistry = GetSetting(StrFolder, StrSubFolder, Key)
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Get Reg"
End Function

Public Function DelRegistry(ByVal Folder As String, Optional Section As String, Optional Key As String)
'Info: Gives ability to delete program settings from registry
'Info: I know its not the best way to do it, but it works does it not?
'Useage: dll.delregistry "Folder", "Section", "Key"
'Useage: You can just delete ALL settings if you used dll.delregistry "Folder"
On Error GoTo fail:
    DeleteSetting Folder, Section, Key
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Delete Reg"
End Function

Public Function ErrorMessage(ByVal ErNumber As String, ByVal ErDesc As String, ByVal ErSource As String, ByVal AppFileName As String, ByVal WhatAreYouTryingToDO As String)
'Info: Sorry for the sloppyness of this one but it does work, makes debugging errors a bit better
'Info: But this is used for your program if instead of using on error resume next just do the following
'Info: Below i will simulate HOW to create and test this error debugging
'Info: Put what is below in something like public sub form_load() or something to test it
'Useage:
'
'on error goto ermsg:
'   dim Var as integer
'Var = 9000000000000000
'ermsg:
'   dll.errormessage err.Number, err.Description, err.Source, app.EXEName, "IF program crashes, put WHAT you have the rutine try to do"
On Error GoTo fail:
With ErrorForm
    .Show
    .Errbox.SelFontSize = 8
    .Errbox.SelColor = vbBlack
    .Errbox.SelText = "Filename: "
    .Errbox.SelColor = vbBlack
    .Errbox.SelText = AppFileName & vbCrLf
    
    .Errbox.SelFontSize = 8
    .Errbox.SelColor = vbBlack
    .Errbox.SelText = "Error Source: "
    .Errbox.SelColor = vbBlack
    .Errbox.SelText = ErSource & vbCrLf
    
    .Errbox.SelFontSize = 8
    .Errbox.SelColor = vbBlack
    .Errbox.SelText = "Error Number: "
    .Errbox.SelColor = vbBlack
    .Errbox.SelText = ErNumber & vbCrLf
    
    .Errbox.SelFontSize = 8
    .Errbox.SelColor = vbBlack
    .Errbox.SelText = "Error Desc: "
    .Errbox.SelColor = vbBlack
    .Errbox.SelText = ErDesc & vbCrLf
    
    .Errbox.SelFontSize = 8
    .Errbox.SelColor = vbBlack
    .Errbox.SelText = "How it happend: "
    .Errbox.SelColor = vbBlack
    .Errbox.SelText = WhatAreYouTryingToDO & vbCrLf
End With
Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "debug window"
End Function

Public Function CreateSplashScreen(ByVal IntVal As Integer, ByVal Color As String, ByVal Company As String, ByVal CopyRight As String, ByVal Warning As String, ByVal Platform As String, ByVal Program As String, ByVal Version As String)
'Info: Gives ability to create such a simple splash screen for your program
'Info: Very simple but it does the job right?
'Info: The 5000 of the Useage desc is the interval of the timer
'Info: Meaning how long it should stay open for before closing on its own
'Info: Default is 4 seconds
'Useage: dll.spashscreen 5000, "Licenced to?", "Company name", "Copyright info", "Warning like a disclaimer", "What operating systems it can run on", "What the program is called", "Version"
On Error GoTo fail:
With frmSplash
    .Show
    .BackColor = Color
    .Frame1.BackColor = Color
    .Timer1.Interval = IntVal
    .lblCompany = Company
    .lblCopyright = CopyRight
    .lblLicenseTo = Computername
    .lblPlatform = Platform
    .lblProductName = Program
    .lblVersion = "Version: " & Version
    .lblWarning = "Warning... " & Warning
End With
Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "createsplashscreen"
End Function

Public Function CreateAboutScreen(ByVal Caption As String, ByVal ProgramTitle As String, ByVal desc As String, ByVal Version As String, ByVal Warning As String)
'Info: Gives ability to create a simple about program dialog
'Useage: dll.createaboutscreen "Caption Text", "Title of program", "Description", "Version", "Disclaimer"
On Error GoTo fail:
With frmAbout
    .Show
    .Caption = Caption
    .lbltitle = ProgramTitle
    .lblDescription = desc
    .lblDisclaimer = "Warning... " & Warning
    .lblVersion = "Version: " & Version
End With
Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "createaboutscreen"
End Function

Public Function CenterForm()
'Info: Allows you to center your program on anyones screen
'Useage: dll.centerform
On Error GoTo fail:
Dim X, Y As Integer
Dim Frm As Form, Src As Screen
    X = Src.Width \ 2 - Frm.Width \ 2
    Y = Src.Height \ 2 - Frm.Height \ 2
Frm.Move X, Y
Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "centerform"
End Function

Public Function ClickFormMove()
'Info: Gives ability to click anywhere on a form and move it
'Info: Anywhere on form_mousedown do the following
'Useage: dll.clickformmove
On Error GoTo fail:
Dim Frm As Form
    releaseCapture
    Call SendMessage(Frm.hWnd, &HA1, 2, 0&)
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "clickformmove"
End Function

Public Function ListBoxSave()
On Error GoTo fail:
    Dim free%, File As String, List As ListBox, SaveList As Long
    free = FreeFile

        Open File For Output As #free
            If FileCheck(File) = False Then Exit Function
                For SaveList& = 0 To List.ListCount - 1
                    Print #free, List.List(SaveList&)
                Next SaveList&
        Close #free
        Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "listboxsave"
End Function

Public Function ListBoxLoad()
'USAGE: LOADLIST(LIST1,"STUFF.LST")
'THAT WILL LOAD THE CONTENTS OF STUFF.LST
On Error GoTo fail:
    Dim free%, G$, File As String, ListBox As ListBox
    free = FreeFile
        If FileCheck(File) = False Then Exit Function
            ListBox.Clear
                Open File For Input As #free
                    Do Until EOF(free)
                        Line Input #free, G$
                            ListBox.AddItem G$
                    Loop
                Close free
                Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "listboxload"
End Function

Public Function PlayWAV(File As String)
'Info: Allows you to play *.wav files
'Useage: dll.playwav("Filename.wav")
On Error GoTo fail:
Dim WaveFile As String
            
    If FileCheck(File) = False Then
        Exit Function
    Else
        WaveFile = sndPlaySound("FILE", SND_ASYNC)
    End If
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "play wav"
End Function

Public Function StopWAV()
'Info: If you have a long wav playing, this will allow you to stop it
'Useage: dll.stopwav
On Error GoTo fail:
Dim WavFile As String
    WavFile = sndPlaySound("", SND_ASYNC)
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "stop wav"
End Function

Public Function CtrlAltDel_Disable()
'Info: Disable CTRL ALT DEL
'Info: To use this just put this in Form_Load
'Useage: dll.ctrlaltdel_disable
On Error GoTo fail:
Dim Cnt$, Var$
    Var = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, Cnt, 0)
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "ctrlaltdel disable"
End Function

Public Function CtrlAltDel_Enable()
'Info: Enable CTRL ALT DEL
'Info: To use this just put this in Form_Load
'Useage: dll.ctrlaltdel_enable
On Error GoTo fail:
Dim Cnt$, Var$
    Var = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, Cnt, 0)
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "ctrlaltdel enable"
End Function

Public Function Encryption(Text, Types)
'Info: Just a simple encryption method
'Useage: To Encrypt, dll.encryption("Text to encrypt", 0)
'Useage: To De-Crypt, dll.encryption("Text to decrypt", 1)
On Error GoTo fail:
Dim Cnt%, Var$, Pro$
For Cnt = 1 To Len(Text)
    If Types = 0 Then
        Var$ = Asc(Mid(Text, Cnt, 1)) - 1
    Else
        Var$ = Asc(Mid(Text, Cnt, 1)) + 1
    End If

    Pro$ = Pro$ & Chr(Var$)
Next Cnt
    Encryption = Pro$
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "encryption"
End Function

Public Function OpenProgram(OpenWhat As String)
'Info: Gives ability to open any program
'Usage: openprogram "c:\appname.exe"
'Usage: openprogram "http://www.pscode.com/vb"
'Usage: openprogram "mailto: email@server.com"
On Error GoTo fail:
Dim Frm As Form
    ShellExecute Frm.hWnd, "Open", OpenWhat, "", &O0, 1
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "openprogram"
End Function

Public Function SendEmail(ByVal Email As String, Optional Subject As String, Optional Body As String)
'Info: pretty much self explanitory, gives user ability to send an e-mail
'Useage: dll.sendemail "patrick-george@rogers.com", "User of DLLs", "Enter message here"
On Error GoTo fail:
    OpenProgram "mailto: " & Email & "?subject=" & Subject & "?body=" & Body
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "sendemail"
End Function

Public Function GotoWebsite(ByVal Site As String)
'Info: Dunno how much more simple this can get, but it opens a website to where ever you tell it to go
'Useage: dll.gotowebsite "http://www.pscode.com/vb"
On Error GoTo fail:
    OpenProgram Site
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "gotowebsite"
End Function

Public Function DatabaseOpen(ByVal Filename As String, ByVal TableName As String)
'Info: Load the database when requested
'Useage: dll.databaseopen app.path & "/databasefile.mdb", "TableName"
On Error GoTo fail:

    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(Filename)
    Set rs = db.OpenRecordset(TableName, dbOpenTable)
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "databaseopen"
End Function

Public Function DatabaseAddNew(ByVal Where As String, ByVal What As String)
'Info: Allows you to add new entrys to the database
'Useage: dll.databaseaddnew "firstname", text1.text
On Error GoTo fail:
    rs.AddNew
    rs(Where) = What
    rs.Update
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "databaseaddnew"
End Function

Public Function DatabaseEdit(ByVal Where As String, ByVal What As String)
'Info: Allows you to edit the database
'Useage: dll.databaseedit "firstname", text1.text
On Error GoTo fail:
    rs.Edit
    rs(Where) = What
    rs.Update
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "databaseedit"
End Function

Public Function DatabaseDelete(ByVal What As String, Optional Move As Integer)
'Info: Allows you to delete entry from database
'Useage: dll.databasedelete "firstname", 3 (To delete ID 3 from the database (It is optional))
On Error GoTo fail:
If Move Then
    rs.Move Move
End If
    rs.Delete
    rs.Update
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "databasedelete"
End Function

Public Function DatabaseMove(ByVal X As Integer)
'Info: Allows you to move X amount in the database
'Useage: dll.databasemove 3
On Error GoTo fail:
    rs.Move X
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "databasemove"
End Function

Public Function DatabaseMoveNext()
'Info: Allows you to move to the next entry in the database
'Useage: dll.databasemovenext
On Error GoTo fail:
    rs.MoveNext
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "databasemovenext"
End Function

Public Function DatabaseMoveLast()
'Info: Allows you to move to the last entry in the database
'Useage: dll.databasemovelast
On Error GoTo fail:
    rs.MoveLast
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "DatabaseMoveLast"
End Function

Public Function DatabaseMovePrevious()
'Info: Allows you to move from current posistion in the database to the previous entry
'Useage: dll.databasemoveprevious
On Error GoTo fail:
    rs.MovePrevious
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "DatabaseMovePrevious"
End Function

Public Function DatabaseCount() As String
'Info: Allows you to count the amount of entrys in the database
'Useage: dll.databasecount
On Error GoTo fail:
    DatabaseCount = rs.RecordCount
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Databasecount"
End Function

Public Function DatabaseEOF() As String
'Info: Self explanitory really.. "End of file" EOF
'Useage: dll.databaseeof
On Error Resume Next
    DatabaseEOF = rs.EOF
End Function

Public Function DatabasePercent() As String
'Info: Good for when using on a progressbar you can show percentage of how much the database is loaded
'Useage: dll.databasepercent
On Error GoTo fail:
    DatabasePercent = Int(rs.PercentPosition)
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Trying to get database percent position"
End Function

Public Function DatabaseClose()
'Info: Closes the database, normally put this on an UNLOAD form or EXIT area
'Useage: dll.databaseclose
On Error GoTo fail:
    rs.Close
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Attempting to remove spaces"
End Function

Public Function DatabaseRead(ByVal What As String, Optional Move As Integer) As String
'Info: Reads each entry from the database then moves to the next
'Useage: text1.text = dll.databaseread("firstname")
On Error GoTo fail:

    If rs(What) = "" Then
        rs.MoveFirst
    End If
    
    If Move Then
        rs.Move Move
    End If
        DatabaseRead = rs(What)
    Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Reading Database"
End Function

Public Function DatabaseMoveFirst()
    'move database to first record
    
        rs.MoveFirst
           
End Function



Public Function RemoveSpaces(ByVal Area As String)
'Info: Allows you to remove spaces from a text box or listbox "Hello and how are you" would be "helloandhowareyou"
'Useage: dll.removespaces text1
'Useage: dll.removespaces list1
On Error GoTo fail:

If Left$(Area, 4) = "text" Then
    Dim Text As TextBox
    Dim Cnt As String, i As Integer
Cnt = ""
    For i = 1 To Len(Text)
        If (Mid(Text, i, 1) <> " ") Then
            Cnt = Cnt + Mid(Text, i, 1)
        End If
    Next i
Text = Cnt
ElseIf Left$(Area, 4) = "list" Then
    Dim List As ListBox
        Dim Cnt1 As String, i1 As Integer
Cnt1 = ""
    For i1 = 1 To Len(List)
        If (Mid(List, i1, 1) <> " ") Then
            Cnt1 = Cnt1 + Mid(List, i1, 1)
        End If
    Next i1
List = Cnt1
End If

fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Attempting to remove spaces"
End Function


Public Function AccessDeniedMessage(Optional Message As String)
On Error GoTo fail:
    If Message Then
        MsgBox Message, vbInformation, "Access Denied"
        Exit Function
    End If
        MsgBox "You do not have authorized access." & vbCrLf & vbCrLf & "Access is denied!", vbCritical, "Access Denied"
        Exit Function
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Access denied message"
End Function

Public Function CreateCustomMessageBox(ByVal Caption As String, ByVal title As String, ByVal desc As String)
On Error GoTo fail:
    If Len(desc) > 250 Then
        MsgBox "You cannot have custom messages greater then 100 chr's", vbInformation, "Custom Message"
        Exit Function
    End If
With FrmMsgbox
    .Show
    .Caption = Caption
    .lbltitle = title
    .lbldesc = desc
    Exit Function
End With
fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Loading custom message box"
End Function


Public Function CreateStopMsg(ByVal DlgTitle As String, ByVal StopMessage As String)
'Create stop message like a "Login failed" type of message
'Usage: CreateStopMsg("Title", "Message")

On Error GoTo fail:

    If frmstop.stopmsg.Caption = "" Then
        MsgBox "Improper use of CreateStopMsg. " & vbCrLf & "Usage: CreateStopMsg(Title, Message)", vbCritical, "Error"
    End If
    
    With frmstop
        .Show
        .Caption = DlgTitle
    End With
    
    frmstop.stopmsg.Caption = StopMessage
    Exit Function
    
    

fail:
    ErrorMessage Err.Number, Err.Description, Err.Source, App.EXEName, "Custom stop message"
End Function


Public Function GetOSVersion() As String

'Tries to get Windows OS Version
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
Dim getVersion As String
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
With osinfo
    Select Case .dwPlatformId
        Case 1
            If .dwMinorVersion = 0 Then
                getVersion = "Windows 95"
            ElseIf .dwMinorVersion = 10 Then
                getVersion = "Windows 98"
            End If
        Case 2
            If .dwMajorVersion = 3 Then
                getVersion = "Windows NT 3.51"
            ElseIf .dwMajorVersion = 4 Then
                getVersion = "Windows NT 4.0"
            ElseIf .dwMajorVersion = 5 Then
                getVersion = "Windows XP"
            End If
        Case Else
            getVersion = "Failed"
    End Select
End With


End Function

























