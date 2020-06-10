Option Explicit

Dim fmo, smo, ssm

Set fmo = new Filemanager
Set smo = new ScriptManager
Set ssm = new StringSanitizeManager

Class Filemanager

    Private FileSystemObject
    Private FilePath
    Private FileContent
    Private FileObject

    Public Property Let FileContentObj(myfso)
        FileContent = myfso
    End Property

    Public Property Get FileContentObj()
        FileContentObj = FileContent
    End Property

    Private Sub Class_Initialize()

        Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
        Set FileObject = Nothing
        FileContent = ""
        FilePath = FileSystemObject.BuildPath(FileSystemObject.GetAbsolutePathName("."), "config.ini")

        ConfigFileInit

    End Sub

    Private Sub Class_Terminate()

    End Sub

    Public Sub ConfigFileInit()

        If FileSystemObject.FileExists(FilePath) Then
            GetRouterAddress
        Else
            CreateConfigFile
        End If

    End Sub

    Private Sub CreateConfigFile()

        Set FileObject = FileSystemObject.CreateTextFile(FilePath, True)
        FileContent = "http://192.168.1.1"
        FileObject.WriteLine FileContent
        FileObject.Close

    End Sub

    Private Sub GetRouterAddress()

        Set FileObject = FileSystemObject.OpenTextFile(FilePath, 1)
        FileContent = FileObject.ReadAll()
        FileObject.Close

    End Sub

    Public Sub WriteRouterAddress()

        Set FileObject = FileSystemObject.OpenTextFile(FilePath, 2)
        FileContent = document.getElementById("RouterAddressInput").value
        FileObject.WriteLine FileContent
        FileObject.Close

    End Sub

End Class

Class ScriptManager

    Private ScriptHostObject
    Private ScriptHostExecObject
    Private ScriptHostRunning
    Private ScriptHostFinished
    Private ScriptHostFailed
    Private ScriptHostExecString
    Private ScriptHostRunArgumentsList
    Private StringCommand
    Private ShellAppExecParameters
    Private ShellAppObj

    Public Property Let ScriptHostObj(mysho)
        ScriptHostObject = mysho
    End Property

    Public Property Get ScriptHostObj()
        ScriptHostObj = ScriptHostObject
    End Property

    Sub OpenRouterAddress(ByVal sanitizedAddress)

        ScriptHostObject.Run sanitizedAddress, 1, False

    End Sub

    Sub ShellExecWinDiagnosticTool()

        ShellAppObj.ShellExecute "msdt.exe", ShellAppExecParameters, "", "runas", 1

    End Sub

    Sub CheckInternetConnection(ByVal stringManager)

        Set ScriptHostExecObject = ScriptHostObject.Exec(ScriptHostExecString)

        Do While ScriptHostExecObject.Status = ScriptHostRunning

            ScriptHostObject.Run "timeout /t 1", 0, True

        Loop

        Select Case ScriptHostExecObject.Status

            Case ScriptHostFinished

                stringManager.StringOutputObj = ScriptHostExecObject.StdOut.ReadAll()

            Case ScriptHostFailed

                stringManager.StringOutputObj = ScriptHostExecObject.StdErr.ReadAll()

        End Select

        stringManager.UpdateConnectionStatus()

    End Sub

    Sub ResetInternetConnection()

        If document.getElementById("gridCheck").checked Then

            ScriptHostRunArgumentsList.Add "shutdown /f /r /t 0"

        Else

            ScriptHostRunArgumentsList.Remove("shutdown /f /r /t 0")

        End If

        For Each StringCommand in ScriptHostRunArgumentsList

            'MsgBox StringCommand
            ScriptHostObject.Run StringCommand, 0, True

        Next

    End Sub

    Private Sub Class_Initialize()
        
        Set ScriptHostObject = CreateObject("WScript.Shell")
        Set ShellAppObj = CreateObject("Shell.Application")

        ScriptHostRunning = 0
        ScriptHostFinished = 1
        ScriptHostFailed = 2

        ScriptHostExecString = "ping google.de"
        StringCommand = ""
        ShellAppExecParameters = "/id NetworkDiagnosticsWeb"

        Set ScriptHostRunArgumentsList = CreateObject("System.Collections.ArrayList")

        CommandListInit

    End Sub

    Private Sub CommandListInit()

        ScriptHostRunArgumentsList.Add "ipconfig /flushdns"
        ScriptHostRunArgumentsList.Add "ipconfig /registerdns"
        ScriptHostRunArgumentsList.Add "ipconfig /release"
        ScriptHostRunArgumentsList.Add "ipconfig /renew"
        ScriptHostRunArgumentsList.Add "netsh winsock reset"
        ScriptHostRunArgumentsList.Add "netsh winsock reset catalog"
        ScriptHostRunArgumentsList.Add "netsh int ipv4 reset"
        ScriptHostRunArgumentsList.Add "netsh int ipv6 reset"

        'ScriptHostRunArgumentsList.Add "ping google.de"

    End Sub

    Private Sub Class_Terminate()

    End Sub

End Class

Class StringSanitizeManager

    Private StringOutput
    Private RegexObj
    Private RegExMatchObj

    Public Property Let StringOutputObj(mystro)
        StringOutput = mystro
    End Property

    Public Property Get StringOutputObj()
        StringOutputObj = StringOutput
    End Property

    Private Sub Class_Initialize()
        
        StringOutput = ""
        
        Set RegexObj = New RegExp
        
        With RegexObj
            .IgnoreCase = True
            .Global = False
        End With

    End Sub

    Private Sub Class_Terminate()

    End Sub

    Sub UpdateConnectionStatus()

        If InStr(StringOutput, "Reply") > 0 Or InStr(StringOutput, "Antwort") Then

            With document.getElementById("checkResult")

                .innerText = "Verbindung erfolgreich hergestellt"
                .style.color = "green"

            End With

        Else

            With document.getElementById("checkResult")

                .innerText = "Verbindungsaufbau nicht möglich"
                .style.color = "red"

            End With

        End If

    End Sub

    Function RouterAddressInputValidator(ByVal strRouterAddress)

        StringOutput = strRouterAddress
        
        ' Pattern checks for http:// or https:// + subdomain?.?domain.suffix or ip 1-3.1-3.1-3.1-3
        RegexObj.Pattern = "(((http:\/\/)|(https:\/\/))(([A-z]+\.{1}[A-z]+\.?[A-z]+)|([0-9]{1,3}\.{1}[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3})))"
        
        Set RegExMatchObj = RegexObj.Execute(StringOutput)

        If RegexObj.Test(StringOutput) Then

            RouterAddressInputValidator = True

        Else

            RouterAddressInputValidator = False

        End If

    End Function

End Class


Sub AppInit

    FormInit

End Sub

Sub FormInit

    document.getElementById("RouterAddressInput").value = fmo.FileContentObj

End Sub

Sub PopRouterWebInterface

    fmo.FileContentObj = document.getElementById("RouterAddressInput").value

    'MsgBox "Routeraddresse wird aufgerufen und gespeichert"

    If ssm.RouterAddressInputValidator(fmo.FileContentObj) Then
        
        fmo.WriteRouterAddress
        smo.OpenRouterAddress(fmo.FileContentObj)
    
    Else
    
        MsgBox "Bitte die auf der Rückseite des Routers abgedruckte Adresse im Format: http://192.168.0.1 oder http://fritzbox.homelan eingeben!"
    
    End If

End Sub

Sub CheckConnection

    'MsgBox "Verbindung wird getestet"
    'document.getElementById("checkResult").innerText = "OK"

    smo.CheckInternetConnection(ssm)

End Sub

Sub ResetConnection

    'MsgBox "Verbindung wird zurückgesetzt"

    smo.ResetInternetConnection

End Sub

Sub OpenWinNetTrblshooter

    smo.ShellExecWinDiagnosticTool

End Sub