Option Explicit

Const plink_cmd = "plink -v -ssh -batch user@100.100.100.100 -pw ??? query session"
Const shell_exec_finished = 1
Const shell_exec_failed = 2

Dim shell_std_output, admin_id, mstc_cmd, oWshShell, oWshShellExec, tmp_string, admin_id_regex, admin_id_matches

Set oWshShell = CreateObject("WScript.Shell")
Set oWshShellExec = oWshShell.Exec(plink_cmd)

' Checks Shell Execution Status every 100ms until it changes
Do While oWshShellExec.Status = 0
    Wscript.Sleep 100
Loop

If oWshShellExec.Status = shell_exec_finished  Then

    shell_std_output = oWshShellExec.StdOut.ReadAll

    ' Get admin_id based on shell_std_output
    GetAdminID

    mstc_cmd = "mstsc.exe /v:100.100.100.100 /shadow:" & admin_id & "/f /control /noConsentPrompt"

    StartRDPSession(mstc_cmd)

ElseIf oWshShellExec = shell_exec_failed Then

    shell_std_output = oWshShellExec.StdErr.ReadAll

    Wscript.Echo shell_std_output

Else

    shell_std_output = "WScript.Shell returned with status " & oWshShellExec.Status

    Wscript.Echo shell_std_output

End If

Sub GetAdminID()
    
    tmp_string = Replace(shell_std_output, " ", "")

    REM ID can be one digit or more

    Set admin_id_regex = New RegExp

    With admin_id_regex

        .Pattern = "admin([0-9]{1,2})"
        .IgnoreCase = True
        .Global = False

    End With

    Set admin_id_matches = admin_id_regex.Execute(tmp_string)

    admin_id = Replace(admin_id_matches.Item(0), "admin", "")

End Sub

Sub StartRDPSession(ByVal mstsc_command)

    oWshShell.Exec(mstsc_command)

End Sub