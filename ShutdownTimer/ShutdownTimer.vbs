Option Explicit

Dim oShutdown
Const strShutdownAbortCmd = "shutdown -a"

Class cShutdown

    Private pWshShell
    Private pShutdownStr
    Private pSelectedIndex
    Private pValidCustomInput
    Private pCustomInput
    Private pCustomInputSanitized
    Private pInputRegex
    Private pInputRegexMatch
    Private pInStrPos
    private pInputInSeconds
    private pIntCounterVar
    private pCounterHours
    private pCounterMinutes
    private pCounterSeconds
    private pCounterStr
    private ExitCondition

    Sub ShutdownAbort

        pWshShell.Run(strShutdownAbortCmd)

    End Sub

    Sub ExecShutdown

        pShutdownStr = "shutdown -s -t " & pInputInSeconds & " -f"
        pWshShell.Run(pShutdownStr)

    End Sub

    Sub SanitizeInput

        ' Please keep in mind on a German system we need a commata here for calc to work instead of an amurrican dot
        pInStrPos = Instr(pCustomInput, ".")

        If pInStrPos > 0 Then

            pCustomInputSanitized = Replace(pCustomInput, ".",",")

        Else

            pCustomInputSanitized = pCustomInput

        End If

    End Sub

    Sub SetCountdown

        pIntCounterVar = pInputInSeconds

        ExitCondition = False

        'document.getElementById("TimeStr").innerText = pInputInSeconds

        Do While pIntCounterVar > 0

            If ExitCondition Then
                Exit Do
            End If

            CountDownCalc

            pIntCounterVar = pIntCounterVar - 1
            pInputInSeconds = pInputInSeconds - 1

            pWshShell.Run "timeout /t 1", 0, True

        Loop
        

    End Sub

    Sub CountDownCalc
        ' For the exact formula please take a look at Countdown.vbs
        pCounterHours = pInputInSeconds / 3600 ' 60 * 60
        pCounterMinutes = (pCounterHours - Int(pCounterHours)) * 60
        pCounterSeconds = pInputInSeconds Mod 60

        pCounterStr = Int(pCounterHours) & " Std. " & Int(pCounterMinutes) & " Min. " & pCounterSeconds & " Sek." 

        document.getElementById("TimeStr").innerText = pCounterStr

    End Sub

    Function ValidInput(custom_input)

        Set pInputRegex = New RegExp

        With pInputRegex
            ' Match 1 digit then commata or dot separator and then 1 to 2 digits for time calc
            .Pattern = "^\d{1}((\.|\,)\d{1,2})?$"
            .IgnoreCase = True
            .Global = False

        End With

        'Set pInputRegexMatch = pInputRegex.Execute(custom_input)

        If pInputRegex.Test(custom_input) Then

            ValidInput = True

        Else

            ValidInput = False

        End If

        'MsgBox pInputRegexMatch.Item(0)


    End Function

    Sub InitShutdown

        pSelectedIndex = document.getElementById("TimeSelectOptions").selectedIndex
        pCustomInput = document.getElementById("CustomTimeInput").value

        If pSelectedIndex <> 0 Then

            pInputInSeconds = document.getElementById("TimeSelectOptions").value
        
            'MsgBox pInputInSeconds
            ExecShutdown
            SetCountdown

        ElseIf pCustomInput <> "" Then

            pValidCustomInput = ValidInput(pCustomInput)

            If pValidCustomInput Then

                SanitizeInput

                pInputInSeconds = pCustomInputSanitized * 3600

                'MsgBox pInputInSeconds
                ExecShutdown
                SetCountdown

            Else

                MsgBox "Bitte eine gültige Zahl eingeben z.B. 3 oder 1.5 oder 2,4"
                ResetValues

            End If

        Else

            MsgBox "Bitte eine Zeit auswählen oder eingeben."

        End If


    End Sub


    Sub ResetValues

        pCounterStr = "0 Std. 0 Min. 0 Sek."
        pShutdownStr = ""
        pCounterHours = 0
        pCounterMinutes = 0
        pCounterSeconds = 0
        pIntCounterVar = 0
        pInStrPos = 0
        pSelectedIndex = 0
        pInputInSeconds = 0
        pCustomInput = ""
        pCustomInputSanitized = ""
        pValidCustomInput = False
        ExitCondition = True
        
        Set pInputRegex = Nothing
        Set pInputRegexMatch = Nothing

        document.getElementById("TimeSelectOptions").selectedIndex = 0
        document.getElementById("CustomTimeInput").Value = ""
        document.getElementById("TimeStr").innerText = pCounterStr

    End Sub

    Private Sub Class_Initialize

        Set pWshShell = CreateObject("WScript.Shell")

        ResetValues

    End Sub

    Private Sub Class_Terminate   

        Set pWshShell = Nothing

    End Sub

End Class

Sub AppInit

    Set oShutdown = New cShutdown

End Sub

Sub AbortShutdown

    oShutdown.ResetValues
    oShutdown.ShutdownAbort

End Sub

Sub SetShutdownTimer

    oShutdown.InitShutdown

End Sub

Sub CustomTimeInputChange

    document.getElementById("TimeSelectOptions").selectedIndex = 0

End Sub

Sub TimeSelectOptionsChange

    document.getElementById("CustomTimeInput").Value = ""

End Sub