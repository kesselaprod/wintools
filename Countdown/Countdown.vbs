Option Explicit

Dim SubList, mAddNum, mAddTen, mAddHun, mSubNum, mSubTen, mSubHun, ExitCondition

Set mSubNum = GetRef("SubNum")
Set mSubTen = GetRef("SubTen")
Set mSubHun = GetRef("SubHun")

Set mAddNum = GetRef("AddNum")
Set mAddTen = GetRef("AddTen")
Set mAddHun = Getref("AddHun")

Set SubList = CreateObject("System.Collections.ArrayList")

SubList.Add mSubNum
SubList.Add mSubTen
SubList.Add mSubHun

SubList.Add mAddNum
SubList.Add mAddTen
SubList.Add mAddHun

'Sub Pong

 '   MsgBox "Pong"

'End Sub

Sub AppInit

    'window.setTimeout "Pong()", 1000

End Sub

Sub AddNum

    Dim intNum

    intNum = document.getElementById("NumerInput").value

    intNum = intNum + 1

    document.getElementById("NumerInput").value = intNum

End Sub

Sub AddTen

    Dim intNum

    intNum = document.getElementById("NumerInput").value

    intNum = intNum + 10

    document.getElementById("NumerInput").value = intNum

End Sub

Sub AddHun

    Dim intNum

    intNum = document.getElementById("NumerInput").value

    intNum = intNum + 100

    document.getElementById("NumerInput").value = intNum

End Sub

Sub SubHun

    Dim intNum

    intNum = document.getElementById("NumerInput").value

    intNum = intNum - 100

    document.getElementById("NumerInput").value = intNum

End Sub

Sub SubTen

    Dim intNum

    intNum = document.getElementById("NumerInput").value

    intNum = intNum - 10

    document.getElementById("NumerInput").value = intNum

End Sub

Sub SubNum

    Dim intNum

    intNum = document.getElementById("NumerInput").value

    intNum = intNum - 1

    document.getElementById("NumerInput").value = intNum

End Sub

Sub Calc

    Dim pSeconds, pMinutes, pHours, strTimerVal, pSecMod, pMinMod

    pSeconds = document.getElementById("NumerInput").value

    'pMinutes = pSeconds / 60

    pHours = pMinutes / 60

    pSecMod = pSeconds Mod 60

    pMinMod = (pHours - Int(pHours)) * 60

    'strTimerVal = vbNewLine & Int(pHours) & "/" & pHours & " iStd/Std " & Int(pMinutes) & "/" & pMinutes & " iMin/Min " & Int(pSeconds) & "/" & pSeconds &" iSec/Sec"

    'strTimerVal = "iStd: " & Int(pHours) & vbNewLine & "Std: " & pHours & vbNewLine & "iMin: " & Int(pMinutes) & vbNewLine & "Min: " & pMinutes & vbNewLine & "iSec: " & Int(pSeconds) & vbNewLine & "Sec: " & pSeconds & vbNewLine & vbNewLine & "pSecMod: " & pSecMod & vbNewLine & "ipMinMod: " & Int(pMinMod)

    strTimerVal = Int(pHours) & " Std." & vbNewLine & Int(pMinMod) & " Min." & vbNewLine & pSecMod & " Sec." 

    'document.getElementById("timer").innerText = document.getElementById("timer").innerText & strTimerVal  
    document.getElementById("timer").innerText = strTimerVal



End Sub

Sub Reset

    document.getElementById("NumerInput").value = "4800"
    document.getElementById("timer").innerText = ""
    document.getElementById("TimeInput").value = "10"
    document.getElementById("MethodSelect").selectedIndex = 0
    ExitCondition = True

End Sub

Sub GetStarted

    Dim iCounter, oWshShell, SubNum

    Set oWshShell = CreateObject("WScript.Shell")

    iCounter = document.getElementById("TimeInput").value
    SubNum = document.getElementById("MethodSelect").selectedIndex
    ExitCondition = False

    Do While iCounter > 0

        If ExitCondition Then
            Exit Do
        End If

        Eval(SubList.Item(SubNum))
        Calc
        'window.setTimeout "Calc()", 1000 not blocking thread

        oWshShell.Run "timeout /t 1", 0, True

        iCounter = iCounter - 1

    Loop

End Sub