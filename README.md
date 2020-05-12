# MSHTA_Tools
Combination of MS HTA, VBScript and JavaScript with a HTML/CSS framework (Bootstrap); some simple tools I made for fun. Uploaded them for backup purposes. By far not the best code but it works. You'd better off using Powershell, classic cmd/bat or some c#/visual basic .net winforms stuff if you plan to deploy windows apps. VBS with WSH/HTA is considered deprecated. For me it's an easy and elegant way to build some simple windows apps and design their GUI using popular frameworks. The interface is held in my native language (German). Get in touch with me if there are any translation issues or questions.

## About

The following apps have been split up into different parts (.hta, .html, .js, .cs, .vbs) for better maintainability:

**TalkToMe** utilizes SAPI.SpVoice with the ability to enter text and change speaker voices.

**ShutdownTimer** schedules a shutdown within a given time span. It includes a countdown as well.

---

The next apps demonstrates a typical (nearly) 'all code in one file' case which (nevertheless) loads .vbs externally:

**Countdown** should actually be renamed to 'Count' because it has the power to count up and down. It was intentionally created for me to learn how to implement counter logic. It contains some dumb maths, please ignore this part.

## Features

My HTA apps are using the Internet Explorer 10 so I set: _(* this is also necessary for frameworks/js to work properly)_

```
<meta http-equiv="X-UA-Compatible" content="IE=10">
```

You can either put all of your code into the .hta file or separate them as I did and then include them with (notice the **application=yes** attribute):

```
<iframe class="embed-responsive-item" src="ShutdownTimer.html" application="yes" frameBorder="0" height="100%" width="100%" style="position:absolute;"></iframe>
```

Please be aware of scripts be sticked to event handlers like:
```
<body onload="vbscript:AppInit()">
```

Due to restrictions for mshta.exe it is impossible to run WScript.Sleep so a timeout could be realized using:
```
oWshShell.Run "timeout /t 1", 0, True
```

Countdown.vbs makes use of the .net ArrayList. This list stores Subprocedures with the exact same index as the select option and gets evaluated in the while loop (which contains an event based exit condition):
```
Set mSubNum = GetRef("SubNum")
Set SubList = CreateObject("System.Collections.ArrayList")
SubList.Add mSubNum
[...]
Do While iCounter > 0

        If ExitCondition Then
            Exit Do
        End If

        Eval(SubList.Item(SubNum))
        Calc
        
        'window.setTimeout "Calc()", 1000 'not in use, not blocking thread, round braces required for vbs sub calling

        ...

    Loop
```

Shutdown.vbs initializes an (overloaded) vbscript class, replaces the american dot with a commata for calc and uses a regex pattern to match "digit,|.digit(1-2)"
```
...
Class cShutdown
...
pCustomInputSanitized = Replace(pCustomInput, ".",",")
...
.Pattern = "^\d{1}((\.|\,)\d{1,2})?$"
...

```

### Last but not least

RDPShadow_Session.vbs is a simple script to remotely pull the admin id with plink for RDP shadow session use and initialize the session. Replace:
***100.100.100.100*** with the remote ip
***user*** with the remote user name
***???*** with the appropriate password for the user

Please note that there has to be a rdpwrapper/terminal server running on the remote machine (with proper policy settings set) and putty to shadow grab and control a Remote Desktop Protocol Session in Windows 10.
