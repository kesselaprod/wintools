Option Explicit

Dim oSapiVoice

Class cSapiVoice

    Private pstrVoiceText
    Private pintVoiceIndex
    Private pSapiObject

    Property Let VoiceText(strVoiceText)
        pstrVoiceText = strVoiceText
    End Property

    Property Get VoiceText
        VoiceText = pstrVoiceText
    End Property

    Property Let VoiceIndex(intVoiceIndex)
        pintVoiceIndex = intVoiceIndex
    End Property

    Sub Hello
        MsgBox "Hello"
    End Sub

    Sub RunTextToSpeech

        Set pSapiObject = CreateObject("SAPI.SpVoice")
        Set pSapiObject.Voice = pSapiObject.GetVoices.Item(pintVoiceIndex)
        pSapiObject.Speak pstrVoiceText

    End Sub

    Private Sub Class_Initialize
        
        pstrVoiceText = ""
        pintVoiceIndex = 0

    End Sub

    Private Sub Class_Terminate   

    End Sub

End Class

Sub AppInit

    Set oSapiVoice = New cSapiVoice

    'MsgBox document

End Sub

Sub eSelectionChanged

    Dim oSelectionIndex

    oSelectionIndex = document.getElementById("ChangeVoiceSelect").selectedIndex

    oSapiVoice.VoiceIndex = oSelectionIndex

    'MsgBox oSelectionIndex

End Sub

Sub ResetForm

    document.getElementById("ChangeVoiceSelect").selectedIndex = 0
    document.getElementById("TextToVoiceArea").Value = ""
    
    oSapiVoice.VoiceIndex = 0
    oSapiVoice.VoiceText = ""

End Sub

Sub SubmitForm

    oSapiVoice.VoiceText = document.getElementById("TextToVoiceArea").Value

    If oSapiVoice.VoiceText <> "" Then

        oSapiVoice.RunTextToSpeech

    Else

        MsgBox "Bitte zuerst Text eingeben."

    End If

End Sub