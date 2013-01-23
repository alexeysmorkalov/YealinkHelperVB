Attribute VB_Name = "CallBacks"
Option Explicit

Public Sub YL_CallBack(ByVal wParam As Long, ByVal lParam As Long, ByVal AppData As Long)
    FormMain.AddLog "* CallBack *"
    Select Case wParam
    Case YL_CALLBACK_MSG_USBPHONE_VERSION: FormMain.AddLog "Phone version: " & YL_EvaluatePhoneVersion(lParam)
    Case YL_CALLBACK_GEN_KEYBUF_CHANGED: FormMain.AddLog "Keyboard buffer changed"
    Case YL_CALLBACK_GEN_HANGUP: FormMain.AddLog "Hang up"
    Case YL_CALLBACK_GEN_OFFHOOK: FormMain.AddLog "Off hook"
    Case YL_CALLBACK_GEN_KEYDOWN: FormMain.AddLog "Keydown: " & EvaluateKey(lParam)
    Case YL_CALLBACK_GEN_KEYBUF_CHANGED: FormMain.AddLog "Keyboard buffer changed"
    Case YL_CALLBACK_MSG_WARNING: FormMain.AddLog "a Warning"
    Case YL_CALLBACK_MSG_ERROR: FormMain.AddLog "an Error"
    Case Else: FormMain.AddLog "other message"
    End Select
    
End Sub

