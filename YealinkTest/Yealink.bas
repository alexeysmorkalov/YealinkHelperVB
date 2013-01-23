Attribute VB_Name = "Yealink"
Option Explicit

' Declarations
Public Const VERSION_B2K = &H520

Public Const YL_IOCTL_OPEN_DEVICE = &H0
Public Const YL_IOCTL_CLOSE_DEVICE = &HFFFF

Public Const YL_CALLBACK_MSG_WARNING = -2
Public Const YL_CALLBACK_MSG_ERROR = -1
Public Const YL_CALLBACK_MSG_USBPHONE_VERSION = 0
Public Const YL_CALLBACK_MSG_USBPHONE_SERIALNO = 1
Public Const YL_CALLBACK_GEN_KEYBUF_CHANGED = 2
Public Const YL_CALLBACK_GEN_KEYDOWN = 3
Public Const YL_CALLBACK_GEN_OFFHOOK = 4
Public Const YL_CALLBACK_GEN_HANGUP = 5
Public Const YL_CALLBACK_GEN_INUSB = 6
Public Const YL_CALLBACK_GEN_INPSTN = 7
Public Const YL_CALLBACK_GEN_PSTNRING_START = 8
Public Const YL_CALLBACK_GEN_PSTNRING_STOP = 9
Public Const YL_CALLBACK_GEN_CALLERID = 10

Public Const YL_IOCTL_GEN_READY = &H1001
Public Const YL_IOCTL_GEN_UNREADY = &H1002
Public Const YL_IOCTL_GEN_CALLIN = &H1003
Public Const YL_IOCTL_GEN_CALLOUT = &H1004
Public Const YL_IOCTL_GEN_TALKING = &H1006

Public Const YL_IOCTL_GEN_GOTOUSB = &H1010
Public Const YL_IOCTL_GEN_GOTOPSTN = &H1011
Public Const YL_IOCTL_GEN_DEFAULTPSTN = &H1012
Public Const YL_IOCTL_GEN_DEFAULTUSB = &H1013
Public Const YL_IOCTL_GEN_ONLYUSB = &H1014

Public Const YL_IOCTL_OPEN_SIGNAL = &H107
Public Const YL_IOCTL_CLOSE_SIGNAL = &H207

Public Const KEY_0 = &H80
Public Const KEY_1 = &H81
Public Const KEY_2 = &H82
Public Const KEY_3 = &H83
Public Const KEY_4 = &H84
Public Const KEY_5 = &H85
Public Const KEY_6 = &H86
Public Const KEY_7 = &H87
Public Const KEY_8 = &H88
Public Const KEY_9 = &H89
Public Const KEY_STAR = &H8B
Public Const KEY_POUND = &H8C
Public Const KEY_SEND = &H91

' RETURN VALUE & ERROR CODE
Public Const YL_RETURN_SUCCESS = 0
Public Const YL_RETURN_NO_FOUND_HID = 1
Public Const YL_RETURN_HID_ISOPENED = 2
Public Const YL_RETURN_HID_NO_OPEN = 3
Public Const YL_RETURN_MAP_ERROR = 4
Public Const YL_RETURN_DEV_VERSION_ERROR = 5
Public Const YL_RETURN_HID_COMM_ERROR = 6
Public Const YL_RETURN_COMMAND_INVALID = 7

Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

Public Declare Function YL_DeviceIoControl Lib "YLHelper.dll" Alias "YL_DeviceIoControl_" (ByVal IoControlCode As Long, _
    ByVal InBufferAddr As Long, ByVal InBufferSize As Long, _
    ByVal OutBufferAddr As Long, ByVal OutBufferSize As Long) As Long
    
' Opens the device
Public Declare Function YL_Open Lib "YLHelper.dll" (ByVal IoControlCode As Long) As Long
' Closes the device
Public Declare Function YL_Close Lib "YLHelper.dll" () As Long
' Be Ready (Log In)
Public Declare Function YL_BeReady Lib "YLHelper.dll" () As Long
' Be Not Ready (Log Off)
Public Declare Function YL_BeNotReady Lib "YLHelper.dll" () As Long
' Phone ringin
Public Declare Function YL_Calling Lib "YLHelper.dll" () As Long
' Get Ringing
Public Declare Function YL_GetRinging Lib "YLHelper.dll" () As Long
Public Declare Function YL_Talking Lib "YLHelper.dll" () As Long
Public Declare Function YL_SwitchToUsbChannels Lib "YLHelper.dll" () As Long
Public Declare Function YL_SwitchToPstnChannels Lib "YLHelper.dll" () As Long
Public Declare Function YL_DefaultInUsbChannels Lib "YLHelper.dll" () As Long
Public Declare Function YL_DefaultInPstnChannels Lib "YLHelper.dll" () As Long
Public Declare Function YL_UseUsbChannelsOnly Lib "YLHelper.dll" () As Long

Public Function YL_EvaluateResult(ByVal code As Long) As String
    Select Case code
    Case YL_RETURN_SUCCESS: YL_EvaluateResult = "Success"
    Case YL_RETURN_NO_FOUND_HID: YL_EvaluateResult = "Device not found"
    Case YL_RETURN_HID_ISOPENED: YL_EvaluateResult = "Device is already opened"
    Case YL_RETURN_HID_NO_OPEN: YL_EvaluateResult = "Device not opened"
    Case YL_RETURN_MAP_ERROR: YL_EvaluateResult = "Map error"
    Case YL_RETURN_DEV_VERSION_ERROR: YL_EvaluateResult = "Device version error"
    Case YL_RETURN_HID_COMM_ERROR: YL_EvaluateResult = "Device communication error"
    Case YL_RETURN_COMMAND_INVALID: YL_EvaluateResult = "Invalid command"
    Case Else: YL_EvaluateResult = "Unknown answer"
    End Select
End Function

Public Function YL_EvaluatePhoneVersion(ByVal code As Long) As String
    If code >= &H100 And code <= &H199 Then
        YL_EvaluatePhoneVersion = "USB-P1K"
    ElseIf code >= &H200 And code <= &H299 Then
        YL_EvaluatePhoneVersion = "USB-P2K or USB-P3K"
    ElseIf code >= &H300 And code <= &H399 Then
        YL_EvaluatePhoneVersion = "USB-V1K"
    ElseIf code >= &H500 And code <= &H519 Then
        YL_EvaluatePhoneVersion = "USB-B1K"
    ElseIf code >= &H520 And code <= &H539 Then
        YL_EvaluatePhoneVersion = "USB-B2K"
    ElseIf code >= &H600 And code <= &H699 Then
        YL_EvaluatePhoneVersion = "USB-T1K"
    Else
        YL_EvaluatePhoneVersion = "Unknown"
    End If
End Function

Public Function EvaluateKey(ByVal key As Long)
    Select Case key
    Case KEY_0: EvaluateKey = "0"
    Case KEY_1: EvaluateKey = "1"
    Case KEY_2: EvaluateKey = "2"
    Case KEY_3: EvaluateKey = "3"
    Case KEY_4: EvaluateKey = "4"
    Case KEY_5: EvaluateKey = "5"
    Case KEY_6: EvaluateKey = "6"
    Case KEY_7: EvaluateKey = "7"
    Case KEY_8: EvaluateKey = "8"
    Case KEY_9: EvaluateKey = "9"
    Case KEY_STAR: EvaluateKey = "*"
    Case KEY_POUND: EvaluateKey = "#"
    Case KEY_SEND: EvaluateKey = "Send"
    Case Else: EvaluateKey = "Unknown key"
    End Select
End Function
