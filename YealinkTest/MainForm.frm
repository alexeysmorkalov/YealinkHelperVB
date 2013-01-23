VERSION 5.00
Begin VB.Form FormMain 
   Caption         =   "Test program for Yealink phone"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUseUSBOnly 
      Caption         =   "Use USB Channels only"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdDefaultInPtsnChannels 
      Caption         =   "Default in PTSN Channels"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton cmdDefaultInUsbChannels 
      Caption         =   "Default in USB Channels"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdSwitchToPstnChannels 
      Caption         =   "Switch to PTSN channels"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdSwitchToUsb 
      Caption         =   "Switch to USB channels"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton cmdTalking 
      Caption         =   "Talking"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdGetRingin 
      Caption         =   "Get Ringin"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdCalling 
      Caption         =   "Phone ringin"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdBeNotReady 
      Caption         =   "Be No Ready (Log out)"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdBeReady 
      Caption         =   "Be Ready (Log In)"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close Device"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Device"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtLog 
      Height          =   5295
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Log"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AddLog(text As String)
    If txtLog.text <> "" Then
        txtLog.text = txtLog.text + vbCrLf
    End If
    txtLog.text = txtLog.text + text
End Sub

Private Sub cmdBeNotReady_Click()
    AddLog "cmdBeNotReady_Click:"
    AddLog YL_EvaluateResult(YL_BeNotReady())
End Sub

Private Sub cmdBeReady_Click()
    AddLog "cmdBeReady_Click:"
    AddLog YL_EvaluateResult(YL_BeReady())
End Sub

Private Sub cmdCalling_Click()
    AddLog "cmdCalling_Click:"
    AddLog YL_EvaluateResult(YL_Calling())
End Sub

Private Sub cmdDefaultInPtsnChannels_Click()
    AddLog "cmdDefaultInPtsnChannels_Click:"
    AddLog YL_EvaluateResult(YL_DefaultInPstnChannels())
End Sub

Private Sub cmdDefaultInUsbChannels_Click()
    AddLog "cmdDefaultInUsbChannels_Click:"
    AddLog YL_EvaluateResult(YL_DefaultInUsbChannels())
End Sub

Private Sub cmdGetRingin_Click()
    AddLog "cmdGetRingin_Click:"
    AddLog YL_EvaluateResult(YL_GetRinging())
End Sub

Private Sub cmdOpen_Click(Index As Integer)
    AddLog "cmdOpen_Click:"
    Dim Res As Long
    Res = YL_Open(AddressOf YL_CallBack)
    AddLog YL_EvaluateResult(Res)
End Sub

Private Sub cmdClose_Click(Index As Integer)
    AddLog "cmdClose_Click:"
    AddLog YL_EvaluateResult(YL_Close())
End Sub

Private Sub cmdSwitchToPstnChannels_Click()
    AddLog "cmdSwitchToPstnChannels_Click:"
    AddLog YL_EvaluateResult(YL_SwitchToPstnChannels())
End Sub

Private Sub cmdSwitchToUsb_Click()
    AddLog "cmdSwitchToUsb_Click:"
    AddLog YL_EvaluateResult(YL_SwitchToUsbChannels())
End Sub

Private Sub cmdTalking_Click()
    AddLog "cmdTalking_Click:"
    AddLog YL_EvaluateResult(YL_Talking())
End Sub

Private Sub cmdUseUSBOnly_Click()
    AddLog "cmdUseUSBOnly_Click:"
    AddLog YL_EvaluateResult(YL_UseUsbChannelsOnly())
End Sub
