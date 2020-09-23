VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form ICQSender 
   Caption         =   "ICQ Sending Demo by Paul O'Brien"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Password_Text 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Enter Your ICQ Password Here"
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Login_Text 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Enter Your ICQ Login Here"
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Send_Command 
      Caption         =   "&Send"
      Height          =   2295
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Message_Text 
      Alignment       =   2  'Center
      Height          =   1125
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Tag             =   "Enter Body Text Here"
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Destination_Text 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Enter the Destination Number Here"
      Top             =   960
      Width           =   3255
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   240
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label DestinationInfo_Label 
      Caption         =   "Note: Enter Destination in International Format, Minus + Sign."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   4455
   End
End
Attribute VB_Name = "ICQSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Send_Command_Click()

Destination = Destination_Text.Text
Message = Message_Text.Text
Login = Login_Text.Text
Password = Password_Text.Text

RetVal = Inet.OpenURL("http://web.icq.com/karma/dologin/1,,,00.html?uService=1&tophone=" & Destination & "&uReturnPath=/sms/thanks&msg=" & Message & "&uLogin=" & Login & "&uPassword=" & Password)

If InStr(RetVal, "Your message has been successfully sent") Then
    junk = MsgBox("Your Message Has Been Sent", vbInformation + vbOKOnly)
Else
    junk = MsgBox("Your Message Could Not Be Sent", vbCritical + vbOKOnly)
End If

End Sub
