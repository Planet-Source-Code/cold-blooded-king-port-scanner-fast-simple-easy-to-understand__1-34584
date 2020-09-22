VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Port Scanner"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "Scan Status Bar"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2880
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "60000"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "1"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Long

Private Sub Command1_Click()
For x = Text2.Text To Text3.Text
Winsock1.Close
Winsock1.Connect Text1.Text, x
Text4.Text = "Scanning Port: " & x
DoEvents
Next x
End Sub

Private Sub Command2_Click()
Dim stop_port
stop_port = x
x = Text3.Text
Text4.Text = "Stop at Port: " & stop_port
End Sub

Private Sub Form_Load()
Text1.Text = Winsock1.LocalIP
End Sub

Private Sub Winsock1_Connect()
List1.AddItem Winsock1.RemotePort & " is Opening"
End Sub

