VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "WINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bounce Gate"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "BounceGate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Socket Control"
      Height          =   1215
      Left            =   3840
      TabIndex        =   8
      Top             =   120
      Width           =   1455
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000000&
         Height          =   255
         Left            =   240
         Top             =   720
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   255
         Left            =   240
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "        Kill"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "      Listen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Text            =   "81"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   "23"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "cyberspace.org"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   5535
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   2520
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2520
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Listen Port:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Return Buffer:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote Bounce Port:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server To Bounce User To:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Winsock1.LocalPort = Text4.Text

Winsock1.Listen
Text1.Text = Text1.Text + vbNewLine + "Listening..."

End Sub

Private Sub Label4_Click()
Winsock1.Close
Form_Load
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label4.BackColor = RGB(32, 32, 32)
Label4.ForeColor = vbWhite

End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackColor = RGB(192, 192, 192)
Label4.ForeColor = vbBlack

End Sub

Private Sub Label5_Click()
Winsock1.Close
Winsock2.Close
Text1.Text = Text1.Text + vbNewLine + "Killed connections!"

End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = RGB(32, 32, 32)
Label5.ForeColor = vbWhite

End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = RGB(192, 192, 192)
Label5.ForeColor = vbBlack


End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1)

End Sub

Private Sub Winsock1_Close()
Winsock1.Close

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
Text1.Text = Text1.Text + vbNewLine + "Connection from: " & Winsock1.RemoteHostIP
Text1.Text = Text1.Text + vbNewLine + "Executing Jump Commands!"
Winsock1.SendData "Gate Initialized!" 'make pretty
Winsock1.SendData vbNewLine
Winsock1.SendData "Locking Gate and preceeding with terminal instructions"
Winsock1.SendData vbNewLine
Winsock1.SendData "Welcome to the Personal Gate System"
Winsock1.SendData "Setting Intergallactic star quardinates to: " & Text2.Text & ":" & Text3.Text ' impress the user
Winsock2.RemoteHost = Text2.Text
Winsock2.RemotePort = Text3.Text
Winsock2.Connect ' connect to jump quardinates


End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errorget1
Dim thedata As String
Winsock1.GetData thedata, vbString
Winsock2.SendData thedata
Text1.Text = Text1.Text + vbNewLine + "Client To Server: " & thedata

errorget1:
'who cares

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo errorerror
Winsock1.SendData Description
Winsock1.Close
Exit Sub
errorerror:
MsgBox Description, vbCritical

End Sub

Private Sub Winsock2_Close()
Winsock2.Close

End Sub

Private Sub Winsock2_Connect()
On Error GoTo errorstoop
Winsock1.SendData vbNewLine

Winsock1.SendData "Connected!"
Winsock1.SendData vbNewLine
errorstoop:
'a
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errorstringy
Dim data As String
Winsock2.GetData data, vbString
Winsock1.SendData data
Text1.Text = Text1.Text + vbNewLine + "Server To Client: " & data
errorstringy:
'a

End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock2.Close

End Sub
