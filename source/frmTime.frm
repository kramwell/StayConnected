VERSION 5.00
Begin VB.Form frmTime 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   Icon            =   "frmTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
      Begin VB.Label Label4 
         Caption         =   "The Connection to Recheck is "
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   2055
      Begin VB.CommandButton Command2 
         Caption         =   "Activate"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1575
         Begin VB.CommandButton cmdGo 
            Caption         =   ".:Go:."
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtRecheck 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Seconds"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   1335
         End
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Click Hide to Keep"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Default is Every 10 Seconds"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Enter a number (in seconds) to Recheck the Connection"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "This Message Box will only appear once"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intnewnum As Integer

Const Internet_Autodial_Force_Unattended As Long = 2



Private Declare Function InternetAutodial Lib "wininet.dll" _
        (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Private Declare Function InternetAutodialHangup Lib "wininet.dll" _
        (ByVal dwReserved As Long) As Long
        
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" _
        Alias "InternetGetConnectedStateExA" (ByVal lpdwFlags As Long, _
        ByVal lpszConnectionName As String, ByVal dwNameLen As Long, _
        ByVal dwReserved As Long) As Long
        
        

Private Sub cmdGo_Click()

intnewnum = txtRecheck.Text

frmConnection.Label3.Caption = intnewnum
Hide
frmConnection.Enabled = True
End Sub

Private Sub Command1_Click()
Hide
frmConnection.Enabled = True
End Sub

Private Function fConnectStatus(sConnection As String, sstatus As String) As Boolean
'
' Determine if currently connected to the internet.
'
' Although the recommended method, doesn't seem reliable.
'
Dim sName  As String
Dim lFlags As Long
Dim l      As Long
Dim iPos   As Integer

sName = Space$(513)
l = InternetGetConnectedStateEx(lFlags, sName, 512, 0&)
iPos = InStr(sName, vbNullChar)
If iPos > 0 Then
    sConnection = Left$(sName, iPos - 1)
ElseIf Not sName = String$(513, 0) Then
    sConnection = sName
End If
fConnectStatus = (l = 1)

End Function

Private Sub Command2_Click()
txtRecheck.Enabled = True
cmdGo.Enabled = True
End Sub

Private Sub Form_Load()
Dim sConnection As String
Dim sstatus As String

        Call fConnectStatus(sConnection, sstatus)

 Label3.Caption = sConnection
End Sub

Private Sub txtRecheck_GotFocus()
cmdGo.Default = True
End Sub
