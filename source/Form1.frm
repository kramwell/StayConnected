VERSION 5.00
Begin VB.Form frmConnection 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3690
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleMode       =   0  'User
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Connect Time"
      Height          =   1815
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton Command4 
         Caption         =   "Hide in Tray"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "^"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblTime 
         Caption         =   "00:00:00"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   3495
      Begin VB.Label lblfff 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Frame Frame4 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   3495
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   4560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   4560
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
      Begin VB.CheckBox chkStayconnected 
         Caption         =   "Stay Connected"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   3495
      Begin VB.Label lblConntime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label3 
      Caption         =   "10"
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by KramWell.com - 10/OCT/2006
'Program to stay online when you are connected to a (dial-up) modem.

Option Explicit

Dim intnewnum As Integer

Dim intseconds As Integer

Const Internet_Autodial_Force_Unattended As Long = 2



Private Declare Function InternetAutodial Lib "wininet.dll" _
        (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Private Declare Function InternetAutodialHangup Lib "wininet.dll" _
        (ByVal dwReserved As Long) As Long
        
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" _
        Alias "InternetGetConnectedStateExA" (ByVal lpdwFlags As Long, _
        ByVal lpszConnectionName As String, ByVal dwNameLen As Long, _
        ByVal dwReserved As Long) As Long
        
Private Const INTERNET_CONNECTION_MODEM = &H1&
Private Const INTERNET_CONNECTION_LAN = &H2&
Private Const INTERNET_CONNECTION_PROXY = &H4&
Private Const INTERNET_CONNECTION_BUSY = &H8&
Private Const INTERNET_RAS_INSTALLED = &H10&
Private Const INTERNET_CONNECTION_OFFLINE = &H20&
Private Const INTERNET_CONNECTION_CONFIGURED = &H40&

Dim intCounter As Integer
Dim counter As Integer

Dim intTime As Integer

Dim intNumTimes As Integer

Dim count1st As Integer

Dim dtedate As Date

'start

Const MAX_TOOLTIP As Integer = 64
Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
Const NIF_TIP = &H4
Const NIM_ADD = &H0
Const NIM_DELETE = &H2
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206

Private Type NOTIFYICONDATA
    cbSize           As Long
    hwnd             As Long
    uID              As Long
    uFlags           As Long
    uCallbackMessage As Long
    hIcon            As Long
    szTip            As String * MAX_TOOLTIP
End Type
Private nfIconData As NOTIFYICONDATA

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long



Private Sub chkStayconnected_Click()
Dim sstatus     As String
Dim sConnection As String



If count1st = 0 Then
count1st = count1st + 1
frmTime.Show
chkStayconnected.Value = 0
frmConnection.Enabled = False


ElseIf count1st > 0 Then
    
    If chkStayconnected.Value = 1 Then
    
        Timer1.Enabled = True   'starts the timer by enabling it

        
        Screen.MousePointer = vbHourglass
                
        Call InternetAutodial(Internet_Autodial_Force_Unattended, 0&)
        DoEvents

        Call fConnectStatus(sConnection, sstatus)
        Screen.MousePointer = vbDefault
        

        ElseIf chkStayconnected.Value = 0 Then
            Timer1.Enabled = False
            counter = 0
            

        
End If
End If
End Sub

Private Sub cmdConnect_Click()
Dim sstatus     As String
Dim sConnection As String

        Screen.MousePointer = vbHourglass
                
        Call InternetAutodial(Internet_Autodial_Force_Unattended, 0&)
        DoEvents

        Call fConnectStatus(sConnection, sstatus)
        Screen.MousePointer = vbDefault
        
dtedate = Format(Now(), "hh:mm:ss")
lblTime.Caption = dtedate

    chkStayconnected.Enabled = True
    cmdConnect.Enabled = False
    
Timer1.Enabled = False


End Sub

Private Sub cmdDisconnect_Click()

Screen.MousePointer = vbHourglass
Call InternetAutodialHangup(0&)
Screen.MousePointer = vbDefault

Timer1.Enabled = False

    chkStayconnected.Enabled = False
    chkStayconnected.Value = 0
    cmdConnect.Enabled = True
    
intNumTimes = 0


Label6.Caption = ""
lblfff.Caption = ""
lblConntime.Caption = ""

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

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Command2.Visible = False
Command3.Visible = True
frmConnection.Height = 3500
End Sub

Private Sub Command3_Click()
Command3.Visible = False
Command2.Visible = True
frmConnection.Height = 2430
End Sub

Private Sub Command4_Click()

With nfIconData
    .hwnd = Me.hwnd
    .uID = Me.Icon
    .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon.Handle
    .szTip = "StayConnected - KramWell.com" & vbNullChar
    .cbSize = Len(nfIconData)
End With
Call Shell_NotifyIcon(NIM_ADD, nfIconData)
'hides it
frmConnection.Hide
End Sub


Private Sub Form_Load()
Dim sConnection As String
Dim sstatus As String

Call fConnectStatus(sConnection, sstatus)
frmConnection.Caption = "  " + sConnection

intNumTimes = 0
intnewnum = 10
End Sub

Private Sub Timer1_Timer()
Dim sstatus     As String
Dim sConnection As String



intnewnum = Label3.Caption

  counter = counter + 1   'we set the counter to count here
  intseconds = counter - intnewnum
  
  Label6.Caption = intseconds & " Seconds until retrying Connection"

  
  If counter = intnewnum Then
  
        Timer2.Enabled = True
  
        Screen.MousePointer = vbHourglass
                        
        Call InternetAutodial(Internet_Autodial_Force_Unattended, 0&)
        DoEvents

        Call fConnectStatus(sConnection, sstatus)
        Screen.MousePointer = vbDefault


Timer2.Enabled = False
intTime = intCounter
intCounter = 0

counter = 0
  
End If

If intTime > 0 Then
intNumTimes = intNumTimes + 1
lblConntime.Caption = "You Have been Reconnected " & intNumTimes & " Times"
    lblfff.Caption = "It took " & intTime & " Seconds to Connect"
intTime = 0

dtedate = Format(Now(), "hh:mm:ss")
lblTime.Caption = dtedate

End If


End Sub

Private Sub Timer2_Timer()

    intCounter = intCounter + 1




End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lMsg As Single
'
' Determine the event that happened to the System Tray icon.
' Left clicking the icon displays a message box.
' Right clicking the icon creates an instance of an object from an
' ActiveX Code component then invokes a method to display a message.
'
lMsg = X / Screen.TwipsPerPixelX
Select Case lMsg
    Case WM_LBUTTONUP
            Show
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)

            
    Case WM_RBUTTONUP
    


End Select
End Sub

