VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Scrolling System Tray Icon - by Jonathan Roach"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "&Start scrolling"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Timer scrollTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   3480
   End
   Begin VB.TextBox txtMsg 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text you want scrolled..."
      Top             =   2040
      Width           =   3615
   End
   Begin VB.PictureBox modIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   5880
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox srcIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   5520
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sdsupport@gto.net"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   1350
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments, Votes and email are always welcome."
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Main.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Main.frx":0095
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0151
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Scrolling Icon in Systray Demo
'by Jonathan Roach - sdsupport@gto.net
'
'Hope you can make use of this code, feel free to email me
'if you have any questions.
'
'Variables related to the tracking of the
'scrolled text and it's positions
Dim xPos As Integer
Dim yPos As Integer
Dim ScrollStart As Integer
Dim ScrollEnd As Integer
Dim Msg As String

Private Sub SetScroll()
'Sets up the scrolling related information
Msg = txtMsg.Text
modIcon.FontSize = 8
modIcon.FontBold = True
ScrollStart = modIcon.ScaleWidth + modIcon.TextWidth(Msg) - 115
ScrollEnd = 0 - modIcon.TextWidth(Msg)
modIcon.Picture = srcIcon.Picture
xPos = ScrollStart
End Sub

Private Sub cmdBegin_Click()
'Initialize things for the scroll effect
SetScroll
'Enable the scroll timer
scrollTimer.Enabled = True
'Disable this command button
cmdBegin.Enabled = False
End Sub

Private Sub cmdQuit_Click()
'This is to exit, clean up everything
scrollTimer.Enabled = False
Unload Form1
End
End Sub

Private Sub Form_Load()
Me.Show
'Load the srcIcon picturebox with the icon
'from the imagelist control
srcIcon.Picture = ImageList1.ListImages(1).Picture
'Setup the data for our icon in the systray
icoDat.cbSize = Len(icoDat)
icoDat.hWnd = Form1.hWnd
icoDat.uId = vbNull
icoDat.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
icoDat.uCallBackMessage = WM_MOUSEMOVE
icoDat.hIcon = srcIcon.Picture
icoDat.szTip = "Scrolling Taskbar Icon Example" & vbNullChar

'Call the Shell_NotifyIcon function to add the icon to the taskbar
'status area.
Shell_NotifyIcon NIM_ADD, icoDat
End Sub

Private Sub Form_Terminate()
' Remove the tray icon.
Shell_NotifyIcon NIM_DELETE, icoDat
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Remove the tray icon.
Shell_NotifyIcon NIM_DELETE, icoDat
End Sub

Private Sub scrollTimer_Timer()
'This handles the positioning of the
'text on the icon, as well as saving the
'new image to the imagelist control and
're-outputting it to the system tray.
If xPos >= ScrollEnd Then
    modIcon.Cls
    modIcon.CurrentX = xPos
    modIcon.CurrentY = 1
    modIcon.Print Msg
    ImageList1.ListImages.Add 2, , modIcon.Image
    xPos = xPos - 1
    srcIcon.Picture = ImageList1.ListImages(2).ExtractIcon
    icoDat.hIcon = srcIcon.Picture
    Shell_NotifyIcon NIM_MODIFY, icoDat
Else
'The scroll has completed so reset everything
    scrollTimer.Enabled = False
    xPos = ScrollStart
    cmdBegin.Enabled = True
End If
End Sub
