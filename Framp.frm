VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Framp 
   BorderStyle     =   0  'None
   Caption         =   "Framp"
   ClientHeight    =   4215
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4830
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4320
      Top             =   960
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   582
      _Version        =   393216
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4471
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Music Directory"
      TabPicture(0)   =   "Framp.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "File1"
      Tab(0).Control(1)=   "Drive1"
      Tab(0).Control(2)=   "Dir1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Media Player"
      TabPicture(1)   =   "Framp.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Check1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Slider2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "About"
      TabPicture(2)   =   "Framp.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(2)=   "Label2"
      Tab(2).ControlCount=   3
      Begin MSComctlLib.Slider Slider2 
         Height          =   1335
         Left            =   3360
         TabIndex        =   16
         Top             =   840
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   2355
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   10
         Min             =   -2500
         Max             =   0
         TickStyle       =   3
         TextPosition    =   1
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mute"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   -74760
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   -72840
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Created by Faiz Rachiemy"
         Height          =   375
         Left            =   -74760
         TabIndex        =   14
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "@frachiemy"
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "faiz.rachiemy@gmail.com      frachiemy.com"
         Height          =   615
         Left            =   -74760
         TabIndex        =   12
         Top             =   1560
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "_"
      Height          =   255
      Left            =   4320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      Height          =   255
      Left            =   4560
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   615
      Left            =   2280
      TabIndex        =   9
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      _Version        =   393216
      TickStyle       =   3
   End
   Begin VB.Label lblEndTrack 
      Height          =   135
      Left            =   2280
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblEndTrack2 
      Height          =   135
      Left            =   4200
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblFile 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Framp"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "Framp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim File As String
Dim Kode As Boolean
Dim EndTrack As Long
Option Explicit
Dim MoveScreen As Boolean, color As Long, flag As Byte
Dim MousX, MousY, CurrX, CurrY As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Call Mute(True)
Else
    Call Mute(False)
End If
End Sub

Private Sub Command1_Click()
End
End Sub
Private Sub Command2_Click()
Me.WindowState = 1
End Sub

Private Sub Dir1_Change()
 File1.Path = Dir1.Path
 File1.FileName = "*.MP3"
End Sub
Private Sub Drive1_Change()
 On Error GoTo Perangkap
 Dir1.Path = Drive1.Drive
Perangkap:
  Select Case Err
    Case 68
      MsgBox "Can't Access Drive " & Drive1.Drive, vbOKOnly + vbCritical, "Scope Error"
      Drive1.Refresh
    Case 0
      Exit Sub
  End Select
End Sub

Private Sub File1_Click()
  MMControl1.Command = "Close"
  MMControl1.Refresh
  lblFile.Caption = File1.FileName
  Play
End Sub

Private Sub File1_DblClick()
 Play
 Slide
 MMControl1.Command = "Play"
End Sub
Sub Play()
  File = File1.Path & "\" & File1.FileName
  If Mid(File, 3, 1) = "\" And Mid(File, 4, 1) = "\" Then
    File = Left(File1.Path, 3) & File1.FileName
   Else
    File = File1.Path & "\" & File1.FileName
  End If
  MMControl1.FileName = File
  MMControl1.Command = "Open"
  EndTrack = MMControl1.TrackLength
  If EndTrack = 0 Then
    MsgBox "Soory Can't Play in this Application", vbOKOnly + vbCritical, "Scope Error"
     Else
    lblEndTrack2.Caption = EndTrack
  End If
End Sub

Private Sub Form_Load()
  File1.FileName = "*.MP3"
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
 If Kode = True Then Exit Sub
  If MMControl1.TrackLength = MMControl1.Position Then
    If File1.ListCount = File1.ListIndex Then
       MMControl1.Command = "Close"
    Else
      With File1
      .ListIndex = .ListIndex + 1
      End With
      Slide
      MMControl1.Command = "Play"
    End If
 End If
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
 Play
 Slide
End Sub

Private Sub MMControl1_StopClick(Cancel As Integer)
  MMControl1.Refresh
  MMControl1.Command = "Close"
  Kode = True
End Sub
Private Sub Slider2_Click()
    Play.volume = Slider2.Value - 2500
End Sub

Private Sub Slider2_Scroll()
    Play.volume = Slider2.Value - 2500
End Sub

Private Sub Timer1_Timer()
 Slider1.Value = MMControl1.Position
 lblEndTrack.Caption = MMControl1.Position
End Sub
Sub Slide()
  Slider1.Min = 0
  Slider1.Max = Val(MMControl1.TrackLength)
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Me.WindowState <> 2 Then
    GeserJendela Me
    End If
End Sub

Private Sub Picture2_GotFocus()

End Sub
