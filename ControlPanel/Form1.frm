VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   18
      Tab             =   11
      TabHeight       =   520
      WordWrap        =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Acessibility"
      TabPicture(0)   =   "Form1.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "btnAcess(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "btnAcess(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnAcess(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnAcess(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnAcess(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Add/Remove Programs"
      TabPicture(1)   =   "Form1.frx":0844
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btnAdrm(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "btnAdrm(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "btnAdrm(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Date/Time"
      TabPicture(2)   =   "Form1.frx":099E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(2)"
      Tab(2).Control(1)=   "Command5"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Display"
      TabPicture(3)   =   "Form1.frx":0AF8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(13)"
      Tab(3).Control(1)=   "btnDisp(0)"
      Tab(3).Control(2)=   "btnDisp(1)"
      Tab(3).Control(3)=   "btnDisp(2)"
      Tab(3).Control(4)=   "btnDisp(3)"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Fonts"
      TabPicture(4)   =   "Form1.frx":0C52
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(6)"
      Tab(4).Control(1)=   "btnMkpf(3)"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Game Controls"
      TabPicture(5)   =   "Form1.frx":0DAC
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command6"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label1(3)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Internet Options"
      TabPicture(6)   =   "Form1.frx":0F06
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "btnIO(5)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "btnIO(4)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "btnIO(3)"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "btnIO(2)"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "btnIO(1)"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "btnIO(0)"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "Label1(17)"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).ControlCount=   7
      TabCaption(7)   =   "Keyboard"
      TabPicture(7)   =   "Form1.frx":1060
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "btnMkpf(1)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Label1(14)"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Modems"
      TabPicture(8)   =   "Form1.frx":11BA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Command8"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "Label1(12)"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Mouse"
      TabPicture(9)   =   "Form1.frx":1554
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "btnMkpf(0)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "Label1(15)"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "Multimedia"
      TabPicture(10)  =   "Form1.frx":16AE
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "btnMul(4)"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).Control(1)=   "btnMul(3)"
      Tab(10).Control(1).Enabled=   0   'False
      Tab(10).Control(2)=   "btnMul(2)"
      Tab(10).Control(2).Enabled=   0   'False
      Tab(10).Control(3)=   "btnMul(1)"
      Tab(10).Control(3).Enabled=   0   'False
      Tab(10).Control(4)=   "btnMul(0)"
      Tab(10).Control(4).Enabled=   0   'False
      Tab(10).Control(5)=   "Label1(11)"
      Tab(10).Control(5).Enabled=   0   'False
      Tab(10).ControlCount=   6
      TabCaption(11)  =   "Network"
      TabPicture(11)  =   "Form1.frx":1808
      Tab(11).ControlEnabled=   -1  'True
      Tab(11).Control(0)=   "Label1(16)"
      Tab(11).Control(0).Enabled=   0   'False
      Tab(11).Control(1)=   "Command1"
      Tab(11).Control(1).Enabled=   0   'False
      Tab(11).ControlCount=   2
      TabCaption(12)  =   "Passwords"
      TabPicture(12)  =   "Form1.frx":1BA2
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Command7"
      Tab(12).Control(1)=   "Label1(7)"
      Tab(12).ControlCount=   2
      TabCaption(13)  =   "Printers"
      TabPicture(13)  =   "Form1.frx":213C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "btnMkpf(2)"
      Tab(13).Control(1)=   "Label1(8)"
      Tab(13).ControlCount=   2
      TabCaption(14)  =   "Regional Settings"
      TabPicture(14)  =   "Form1.frx":2296
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "Label1(9)"
      Tab(14).Control(1)=   "Command9(0)"
      Tab(14).Control(2)=   "Command9(1)"
      Tab(14).Control(3)=   "Command9(2)"
      Tab(14).Control(4)=   "Command9(3)"
      Tab(14).Control(5)=   "Command9(4)"
      Tab(14).ControlCount=   6
      TabCaption(15)  =   "Sounds"
      TabPicture(15)  =   "Form1.frx":23F0
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "Command2"
      Tab(15).Control(0).Enabled=   0   'False
      Tab(15).Control(1)=   "Label1(10)"
      Tab(15).Control(1).Enabled=   0   'False
      Tab(15).ControlCount=   2
      TabCaption(16)  =   "Add New Hardware"
      TabPicture(16)  =   "Form1.frx":254A
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "Command3"
      Tab(16).Control(0).Enabled=   0   'False
      Tab(16).Control(1)=   "Label1(4)"
      Tab(16).Control(1).Enabled=   0   'False
      Tab(16).ControlCount=   2
      TabCaption(17)  =   "Mail and Fax"
      TabPicture(17)  =   "Form1.frx":28E4
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "Command4"
      Tab(17).Control(0).Enabled=   0   'False
      Tab(17).Control(1)=   "Label1(5)"
      Tab(17).Control(1).Enabled=   0   'False
      Tab(17).ControlCount=   2
      Begin VB.CommandButton Command9 
         Caption         =   "&Date"
         Height          =   495
         Index           =   4
         Left            =   -74880
         Picture         =   "Form1.frx":2A3E
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   4560
         Width           =   2175
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Time"
         Height          =   495
         Index           =   3
         Left            =   -74880
         Picture         =   "Form1.frx":2B88
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Currency"
         Height          =   495
         Index           =   2
         Left            =   -74880
         Picture         =   "Form1.frx":2CD2
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Number"
         Height          =   495
         Index           =   1
         Left            =   -74880
         Picture         =   "Form1.frx":2E1C
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Passwords"
         Height          =   495
         Left            =   -74880
         Picture         =   "Form1.frx":2F66
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CommandButton btnMul 
         Caption         =   "&Devices"
         Height          =   495
         Index           =   4
         Left            =   -74880
         Picture         =   "Form1.frx":34F0
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   4560
         Width           =   2055
      End
      Begin VB.CommandButton btnMul 
         Caption         =   "&CD Music"
         Height          =   495
         Index           =   3
         Left            =   -74880
         Picture         =   "Form1.frx":363A
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton btnMul 
         Caption         =   "&MIDI"
         Height          =   495
         Index           =   2
         Left            =   -74880
         Picture         =   "Form1.frx":3784
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton btnMul 
         Caption         =   "&Video"
         Height          =   495
         Index           =   1
         Left            =   -74880
         Picture         =   "Form1.frx":3B0E
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CommandButton btnMul 
         Caption         =   "&Audio"
         Height          =   495
         Index           =   0
         Left            =   -74880
         Picture         =   "Form1.frx":3C58
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CommandButton btnIO 
         Caption         =   "&Advanced"
         Height          =   495
         Index           =   5
         Left            =   -74880
         Picture         =   "Form1.frx":3FE2
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   5160
         Width           =   2295
      End
      Begin VB.CommandButton btnIO 
         Caption         =   "&Programs"
         Height          =   495
         Index           =   4
         Left            =   -74880
         Picture         =   "Form1.frx":412C
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CommandButton btnIO 
         Caption         =   "Conne&ctions"
         Height          =   495
         Index           =   3
         Left            =   -74880
         Picture         =   "Form1.frx":44B6
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3960
         Width           =   2295
      End
      Begin VB.CommandButton btnIO 
         Caption         =   "Cont&ent"
         Height          =   495
         Index           =   2
         Left            =   -74880
         Picture         =   "Form1.frx":4840
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3360
         Width           =   2295
      End
      Begin VB.CommandButton btnIO 
         Caption         =   "&Security"
         Height          =   495
         Index           =   1
         Left            =   -74880
         Picture         =   "Form1.frx":498A
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2760
         Width           =   2295
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Regional Settings"
         Height          =   495
         Index           =   0
         Left            =   -74880
         Picture         =   "Form1.frx":4AD4
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Network Pro&perties"
         Height          =   495
         Left            =   120
         Picture         =   "Form1.frx":4C1E
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Modems Properties"
         Height          =   495
         Left            =   -74880
         Picture         =   "Form1.frx":4FA8
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton btnIO 
         Caption         =   "Ge&neral"
         Height          =   495
         Index           =   0
         Left            =   -74880
         Picture         =   "Form1.frx":5332
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CommandButton btnMkpf 
         Caption         =   "&Fonts"
         Height          =   495
         Index           =   3
         Left            =   -74880
         Picture         =   "Form1.frx":547C
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton btnMkpf 
         Caption         =   "&Printers"
         Height          =   495
         Index           =   2
         Left            =   -74880
         Picture         =   "Form1.frx":55C6
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton btnMkpf 
         Caption         =   "&Keyboard"
         Height          =   495
         Index           =   1
         Left            =   -74880
         Picture         =   "Form1.frx":5710
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton btnMkpf 
         Caption         =   "&Mouse"
         Height          =   495
         Index           =   0
         Left            =   -74880
         Picture         =   "Form1.frx":585A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton btnDisp 
         Caption         =   "&Settings"
         Height          =   495
         Index           =   3
         Left            =   -74880
         Picture         =   "Form1.frx":59A4
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3960
         Width           =   2415
      End
      Begin VB.CommandButton btnDisp 
         Caption         =   "Ap&pearance"
         Height          =   495
         Index           =   2
         Left            =   -74880
         Picture         =   "Form1.frx":5AEE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CommandButton btnDisp 
         Caption         =   "Screen &Saver"
         Height          =   495
         Index           =   1
         Left            =   -74880
         Picture         =   "Form1.frx":5E78
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton btnDisp 
         Caption         =   "B&ackground Properties"
         Height          =   495
         Index           =   0
         Left            =   -74880
         Picture         =   "Form1.frx":6202
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Joys&tick Options"
         Height          =   495
         Left            =   -74880
         Picture         =   "Form1.frx":658C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Date/Time"
         Height          =   495
         Left            =   -74880
         Picture         =   "Form1.frx":66D6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CommandButton btnAdrm 
         Caption         =   "St&artup Disk"
         Height          =   495
         Index           =   2
         Left            =   -74880
         Picture         =   "Form1.frx":6820
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CommandButton btnAdrm 
         Caption         =   "Windows &Setup"
         Height          =   495
         Index           =   1
         Left            =   -74880
         Picture         =   "Form1.frx":696A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton btnAdrm 
         Caption         =   "In&stall/Uninstall"
         Height          =   495
         Index           =   0
         Left            =   -74880
         Picture         =   "Form1.frx":6AB4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CommandButton btnAcess 
         Caption         =   "Gene&ral"
         Height          =   495
         Index           =   4
         Left            =   -74880
         Picture         =   "Form1.frx":6E3E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4440
         Width           =   2295
      End
      Begin VB.CommandButton btnAcess 
         Caption         =   "Mo&use"
         Height          =   495
         Index           =   3
         Left            =   -74880
         Picture         =   "Form1.frx":7230
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3840
         Width           =   2295
      End
      Begin VB.CommandButton btnAcess 
         Caption         =   "D&isplay"
         Height          =   495
         Index           =   2
         Left            =   -74880
         Picture         =   "Form1.frx":737A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3240
         Width           =   2295
      End
      Begin VB.CommandButton btnAcess 
         Caption         =   "S&ound"
         Height          =   495
         Index           =   1
         Left            =   -74880
         Picture         =   "Form1.frx":74C4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CommandButton btnAcess 
         Caption         =   "Key&board"
         Height          =   495
         Index           =   0
         Left            =   -74880
         Picture         =   "Form1.frx":760E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Mail And &Fax"
         Height          =   495
         Left            =   -74880
         Picture         =   "Form1.frx":7758
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add New &Hardware"
         Height          =   495
         Left            =   -74880
         Picture         =   "Form1.frx":78A2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "S&ounds"
         Height          =   495
         Left            =   -74880
         Picture         =   "Form1.frx":7C2C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Internet Explorer Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1095
         Index           =   17
         Left            =   -72360
         TabIndex        =   43
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Network Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   16
         Left            =   2400
         TabIndex        =   42
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Mouse Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   15
         Left            =   -72600
         TabIndex        =   41
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Keyboard Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   14
         Left            =   -72600
         TabIndex        =   40
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Display Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   13
         Left            =   -72360
         TabIndex        =   39
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Modem Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   12
         Left            =   -72360
         TabIndex        =   38
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Multimedia Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   11
         Left            =   -72360
         TabIndex        =   37
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Sounds Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   10
         Left            =   -72600
         TabIndex        =   36
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Regional Settings Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1215
         Index           =   9
         Left            =   -72600
         TabIndex        =   35
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Printers Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   8
         Left            =   -72600
         TabIndex        =   34
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Passwords Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   7
         Left            =   -72360
         TabIndex        =   33
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Fonts Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   6
         Left            =   -72480
         TabIndex        =   28
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Mail and Fax Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   5
         Left            =   -72360
         TabIndex        =   27
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Add New Hardware."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   4
         Left            =   -72600
         TabIndex        =   26
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Joystick Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   3
         Left            =   -72480
         TabIndex        =   17
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Date/Time  Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   2
         Left            =   -72480
         TabIndex        =   15
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Add/Remove Options Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Index           =   1
         Left            =   -72360
         TabIndex        =   14
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "This Section is for Acessibility Options."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   855
         Index           =   0
         Left            =   -72480
         TabIndex        =   9
         Top             =   2040
         Width           =   3135
      End
   End
   Begin VB.Image Image1 
      Height          =   6510
      Left            =   0
      Picture         =   "Form1.frx":7D76
      Top             =   0
      Width           =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi folks! this is one code that you will really enjoy if you are ia fix like
'I was a few days back when I wanted to call control panel functions and that
'too the appropriate one ... my lil niece (who has just strted VB)  asked me
'if I can call the Display option of Control Panel and then when it opens it
'shows the ScreenSaver Tab .. well I did manage to call the Display Properties
'but never got the ScreenSaver tab to start as default .. I made this code for
'helping those who needs info about Displaying the indicated Windows settings
'dialog.
'Last but not the least this is first code I am dedicating to someone. I dedicate
'this code to my lovely lil niece Sanu as she was the first one to put me thinking
'about the stuff ....

'Please vote for me if you think it comes of any help to you ..
'Thank you for downloading ...
'Biswajyoti Das (askbiswa@hotmail.com)

Private Sub btnAcess_Click(index As Integer)
Call mdAcess(index + 1)
End Sub

Private Sub btnAdrm_Click(index As Integer)
Call mdAddRem(index + 1)
End Sub

Private Sub btnDisp_Click(index As Integer)
Call mdDisp(index)
End Sub

Private Sub btnIO_Click(index As Integer)
Shell " rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,," & index
End Sub

Private Sub btnMkpf_Click(index As Integer)
Call mdmkpf(index)
End Sub

Private Sub btnMul_Click(index As Integer)
Call mdMul(index)
End Sub

Private Sub Command1_Click()
mdNetwork
End Sub

Private Sub Command2_Click()
Call mdSound
End Sub

Private Sub Command3_Click()
Call mdAddNH
End Sub

Private Sub Command4_Click()
Call mdMaFax
End Sub

Private Sub Command5_Click()
Call mdDT
End Sub

Private Sub Command7_Click()
Call mdPass
End Sub

Private Sub Command8_Click()
Call mdModem
End Sub

Private Sub Command9_Click(index As Integer)
Call mdIntl(index)
End Sub

Private Sub Image1_Click()
Shell "start mailto:askbiswa@hotmail.com"
End Sub
