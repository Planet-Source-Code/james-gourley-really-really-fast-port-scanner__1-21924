VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{CC1E317A-3102-11D1-816E-00A024E95548}#5.0#0"; "VBUPROGRESS.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Scanner"
   ClientHeight    =   5412
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5412
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Maximum Scan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   5184
      TabIndex        =   54
      Top             =   0
      Width           =   2580
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Refresh Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   3876
      TabIndex        =   53
      Top             =   360
      Width           =   1284
   End
   Begin VBUProgressBarControl.VBUProgress VBUProgress1 
      Height          =   288
      Left            =   2760
      Top             =   684
      Width           =   2376
      _ExtentX        =   4191
      _ExtentY        =   508
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":030A
      FillStyle       =   1
      CaptionStyle    =   1
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1752
      ItemData        =   "Form1.frx":0326
      Left            =   5160
      List            =   "Form1.frx":0328
      TabIndex        =   50
      Top             =   288
      Width           =   2616
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   1332
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   384
      Top             =   1332
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   48
      Top             =   1332
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4080
      ItemData        =   "Form1.frx":032A
      Left            =   36
      List            =   "Form1.frx":032C
      TabIndex        =   3
      Top             =   1308
      Width           =   5112
   End
   Begin VB.Frame Frame2 
      Caption         =   "Advanced Status"
      Height          =   3036
      Left            =   5160
      TabIndex        =   34
      Top             =   2064
      Width           =   2628
      Begin VB.Label Label38 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1008
         TabIndex        =   52
         Top             =   1416
         Width           =   1476
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Errors :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   36
         TabIndex        =   51
         Top             =   1428
         Width           =   912
      End
      Begin VB.Label Label36 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1404
         TabIndex        =   48
         Top             =   1992
         Width           =   732
      End
      Begin VB.Label Label35 
         Caption         =   "Active Winsocks :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   108
         TabIndex        =   47
         Top             =   2040
         Width           =   1296
      End
      Begin VB.Label Label23 
         Caption         =   "Ports Found :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   24
         TabIndex        =   46
         Top             =   1152
         Width           =   960
      End
      Begin VB.Label Label22 
         Caption         =   "Left :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   504
         TabIndex        =   45
         Top             =   612
         Width           =   360
      End
      Begin VB.Label Label33 
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   972
         TabIndex        =   44
         Top             =   2628
         Width           =   1596
      End
      Begin VB.Label Label32 
         Caption         =   "End Time :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   156
         Left            =   96
         TabIndex        =   43
         Top             =   2676
         Width           =   840
      End
      Begin VB.Label Label31 
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   42
         Top             =   2304
         Width           =   1608
      End
      Begin VB.Label Label30 
         Caption         =   "Start Time :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   84
         TabIndex        =   41
         Top             =   2328
         Width           =   864
      End
      Begin VB.Label Label29 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1008
         TabIndex        =   40
         Top             =   1140
         Width           =   972
      End
      Begin VB.Label Label28 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   1008
         TabIndex        =   39
         Top             =   828
         Width           =   912
      End
      Begin VB.Label Label27 
         Caption         =   "Timed Out :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   24
         TabIndex        =   38
         Top             =   840
         Width           =   852
      End
      Begin VB.Label Label26 
         Caption         =   "30000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   1008
         TabIndex        =   37
         Top             =   588
         Width           =   972
      End
      Begin VB.Label Label25 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   996
         TabIndex        =   36
         Top             =   348
         Width           =   1032
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Scanned :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   84
         TabIndex        =   35
         Top             =   384
         Width           =   744
      End
   End
   Begin VB.ListBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      ItemData        =   "Form1.frx":032E
      Left            =   2280
      List            =   "Form1.frx":0371
      TabIndex        =   33
      Top             =   996
      Width           =   2868
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4092
      TabIndex        =   22
      Top             =   0
      Width           =   1068
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Scan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3132
      TabIndex        =   21
      Top             =   0
      Width           =   948
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1152
      TabIndex        =   8
      Text            =   "100"
      Top             =   360
      Width           =   444
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   2160
      TabIndex        =   7
      Text            =   "30000"
      Top             =   684
      Width           =   576
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   1284
      TabIndex        =   5
      Text            =   "1"
      Top             =   684
      Width           =   576
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2508
      TabIndex        =   2
      Top             =   0
      Width           =   636
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   972
      TabIndex        =   0
      Top             =   -12
      Width           =   1524
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   24
      TabIndex        =   10
      Top             =   3696
      Width           =   5112
      Begin VB.Label Label20 
         Caption         =   "0"
         Height          =   360
         Left            =   1632
         TabIndex        =   29
         Top             =   1128
         Width           =   1608
      End
      Begin VB.Label Label19 
         Caption         =   "Ports Found :"
         Height          =   336
         Left            =   108
         TabIndex        =   28
         Top             =   1104
         Width           =   1440
      End
      Begin VB.Label Label123 
         Caption         =   "Started Scan At :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   132
         TabIndex        =   27
         Top             =   852
         Width           =   1224
      End
      Begin VB.Label Label16 
         Caption         =   "- - - - - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1380
         TabIndex        =   26
         Top             =   888
         Width           =   876
      End
      Begin VB.Label Label17 
         Caption         =   "Scan Ended At :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   2760
         TabIndex        =   25
         Top             =   888
         Width           =   1236
      End
      Begin VB.Label Label18 
         Caption         =   "- - - - - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   4020
         TabIndex        =   24
         Top             =   900
         Width           =   972
      End
      Begin VB.Label Label15 
         Caption         =   "Errors Reported"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   20
         Top             =   564
         Width           =   2136
      End
      Begin VB.Label Label14 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   144
         TabIndex        =   19
         Top             =   552
         Width           =   252
      End
      Begin VB.Label Label13 
         Caption         =   "Active Winsock Controls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   3708
         TabIndex        =   18
         Top             =   564
         Width           =   1788
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   3408
         TabIndex        =   17
         Top             =   552
         Width           =   276
      End
      Begin VB.Label Label11 
         Caption         =   "Currently"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   2748
         TabIndex        =   16
         Top             =   564
         Width           =   660
      End
      Begin VB.Label Label10 
         Caption         =   "Has Timed Out"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   4188
         TabIndex        =   15
         Top             =   228
         Width           =   1176
      End
      Begin VB.Label Label9 
         Caption         =   "Port Number "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   2736
         TabIndex        =   14
         Top             =   240
         Width           =   948
      End
      Begin VB.Label Label4 
         Caption         =   "Scanning Port Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   72
         TabIndex        =   13
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "- - - - -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1752
         TabIndex        =   12
         Top             =   264
         Width           =   528
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "- - - - -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   3720
         TabIndex        =   11
         Top             =   228
         Width           =   396
      End
   End
   Begin VB.Label Label37 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 0 Ports/sec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   5160
      TabIndex        =   49
      Top             =   5124
      Width           =   2628
   End
   Begin VB.Label Command5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   13.8
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   12
      TabIndex        =   32
      Top             =   996
      Width           =   2232
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5148
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Line Line1 
      X1              =   5136
      X2              =   -60
      Y1              =   648
      Y2              =   648
   End
   Begin VB.Shape Shape2 
      Height          =   180
      Left            =   12
      Top             =   1056
      Width           =   156
   End
   Begin VB.Shape Shape1 
      Height          =   180
      Left            =   12
      Top             =   732
      Width           =   156
   End
   Begin VB.Label Command4 
      BackStyle       =   0  'Transparent
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   13.8
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   24
      TabIndex        =   31
      Top             =   660
      Width           =   2724
   End
   Begin VB.Label Label21 
      Caption         =   "Scan Common Ports :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   264
      TabIndex        =   30
      Top             =   984
      Width           =   1968
   End
   Begin VB.Label Label6 
      Caption         =   "Try To Keep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   12
      TabIndex        =   23
      Top             =   360
      Width           =   1140
   End
   Begin VB.Label Label7 
      Caption         =   "Active Winsock Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1632
      TabIndex        =   9
      Top             =   360
      Width           =   2208
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1860
      TabIndex        =   6
      Top             =   684
      Width           =   372
   End
   Begin VB.Label Label2 
      Caption         =   "Scan From "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   252
      TabIndex        =   4
      Top             =   684
      Width           =   1224
   End
   Begin VB.Label Label1 
      Caption         =   "Address :"
      Height          =   300
      Left            =   36
      TabIndex        =   1
      Top             =   0
      Width           =   888
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Index As Double
Public DelIndex As Double
Public oldSpeed As Long
Public seconds As Long

Public Function GetPort(PortName As String)
If PortName = "SyStat" Then Index = 11
If PortName = "NetStat" Then Index = 15
If PortName = "FTP" Then Index = 21
If PortName = "SSH" Then Index = 22
If PortName = "TelNet" Then Index = 23
If PortName = "STMP" Then Index = 25
If PortName = "Whois" Then Index = 43
If PortName = "Name Server" Then Index = 53
If PortName = "Finger" Then Index = 79
If PortName = "HTTP" Then Index = 80
If PortName = "Pop3" Then Index = 110
If PortName = "Ident" Then Index = 113
If PortName = "Secure HTTP" Then Index = 400
If PortName = "Wingate" Then Index = 1080
If PortName = "SubSeven (1243)" Then Index = 1243
If PortName = "DipStix" Then Index = 2002
If PortName = "NetBus" Then Index = 12345
If PortName = "SubSeven (27374)" Then Index = 27374
If PortName = "Back Orifice" Then Index = 31337
End Function

Private Sub Check1_Click()
Select Case Check1.Value
Case 1
Timer2.Enabled = True
Case 0
Timer2.Enabled = False
End Select
End Sub

Private Sub Command1_Click()
On Error GoTo y
Label31.Caption = Time
Label16.Caption = Time
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
If Command4.Caption = "ü" Then
VBUProgress1.Max = Text3.Text - Text2.Text
seconds = Left(Right(Time, 5), 2)
oldSpeed = Text3.Text
Index = Text2.Text
Index2 = Text3.Text
If Index = 1 Then Else Load Winsock1(Text2.Text)
Label5.Caption = Index
DelIndex = Text2.Text
Timer1.Enabled = True
Else
Load Winsock1(11)
Winsock1(11).Connect Text1.Text, 11
Load Winsock1(15)
Winsock1(15).Connect Text1.Text, 15
Load Winsock1(21)
Winsock1(21).Connect Text1.Text, 21
Load Winsock1(22)
Winsock1(22).Connect Text1.Text, 22
Load Winsock1(23)
Winsock1(23).Connect Text1.Text, 23
Load Winsock1(25)
Winsock1(25).Connect Text1.Text, 25
Load Winsock1(43)
Winsock1(43).Connect Text1.Text, 43
Load Winsock1(53)
Winsock1(53).Connect Text1.Text, 53
Load Winsock1(79)
Winsock1(79).Connect Text1.Text, 79
Load Winsock1(80)
Winsock1(80).Connect Text1.Text, 80
Load Winsock1(110)
Winsock1(110).Connect Text1.Text, 110
Load Winsock1(113)
Winsock1(113).Connect Text1.Text, 113
Load Winsock1(139)
Winsock1(139).Connect Text1.Text, 139
Load Winsock1(400)
Winsock1(400).Connect Text1.Text, 400
Load Winsock1(1080)
Winsock1(1080).Connect Text1.Text, 1080
Load Winsock1(1243)
Winsock1(1243).Connect Text1.Text, 1243
Load Winsock1(2002)
Winsock1(2002).Connect Text1.Text, 2002
Load Winsock1(8080)
Winsock1(8080).Connect Text1.Text, 8080
Load Winsock1(12345)
Winsock1(12345).Connect Text1.Text, 12345
Load Winsock1(27374)
Winsock1(27374).Connect Text1.Text, 27374
Load Winsock1(31337)
Winsock1(31337).Connect Text1.Text, 31337
Label5.Caption = Winsock1.UBound
Label12.Caption = Combo1.ListCount
Timer2.Interval = Text4.Text * 10
Command2_Click
End If
Exit Sub
y:
If Err.Number = 10050 Then
List2.AddItem "Error Connecting To Port " & Winsock1.UBound
Label14.Caption = Label14.Caption + 1
Resume Next
End If
End Sub

Private Sub Command2_Click()
Label37.Caption = " 0 Ports/Sec"
Label18.Caption = Time
Label33.Caption = Time
Timer1.Enabled = False
Winsock1(1).Close
Timer2.Enabled = True
If Command4.Caption = "ü" Then
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Command1.Enabled = True
Command3.Enabled = True
VBUProgress1.Value = 0
End If
End Sub

Private Sub Command3_Click()
a:
If List1.ListCount > 0 Then
List1.RemoveItem 0
GoTo a
Else
GoTo b
End If
b:
If List2.ListCount > 0 Then
List2.RemoveItem 0
GoTo b
Else
GoTo x:
End If
x:
Index = 0
DelIndex = 0
Label31.Caption = "- - - - -"
Label33.Caption = "- - - - -"
Label25.Caption = Text2.Text
Label26.Caption = Text3.Text
Label28.Caption = "0"
Label38.Caption = "0"
If Timer1.Enabled = False Then
Label14.Caption = "0"
Label20.Caption = "0"
Label16.Caption = "- - - - -"
Label18.Caption = "- - - - -"
Label5.Caption = "- - - - -"
Label8.Caption = "- - - - -"
End If
End Sub

Private Sub Command4_Click()
Command4.Caption = "ü"
Command5.Caption = ""
End Sub

Private Sub Command5_Click()
Command5.Caption = "ü"
Command4.Caption = ""
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command6_Click()
If Label36.Caption > 1 Then
Timer2_Timer
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
End If
End Sub

Private Sub Form_Load()
Index = 1
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label22_Click()
Select Case Label22.Caption
Case "ü"
Label22.Caption = ""
Case ""
Label22.Caption = "ü"
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
On Error GoTo y
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command1.Enabled = False
Command3.Enabled = False
If Left(Right(Time, 5), 2) = seconds + 1 Then
Label37.Caption = " " & -1 * (Label26.Caption - oldSpeed) & " Ports/Sec"
oldSpeed = Label26.Caption
seconds = Left(Right(Time, 5), 2)
End If
Label38.Caption = List2.ListCount
Label26.Caption = Text3.Text - Label5.Caption
Label20.Caption = List1.ListCount
Label5.Caption = Index
Label25.Caption = Index
Label8.Caption = DelIndex
Label28.Caption = DelIndex
Label29.Caption = Label20.Caption
Label36.Caption = Label12.Caption
Label33.Caption = Time
If Command4.Caption = "ü" Then
VBUProgress1.Value = Label5.Caption - Text2.Text
Label12.Caption = Index - DelIndex
Winsock1(Index).Connect Text1.Text, Index
If Index = Text3.Text Then Command2_Click
If Label22.Caption = "ü" Then If Index >= Index2 Then Command2_Click
Index = Index + 1
Load Winsock1(Index)

If Label12.Caption >= Text4.Text - 1 Then
Label8.Caption = DelIndex
Unload Winsock1(DelIndex)
DelIndex = DelIndex + 1
If DelIndex >= Text3.Text Then Exit Sub
End If
Exit Sub

Else

End If

y:
If Err.Number = 10050 Then
List2.AddItem "Error Connecting To Port " & Index
If Command4.Caption = "ü" Then
If Index = Text3.Text Then Command2_Click
Index = Index + 1
Load Winsock1(Index)
End If
Label14.Caption = Label14.Caption + 1
End If
Resume Next
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Label36.Caption = "0" Then Timer2.Enabled = False
If Label36.Caption = "1" Then Timer2.Enabled = False
Timer2.Interval = 1
If Left(Right(Time, 5), 2) = seconds + 1 Then
Label37.Caption = " " & -1 * (Label26.Caption - oldSpeed) & " Ports/Sec"
oldSpeed = Label26.Caption
seconds = Left(Right(Time, 5), 2)
End If
Label38.Caption = List2.ListCount
Label26.Caption = Text3.Text - Label5.Caption
Label20.Caption = List1.ListCount
Label5.Caption = Index
Label25.Caption = Index
Label8.Caption = DelIndex
Label28.Caption = DelIndex
Label29.Caption = Label20.Caption
Label36.Caption = Label12.Caption
Label33.Caption = Time
If Timer1.Enabled = False Then
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Command3.Enabled = True
Command1.Enabled = True
End If
If Command4.Caption = "ü" Then
If DelIndex >= Index Then Timer2.Enabled = False
Label12.Caption = Index - DelIndex
Label8.Caption = DelIndex
Unload Winsock1(DelIndex)
DelIndex = DelIndex + 1
Else
Unload Winsock1(11)
Unload Winsock1(15)
Unload Winsock1(21)
Unload Winsock1(22)
Unload Winsock1(23)
Unload Winsock1(25)
Unload Winsock1(43)
Unload Winsock1(53)
Unload Winsock1(79)
Unload Winsock1(80)
Unload Winsock1(110)
Unload Winsock1(113)
Unload Winsock1(139)
Unload Winsock1(400)
Unload Winsock1(1080)
Unload Winsock1(1243)
Unload Winsock1(2002)
Unload Winsock1(8080)
Unload Winsock1(12345)
Unload Winsock1(27374)
Unload Winsock1(31337)
Label12.Caption = "0"
End If
End Sub

Private Sub Winsock1_Connect(Index As Integer)
List1.AddItem "Port Found : " & Winsock1(Index).RemotePort
Winsock1(Index).Close
End Sub

