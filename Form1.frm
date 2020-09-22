VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  '³æ½u©T©w¤u¨ãµøµ¡
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   532
   ScaleMode       =   3  '¹³¯À
   ScaleWidth      =   775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '¨t²Î¹w³]­È
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.PictureBox P1 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1395
      Index           =   1
      Left            =   8400
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   91
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   130
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox missile 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   570
      IMEMode         =   2  'Ãö³¬
      Index           =   1
      Left            =   4200
      Picture         =   "Form1.frx":8B9A
      ScaleHeight     =   36
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   189
      TabIndex        =   31
      Top             =   960
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.PictureBox missile 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      IMEMode         =   2  'Ãö³¬
      Index           =   0
      Left            =   4200
      Picture         =   "Form1.frx":DBBC
      ScaleHeight     =   39
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   189
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.PictureBox hpp 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   390
      IMEMode         =   2  'Ãö³¬
      Left            =   10680
      Picture         =   "Form1.frx":13286
      ScaleHeight     =   24
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   22
      TabIndex        =   29
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox e1 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1590
      IMEMode         =   2  'Ãö³¬
      Index           =   3
      Left            =   5640
      Picture         =   "Form1.frx":13928
      ScaleHeight     =   104
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   155
      TabIndex        =   28
      Top             =   3120
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.PictureBox e1 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   630
      IMEMode         =   2  'Ãö³¬
      Index           =   2
      Left            =   6120
      Picture         =   "Form1.frx":1F78A
      ScaleHeight     =   40
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   50
      TabIndex        =   27
      Top             =   2280
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox p2 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   285
      IMEMode         =   2  'Ãö³¬
      Left            =   4920
      Picture         =   "Form1.frx":20F8C
      ScaleHeight     =   17
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   37
      TabIndex        =   26
      Top             =   2760
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox e1 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1365
      IMEMode         =   2  'Ãö³¬
      Index           =   1
      Left            =   3840
      Picture         =   "Form1.frx":2173E
      ScaleHeight     =   89
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   101
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   210
      IMEMode         =   2  'Ãö³¬
      Index           =   0
      Left            =   4920
      Picture         =   "Form1.frx":28130
      ScaleHeight     =   12
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   58
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      IMEMode         =   2  'Ãö³¬
      Index           =   1
      Left            =   5880
      Picture         =   "Form1.frx":289B2
      ScaleHeight     =   15
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   44
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1530
      Index           =   2
      Left            =   1440
      Picture         =   "Form1.frx":291B0
      ScaleHeight     =   100
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   97
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox Laser 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   705
      IMEMode         =   2  'Ãö³¬
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":30402
      ScaleHeight     =   45
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   183
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox e1 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1170
      IMEMode         =   2  'Ãö³¬
      Index           =   0
      Left            =   9000
      Picture         =   "Form1.frx":3654C
      ScaleHeight     =   76
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   92
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox Boom 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      IMEMode         =   2  'Ãö³¬
      Index           =   5
      Left            =   3960
      Picture         =   "Form1.frx":3B77E
      ScaleHeight     =   31
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   30
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Boom 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   450
      IMEMode         =   2  'Ãö³¬
      Index           =   4
      Left            =   3840
      Picture         =   "Form1.frx":3C2E4
      ScaleHeight     =   28
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   29
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Boom 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   420
      IMEMode         =   2  'Ãö³¬
      Index           =   3
      Left            =   3120
      Picture         =   "Form1.frx":3CCC6
      ScaleHeight     =   26
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   29
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Boom 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      IMEMode         =   2  'Ãö³¬
      Index           =   2
      Left            =   3120
      Picture         =   "Form1.frx":3D5F8
      ScaleHeight     =   23
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   26
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Boom 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   300
      IMEMode         =   2  'Ãö³¬
      Index           =   1
      Left            =   3120
      Picture         =   "Form1.frx":3DD6A
      ScaleHeight     =   18
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   18
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Boom 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   120
      IMEMode         =   2  'Ãö³¬
      Index           =   0
      Left            =   3120
      Picture         =   "Form1.frx":3E19C
      ScaleHeight     =   6
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   6
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox Sword 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   165
      IMEMode         =   2  'Ãö³¬
      Index           =   10
      Left            =   1680
      Picture         =   "Form1.frx":3E256
      ScaleHeight     =   9
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   72
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.PictureBox Sword 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   465
      IMEMode         =   2  'Ãö³¬
      Index           =   9
      Left            =   1320
      Picture         =   "Form1.frx":3EA30
      ScaleHeight     =   29
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   70
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox Sword 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   690
      IMEMode         =   2  'Ãö³¬
      Index           =   8
      Left            =   1200
      Picture         =   "Form1.frx":40276
      ScaleHeight     =   44
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   63
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Sword 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   645
      IMEMode         =   2  'Ãö³¬
      Index           =   7
      Left            =   0
      Picture         =   "Form1.frx":423B8
      ScaleHeight     =   41
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   72
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.PictureBox Sword 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   585
      IMEMode         =   2  'Ãö³¬
      Index           =   6
      Left            =   0
      Picture         =   "Form1.frx":44692
      ScaleHeight     =   37
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   74
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Sword 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      IMEMode         =   2  'Ãö³¬
      Index           =   5
      Left            =   0
      Picture         =   "Form1.frx":46734
      ScaleHeight     =   39
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   77
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox Sword 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   555
      IMEMode         =   2  'Ãö³¬
      Index           =   4
      Left            =   120
      Picture         =   "Form1.frx":48ACE
      ScaleHeight     =   35
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   53
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox Sword 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   165
      IMEMode         =   2  'Ãö³¬
      Index           =   3
      Left            =   120
      Picture         =   "Form1.frx":4A0F0
      ScaleHeight     =   9
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   27
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox Sword 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   240
      IMEMode         =   2  'Ãö³¬
      Index           =   2
      Left            =   120
      Picture         =   "Form1.frx":4A426
      ScaleHeight     =   14
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   28
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Sword 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   510
      IMEMode         =   2  'Ãö³¬
      Index           =   1
      Left            =   120
      Picture         =   "Form1.frx":4A900
      ScaleHeight     =   32
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   34
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Sword 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   300
      IMEMode         =   2  'Ãö³¬
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":4B642
      ScaleHeight     =   18
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   46
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox S1 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   135
      IMEMode         =   2  'Ãö³¬
      Index           =   0
      Left            =   4680
      Picture         =   "Form1.frx":4C05C
      ScaleHeight     =   7
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   65
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox ghj 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1530
      IMEMode         =   2  'Ãö³¬
      Left            =   2640
      Picture         =   "Form1.frx":4C5FA
      ScaleHeight     =   100
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   97
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1530
      Index           =   0
      Left            =   7440
      Picture         =   "Form1.frx":5384C
      ScaleHeight     =   100
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   97
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

