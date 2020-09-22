VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  '¥­­±
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  '³æ½u©T©w¤u¨ãµøµ¡
   Caption         =   "G.O.D. of Battle !!"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9630
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   491
   ScaleMode       =   3  '¹³¯À
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.PictureBox Picture3 
      Appearance      =   0  '¥­­±
      BackColor       =   &H00000000&
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9615
      TabIndex        =   2
      Top             =   7080
      Width           =   9615
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Sound Effect"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8160
         TabIndex        =   5
         Top             =   0
         Value           =   1  '®Ö¨ú
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   7200
         TabIndex        =   3
         Text            =   "30"
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H00000000&
         BackStyle       =   0  '³z©ú
         Caption         =   "START THE GAME "
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   230
         TabIndex        =   7
         Top             =   0
         Width           =   4455
      End
      Begin VB.Label Label3 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H00000000&
         BackStyle       =   0  '³z©ú
         Caption         =   "START THE GAME "
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   10
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Time Delay ( For fast computer ) :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   0
         Width           =   9615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8040
      Top             =   4320
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   4800
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   120
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   400
      TabIndex        =   1
      Top             =   -120
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '¥­­±
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   0
      Picture         =   "Form2.frx":232C2
      ScaleHeight     =   240
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4800
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x
Dim Y
Dim h
Dim w
Dim v

Private Sub hh()
Do
Me.Cls
StretchBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcPaint
'StretchBlt Me.hdc, 61, 12, 520, 156, Picture3.hdc, 0, 0, 400, 120, vbSrcAnd
StretchBlt Me.hdc, 61, 12, 520, 156, Picture2.hdc, 0, 0, 400, 120, vbSrcPaint
       x = -w / 2 + Me.ScaleWidth / 2
       Y = -h / 2 + 12 + 78
       If h > 156 Then
       h = h - 10
       Else
       h = 156
       End If
        If w > 520 Then
       w = w - 20
       Else
       w = 520
       End If
       
       'StretchBlt Me.hdc, x, y, w, h, Picture3.hdc, 0, 0, 400, 120, vbSrcAnd
   StretchBlt Me.hdc, x, Y, w, h, Picture2.hdc, 0, 0, 400, 120, vbSrcPaint
   Call delay(30)
       DoEvents
               Loop
           
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlack
End Sub

Private Sub Form_Load()
If v <> 1 Then
Form3.OLE2.AppIsRunning = True
Form3.OLE2.DoVerb
End If
'hh
h = Picture2.ScaleHeight * 3
w = Picture2.ScaleWidth * 3
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlack
End Sub

Private Sub Form_Unload(Cancel As Integer)
If v <> 1 Then End
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlack


End Sub

Private Sub Label2_Click()
v = 1
dtime = Val(Text1.Text)
If Check1.Value = 0 Then
Sound = 1
Form3.OLE2.AppIsRunning = False
Else
Form3.OLE2.AppIsRunning = False
Form3.OLE1.AppIsRunning = True
Form3.OLE1.DoVerb
End If
picbg.Show
Unload Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbRed
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlack
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlack
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlack
End Sub

Private Sub Timer1_Timer()
hh
Timer1.Enabled = False
End Sub
