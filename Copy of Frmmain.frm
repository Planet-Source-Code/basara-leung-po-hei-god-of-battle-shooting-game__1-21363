VERSION 5.00
Begin VB.Form picbg 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "G.O.D. of Battle!!            (http://www.gamemagichk.com)"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   Icon            =   "Copy of Frmmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Copy of Frmmain.frx":08CA
   ScaleHeight     =   493
   ScaleMode       =   3  '¹³¯À
   ScaleWidth      =   644
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7395
      ScaleWidth      =   9675
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.Label Label8 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "http://www.gamemagichk.com"
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   5880
         Width           =   9615
      End
      Begin VB.Label Label7 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "It is just the beginning , God.Of.Battle. Full Verion is coming soon !!"
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   5520
         Width           =   9615
      End
      Begin VB.OLE OLE3 
         Class           =   "mplayer"
         Height          =   1335
         Left            =   0
         OleObjectBlob   =   "Copy of Frmmain.frx":0923
         TabIndex        =   17
         Top             =   -9999
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OLE OLE4 
         Class           =   "mplayer"
         Height          =   1335
         Left            =   2880
         OleObjectBlob   =   "Copy of Frmmain.frx":3B3B
         TabIndex        =   16
         Top             =   -9999
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "GAME OVER"
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
            Size            =   72
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   3255
         Index           =   1
         Left            =   -120
         TabIndex        =   15
         Top             =   840
         Width           =   9615
      End
      Begin VB.Label Label6 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "GAME OVER"
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
            Size            =   72
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3255
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   960
         Width           =   9615
      End
      Begin VB.Image Image8 
         Height          =   7305
         Left            =   0
         Picture         =   "Copy of Frmmain.frx":6D53
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9615
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   255
      Index           =   1
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9735
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   1
         Left            =   6240
         TabIndex        =   12
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   1
         Left            =   5880
         Picture         =   "Copy of Frmmain.frx":7376D
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  '³z©ú
         ForeColor       =   &H00C0C0C0&
         Height          =   855
         Index           =   1
         Left            =   4200
         TabIndex        =   11
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   1
         Left            =   9120
         TabIndex        =   10
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   1
         Left            =   8040
         TabIndex        =   9
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   1
         Left            =   7080
         TabIndex        =   8
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   255
         Index           =   1
         Left            =   6720
         Picture         =   "Copy of Frmmain.frx":73E0F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   135
         Index           =   1
         Left            =   8520
         Picture         =   "Copy of Frmmain.frx":745C1
         Stretch         =   -1  'True
         Top             =   0
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   1
         Left            =   7560
         Picture         =   "Copy of Frmmain.frx":7A70B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Shape whp 
         BackStyle       =   1  '¤£³z©ú
         Height          =   120
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   255
      Index           =   0
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9735
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   9735
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   0
         Left            =   6240
         TabIndex        =   6
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image4 
         Height          =   255
         Index           =   0
         Left            =   5880
         Picture         =   "Copy of Frmmain.frx":7F72D
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  '³z©ú
         ForeColor       =   &H00C0C0C0&
         Height          =   855
         Index           =   0
         Left            =   4200
         TabIndex        =   5
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   0
         Left            =   9120
         TabIndex        =   4
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   0
         Left            =   8040
         TabIndex        =   3
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   0
         Left            =   7080
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   255
         Index           =   0
         Left            =   6720
         Picture         =   "Copy of Frmmain.frx":7FDCF
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   135
         Index           =   0
         Left            =   8520
         Picture         =   "Copy of Frmmain.frx":80581
         Stretch         =   -1  'True
         Top             =   0
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   0
         Left            =   7560
         Picture         =   "Copy of Frmmain.frx":866CB
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Shape whp 
         BackStyle       =   1  '¤£³z©ú
         Height          =   120
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox x 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   7455
      Left            =   9600
      Picture         =   "Copy of Frmmain.frx":8B6ED
      ScaleHeight     =   493
      ScaleMode       =   3  '¹³¯À
      ScaleWidth      =   660
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   9960
   End
   Begin VB.Timer tmrTricker 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9960
      Top             =   360
   End
End
Attribute VB_Name = "picbg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Stage
Dim Fun(0 To 1)
Dim Las(0 To 1)
Dim Mis(0 To 1)

Dim Vol
Dim Boss
Dim Exp
Dim Hhp
Dim Hhpx
Dim Hhpy
Dim Hhpp(0 To 1)
'1P
Dim P1X(0 To 1)
Dim P1Y(0 To 1)
Dim P1hp(0 To 1)
'2p
Dim P2X(0 To 2)
Dim P2Y(0 To 2)
Dim P2hp(0 To 2)
Dim P2s
'enemy
Dim E1X(0 To 50)
Dim E1Y(0 To 50)
Dim E1S(0 To 50)
Dim E1hp(0 To 50)
Dim E1p(0 To 50)
'e shot
Dim E2X(0 To 50)
Dim E2Y(0 To 50)
Dim E2S(0 To 50)
Dim E2V(0 To 50)
'move
Dim P1mX(0 To 1)
Dim P1mY(0 To 1)
'att
Dim P1s(0 To 60)
Dim P1sY(0 To 60)
Dim P1sX(0 To 60)
Dim P1sV(0 To 60)
'laser
Dim L1sY(0 To 1)
Dim L1sX(0 To 1)
Dim L1sV(0 To 1)
'var
Dim Ee
Dim EeW(0 To 50)
Dim EeH(0 To 50)
Dim nNn
Dim nN
Dim nNnn
Dim mM(0 To 1)
Dim oO
Dim mMm(0 To 50)
'Sword
Dim s1X(0 To 1)
Dim s1V(0 To 1)
Dim s1Y(0 To 1)
'missile
Dim Ms1X(0 To 11)
Dim Ms1V(0 To 11)
Dim Ms1Y(0 To 11)
Dim Ms1S(0 To 11)
'boom
Dim Bs1X(0 To 50)
Dim Bs1V(0 To 50)
Dim Bs1Y(0 To 50)
'boom2
Dim Bs2X(0 To 50)
Dim Bs2V(0 To 50)
Dim Bs2Y(0 To 50)
'star
Dim Stars1(60, 2) As Long



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyUp
  P1mY(0) = 0
  Case vbKeyDown
  P1mY(0) = 0
  Case vbKeyLeft
  P1mX(0) = 0
  Case vbKeyRight
  P1mX(0) = 0
  Case vbKeyNumpad1
  P1s(0) = 0
  Case vbKeyW
  P1mY(1) = 0
  Case vbKeyS
  P1mY(1) = 0
  Case vbKeyA
  P1mX(1) = 0
  Case vbKeyD
  P1mX(1) = 0
  Case vbKeyG
  P1s(1) = 0

          End Select
End Sub
Private Sub Form_Keydown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyUp
  If P1hp(0) > 0 Then P1mY(0) = 1
  Case vbKeyW
  If P1hp(1) > 0 Then P1mY(1) = 1
  Case vbKeyDown
  If P1hp(0) > 0 Then P1mY(0) = 2
      Case vbKeyS
    If P1hp(1) > 0 Then P1mY(1) = 2
  Case vbKeyLeft
    If P1hp(0) > 0 Then P1mX(0) = 1
      Case vbKeyA
    If P1hp(1) > 0 Then P1mX(1) = 1
  Case vbKeyRight
If P1hp(0) > 0 Then P1mX(0) = 2
  Case vbKeyD
If P1hp(1) > 0 Then P1mX(1) = 2
  Case vbKeyNumpad6
  If Hhpp(0) >= 1 And P1hp(0) > 0 Then
   If Sound = 0 Then sndPlaySound App.Path & "\data\ut4.mid", SND_ASYNC
  Hhpp(0) = Hhpp(0) - 1
  P1hp(0) = P1hp(0) + 30
     If P1hp(0) > 200 Then P1hp(0) = 200
  End If
  Case vbKeyU
  If Hhpp(1) >= 1 And P1hp(1) > 0 Then
   If Sound = 0 Then sndPlaySound App.Path & "\data\ut4.mid", SND_ASYNC
  Hhpp(1) = Hhpp(1) - 1
  P1hp(1) = P1hp(1) + 30
     If P1hp(1) > 200 Then P1hp(1) = 200
  End If
  Case vbKeyNumpad1
  If s1V(0) = 0 Then P1s(0) = 1
  Case vbKeyG
  If s1V(1) = 0 Then P1s(1) = 1
  Case vbKeyNumpad2
 If s1V(0) <> 1 Then
 s1V(0) = 1
  If Sound = 0 Then sndPlaySound App.Path & "\data\ut6.mid", SND_ASYNC
  End If
  Case vbKeyH
 If s1V(1) < 1 Then
 s1V(1) = 1
  If Sound = 0 Then sndPlaySound App.Path & "\data\ut6.mid", SND_ASYNC
  End If
  Case vbKeyNumpad5
If L1sV(0) = 0 And Las(0) >= 1 Then
 If Sound = 0 Then sndPlaySound App.Path & "\data\ut2.mid", SND_ASYNC
L1sX(0) = P1X(0) + 85
L1sY(0) = P1Y(0) + 63
L1sV(0) = 1

  Las(0) = Las(0) - 1
End If
  Case vbKeyF12
  P1hp(1) = 100
  P1hp(0) = 100
  Case vbKeyF1
  
  Las(0) = 99
  Mis(0) = 99
  Fun(0) = 99
  Hhpp(0) = 99
  Case vbKeyF2
  P1hp(0) = 0
    Case vbKeyF3
    Exp = 9999
  Case vbKeyNumpad4
  If Fun(0) >= 1 Then
  If P2hp(1) <= 0 Then
P2X(1) = P1X(0) + 85
P2Y(1) = P1Y(0) + 63
P2hp(1) = 5
Fun(0) = Fun(0) - 1
Else
  If P2hp(0) <= 0 Then
P2X(0) = P1X(0) + 85
P2Y(0) = P1Y(0) + 63
P2hp(0) = 5
Fun(0) = Fun(0) - 1
End If
End If
End If
  Case vbKeyNumpad3
  If Ms1V(0) <= 0 And Mis(0) >= 1 Then
   If Sound = 0 Then sndPlaySound App.Path & "\data\ut7.mid", SND_ASYNC
  Ms1V(0) = 10
  Mis(0) = Mis(0) - 1

  For i = 0 To 9
  Ms1S(i) = Int(Rnd * 41 + 10)
  Ms1Y(i) = 40 + i * 40
  Ms1X(i) = -200
  Next
  End If
          End Select
End Sub
Private Sub Form_Load()
Hhpp(0) = 1
Hhpp(1) = 1
Ee = 3
Exp = 0
Stage = 1
P1hp(0) = 200
Las(1) = 1
Mis(1) = 1
Fun(1) = 1
Las(0) = 1
Mis(0) = 1
Fun(0) = 1
P1sV(0) = 1
Kof2 = Kof1 + x.ScaleWidth
P1X(0) = 50
P1Y(0) = 168
        For n = 1 To 60
        Stars1(n, 1) = Int(Rnd * picbg.ScaleWidth + Screen.TwipsPerPixelX)
        Stars1(n, 2) = Int(Rnd * x.ScaleHeight + Screen.TwipsPerPixelY)
         Ms1X(i) = 9999
    Next
tmrTricker.Enabled = True
End Sub
Sub enemy()
For i = 0 To Ee
Randomize
If E1hp(3) <= 0 And Vol > 15 And Stage = 1 And E1p(3) = 3 Then Exit Sub
If E1hp(i) <= 0 Then

E1Y(i) = Rnd * x.ScaleHeight
E1X(i) = Me.ScaleWidth
If Stage = 1 And Vol > 15 Then
'Ee = 3

If i <= 2 Then
If E1hp(3) <= 700 And E1S(3) = 4 Then
E1S(i) = 3
E1p(i) = 2
E1hp(i) = 30
End If
Else
If i = 3 Then

E1p(3) = 3
E1S(3) = 4
E1hp(3) = 1400
E2S(3) = 0
End If
End If
End If
If Stage = 1 And Vol <= 15 Then
Randomize
If Vol > 10 Then Ee = 5
If Rnd > 0.5 Then
If Vol > 10 And Rnd > 0.6 Then
E1S(i) = 3
E1p(i) = 2
E1hp(i) = 30
Else
E1S(i) = 0
E1p(i) = 0
E1hp(i) = 70
End If
Else

If Rnd < 0.25 Then
E1S(i) = 1
E1hp(i) = 50
Else

E1S(i) = 2
E1hp(i) = 50
End If

E1p(i) = 1

End If




End If
End If
Next


End Sub
Private Sub Game()


 Do

If Me.Visible = False Or Picture2.Visible = True Then Exit Sub
For i = 0 To 1
If P1hp(i) <= 0 And Picture1(i).Visible = True Then Picture1(i).Visible = False
If P1hp(i) > 0 And Picture1(i).Visible = False Then
Picture1(i).Visible = True
whp(i).Visible = True
End If
'hp

If Hhp = 1 Then Hhpx = Hhpx - 10
If Picture1(i).Visible = True Then
If Label5(i).Caption <> " x " & Int(Hhpp(i)) Then Label5(i).Caption = " x " & Int(Hhpp(i))
If Label1(i).Caption <> " x " & Int(Fun(i)) Then Label1(i).Caption = " x " & Int(Fun(i))
If Label2(i).Caption <> " x " & Int(Mis(i)) Then Label2(i).Caption = " x " & Int(Mis(i))
If Label3(i).Caption <> " x " & Int(Las(i)) Then Label3(i).Caption = " x " & Int(Las(i))
If Label4(i).Caption <> " Lv " & (1 + Int(Exp / 500)) & " Exp " & Exp Then Label4(i).Caption = " Lv " & (1 + Int(Exp / 500)) & " Exp " & Exp
If Mis(i) <= 2 Then Mis(i) = Mis(i) + 0.001


If Las(i) <= 2 Then Las(i) = Las(i) + 0.002


If Fun(i) <= 2 Then Fun(i) = Fun(i) + 0.004
End If

Next

For i = 0 To Ee
EeW(i) = Form1.e1(E1p(i)).ScaleWidth
EeH(i) = Form1.e1(E1p(i)).ScaleHeight
Next
' enemy move



For i = 0 To Ee
If E1S(i) = 4 Then
If E1Y(i) >= P1Y(0) Then E1Y(i) = E1Y(i) - 5
If E1Y(i) < P1Y(0) Then E1Y(i) = E1Y(i) + 5
If E1X(i) >= Me.ScaleWidth - EeW(i) Then
E1X(i) = E1X(i) - 5
Else
If Ee <> 3 Then Ee = 3
End If
End If


If E1S(i) = 3 Then
If E1Y(i) > P1Y(0) + 50 Then E1Y(i) = E1Y(i) - 8
If E1Y(i) < P1Y(0) + 50 Then E1Y(i) = E1Y(i) + 8
End If

If E1S(i) = 0 Or E1S(i) = 3 Then
E1X(i) = E1X(i) - 5 * (E1S(i) + 1)
If E1X(i) < 0 - EeW(i) Then E1hp(i) = 0
End If

If E1S(i) = 1 Then
E1X(i) = E1X(i) - 8
If E1X(i) < 0 - EeW(i) Then E1hp(i) = 0
E1Y(i) = E1Y(i) - 8
If E1Y(i) <= 0 Then E1S(i) = 2
End If

If E1S(i) = 2 Then
E1X(i) = E1X(i) - 8
If E1X(i) < 0 - EeW(i) Then E1hp(i) = 0
E1Y(i) = E1Y(i) + 8
If E1Y(i) >= x.ScaleHeight - EeH(i) Then E1S(i) = 1
End If

Next
'Attack
nN = nN + 0.3
If nN > 50 Then nN = 1
If P1s(0) = 1 And s1V(0) <> 1 And P1hp(0) > 0 Then

If P1sV(Int(nN)) = 0 Then
P1sX(Int(nN)) = P1X(0) + 85
P1sY(Int(nN)) = P1Y(0) + 63
P1sV(Int(nN)) = 1
If Sound = 0 Then sndPlaySound App.Path & "\data\ut5.mid", SND_ASYNC Or SND_NOSTOP
End If

End If
oO = oO + 0.3
If oO > 60 Then oO = 51
If oO < 51 Then oO = 51
If P1s(1) = 1 And s1V(1) <> 1 And P1hp(1) > 0 Then

If P1sV(Int(oO)) = 0 Then
P1sX(Int(oO)) = P1X(1) + 85
P1sY(Int(oO)) = P1Y(1) + 53
P1sV(Int(oO)) = 1
If Sound = 0 Then sndPlaySound App.Path & "\data\ut5.mid", SND_ASYNC Or SND_NOSTOP
End If

End If
'e att
For i = 0 To Ee
nNnn = nNnn + 1
If E1S(i) = 0 Or E1S(i) = 3 Or E1S(i) = 4 Then
If E1hp(i) > 0 And E2V(i) = 0 Then
If Sound = 0 Then sndPlaySound App.Path & "\data\ut5.mid", SND_ASYNC Or SND_NOSTOP
E2V(i) = 1
E2X(i) = E1X(i) - 50
E2Y(i) = E1Y(i) + EeH(i) / 2
If E1p(i) = 0 Then E2S(i) = 0
If E1p(i) = 2 Then E2S(i) = 1
If E1S(i) = 4 Then

If nNnn > 50 Then nNnn = 1

E2X(50 - Int(nNnn)) = E1X(i)
E2Y(50 - Int(nNnn)) = E1Y(i) + 50
E2V(50 - Int(nNnn)) = 1


End If
End If

End If
Next
'p2att


For i = 0 To 1
nNn = nNn + 0.1
If nNn >= 2 Then nNn = 1
If nNn = 1 Or nNn = 1.1 Then
If P2hp(i) > 0 Then
If P1sV(50 - Int(nN) - i) = 0 Then
P1sX(50 - Int(nN) - i) = P2X(i)
P1sY(50 - Int(nN) - i) = P2Y(i) + 5
P1sV(50 - Int(nN) - i) = 1
If Sound = 0 Then sndPlaySound App.Path & "\data\ut5.mid", SND_ASYNC Or SND_NOSTOP
End If
End If
End If
Next



'Att Move
For i = 1 To 60
If P1sV(i) = 1 Then P1sX(i) = P1sX(i) + 25

If P1sX(i) >= Me.ScaleWidth Then P1sV(i) = 0
Next
If L1sV(0) = 1 Then L1sX(0) = L1sX(0) + 50
If L1sX(0) > Me.ScaleWidth Then L1sV(0) = 0
'Move
For i = 0 To 1
If P1mY(i) = 1 And P1Y(i) > 0 And P1hp(i) > 0 Then P1Y(i) = P1Y(i) - 10 - Int(Exp / 500) / 2
If P1mY(i) = 2 And P1Y(i) < -Form1.P1(i).ScaleHeight + x.ScaleHeight And P1hp(i) > 0 Then P1Y(i) = P1Y(i) + 10 + Int(Exp / 500) / 2
If P1mX(i) = 1 And P1X(i) > 0 And P1hp(i) > 0 Then P1X(i) = P1X(i) - 10 - Int(Exp / 500) / 2
If P1mX(i) = 2 And P1X(i) < -Form1.P1(i).ScaleWidth + Me.ScaleWidth And P1hp(i) > 0 Then P1X(i) = P1X(i) + 10 + Int(Exp / 500) / 2
    Next
    Randomize
        With picbg
        'Background
                If E1hp(3) <= 0 And Vol > 15 And Stage = 1 And E1p(3) = 3 Or P1hp(0) + P1hp(1) <= 0 Then
        If P1hp(0) > 0 Then
        Boss = 1
        P1X(0) = P1X(0) + 15
                    Bs2V(Int(nN)) = 1
       Bs2X(Int(nN)) = E1X(3) + 10
       Bs2Y(Int(nN)) = E1Y(3) + 10
        Else
               Bs2V(Int(nN)) = 1
       Bs2X(Int(nN)) = P1X(0)
       Bs2Y(Int(nN)) = P1Y(0) + 10
        End If
                Else
                  If Kof1 <= 0 Then
BitBlt .hdc, Kof1 + x.ScaleWidth - 20, 0, -Kof1 + 20, x.ScaleHeight, x.hdc, 0, 0, vbSrcCopy
Else
BitBlt .hdc, Kof1 - x.ScaleWidth + 20, 0, -Kof1 + 2 * x.ScaleWidth - 20, x.ScaleHeight, x.hdc, 0, 0, vbSrcCopy
End If
    BitBlt .hdc, Kof1, 0, -Kof1 + x.ScaleWidth, x.Height, x.hdc, 0, 0, vbSrcCopy
    End If
         'Star
         For n = 1 To 10
        StarsMove n, 0.1 + Boss
    Next
    For n = 11 To 20
        StarsMove n, 0.2 + Boss
    Next
    For n = 21 To 30
        StarsMove n, 0.3 + Boss
    Next
    For n = 31 To 40
        StarsMove n, 0.4
    Next

 If Kof1 + .ScaleWidth < 0 Then
 Kof1 = .ScaleWidth - 20
Vol = Vol + 1
   If P1hp(0) + P1hp(1) <= 0 Then
   Picture2.Visible = True
If Sound = 0 Then
Form3.OLE1.AppIsRunning = False
    OLE3.AppIsRunning = True
  OLE3.DoVerb
     'Me.SetFocus
     End If
   End If
   If E1hp(3) <= 0 And Vol > 15 And Stage = 1 And E1p(3) = 3 Then
   If Sound = 0 Then
     Form3.OLE1.AppIsRunning = False
OLE4.AppIsRunning = True
OLE4.DoVerb
     Me.SetFocus
     End If
   Picture2.Visible = True
  Label6(0).Caption = "You Win !!"
    Label6(1).Caption = "You Win !!"
   End If
End If
enemy
     Kof1 = Kof1 - 10
           'Draw p2
           For i = 0 To 1
     If P2hp(i) > 0 Then
    ' BitBlt .hdc, P2X(i), P2Y(i), Form1.p2.ScaleWidth, Form1.p2.ScaleHeight, Form1.p3.hdc, 0, 0, vbSrcAnd
  BitBlt .hdc, P2X(i), P2Y(i), Form1.p2.ScaleWidth, Form1.p2.ScaleHeight, Form1.p2.hdc, 0, 0, vbSrcPaint
End If
     Next
           'Draw p
           
       ' BitBlt .hdc, P1X(0), P1Y(0), Form1.P1(0).ScaleWidth, Form1.P1(0).ScaleHeight, Form1.P1(1).hdc, 0, 0, vbSrcAnd
        If P1hp(0) > 0 Then
        If P1mX(0) + P1mY(0) > 0 Then
                BitBlt .hdc, P1X(0), P1Y(0), Form1.P1(0).ScaleWidth, Form1.P1(0).ScaleHeight, Form1.P1(0).hdc, 0, 0, vbSrcPaint
                Else
        BitBlt .hdc, P1X(0), P1Y(0), Form1.P1(0).ScaleWidth, Form1.P1(0).ScaleHeight, Form1.P1(2).hdc, 0, 0, vbSrcPaint
     End If
     End If
        If P1hp(1) > 0 Then BitBlt .hdc, P1X(1), P1Y(1), Form1.P1(1).ScaleWidth, Form1.P1(1).ScaleHeight, Form1.P1(1).hdc, 0, 0, vbSrcPaint

     'Draw e
     For i = 0 To Ee
     If E1hp(i) > 0 Then
       ' BitBlt .hdc, E1X(i), E1Y(i), EeW(i), EeH(i), Form1.e2(E1p(i)).hdc, 0, 0, vbSrcAnd
        BitBlt .hdc, E1X(i), E1Y(i), EeW(i), EeH(i), Form1.e1(E1p(i)).hdc, 0, 0, vbSrcPaint

End If
Next
     'draw hp
     If Hhp > 0 Then
       ' BitBlt .hdc, E1X(i), E1Y(i), EeW(i), EeH(i), Form1.e2(E1p(i)).hdc, 0, 0, vbSrcAnd
        BitBlt .hdc, Hhpx, Hhpy, 21, 21, Form1.hpp.hdc, 0, 0, vbSrcPaint

End If
'Draw s
For i = 1 To 60
 If P1sV(i) = 1 Then
      ' BitBlt .hdc, P1sX(i), P1sY(i), Form1.S1(0).ScaleWidth, Form1.S1(0).ScaleHeight, Form1.S1(1).hdc, 0, 0, vbSrcAnd
        BitBlt .hdc, P1sX(i), P1sY(i), Form1.S1(0).ScaleWidth, Form1.S1(0).ScaleHeight, Form1.S1(0).hdc, 0, 0, vbSrcPaint
      End If
      Next
For i = 0 To 50
 If E2V(i) = 1 Then
      '  BitBlt .hdc, E2X(i), E2Y(i), Form1.Picture1(E2S(i)).ScaleWidth, Form1.Picture1(E2S(i)).ScaleHeight, Form1.Picture2(E2S(i)).hdc, 0, 0, vbSrcAnd
        BitBlt .hdc, E2X(i), E2Y(i), Form1.Picture1(E2S(i)).ScaleWidth, Form1.Picture1(E2S(i)).ScaleHeight, Form1.Picture1(E2S(i)).hdc, 0, 0, vbSrcPaint

      End If
              If E2V(i) >= 1 Then E2X(i) = E2X(i) - 25
If -E2X(i) + E1X(i) >= Me.ScaleWidth Then E2V(i) = 0
If E1p(i) = 2 And E2S(i) = 1 And E2X(i) <= 0 Then E2V(i) = 0
      Next
 
 If L1sV(0) = 1 Then
        'StretchBlt .hdc, L1sX(0), L1sY(0), Form1.Laser(1).ScaleWidth * 2, Form1.Laser(1).ScaleHeight, Form1.Laser(1).hdc, 0, 0, Form1.Laser(1).ScaleWidth, Form1.Laser(1).ScaleHeight, vbSrcAnd
        StretchBlt .hdc, L1sX(0), L1sY(0), Form1.Laser(0).ScaleWidth * 2, Form1.Laser(0).ScaleHeight, Form1.Laser(0).hdc, 0, 0, Form1.Laser(0).ScaleWidth, Form1.Laser(0).ScaleHeight, vbSrcPaint
      End If
     'Draw sword
     'sword move
     For i = 0 To 1
If i = 0 Then s1X(i) = P1X(i) + 90
If i = 1 Then s1X(i) = P1X(i) + 110
If mM(i) < 5 Then
s1Y(i) = P1Y(i) + 40
Else
s1Y(i) = P1Y(i) + 60 - Form1.Sword(Int(mM(i))).ScaleHeight / 2
End If
's1Y(0) = P1Y(0) + 40


 If s1V(i) = 1 Then
     '   BitBlt .hdc, s1X(0), s1Y(0), Form1.Sword(Int(mM)).ScaleWidth, Form1.Sword(Int(mM)).ScaleHeight, Form1.Sword2(Int(mM)).hdc, 0, 0, vbSrcAnd
        BitBlt .hdc, s1X(i), s1Y(i), Form1.Sword(Int(mM(i))).ScaleWidth, Form1.Sword(Int(mM(i))).ScaleHeight, Form1.Sword(Int(mM(i))).hdc, 0, 0, vbSrcPaint

   If mM(i) <= 5 Then
   mM(i) = mM(i) + 0.6
   Else
   mM(i) = mM(i) + 0.4
   End If


      End If
   If mM(i) >= 11 Then
  s1V(i) = 0
  mM(i) = 0
  End If
 Next
       'missile
      For i = 0 To 10
 If Ms1S(i) > 0 Then
 Ms1X(i) = Ms1X(i) + Ms1S(i)
       ' BitBlt .hdc, Ms1X(i), Ms1Y(i), Form1.missile(1).ScaleWidth, Form1.missile(1).ScaleHeight, Form1.missile(1).hdc, 0, 0, vbSrcAnd
        BitBlt .hdc, Ms1X(i), Ms1Y(i), Form1.missile(0).ScaleWidth, Form1.missile(0).ScaleHeight, Form1.missile(Ms1S(11)).hdc, 0, 0, vbSrcPaint
 If Ms1S(11) = 1 Then
  Ms1S(11) = 0
  Else
   Ms1S(11) = 1
   End If
End If
      
      If Ms1X(i) >= Me.ScaleWidth And Ms1S(i) > 0 Then
      Ms1V(0) = Ms1V(0) - 1
      Ms1S(i) = 0
      End If
      Next
      
      
' hit
  For i = 0 To Ee
  For n = 0 To 60
       If Abs(E1X(i) + EeW(i) / 2 - P1sX(n) - 30) <= EeW(i) / 2 + 30 And Abs(E1Y(i) + EeH(i) / 2 - P1sY(n) - 5) <= EeH(i) / 2 + 5 And P1sV(n) = 1 And E1hp(i) > 0 Then
                E1hp(i) = E1hp(i) - 5 - Int(Exp / 500) / 2
               If E1hp(i) > 0 Then
       If Bs1V(nN) = 0 Then
       Bs1V(Int(nN)) = 1
       Bs1X(Int(nN)) = P1sX(n) + 50
       Bs1Y(Int(nN)) = P1sY(n) + 5
        P1sV(n) = 0
        Exp = Exp + 1
               End If
       Else
                
       If Bs2V(nN) = 0 Then
       Bs2V(Int(nN)) = 1
       Bs2X(Int(nN)) = E1X(i)
       Bs2Y(Int(nN)) = E1Y(i)
        P1sV(n) = 0
        If Sound = 0 Then sndPlaySound App.Path & "\data\ut3.mid", SND_ASYNC
        Exp = Exp + 10
        If Rnd >= 0.95 And Hhp <= 0 And Hhpp(0) + Hhpp(1) < 5 Then
Randomize

Hhp = 1
Hhpx = Me.ScaleWidth
Hhpy = Rnd * (x.ScaleHeight - 50)

End If
               End If
End If

    End If
      Next
      Next
'Shit
  For i = 0 To Ee
  For n = 0 To 1
       If Abs(E1X(i) + EeW(i) / 2 - s1X(n)) <= EeW(i) And Abs(E1Y(i) + EeH(i) / 2 - s1Y(n)) <= EeH(i) And s1V(n) = 1 And E1hp(i) > 0 Then
                E1hp(i) = E1hp(i) - 3 - Int(Exp / 500) / 3
               If E1hp(i) > 0 Then
       If Bs1V(nN) = 0 Then
       Bs1V(Int(nN)) = 1
       Bs1X(Int(nN)) = E1X(i) + EeW(i) / 2
       Bs1Y(Int(nN)) = P1Y(n) + 50
      Exp = Exp + 1
               End If
       Else
                
       If Bs2V(nN) = 0 Then
        If Sound = 0 Then sndPlaySound App.Path & "\data\ut3.mid", SND_ASYNC
       Bs2V(Int(nN)) = 1
       Bs2X(Int(nN)) = E1X(i)
       Bs2Y(Int(nN)) = E1Y(i)
     Exp = Exp + 10
     If Rnd >= 0.95 And Hhp <= 0 And Hhpp(0) + Hhpp(1) < 5 Then
Randomize
Hhp = 1
Hhpx = Me.ScaleWidth
Hhpy = Rnd * (x.ScaleHeight - 50)

End If
               End If
End If

    End If
      Next
      Next
'Mhit
  For i = 0 To Ee
  For n = 0 To 11
       If Abs(E1X(i) + EeW(i) / 2 - Ms1X(n) - 30) <= EeW(i) / 2 + 30 And Abs(E1Y(i) + EeH(i) / 2 - Ms1Y(n) - 5) <= EeH(i) / 2 + 5 And Ms1S(n) >= 1 And E1hp(i) > 0 Then
                E1hp(i) = E1hp(i) - 40 - Int(Exp / 500) * 4
         If Bs2V(nN) = 0 Then
       Bs2V(Int(nN)) = 1
       Bs2X(Int(nN)) = E1X(i)
       Bs2Y(Int(nN)) = E1Y(i)
    Exp = Exp + 1
              Ms1V(0) = Ms1V(0) - 1
              Ms1S(n) = 0
End If

    End If
      Next
      Next
'Lhit
  For i = 0 To Ee
  For n = 0 To 0
       If Abs(E1X(i) + EeW(i) / 2 - L1sX(n) - 300) <= EeW(i) / 2 + 300 And Abs(E1Y(i) + EeH(i) / 2 - L1sY(n) - 10) <= EeH(i) / 2 + 10 And L1sV(n) >= 1 And E1hp(i) > 0 Then
                E1hp(i) = E1hp(i) - 10 - Int(Exp / 500) * 1
         If Bs2V(nN) = 0 Then
       Bs2V(Int(nN)) = 1
       Bs2X(Int(nN)) = E1X(i)
       Bs2Y(Int(nN)) = E1Y(i)
    Exp = Exp + 1
End If

    End If
      Next
      Next
        'E2 hit p
  For i = 0 To 1
  For n = 0 To 50
       If Abs(P1X(i) + Form1.P1(i).ScaleWidth / 2 - E2X(n) - 30) <= Form1.P1(i).ScaleWidth / 2 + 30 And Abs(P1Y(i) + Form1.P1(i).ScaleHeight / 2 - E2Y(n) - 5) <= Form1.P1(i).ScaleHeight / 2 + 5 And E2V(n) = 1 And P1hp(i) > 0 And Ms1V(0) <= 0 Then
                P1hp(i) = P1hp(i) - 5
         If Bs2V(nN) = 0 Then
       Bs2V(Int(nN)) = 1
       Bs2X(Int(nN)) = P1X(i)
       Bs2Y(Int(nN)) = E2Y(n) - 25
E2V(n) = 2
End If

    End If
      Next
    
      'hp hit
      If Abs(P1X(i) + Form1.P1(i).ScaleWidth / 2 - Hhpx - 30) <= Form1.P1(i).ScaleWidth / 2 + 30 And Abs(P1Y(i) + Form1.P1(i).ScaleHeight / 2 - Hhpy - 5) <= Form1.P1(i).ScaleHeight / 2 + 5 And Hhp = 1 And P1hp(i) > 0 Then
      Hhpp(i) = Hhpp(i) + 1
   
      Hhp = 0
      End If
      Next
            'E hit p
  For i = 0 To 1
  For n = 0 To Ee
  If E1S(n) <= 2 And E1S(n) >= 1 Or E1S(n) = 4 Then
       If Abs(P1X(i) + Form1.P1(i).ScaleWidth / 2 - E1X(n) - EeW(n) / 2) <= Form1.P1(i).ScaleWidth / 2 + EeW(n) / 2 And Abs(P1Y(i) + Form1.P1(i).ScaleHeight / 2 - E1Y(n) - EeH(n) / 2) <= Form1.P1(0).ScaleWidth / 2 + EeH(n) / 2 And E1hp(n) >= 1 And P1hp(i) > 0 And Ms1V(0) <= 0 Then
                P1hp(i) = P1hp(i) - 1.2
         If Bs2V(nN) = 0 Then
       Bs2V(Int(nN)) = 1
       Bs2X(Int(nN)) = P1X(i)
       Bs2Y(Int(nN)) = E1Y(n)

End If
End If
    End If
      Next
      Next
      'E2 hit p2
  For i = 0 To 1
  For n = 0 To 50
       If Abs(P2X(i) - E2X(n)) <= 30 And Abs(P2Y(i) - E2Y(n)) <= 30 And E2V(n) = 1 And P2hp(i) > 0 And Ms1V(0) <= 0 Then
              If i = 0 And P2s = 1 Then
              P2hp(i) = P2hp(i) - 2
              Else
              P2hp(i) = P2hp(i) - 3
              End If
         If Bs2V(nN) = 0 Then
         Bs2V(Int(nN)) = 1
       Bs2X(Int(nN)) = P2X(i)
       Bs2Y(Int(nN)) = E2Y(n) - 5
E2V(n) = 2
End If

    End If
      Next
      Next
            'E hit p2
  For i = 0 To 1
  For n = 0 To Ee
    If Abs(P2X(i) + Form1.p2.ScaleWidth / 2 - E1X(n) - EeW(n) / 2) <= Form1.p2.ScaleWidth / 2 + EeW(n) / 2 And Abs(P2Y(i) + Form1.p2.ScaleHeight / 2 - E1Y(n) - EeH(n) / 2) <= Form1.p2.ScaleWidth / 2 + EeH(n) / 2 And E1hp(n) >= 1 And P2hp(i) > 0 And E1S(n) <= 2 And E1S(n) >= 1 And Ms1V(0) <= 0 Then
                If i = 0 And P2s = 1 Then
              P2hp(i) = P2hp(i) - 0.2
              Else
              P2hp(i) = P2hp(i) - 0.3
              End If
         If Bs2V(nN) = 0 Then
       Bs2V(Int(nN)) = 1
       Bs2X(Int(nN)) = P2X(i)
       Bs2Y(Int(nN)) = E1Y(n) - 5

End If

    End If
      Next
      Next
            ' Boom

            For i = 0 To 50

 If Bs1V(i) = 1 Then
  mMm(i) = mMm(i) + 0.4
   If mMm(i) >= 6 Then
  Bs1V(i) = 0
  mMm(i) = 0
  
  End If
'BitBlt .hdc, Bs1X(i), Bs1Y(i), Form1.Boom(Int(mMm(i))).ScaleWidth, Form1.Boom(Int(mMm(i))).ScaleHeight, Form1.Boom2(Int(mMm(i))).hdc, 0, 0, vbSrcAnd
     BitBlt .hdc, Bs1X(i), Bs1Y(i), Form1.Boom(Int(mMm(i))).ScaleWidth, Form1.Boom(Int(mMm(i))).ScaleHeight, Form1.Boom(Int(mMm(i))).hdc, 0, 0, vbSrcPaint
         

      End If

      Next
            ' Boom2

            For i = 0 To 50

 If Bs2V(i) = 1 Then
  mMm(i) = mMm(i) + 0.4
   If mMm(i) >= 6 Then
  Bs2V(i) = 0
  mMm(i) = 0
  
  End If
       
         'StretchBlt .hdc, Bs2X(i), Bs2Y(i), Form1.Boom(Int(mMm(i))).ScaleWidth * 3, Form1.Boom(Int(mMm(i))).ScaleHeight * 3, Form1.Boom2(Int(mMm(i))).hdc, 0, 0, Form1.Boom(Int(mMm(i))).ScaleWidth, Form1.Boom(Int(mMm(i))).ScaleHeight, vbSrcAnd
          StretchBlt .hdc, Bs2X(i), Bs2Y(i), Form1.Boom(Int(mMm(i))).ScaleWidth * 3, Form1.Boom(Int(mMm(i))).ScaleHeight * 3, Form1.Boom(Int(mMm(i))).hdc, 0, 0, Form1.Boom(Int(mMm(i))).ScaleWidth, Form1.Boom(Int(mMm(i))).ScaleHeight, vbSrcPaint
 
      End If

      Next
      'hp bar
      For i = 0 To 1
      If P1hp(i) > 0 Then
If whp(i).Width >= (P1hp(i) * 1000 / 50 - 100 / 50) And whp(i).Width <= (P1hp(i) * 1000 / 50 + 100 / 50) And whp(i).BackColor <> &H80000005 Then whp(i).BackColor = &H80000005
If whp(i).Width < (P1hp(i) * 1000 / 50 - 100 / 50) Then
If whp(i).Width < (P1hp(i) * 1000 / 50 - 4001 / 50) Then
whp(i).Width = whp(i).Width + (4000 / 50)
Else
whp(i).Width = whp(i).Width + 1
End If
If whp(i).BackColor <> &H80FF80 Then whp(i).BackColor = &H80FF80
Else
If whp(i).Width > (P1hp(i) * 1000 / 50 + 100 / 50) Then
If whp(i).Width > (P1hp(i) * 1000 / 50 + 1001 / 50) Then
whp(i).Width = whp(i).Width - (1000 / 50)
Else
whp(i).Width = whp(i).Width - 1
End If
If whp(i).BackColor <> &HFF& Then whp(i).BackColor = &HFF&
End If
End If
Else
 P1hp(i) = 0
If whp(i).Visible = True Then
whp(i).Visible = False
Bs2X(Int(nN)) = P1X(i)
Bs2Y(Int(nN)) = P1Y(i)
 If Sound = 0 Then sndPlaySound App.Path & "\data\ut3.mid", SND_ASYNC
 'Picture2.Visible = True
 End If
End If

   
   Next
   

      End With
      If dtime > 0 Then Call delay(dtime)
       DoEvents
               Loop
           
           End Sub
Sub StarsMove(n, Speed)
        If Stars1(n, 1) < 0 Then
            Stars1(n, 1) = picbg.ScaleWidth
            Stars1(n, 2) = Int(Rnd * x.ScaleHeight + Screen.TwipsPerPixelY)
        End If
        Stars1(n, 1) = Stars1(n, 1) - Screen.TwipsPerPixelX * Speed
       picbg.PSet (Stars1(n, 1), Stars1(n, 2)), QBColor(15)
End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image8_Click()
xreturn = Shell("start.exe http://www.gamemagichk.com", 0)
End Sub

Private Sub Label6_Click(Index As Integer)
xreturn = Shell("start.exe http://www.gamemagichk.com", 0)
End Sub

Private Sub Label7_Click()
xreturn = Shell("start.exe http://www.gamemagichk.com", 0)
End Sub

Private Sub Label8_Click()
xreturn = Shell("start.exe http://www.gamemagichk.com", 0)
End Sub

Private Sub tmrTricker_Timer()
Game
tmrTricker.Enabled = False
End Sub

