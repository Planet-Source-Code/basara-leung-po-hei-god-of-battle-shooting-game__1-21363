VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Visible         =   0   'False
   Begin VB.OLE OLE2 
      Class           =   "mplayer"
      Height          =   1335
      Left            =   1800
      OleObjectBlob   =   "Form3.frx":0000
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OLE OLE1 
      Class           =   "mplayer"
      Height          =   1335
      Left            =   0
      OleObjectBlob   =   "Form3.frx":3218
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'OLE1.AppIsRunning = True
'OLE1.DoVerb
End Sub
