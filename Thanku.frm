VERSION 5.00
Begin VB.Form Thanku 
   Caption         =   "Exit"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8280
      Top             =   5160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "THANK YOU"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   4680
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   6720
      Left            =   0
      Picture         =   "Thanku.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11040
   End
End
Attribute VB_Name = "Thanku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Timer1.Interval = Timer1.Interval + 100
If Timer1.Interval = 700 Then
End
End If
End Sub
