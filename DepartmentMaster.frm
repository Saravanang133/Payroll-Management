VERSION 5.00
Begin VB.Form DepartmentMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DEPARTMENT MASTER"
   ClientHeight    =   5010
   ClientLeft      =   5985
   ClientTop       =   4125
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   6135
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "VIEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "MODIFY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "NEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "DEPARTMENT MASTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.Label Label3 
         Caption         =   "Department name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Department Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2295
      End
   End
End
Attribute VB_Name = "DepartmentMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
rs.MoveLast
Text1.Text = rs(0) + 1
Text2.SetFocus
End Sub

Private Sub Command2_Click()
rs.MoveFirst
While Not rs.EOF
If rs(0) = Text1.Text Then
temp = 1
ElseIf rs(2) = Text3.Text Then
temp1 = 1
End If
rs.MoveNext
Wend

If temp = 1 Then
MsgBox "This Department code already Exisits", vbInformation
Exit Sub
ElseIf temp1 = 1 Then
MsgBox "This Department name already Exisits", vbInformation
Exit Sub
Else
rs.AddNew
ins
rs.Update
MsgBox "Record Inserted Successfully", vbInformation
End If

End Sub

Private Sub Command3_Click()
rs.MoveFirst
While Not rs.EOF
If rs(0) = Text1.Text Then
rs.Delete
MsgBox "Record Deleted Successfully", vbInformation
clear
Exit Sub
Else
temp2 = 1
End If
rs.MoveNext
Wend


If temp2 = 1 Then
MsgBox "Record Not Found to delete", vbInformation
End If
End Sub

Private Sub Command4_Click()
rs.MoveFirst
While Not rs.EOF
If rs(0) = Text1.Text Then
ins
MsgBox "Record Updated Successfully", vbInformation
Exit Sub
Else
temp2 = 1
End If
rs.MoveNext
Wend


If temp2 = 1 Then
MsgBox "Record Not Found", vbInformation
End If
End Sub

Private Sub Command5_Click()
clear
End Sub

Private Sub Command6_Click()
rs.MoveFirst
While Not rs.EOF
If rs(0) = Text1.Text Then
display
Exit Sub
Else
temp2 = 1
End If
rs.MoveNext
Wend


If temp2 = 1 Then
MsgBox "Record Not Found", vbInformation
End If

End Sub

Private Sub Form_Load()
Call open_db

rs.CursorLocation = adUseClient
rs.Open "dep_mas", con, adOpenDynamic, adLockOptimistic


End Sub
Sub display()
Text1.Text = rs(0)
Text2.Text = rs(1)
Text3.Text = rs(2)
End Sub
Sub ins()
rs(0) = Text1.Text
rs(1) = Text2.Text
rs(2) = Text3.Text
End Sub
Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call close_db
End Sub

