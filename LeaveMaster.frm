VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form LeaveMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   3480
   ClientTop       =   3525
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6300
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   0
      TabIndex        =   12
      Top             =   3960
      Width           =   6255
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "REPORT"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1080
         Width           =   1455
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   1455
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   1455
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1455
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
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LEAVE MASTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
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
         Left            =   4920
         TabIndex        =   10
         Top             =   3240
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Medical Leave"
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
         Left            =   2880
         TabIndex        =   9
         Top             =   3240
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Casual Leave"
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
         Left            =   2880
         TabIndex        =   8
         Top             =   2520
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2880
         TabIndex        =   7
         Top             =   1800
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   55312385
         CurrentDate     =   40918
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
         Left            =   2880
         TabIndex        =   6
         Top             =   1080
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   5
         Text            =   "Select"
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   11
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Type"
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
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Date"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Employee Code"
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "LeaveMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim temp

Private Sub Combo1_Click()
rs1.MoveFirst
While Not rs1.EOF
If Combo1.Text = rs1(0) Then
Text1.Text = rs1(1)
Exit Sub
End If
rs1.MoveNext
Wend
End Sub

Private Sub Command1_Click()
If temp = 1 Then
rs2.Close
End If
rs2.CursorLocation = adUseClient
rs2.Open "select * from leave_mas where emp_code = " & Combo1.Text, con, adOpenDynamic, adLockOptimistic
temp = 1
Set LeaveMasterReport.DataGrid1.DataSource = rs2
LeaveMasterReport.Show

End Sub

Private Sub Command2_Click()
rs.AddNew
ins
rs.Update
MsgBox "Record Inserted Successfully", vbInformation
End Sub

Private Sub Command3_Click()
rs.MoveFirst
While Not rs.EOF
If rs(0) = Combo1.Text And rs(2) = DTPicker1.Value Then
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
MsgBox "Record Not Found on that date", vbInformation
End If
End Sub

Private Sub Command4_Click()
rs.MoveFirst
While Not rs.EOF
If rs(0) = Combo1.Text Then
ins
MsgBox "Record Updated Successfully", vbInformation
Exit Sub
End If
rs.MoveNext
Wend

End Sub

Private Sub Command5_Click()
clear
End Sub

Private Sub Command6_Click()
rs.MoveFirst
While Not rs.EOF
If rs(2) = DTPicker1.Value And rs(0) = Combo1.Text Then
display
Exit Sub
End If
rs.MoveNext
Wend

End Sub

Private Sub Form_Load()
Call open_db

rs.CursorLocation = adUseClient
rs.Open "leave_mas", con, adOpenDynamic, adLockOptimistic


rs1.CursorLocation = adUseClient
rs1.Open "emp_mas", con, adOpenDynamic, adLockOptimistic


rs1.MoveFirst
While Not rs1.EOF
Combo1.AddItem rs1(0)
rs1.MoveNext
Wend

End Sub
Sub display()
Combo1.Text = rs(0)
Text1.Text = rs(1)
DTPicker1.Value = rs(2)
If rs(3) = "Casual" Then
Option1.Value = True
Else
Option2.Value = True
l = Len(rs(3))

temp = Mid(rs(3), l - 1, 2)
Text2.Text = temp
End If
End Sub
Sub ins()
rs(0) = Combo1.Text
rs(1) = Text1.Text
rs(2) = DTPicker1.Value
If Option1.Value = True Then
rs(3) = "Casual"
Else
rs(3) = "Medical Leave" & " " & Text2.Text
End If
End Sub
Sub clear()
Combo1.Text = "Select"
Text1.Text = ""
DTPicker1.Value = Date
Text2.Text = ""
Option1.Value = False
Option2.Value = False
Combo1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call close_db
End Sub
