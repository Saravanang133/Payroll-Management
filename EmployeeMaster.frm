VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form EmployeeMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Master"
   ClientHeight    =   10830
   ClientLeft      =   2130
   ClientTop       =   2355
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   8955
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   6600
      TabIndex        =   26
      Top             =   1680
      Width           =   2055
      Begin VB.CommandButton Command7 
         Caption         =   "CLOSE"
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
         TabIndex        =   33
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
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
         Left            =   240
         TabIndex        =   32
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
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
         Left            =   240
         TabIndex        =   31
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
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
         Left            =   240
         TabIndex        =   30
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
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
         Left            =   240
         TabIndex        =   29
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
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
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
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
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Female"
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
      Left            =   4800
      TabIndex        =   21
      Top             =   6480
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Male"
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
      Left            =   3360
      TabIndex        =   20
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text6 
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
      Left            =   3360
      TabIndex        =   7
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox Text5 
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
      Left            =   3360
      TabIndex        =   6
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox Text4 
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
      Left            =   3360
      TabIndex        =   5
      Top             =   3720
      Width           =   1935
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
      Left            =   3360
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
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
      Left            =   3360
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "EMPLOYEE DETAILS"
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.Frame Frame3 
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
         Height          =   975
         Left            =   720
         TabIndex        =   34
         Top             =   6960
         Width           =   5535
         Begin VB.OptionButton Option4 
            Caption         =   "Permanent"
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
            Left            =   3360
            TabIndex        =   36
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Temporary"
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
            Left            =   1560
            TabIndex        =   35
            Top             =   360
            Width           =   1455
         End
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
         Left            =   3360
         TabIndex        =   25
         Text            =   "Select"
         Top             =   9000
         Width           =   3015
      End
      Begin VB.TextBox Text7 
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
         Left            =   3360
         TabIndex        =   23
         Top             =   9600
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   3360
         TabIndex        =   22
         Top             =   8280
         Width           =   3015
         _ExtentX        =   5318
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
         Format          =   39190529
         CurrentDate     =   40914
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1095
         Left            =   3360
         TabIndex        =   3
         Top             =   1680
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"EmployeeMaster.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   3360
         TabIndex        =   8
         Top             =   5640
         Width           =   3015
         _ExtentX        =   5318
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
         Format          =   39190529
         CurrentDate     =   40914
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
         Left            =   3360
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label12 
         Caption         =   "Department"
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
         Left            =   480
         TabIndex        =   24
         Top             =   9000
         Width           =   2655
      End
      Begin VB.Label Label11 
         Caption         =   "Salary"
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
         Left            =   480
         TabIndex        =   19
         Top             =   9720
         Width           =   2895
      End
      Begin VB.Label Label10 
         Caption         =   "Date Of Joining"
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
         Left            =   480
         TabIndex        =   18
         Top             =   8280
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Gender"
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
         Left            =   480
         TabIndex        =   17
         Top             =   6480
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "Date Of Birth"
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
         Left            =   480
         TabIndex        =   16
         Top             =   5760
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Experiance"
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
         Left            =   480
         TabIndex        =   15
         Top             =   5040
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Qualification"
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
         Left            =   480
         TabIndex        =   14
         Top             =   4320
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "State"
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
         Left            =   480
         TabIndex        =   13
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "City"
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
         Left            =   480
         TabIndex        =   12
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Address"
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
         Left            =   480
         TabIndex        =   11
         Top             =   1800
         Width           =   2295
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
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   2415
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
         Left            =   480
         TabIndex        =   9
         Top             =   600
         Width           =   2415
      End
   End
End
Attribute VB_Name = "EmployeeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub Command1_Click()
rs.MoveLast
clear
Text1.Text = rs(0) + 1
Text2.SetFocus
End Sub

Private Sub Command2_Click()
rs.MoveFirst
While Not rs.EOF
If rs(0) = Text1.Text Then
temp = 1
End If
rs.MoveNext
Wend

If temp = 1 Then
MsgBox "This Employee code already Exisits", vbInformation
Exit Sub
Else
rs.AddNew
ins
rs.Update
MsgBox "Record Inserted Successfully", vbInformation
End If

End Sub


Sub display()
Text1.Text = rs(0)
Text2.Text = rs(1)
RichTextBox1.Text = rs(2)
Text3.Text = rs(3)
Text4.Text = rs(4)
Text5.Text = rs(5)
Text6.Text = rs(6)
DTPicker1.Value = rs(7)
If rs(8) = "Male" Then
Option1.Value = True
Else
Option2.Value = True
End If
DTPicker2.Value = rs(9)
Combo1.Text = rs(10)
Text7.Text = rs(11)
End Sub
Sub ins()
rs(0) = Text1.Text
rs(1) = Text2.Text
rs(2) = RichTextBox1.Text
rs(3) = Text3.Text
rs(4) = Text4.Text
rs(5) = Text5.Text
rs(6) = Text6.Text
rs(7) = DTPicker1.Value
If Option4.Value = True Then
    rs("type") = 1
Else
    rs("type") = 0
End If
If Option1.Value = True Then
rs(8) = "Male"
Else
rs(8) = "Female"
End If
rs(9) = DTPicker2.Value
rs(10) = Combo1.Text
rs(11) = Text7.Text
End Sub
Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
RichTextBox1.Text = ""
DTPicker1.Value = Now
DTPicker2.Value = Now
Combo1.Text = "Select"
Text1.SetFocus
End Sub

Private Sub Command3_Click()
rs.MoveFirst
While Not rs.EOF
If rs(0) = Text1.Text Then
ins
rs.Update
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


Private Sub Command4_Click()
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
MsgBox "Record Not Found ... to delete", vbInformation
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

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call open_db
rs.CursorLocation = adUseClient
rs.Open "emp_mas", con, adOpenDynamic, adLockOptimistic

rs1.CursorLocation = adUseClient
rs1.Open "dep_mas", con, adOpenDynamic, adLockOptimistic

rs1.MoveFirst
While Not rs1.EOF
Combo1.AddItem rs1(2)
rs1.MoveNext
Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call close_db
End Sub

