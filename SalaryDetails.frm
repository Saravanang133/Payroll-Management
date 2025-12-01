VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SalaryDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALARY DETAILS"
   ClientHeight    =   7200
   ClientLeft      =   1440
   ClientTop       =   1935
   ClientWidth     =   12870
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
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
   ScaleHeight     =   7200
   ScaleWidth      =   12870
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   29
      Top             =   6240
      Width           =   11775
      Begin VB.CommandButton Command9 
         Caption         =   "REPORT"
         Height          =   495
         Left            =   6720
         TabIndex        =   37
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Caption         =   "VIEW"
         Height          =   495
         Left            =   5160
         TabIndex        =   35
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "CLOSE"
         Height          =   495
         Left            =   9840
         TabIndex        =   34
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "CLEAR"
         Height          =   495
         Left            =   8280
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "MODIFY"
         Height          =   495
         Left            =   2040
         TabIndex        =   32
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "DELETE"
         Height          =   495
         Left            =   3600
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ADD"
         Height          =   495
         Left            =   480
         TabIndex        =   30
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SALARY DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton Command5 
         Caption         =   "Calculate Salary"
         Height          =   495
         Left            =   9600
         TabIndex        =   36
         Top             =   5520
         Width           =   2055
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   7920
         TabIndex        =   28
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   2520
         TabIndex        =   27
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   7920
         TabIndex        =   26
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   2520
         TabIndex        =   25
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   7920
         TabIndex        =   24
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   2520
         TabIndex        =   23
         Top             =   3960
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2520
         TabIndex        =   12
         Top             =   3240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         _Version        =   393216
         Format          =   82771969
         CurrentDate     =   40918
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2520
         TabIndex        =   11
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2520
         TabIndex        =   10
         Top             =   1680
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Height          =   420
         Left            =   2520
         TabIndex        =   9
         Text            =   "Select"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         Caption         =   "LEAVE DETAILS"
         Height          =   2895
         Left            =   5760
         TabIndex        =   5
         Top             =   840
         Width           =   5415
         Begin VB.CommandButton Command1 
            Caption         =   "Cal"
            Height          =   495
            Left            =   4680
            TabIndex        =   22
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox Text5 
            Height          =   495
            Left            =   3720
            TabIndex        =   21
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Height          =   495
            Left            =   2760
            TabIndex        =   20
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Height          =   495
            Left            =   2760
            TabIndex        =   19
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Total Leave For Loss Of Pay"
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   2040
            Width           =   3135
         End
         Begin VB.Label Label6 
            Caption         =   "Medical Leave"
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label5 
            Caption         =   "Casual Leave"
            Height          =   375
            Left            =   360
            TabIndex        =   6
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.Label Label13 
         Caption         =   "Net Salary"
         Height          =   495
         Left            =   5760
         TabIndex        =   18
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "PF"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "TA"
         Height          =   495
         Left            =   5760
         TabIndex        =   16
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "DA"
         Height          =   495
         Left            =   360
         TabIndex        =   15
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "HRA"
         Height          =   495
         Left            =   5760
         TabIndex        =   14
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Loss Of Pay"
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Select Month"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Employee Name"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Monthly Salary"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Employee Code"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
   End
End
Attribute VB_Name = "SalaryDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim temp1, type1


Private Sub Combo1_Click()
Dim cas, med, temp
Dim r As Integer
r = 0

cas = med = 0
rs.MoveFirst
While Not rs.EOF
If Combo1.Text = rs(0) Then
Text1.Text = rs(1)
Text2.Text = rs(11)
type1 = rs("type")

If type1 = 1 Then
    Text7.Enabled = True
    Text8.Enabled = True
    Text9.Enabled = True
    Text10.Enabled = True
Else
    Text7.Enabled = False
    Text8.Enabled = False
    Text9.Enabled = False
    Text10.Enabled = False
End If

End If
rs.MoveNext
Wend


End Sub



Private Sub Command1_Click()
Text5.Text = Val(Text3.Text) + Val(Text4.Text)
m = Month(DTPicker1.Value)

If m = 1 Or m = 3 Or m = 5 Or m = 7 Or m = 8 Or m = 10 Or m = 12 Then
sal = Text2.Text / 31
Text6.Text = sal * Text5.Text
Text6.Text = Round(Text6.Text)
Else
sal = Text2.Text / 30
Text6.Text = sal * Text5.Text
Text6.Text = Round(Text6.Text)
End If
End Sub

Private Sub Command2_Click()
rs3.MoveFirst

While Not rs3.EOF
If rs3(0) = Combo1.Text And rs3(1) = Month(DTPicker1.Value) Then
MsgBox ("Salary credited for this Employee , if you want to modify details , use MODIFY button"), vbInformation
Exit Sub
Else
rs3.AddNew
ins
rs3.Update
MsgBox "Record Inserted Successfully", vbInformation
Exit Sub
End If
rs3.MoveNext
Wend
End Sub

Private Sub Command3_Click()
rs3.MoveFirst
While Not rs3.EOF
If rs3(0) = Combo1.Text And rs3(1) = Month(DTPicker1.Value) Then
rs3.Delete
MsgBox "Record Deleted Successfully", vbInformation
clear
Exit Sub
Else
temp2 = 1
End If
rs3.MoveNext
Wend


If temp2 = 1 Then
MsgBox "Record Not Found ... to delete", vbInformation
End If

End Sub

Private Sub Command4_Click()
rs3.MoveFirst
While Not rs3.EOF
If rs3(0) = Combo1.Text Then
ins
MsgBox "Record Updated Successfully", vbInformation
Exit Sub
Else
temp2 = 1
End If
rs3.MoveNext
Wend


If temp2 = 1 Then
MsgBox "Record Not Found", vbInformation
End If
End Sub

Private Sub Command5_Click()

If type1 = 1 Then
    Text7.Text = Text2.Text * 0.01
    Text8.Text = Text2.Text * 0.015
    Text9.Text = Text2.Text * 0.02
    Text10.Text = Text2.Text * 0.01
    Text11.Text = (Val(Text2.Text) + Val(Text7.Text) + Val(Text8.Text) + Val(Text9.Text)) - Val(Text10.Text) - Val(Text6.Text)
Else
    Text11.Text = Val(Text2.Text) - Val(Text6.Text)
End If
End Sub

Private Sub Command6_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Combo1.Text = "Select"
DTPicker1.Value = Date
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Command8_Click()

If type1 = 1 Then
    Text7.Enabled = True
    Text8.Enabled = True
    Text9.Enabled = True
    Text10.Enabled = True
Else
    Text7.Enabled = False
    Text8.Enabled = False
    Text9.Enabled = False
    Text10.Enabled = False
End If


rs3.MoveFirst
While Not rs3.EOF
If rs3(0) = Combo1.Text And rs3(1) = Month(DTPicker1.Value) Then
display
Exit Sub
Else
temp2 = 1
End If
rs3.MoveNext
Wend


If temp2 = 1 Then
MsgBox "Record Not Found", vbInformation
End If

End Sub

Private Sub Command9_Click()
If temp = 1 Then
rs4.Close
End If
rs4.CursorLocation = adUseClient
rs4.Open "select * from sal_det where emp_code = " & Combo1.Text, con, adOpenDynamic, adLockOptimistic
temp = 1
Set LeaveMasterReport.DataGrid1.DataSource = rs4
LeaveMasterReport.Show
End Sub

Private Sub DTPicker1_Change()
Dim cas, med, temp1
cas = med = 0

rs1.Close

rs1.CursorLocation = adUseClient
rs1.Open "select * from leave_mas where emp_code=" & Combo1.Text, con, adOpenDynamic, adLockOptimistic

rs1.MoveFirst
While Not rs1.EOF
rm = Month(rs1(2))
ry = Year(rs1(2))
m = Month(DTPicker1.Value)
Y = Year(DTPicker1.Value)
If rm = m And ry = Y Then
If rs1(3) = "Casual" Then
r = r + 1
cas = CInt(cas) + 1
Else
l = Len(rs1(3))
temp = Mid(rs1(3), l - 1, 2)
med = med + Val(temp)
End If
End If
rs1.MoveNext
Wend

'If cas > 0 Then
'Text3.Text = cas - 1
'Else
'Text3.Text = 0
'End If
Text3.Text = r
If med > 0 Then
Text4.Text = med
Else
Text4.Text = 0
End If
End Sub

Private Sub Form_Load()
Call open_db

rs.CursorLocation = adUseClient
rs.Open "emp_mas", con, adOpenDynamic, adLockOptimistic

rs1.CursorLocation = adUseClient
rs1.Open "leave_mas", con, adOpenDynamic, adLockOptimistic

rs2.CursorLocation = adUseClient
rs2.Open "emp_mas", con, adOpenDynamic, adLockOptimistic

rs3.CursorLocation = adUseClient
rs3.Open "sal_det", con, adOpenDynamic, adLockOptimistic

rs.MoveFirst
While Not rs.EOF
Combo1.AddItem (rs(0))
rs.MoveNext
Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call close_db
End Sub


Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Combo1.Text = "Select"
DTPicker1.Value = Date
End Sub
Sub display()
If type1 = 1 Then
    Combo1.Text = rs3(0)
    DTPicker1.Value = rs3(1)
    Text5.Text = rs3(2)
    Text6.Text = rs3(3)
    Text7.Text = rs3(4)
    Text8.Text = rs3(5)
    Text9.Text = rs3(6)
    Text10.Text = rs3(7)
    Text11.Text = rs3(8)
Else
    Combo1.Text = rs3(0)
    DTPicker1.Value = rs3(1)
    Text5.Text = rs3(2)
    Text6.Text = rs3(3)
    Text11.Text = rs3(8)
End If
End Sub
Sub ins()
If type1 = 1 Then
    rs3(0) = Combo1.Text
    rs3(1) = Month(DTPicker1.Value)
    rs3(2) = Text5.Text
    rs3(3) = Text6.Text
    rs3(4) = Text7.Text
    rs3(5) = Text8.Text
    rs3(6) = Text9.Text
    rs3(7) = Val(Text10.Text)
    rs3(8) = Text11.Text
Else
    rs3(0) = Combo1.Text
    rs3(1) = Month(DTPicker1.Value)
    rs3(2) = Text5.Text
    rs3(3) = Text6.Text
    rs3(8) = Text11.Text
End If
End Sub
