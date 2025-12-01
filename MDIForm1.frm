VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3465
   ClientLeft      =   -90
   ClientTop       =   4905
   ClientWidth     =   12060
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mas 
      Caption         =   "Master"
      Begin VB.Menu dep_mas 
         Caption         =   "Department Master"
      End
      Begin VB.Menu emp_mas 
         Caption         =   "Employee Master"
      End
      Begin VB.Menu leave_mas 
         Caption         =   "Leave Master"
      End
      Begin VB.Menu xit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu trans 
      Caption         =   "Transaction"
      Begin VB.Menu sal_det 
         Caption         =   "Salary Details"
      End
   End
   Begin VB.Menu rep 
      Caption         =   "Report"
      Begin VB.Menu leave_rep 
         Caption         =   "Leave Report"
      End
      Begin VB.Menu sal_rep 
         Caption         =   "Salary Report"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub dep_mas_Click()
DepartmentMaster.Show
End Sub

Private Sub emp_mas_Click()
EmployeeMaster.Show
End Sub

Private Sub leave_mas_Click()
LeaveMaster.Show
End Sub

Private Sub leave_rep_Click()
DataReport1.Show
End Sub

Private Sub sal_det_Click()
SalaryDetails.Show
End Sub

Private Sub sal_rep_Click()
DataReport2.Show
End Sub

Private Sub xit_Click()
Thanku.Show
End Sub
