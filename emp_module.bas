Attribute VB_Name = "emp_module"
Public con As New ADODB.Connection

Public Sub open_db()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\emp.mdb;Persist Security Info=False"
End Sub

Public Sub close_db()
con.Close
End Sub

