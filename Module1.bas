Attribute VB_Name = "Module1"
Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public str As String
Sub main()
    str = "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.Path & "\Firecracker.mdb; Persist Security Info=False"
    cn.Open str
    frmsplash.Show
   ' frmmenu.Show
    
End Sub
Function CHECKTEXT(k As Integer)
Select Case k
        Case 65 To 90, 97 To 122, 8, 32
                 k = k
        Case Else
                 k = 0
End Select
CHECKTEXT = k
End Function
Function CHECKNUM(k As Integer)
Select Case k
        Case 48 To 57, 8
                 k = k
        Case Else
                 k = 0
End Select
CHECKNUM = k
End Function

