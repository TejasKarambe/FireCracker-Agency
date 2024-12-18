VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7065
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Palatino Linotype"
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
   ScaleHeight     =   8805
   ScaleWidth      =   7065
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   495
      Left            =   4200
      TabIndex        =   31
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   495
      Left            =   1440
      TabIndex        =   30
      Top             =   8160
      Width           =   1575
   End
   Begin VB.TextBox txtsell 
      Height          =   450
      Left            =   1440
      TabIndex        =   28
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox txttotal 
      Enabled         =   0   'False
      Height          =   450
      Left            =   4920
      TabIndex        =   26
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txtqty 
      Height          =   450
      Left            =   1440
      TabIndex        =   24
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txtcost 
      Height          =   450
      Left            =   4920
      TabIndex        =   22
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtpack 
      Height          =   450
      Left            =   1440
      TabIndex        =   20
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   6735
      Begin VB.ComboBox cmbname 
         Height          =   450
         Left            =   2400
         TabIndex        =   19
         Top             =   1080
         Width           =   4215
      End
      Begin VB.ComboBox cmbsuppno 
         Height          =   450
         Left            =   2400
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtaddress 
         Enabled         =   0   'False
         Height          =   450
         Left            =   2400
         TabIndex        =   10
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox txtcontact 
         Enabled         =   0   'False
         Height          =   450
         Left            =   2400
         TabIndex        =   9
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtshopname 
         Enabled         =   0   'False
         Height          =   450
         Left            =   2400
         TabIndex        =   8
         Top             =   3000
         Width           =   4215
      End
      Begin VB.TextBox txtsupplier 
         Enabled         =   0   'False
         Height          =   450
         Left            =   2400
         TabIndex        =   7
         Top             =   3600
         Width           =   4335
      End
      Begin VB.TextBox txtcompany 
         Enabled         =   0   'False
         Height          =   450
         Left            =   2400
         TabIndex        =   6
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Label Label10 
         Caption         =   "Supp. No. :"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Name :"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Address :"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Contact No.:"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Shop Name :"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Supplier of :"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Company/Brand :"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   4200
         Width           =   1935
      End
   End
   Begin VB.TextBox txtorderno 
      Enabled         =   0   'False
      Height          =   450
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpdate 
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   83099649
      CurrentDate     =   44975
   End
   Begin VB.Label Label15 
      Caption         =   "Sell Price :"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "Total :"
      Height          =   375
      Left            =   3600
      TabIndex        =   27
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Cost Price :"
      Height          =   375
      Left            =   3600
      TabIndex        =   25
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Quantity :"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Pack Size :"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Purchase Order"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label9 
      Caption         =   "Date :"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Order No. :"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmpurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbname_Click()
str = "Select * from Supplier Where Sname=" & cmbname.Text
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
         cmbsuppno.Text = rs.Fields("SuppNo")
         txtaddress.Text = rs.Fields("Address")
         txtcontact.Text = rs.Fields("Contactno")
         txtshopname.Text = rs.Fields("ShopName")
         txtsupplier.Text = rs.Fields("Supplierof")
         txtcompany.Text = rs.Fields("Company")
         rs.Close
End Sub

Private Sub cmbsuppno_Click()
    str = "Select * from Supplier Where SuppNo=" & cmbsuppno.Text
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
         cmbname.Text = rs.Fields("Sname")
         txtaddress.Text = rs.Fields("Address")
         txtcontact.Text = rs.Fields("Contactno")
         txtshopname.Text = rs.Fields("ShopName")
         txtsupplier.Text = rs.Fields("Supplierof")
         txtcompany.Text = rs.Fields("Company")
         rs.Close
         
End Sub

Private Sub cmdback_Click()
    Unload Me
End Sub

Private Sub cmdnew_Click()
    If cmdnew.Caption = "New" Then
        cmdnew.Caption = "Order"
        str = "Select * from Purchase order by Orderno"
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
        If rs.EOF = True Then
            txtorderno.Text = 1
        Else
            rs.MoveLast
            txtorderno.Text = rs.Fields("Orderno") + 1
        End If
        rs.Close
        Call fillno
        Call fillname
    Else
        str = "Select * from Purchase order by Orderno"
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
        rs.AddNew
        rs.Fields("OrderNo") = txtorderno.Text
        rs.Fields("Odate") = dtpdate.Value
        rs.Fields("Suppno") = cmbsuppno.Text
        rs.Fields("SName") = cmbname.Text
        rs.Fields("Item") = txtsupplier.Text
        rs.Fields("Company") = txtcompany.Text
        rs.Fields("Pack") = txtpack.Text
        rs.Fields("Cost") = txtcost.Text
        rs.Fields("Quantity") = txtqty.Text
        rs.Fields("Total") = txttotal.Text
        rs.Fields("Sell") = txtsell.Text
        rs.Update
        rs.Close
        MsgBox " Item " & txtsupplier.Text & " is Ordered", vbInformation
        Unload Me
    End If
End Sub

Private Sub fillno()
    str = "Select * from Supplier order by SuppNo"
    Set rs = New ADODB.Recordset
    rs.Open str, cn, 1, 3
    While Not rs.EOF
        cmbsuppno.AddItem (rs.Fields("Suppno"))
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub fillname()
     str = "Select * from Supplier order by SuppNo"
    Set rs = New ADODB.Recordset
    rs.Open str, cn, 1, 3
    While Not rs.EOF
        cmbname.AddItem (rs.Fields("Sname"))
        rs.MoveNext
    Wend
    rs.Close
End Sub


Private Sub Form_Load()
    dtpdate.Value = Date
End Sub

Private Sub txtqty_Change()
    txttotal.Text = Val(txtcost.Text) * Val(txtqty.Text)
End Sub
