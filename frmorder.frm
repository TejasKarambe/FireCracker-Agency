VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmorder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10590
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
   ScaleHeight     =   6870
   ScaleWidth      =   10590
   Begin VB.TextBox txtfinal 
      Enabled         =   0   'False
      Height          =   450
      Left            =   8760
      TabIndex        =   31
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      Height          =   495
      Left            =   3000
      TabIndex        =   30
      Top             =   6240
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2370
      Left            =   6960
      TabIndex        =   29
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   2370
      Left            =   8880
      TabIndex        =   28
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox cmbitem 
      Height          =   450
      Left            =   1080
      TabIndex        =   26
      Top             =   5040
      Width           =   3015
   End
   Begin VB.TextBox txtorderno 
      Enabled         =   0   'False
      Height          =   450
      Left            =   2040
      TabIndex        =   17
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   6735
      Begin VB.TextBox txtshopname 
         Enabled         =   0   'False
         Height          =   450
         Left            =   2400
         TabIndex        =   11
         Top             =   3000
         Width           =   4215
      End
      Begin VB.TextBox txtcontact 
         Enabled         =   0   'False
         Height          =   450
         Left            =   2400
         TabIndex        =   10
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtaddress 
         Enabled         =   0   'False
         Height          =   450
         Left            =   2400
         TabIndex        =   9
         Top             =   1800
         Width           =   4215
      End
      Begin VB.ComboBox cmbdistno 
         Height          =   450
         Left            =   2400
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox cmbname 
         Height          =   450
         Left            =   2400
         TabIndex        =   7
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "Shop Name :"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Contact No.:"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Address :"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Name :"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Dist. No. :"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.TextBox txtpack 
      Height          =   450
      Left            =   5520
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtqty 
      Height          =   450
      Left            =   1440
      TabIndex        =   4
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txttotal 
      Enabled         =   0   'False
      Height          =   450
      Left            =   5520
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtsell 
      Height          =   450
      Left            =   1440
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "Order"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   495
      Left            =   8880
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtpdate 
      Height          =   495
      Left            =   4920
      TabIndex        =   18
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   95420417
      CurrentDate     =   44975
   End
   Begin VB.Label Label8 
      Caption         =   "Final Amount :"
      Height          =   375
      Left            =   6960
      TabIndex        =   32
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Item :"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Order No. :"
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Date :"
      Height          =   375
      Left            =   3960
      TabIndex        =   24
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sell Order"
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
      Left            =   3240
      TabIndex        =   23
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label11 
      Caption         =   "Pack Size :"
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Quantity :"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "Total :"
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "Sell Price :"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Width           =   1095
   End
End
Attribute VB_Name = "frmorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbdistno_Click()
     str = "Select * from Distributor Where Distno=" & cmbdistno.Text
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
        cmbname.Text = rs.Fields("Sname")
         txtaddress.Text = rs.Fields("Address")
         txtcontact.Text = rs.Fields("Contactno")
         txtshopname.Text = rs.Fields("ShopName")
         rs.Close
End Sub

Private Sub cmbitem_Click()
     str = "Select * from Purchase Where Item='" & cmbitem.Text & "'"
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
         txtpack.Text = rs.Fields("Pack")
         txtsell.Text = rs.Fields("Sell")
         rs.Close
End Sub

Private Sub cmbname_Click()
    str = "Select * from Distributor Where Sname='" & cmbname.Text & "'"
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
        cmbdistno.Text = rs.Fields("Distno")
         txtaddress.Text = rs.Fields("Address")
         txtcontact.Text = rs.Fields("Contactno")
         txtshopname.Text = rs.Fields("ShopName")
         rs.Close
End Sub

Private Sub cmdadd_Click()
    List1.AddItem (cmbitem.Text)
    List2.AddItem Val(txttotal.Text)
End Sub

Private Sub cmdback_Click()
    Unload Me
End Sub

Private Sub cmdnew_Click()
    Dim K As Integer
    Dim I As Integer
    For I = 0 To List2.ListCount
        If List2.List(I) <> "" Then
            K = K + Val(List2.List(I))
        End If
   
    Next I
    txtfinal.Text = K
    str = "Select * from SellOrder order by Orderno"
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
        rs.AddNew
        rs.Fields("OrderNo") = txtorderno.Text
        rs.Fields("Odate") = dtpdate.Value
        rs.Fields("DistNo") = cmbdistno.Text
        rs.Fields("DName") = cmbname.Text
        rs.Fields("Address") = txtaddress.Text
        rs.Fields("ShopName") = txtshopname.Text
        rs.Fields("Total") = txtfinal.Text
        rs.Update
        rs.Close
        MsgBox "Order from Distributor Is Taken", vbInformation
        Unload Me
    
End Sub

Private Sub Form_Load()
    dtpdate.Value = Date
     str = "Select * from SellOrder order by Orderno"
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
        Call fillitem
End Sub
Private Sub fillno()
    str = "Select * from Distributor order by Distno"
    Set rs = New ADODB.Recordset
    rs.Open str, cn, 1, 3
    While Not rs.EOF
        cmbdistno.AddItem (rs.Fields("Distno"))
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub fillname()
    str = "Select * from Distributor order by Distno"
    Set rs = New ADODB.Recordset
    rs.Open str, cn, 1, 3
    While Not rs.EOF
        cmbname.AddItem (rs.Fields("Sname"))
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub fillitem()
  str = "Select * from Purchase order by OrderNo"
    Set rs = New ADODB.Recordset
    rs.Open str, cn, 1, 3
    While Not rs.EOF
        cmbitem.AddItem (rs.Fields("Item"))
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub txtqty_Change()
    txttotal.Text = Val(txtsell.Text) * Val(txtqty.Text)
End Sub
