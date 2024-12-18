VERSION 5.00
Begin VB.Form frmsupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6615
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
   ScaleHeight     =   6375
   ScaleWidth      =   6615
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   615
      Left            =   3600
      TabIndex        =   16
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   615
      Left            =   840
      TabIndex        =   15
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox txtcompany 
      Height          =   450
      Left            =   2280
      TabIndex        =   14
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox txtsupplier 
      Height          =   450
      Left            =   2280
      TabIndex        =   12
      Top             =   4080
      Width           =   4335
   End
   Begin VB.TextBox txtshopname 
      Height          =   450
      Left            =   2280
      TabIndex        =   10
      Top             =   3480
      Width           =   4215
   End
   Begin VB.TextBox txtcontact 
      Height          =   450
      Left            =   2280
      TabIndex        =   8
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtaddress 
      Height          =   450
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox txtname 
      Height          =   450
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtsuupno 
      Enabled         =   0   'False
      Height          =   450
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Company/Brand :"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Supplier of :"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Shop Name :"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Contact No.:"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Address :"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Name :"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Supp. No. :"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Supplier"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
    Unload Me
End Sub

Private Sub cmdnew_Click()
    If cmdnew.Caption = "New" Then
        cmdnew.Caption = "Save"
        str = "Select * from Supplier order by SuppNo"
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
        If rs.EOF = True Then
            txtsuupno.Text = 1
        Else
            rs.MoveLast
            txtsuupno.Text = rs.Fields("Suppno") + 1
        End If
        rs.Close
    Else
        str = "Select * from Supplier order by SuppNo"
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
        rs.AddNew
        rs.Fields("Suppno") = txtsuupno.Text
        rs.Fields("Sname") = txtname.Text
        rs.Fields("Address") = txtaddress.Text
        rs.Fields("Contactno") = txtcontact.Text
        rs.Fields("ShopName") = txtshopname.Text
        rs.Fields("Supplierof") = txtsupplier.Text
        rs.Fields("Company") = txtcompany.Text
        rs.Update
        rs.Close
        MsgBox " New Supplier is Added", vbInformation
        Unload Me
        
        
    End If
    
End Sub













