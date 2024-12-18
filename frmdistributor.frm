VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdistributor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6645
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
   ScaleHeight     =   6105
   ScaleWidth      =   6645
   Begin MSComCtl2.DTPicker dtpdate 
      Height          =   495
      Left            =   4800
      TabIndex        =   18
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   95092737
      CurrentDate     =   44975
   End
   Begin VB.TextBox txtdeposit 
      Height          =   450
      Left            =   2280
      TabIndex        =   5
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox txtcity 
      Height          =   450
      Left            =   2280
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtdistno 
      Enabled         =   0   'False
      Height          =   450
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtname 
      Height          =   450
      Left            =   2280
      TabIndex        =   0
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox txtaddress 
      Height          =   450
      Left            =   2280
      TabIndex        =   1
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox txtcontact 
      Height          =   450
      Left            =   2280
      TabIndex        =   3
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtshopname 
      Height          =   450
      Left            =   2280
      TabIndex        =   4
      Top             =   3960
      Width           =   4215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Date :"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Deposit Amount :"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "City :"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Distributor"
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
      TabIndex        =   14
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Dist. No. :"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Name :"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Address :"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Contact No.:"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Shop Name :"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "frmdistributor"
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
        str = "Select * from Distributor order by DistNo"
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
        If rs.EOF = True Then
            txtdistno.Text = 1
        Else
            rs.MoveLast
            txtdistno.Text = rs.Fields("Distno") + 1
        End If
        rs.Close
    Else
        str = "Select * from Distributor order by DistNo"
        Set rs = New ADODB.Recordset
        rs.Open str, cn, 1, 3
        rs.AddNew
        rs.Fields("Distno") = txtdistno.Text
        rs.Fields("Ddate") = dtpdate.Value
        rs.Fields("Sname") = txtname.Text
        rs.Fields("Address") = txtaddress.Text
         rs.Fields("City") = txtcity.Text
        rs.Fields("Contactno") = txtcontact.Text
        rs.Fields("ShopName") = txtshopname.Text
        rs.Fields("Deposit") = txtdeposit.Text
        rs.Update
        rs.Close
        MsgBox " New Distributor is Added", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    dtpdate.Value = Date
End Sub

