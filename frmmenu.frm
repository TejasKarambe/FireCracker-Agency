VERSION 5.00
Begin VB.MDIForm frmmenu 
   BackColor       =   &H8000000C&
   Caption         =   "Firecracker Agency"
   ClientHeight    =   6795
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8220
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnunew 
      Caption         =   "New"
      Begin VB.Menu mnusupplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu mnudistributor 
         Caption         =   "Distributor"
      End
   End
   Begin VB.Menu mnuorder 
      Caption         =   "Order"
      Begin VB.Menu mnupurchase 
         Caption         =   "Purchase"
      End
      Begin VB.Menu mnusell 
         Caption         =   "Sell"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnuviewsupp 
         Caption         =   "Supplier"
      End
      Begin VB.Menu mnuviewdistri 
         Caption         =   "Distributor"
      End
      Begin VB.Menu mnuvieworder 
         Caption         =   "Order"
      End
      Begin VB.Menu mnuviewincome 
         Caption         =   "Income"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnudistributor_Click()
    frmdistributor.Show
End Sub

Private Sub mnupurchase_Click()
    frmpurchase.Show
End Sub

Private Sub mnuquit_Click()
    End
End Sub

Private Sub mnusell_Click()
    frmorder.Show
End Sub

Private Sub mnusupplier_Click()
    frmsupplier.Show
End Sub

Private Sub mnuviewdistri_Click()
    frmviewdistributor.Show
End Sub

Private Sub mnuviewincome_Click()
    frmincome.Show
End Sub

Private Sub mnuvieworder_Click()
    frmviewpurchase.Show
End Sub

Private Sub mnuviewsupp_Click()
    frmviewsupplier.Show
End Sub
