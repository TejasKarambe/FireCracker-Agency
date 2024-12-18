VERSION 5.00
Begin VB.Form frmsplash 
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2160
      Top             =   3480
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   615
      Left            =   5400
      TabIndex        =   4
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   735
      Left            =   5400
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Firecracker"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Unload Me
    frmlogin.Show
End Sub
