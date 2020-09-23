VERSION 5.00
Begin VB.Form frmColorButton 
   Caption         =   "Mouse Button Color"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose which button color you would like to set"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      Begin VB.OptionButton optRight 
         Caption         =   "Right"
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optLeft 
         Caption         =   "Left"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmColorButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    
    
    frmCustomColor.Show
    Unload Me
    
End Sub

Private Sub Form_Load()
    frmCustomColor.Visible = False
    blnLeft = True
    blnRight = False
End Sub

Private Sub optLeft_Click()
    blnLeft = True
    blnRight = False
End Sub

Private Sub optRight_Click()
    blnLeft = False
    blnRight = True
End Sub
