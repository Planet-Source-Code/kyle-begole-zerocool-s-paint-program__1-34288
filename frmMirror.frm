VERSION 5.00
Begin VB.Form frmMirror 
   Caption         =   "Mirror"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Frame fraMirror 
      Caption         =   "Choose where you want the mirrors positioned."
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4815
      Begin VB.OptionButton optNone 
         Caption         =   "No Mirrors"
         Height          =   495
         Left            =   3720
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optBoth 
         Caption         =   "Both Mirrors"
         Height          =   495
         Left            =   2600
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optHorizontal 
         Caption         =   "Horizontal Mirrors"
         Height          =   495
         Left            =   1240
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optVertical 
         Caption         =   "Vertical Mirrors"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMirror"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    frmPaint.Enabled = True
    Unload Me
End Sub

Private Sub optBoth_Click()
    blnBothM = True
    blnHorM = False
    blnVertM = False
    frmPaint.HorizontalMirrorLine.Visible = True
    frmPaint.VertMirrorLine.Visible = True
End Sub

Private Sub optHorizontal_Click()
    blnHorM = True
    blnVertM = False
    blnBothM = False
    frmPaint.HorizontalMirrorLine.Visible = True
    frmPaint.VertMirrorLine.Visible = False
End Sub

Private Sub optNone_Click()
    blnHorM = False
    blnVertM = False
    blnBothM = False
    frmPaint.HorizontalMirrorLine.Visible = False
    frmPaint.VertMirrorLine.Visible = False
End Sub

Private Sub optVertical_Click()
    blnVertM = True
    blnHorM = False
    blnBothM = False
    frmPaint.HorizontalMirrorLine.Visible = False
    frmPaint.VertMirrorLine.Visible = True
End Sub
