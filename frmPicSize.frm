VERSION 5.00
Begin VB.Form frmPicSize 
   Caption         =   "Picture Size"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame fraSizeOpt 
      Caption         =   "Picture Size Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton opt9000x11000 
         Caption         =   "9000 x 11,000"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton opt5000x7500 
         Caption         =   "5000 x 7500"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton opt8000x10000 
         Caption         =   "8000 x 10,000"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton opt7300x8150 
         Caption         =   "7300 x 8150"
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPicSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    
    If blnCav1 = True Then
        
        frmPaint.picMainPic.ScaleHeight = 5000
        frmPaint.picMainPic.ScaleWidth = 7500
        frmPaint.HScrPicMainPic.Max = 7500 / 15
        frmPaint.VScrPicMainPic.Max = 5000 / 15
        GoTo ExitOk
        
    ElseIf blnCav2 = True Then
        
        frmPaint.picMainPic.ScaleHeight = 7300
        frmPaint.picMainPic.ScaleWidth = 8150
        frmPaint.HScrPicMainPic.Max = 8150 / 15
        frmPaint.VScrPicMainPic.Max = 7300 / 15
        GoTo ExitOk
        
    ElseIf blnCav3 = True Then
        
        frmPaint.picMainPic.ScaleHeight = 8000
        frmPaint.picMainPic.ScaleWidth = 10000
        frmPaint.HScrPicMainPic.Max = 10000 / 15
        frmPaint.VScrPicMainPic.Max = 8000 / 15
        GoTo ExitOk
        
    ElseIf blnCav4 = True Then
        
        frmPaint.picMainPic.ScaleHeight = 9000
        frmPaint.picMainPic.ScaleWidth = 11000
        frmPaint.HScrPicMainPic.Max = 11000 / 15
        frmPaint.VScrPicMainPic.Max = 9000 / 15
        
    End If

ExitOk:
    
    frmPaint.Enabled = True
    
    Unload Me
    
End Sub

Private Sub opt5000x7500_Click()
    blnCav1 = True
    blnCav2 = False
    blnCav3 = False
    blnCav4 = False
End Sub

Private Sub opt7300x8150_Click()
    blnCav1 = False
    blnCav2 = True
    blnCav3 = False
    blnCav4 = False
End Sub

Private Sub opt8000x10000_Click()
    blnCav1 = False
    blnCav2 = False
    blnCav3 = True
    blnCav4 = False
End Sub

Private Sub opt9000x11000_Click()
    blnCav1 = False
    blnCav2 = False
    blnCav3 = False
    blnCav4 = True
End Sub
