VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDrawingSize 
   Caption         =   "Drawing Size"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDotPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1425
      ScaleWidth      =   1785
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin MSComctlLib.Slider slidSize 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Min             =   1
      Max             =   500
      SelStart        =   1
      Value           =   1
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblSlideVal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Use the slider to determain the size of dot you wish to draw with."
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmDrawingSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer

Private Sub cmdOk_Click()
    frmPaint.Enabled = True
    Unload Me
End Sub


Private Sub Form_Load()
    
    I = 5
    
    On Error GoTo E:
    
    I = InputBox("Please enter a number between 1 and 20 to determine your slide change rate.", "Slide Change Rate")
    
    
    If I > 20 Then
        I = InputBox("Please enter a number between 1 and 20 to determine your slide change rate.", "Slide Change Rate")
    End If
    
    slidSize.LargeChange = I
E:
    
End Sub

Private Sub slidSize_Scroll()
    
    picDotPreview.Cls
    lblSlideVal.Caption = slidSize.value
    intDrawWidth = slidSize.value
    picDotPreview.Circle (907.5, 727.5), intDrawWidth
    
End Sub


