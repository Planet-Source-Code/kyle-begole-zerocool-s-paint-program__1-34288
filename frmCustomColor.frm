VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomColor 
   Caption         =   "Custom Color"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCustColorSample 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1800
      ScaleHeight     =   465
      ScaleWidth      =   1305
      TabIndex        =   9
      Top             =   3120
      Width           =   1335
   End
   Begin MSComctlLib.Slider slidBlue 
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Max             =   255
   End
   Begin MSComctlLib.Slider slidGreen 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Max             =   255
   End
   Begin MSComctlLib.Slider slidRed 
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Max             =   255
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Choose your Paint Color in RGB color mode.The first slider is for Red, second for Green, third for Blue."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Slide each slider to determain the value of each color."
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Blue"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Green"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   495
   End
End
Attribute VB_Name = "frmCustomColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    
    frmPaint.Enabled = True
    blnCustomPaint = True

    If blnLeft = True Then
        
        intRed = slidRed.value
        intGreen = slidGreen.value
        intBlue = slidBlue.value
        
        frmPaint.ForeColorSample.BackColor = RGB(intRed, intGreen, intBlue)
    
    ElseIf blnRight = True Then
        
        BCRed = slidRed.value
        BCGreen = slidGreen.value
        BCBlue = slidBlue.value
        
        frmPaint.BackColorSample.BackColor = RGB(BCRed, BCGreen, BCBlue)

    End If
    
    Unload Me
End Sub

Private Sub slidBlue_Scroll()
    
    SamBlue = slidBlue.value
    
    frmCustomColor.picCustColorSample.BackColor = RGB(SamRed, SamGreen, SamBlue)
    
End Sub

Private Sub slidGreen_Scroll()
    
    SamGreen = slidGreen.value
    
    frmCustomColor.picCustColorSample.BackColor = RGB(SamRed, SamGreen, SamBlue)
    
End Sub

Private Sub slidRed_Scroll()
    
    SamRed = slidRed.value
    frmCustomColor.picCustColorSample.BackColor = RGB(SamRed, SamGreen, SamBlue)
    
End Sub
